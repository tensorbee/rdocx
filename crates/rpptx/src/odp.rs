//! Bounded OpenDocument Presentation import and export for the native facade.

use std::collections::{BTreeMap, HashSet};
use std::fmt::Write as _;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use super::*;

const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const PRESENTATION_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const ODP_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.presentation";
const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 300_000;
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_TABLE_CELLS: usize = 65_536;
const MAX_SLIDE_OBJECTS: usize = 10_000;
const MAX_DIAGNOSTICS: usize = 10_000;

/// One stable location-aware loss report at the ODP conversion boundary.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OdpDiagnostic {
    pub path: String,
    pub message: String,
}

/// A freshly converted presentation and ordered ODP diagnostics.
pub struct OdpReadResult {
    pub presentation: Presentation,
    pub diagnostics: Vec<OdpDiagnostic>,
}

/// Deterministic ODP bytes and ordered lossy-conversion diagnostics.
pub struct OdpWriteResult {
    pub bytes: Vec<u8>,
    pub diagnostics: Vec<OdpDiagnostic>,
}

#[derive(Clone, Copy)]
struct OdpLimits {
    archive: PackageReadLimits,
}

impl OdpLimits {
    const DEFAULT: Self = Self {
        archive: PackageReadLimits {
            max_entries: 4_096,
            max_part_uncompressed_bytes: 64 * 1024 * 1024,
            max_total_uncompressed_bytes: 128 * 1024 * 1024,
        },
    };
}

impl Presentation {
    /// Converts an ODF 1.3 presentation package into a fresh editable deck.
    pub fn from_odp_bytes(bytes: &[u8]) -> Result<OdpReadResult> {
        from_odp_with_limits(bytes, OdpLimits::DEFAULT)
    }

    /// Converts ODP while enforcing caller-selected archive expansion limits.
    pub fn from_odp_bytes_with_limits(
        bytes: &[u8],
        limits: PackageReadLimits,
    ) -> Result<OdpReadResult> {
        from_odp_with_limits(bytes, OdpLimits { archive: limits })
    }

    /// Opens and converts one ODP path through the bounded default profile.
    pub fn open_odp<P: AsRef<Path>>(path: P) -> Result<OdpReadResult> {
        let file = std::fs::File::open(path).map_err(OpcError::from)?;
        let mut bytes = Vec::new();
        file.take(
            OdpLimits::DEFAULT
                .archive
                .max_total_uncompressed_bytes
                .saturating_add(1),
        )
        .read_to_end(&mut bytes)
        .map_err(OpcError::from)?;
        if bytes.len() as u64 > OdpLimits::DEFAULT.archive.max_total_uncompressed_bytes {
            return Err(odp_error(None, 0, "ODP input exceeds the size limit"));
        }
        Self::from_odp_bytes(&bytes)
    }

    /// Serializes the supported presentation subset as deterministic ODF 1.3.
    pub fn to_odp_bytes(&self) -> Result<OdpWriteResult> {
        OdpWriter::new(self).write()
    }

    /// Atomically saves deterministic ODP and returns ordered diagnostics.
    pub fn save_odp<P: AsRef<Path>>(&self, path: P) -> Result<Vec<OdpDiagnostic>> {
        let result = self.to_odp_bytes()?;
        write_atomic_file(
            path.as_ref(),
            &result.bytes,
            "rpptx-odp",
            "invalid ODP file name",
            "could not allocate ODP save staging file",
        )
        .map_err(OpcError::from)?;
        Ok(result.diagnostics)
    }
}

fn odp_error(part: Option<&str>, offset: u64, message: impl Into<String>) -> Error {
    Error::Odp {
        part: part.map(str::to_owned),
        offset,
        message: message.into(),
    }
}

#[derive(Clone, Debug)]
struct XmlName {
    namespace: Option<String>,
    local: String,
}

#[derive(Clone, Debug)]
struct XmlAttribute {
    name: XmlName,
    value: String,
}

#[derive(Clone, Debug)]
enum XmlChild {
    Element(XmlNode),
    Text(String),
}

#[derive(Clone, Debug)]
struct XmlNode {
    name: XmlName,
    attributes: Vec<XmlAttribute>,
    children: Vec<XmlChild>,
}

impl XmlNode {
    fn is(&self, namespace: &str, local: &str) -> bool {
        self.name.namespace.as_deref() == Some(namespace) && self.name.local == local
    }

    fn attr(&self, namespace: Option<&str>, local: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.name.namespace.as_deref() == namespace && attribute.name.local == local
            })
            .map(|attribute| attribute.value.as_str())
    }

    fn elements(&self) -> impl Iterator<Item = &XmlNode> {
        self.children.iter().filter_map(|child| match child {
            XmlChild::Element(element) => Some(element),
            XmlChild::Text(_) => None,
        })
    }

    fn descendants<'a>(&'a self, output: &mut Vec<&'a XmlNode>) {
        for child in self.elements() {
            output.push(child);
            child.descendants(output);
        }
    }
}

fn parse_xml(part: &str, xml: &[u8]) -> Result<XmlNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<XmlNode>::new();
    let mut root = None;
    let mut nodes = 0_usize;
    let mut text_bytes = 0_usize;
    loop {
        let offset = reader.buffer_position();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| odp_error(Some(part), offset, format!("malformed XML: {error}")))?;
        let namespace = namespace_value(namespace, part, offset)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    odp_error(Some(part), offset, "ODP XML node count overflowed")
                })?;
                if nodes > MAX_XML_NODES || stack.len() >= MAX_XML_DEPTH {
                    return Err(odp_error(
                        Some(part),
                        offset,
                        "ODP XML exceeds its structural limits",
                    ));
                }
                stack.push(xml_node(&reader, namespace, &element, part, offset)?);
            }
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    odp_error(Some(part), offset, "ODP XML node count overflowed")
                })?;
                if nodes > MAX_XML_NODES || stack.len() >= MAX_XML_DEPTH {
                    return Err(odp_error(
                        Some(part),
                        offset,
                        "ODP XML exceeds its structural limits",
                    ));
                }
                attach_node(
                    &mut stack,
                    &mut root,
                    xml_node(&reader, namespace, &element, part, offset)?,
                    part,
                    offset,
                )?;
            }
            Event::End(element) => {
                let node = stack.pop().ok_or_else(|| {
                    odp_error(Some(part), offset, "ODP XML end tag has no open element")
                })?;
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref())
                    .map_err(|_| odp_error(Some(part), offset, "end tag is not UTF-8"))?;
                if node.name.namespace != namespace || node.name.local != local {
                    return Err(odp_error(
                        Some(part),
                        offset,
                        "ODP XML end tag does not match its open element",
                    ));
                }
                attach_node(&mut stack, &mut root, node, part, offset)?;
            }
            Event::Text(text) => {
                let decoded = text.xml10_content().map_err(|error| {
                    odp_error(Some(part), offset, format!("invalid XML text: {error}"))
                })?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| {
                        odp_error(
                            Some(part),
                            offset,
                            format!("unresolved XML entity: {error}"),
                        )
                    })?
                    .into_owned();
                append_text(&mut stack, value, &mut text_bytes, part, offset)?;
            }
            Event::CData(text) => {
                let value = text
                    .xml10_content()
                    .map_err(|error| {
                        odp_error(Some(part), offset, format!("invalid XML CDATA: {error}"))
                    })?
                    .into_owned();
                append_text(&mut stack, value, &mut text_bytes, part, offset)?;
            }
            Event::DocType(_) => {
                return Err(odp_error(
                    Some(part),
                    offset,
                    "DTD declarations are forbidden",
                ));
            }
            Event::GeneralRef(_) => {
                return Err(odp_error(
                    Some(part),
                    offset,
                    "unresolved XML entity is unsupported",
                ));
            }
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(odp_error(
            Some(part),
            xml.len() as u64,
            "unclosed XML element",
        ));
    }
    root.ok_or_else(|| odp_error(Some(part), 0, "ODP XML has no root element"))
}

fn namespace_value(
    namespace: ResolveResult<'_>,
    part: &str,
    offset: u64,
) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(namespace) => Ok(Some(
            std::str::from_utf8(namespace.as_ref())
                .map_err(|_| odp_error(Some(part), offset, "namespace URI is not UTF-8"))?
                .to_owned(),
        )),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(_) => Err(odp_error(
            Some(part),
            offset,
            "ODP XML uses an unresolved namespace prefix",
        )),
    }
}

fn xml_node(
    reader: &NsReader<&[u8]>,
    namespace: Option<String>,
    element: &BytesStart<'_>,
    part: &str,
    offset: u64,
) -> Result<XmlNode> {
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(|_| odp_error(Some(part), offset, "element local name is not UTF-8"))?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            odp_error(
                Some(part),
                offset,
                format!("malformed XML attribute: {error}"),
            )
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_value(namespace, part, offset)?;
        let local = std::str::from_utf8(local.as_ref())
            .map_err(|_| odp_error(Some(part), offset, "attribute local name is not UTF-8"))?
            .to_owned();
        if !seen.insert((namespace.clone(), local.clone())) {
            return Err(odp_error(
                Some(part),
                offset,
                "duplicate expanded XML attribute",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                odp_error(
                    Some(part),
                    offset,
                    format!("invalid XML attribute value: {error}"),
                )
            })?
            .into_owned();
        attributes.push(XmlAttribute {
            name: XmlName { namespace, local },
            value,
        });
    }
    Ok(XmlNode {
        name: XmlName { namespace, local },
        attributes,
        children: Vec::new(),
    })
}

fn attach_node(
    stack: &mut [XmlNode],
    root: &mut Option<XmlNode>,
    node: XmlNode,
    part: &str,
    offset: u64,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlChild::Element(node));
    } else if root.replace(node).is_some() {
        return Err(odp_error(
            Some(part),
            offset,
            "ODP XML has multiple root elements",
        ));
    }
    Ok(())
}

fn append_text(
    stack: &mut [XmlNode],
    value: String,
    retained: &mut usize,
    part: &str,
    offset: u64,
) -> Result<()> {
    if value.is_empty() || (stack.is_empty() && value.chars().all(char::is_whitespace)) {
        return Ok(());
    }
    *retained = retained
        .checked_add(value.len())
        .filter(|value| *value <= MAX_TEXT_BYTES)
        .ok_or_else(|| odp_error(Some(part), offset, "ODP XML exceeds the text limit"))?;
    let parent = stack
        .last_mut()
        .ok_or_else(|| odp_error(Some(part), offset, "ODP XML has text outside its root"))?;
    parent.children.push(XmlChild::Text(value));
    Ok(())
}

fn read_archive(bytes: &[u8], limits: PackageReadLimits) -> Result<BTreeMap<String, Vec<u8>>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| odp_error(None, 0, format!("invalid ODP ZIP: {error}")))?;
    if archive.len() > limits.max_entries {
        return Err(odp_error(None, 0, "ODP exceeds the entry limit"));
    }
    let mut entries = BTreeMap::new();
    let mut total = 0_u64;
    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .map_err(|error| odp_error(None, 0, format!("cannot read ODP entry: {error}")))?;
        let name = entry
            .name()
            .map_err(|error| odp_error(None, 0, format!("cannot decode ODP entry name: {error}")))?
            .into_owned();
        validate_entry_name(&name)?;
        if entry.encrypted() {
            return Err(odp_error(Some(&name), 0, "unsupported ODP archive entry"));
        }
        if entry.is_dir() {
            continue;
        }
        if !entry.is_file() {
            return Err(odp_error(Some(&name), 0, "unsupported ODP archive entry"));
        }
        if !matches!(
            entry.compression(),
            CompressionMethod::Stored | CompressionMethod::Deflated
        ) {
            return Err(odp_error(
                Some(&name),
                0,
                "unsupported ODP compression method",
            ));
        }
        if entry.size() > limits.max_part_uncompressed_bytes {
            return Err(odp_error(Some(&name), 0, "ODP part exceeds the size limit"));
        }
        total = total
            .checked_add(entry.size())
            .filter(|total| *total <= limits.max_total_uncompressed_bytes)
            .ok_or_else(|| odp_error(Some(&name), 0, "ODP exceeds the total size limit"))?;
        let capacity = usize::try_from(entry.size())
            .map_err(|_| odp_error(Some(&name), 0, "ODP part size is not addressable"))?;
        let mut content = Vec::with_capacity(capacity);
        entry.read_to_end(&mut content).map_err(|error| {
            odp_error(Some(&name), 0, format!("cannot expand ODP entry: {error}"))
        })?;
        if entries.insert(name.clone(), content).is_some() {
            return Err(odp_error(Some(&name), 0, "duplicate ODP archive entry"));
        }
    }
    let mut first = archive
        .by_index(0)
        .map_err(|error| odp_error(None, 0, format!("cannot read ODP mimetype: {error}")))?;
    let first_name = first
        .name()
        .map_err(|error| odp_error(None, 0, format!("cannot decode ODP mimetype name: {error}")))?;
    if first_name != "mimetype" || first.compression() != CompressionMethod::Stored {
        return Err(odp_error(
            Some("mimetype"),
            0,
            "ODP mimetype must be the first stored entry",
        ));
    }
    let mut mimetype = Vec::new();
    first.read_to_end(&mut mimetype).map_err(OpcError::from)?;
    if mimetype != ODP_MIMETYPE {
        return Err(odp_error(
            Some("mimetype"),
            0,
            "ODP mimetype is missing or invalid",
        ));
    }
    Ok(entries)
}

fn validate_entry_name(name: &str) -> Result<()> {
    let path = name.strip_suffix('/').unwrap_or(name);
    if path.is_empty()
        || name.starts_with('/')
        || name.starts_with('\\')
        || name.contains('\\')
        || name.contains('\0')
        || path
            .split('/')
            .any(|component| component.is_empty() || component == "." || component == "..")
    {
        return Err(odp_error(Some(name), 0, "unsafe ODP archive entry name"));
    }
    Ok(())
}

#[derive(Clone, Copy, Debug)]
struct FrameRect {
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
}

#[derive(Debug)]
enum OdpObject {
    Text(FrameRect, String),
    Shape(FrameRect, String),
    Table(FrameRect, Vec<Vec<String>>),
    Image(FrameRect, String),
    Unsupported(String),
}

#[derive(Debug, Default)]
struct OdpPage {
    name: Option<String>,
    objects: Vec<OdpObject>,
    notes: Option<String>,
}

fn from_odp_with_limits(bytes: &[u8], limits: OdpLimits) -> Result<OdpReadResult> {
    let entries = read_archive(bytes, limits.archive)?;
    let content = entries
        .get("content.xml")
        .ok_or_else(|| odp_error(Some("content.xml"), 0, "required ODP content is missing"))?;
    if let Some(manifest) = entries.get("META-INF/manifest.xml") {
        let root = parse_xml("META-INF/manifest.xml", manifest)?;
        if !root.is(MANIFEST_NS, "manifest") {
            return Err(odp_error(
                Some("META-INF/manifest.xml"),
                0,
                "invalid ODP manifest root",
            ));
        }
        let mut descendants = Vec::new();
        root.descendants(&mut descendants);
        if descendants
            .iter()
            .any(|node| node.is(MANIFEST_NS, "encryption-data"))
        {
            return Err(odp_error(
                Some("META-INF/manifest.xml"),
                0,
                "encrypted ODP content is unsupported",
            ));
        }
    }
    let root = parse_xml("content.xml", content)?;
    let pages = parse_pages(&root)?;
    let mut presentation = Presentation::new()?;
    let mut diagnostics = Vec::new();
    if entries.contains_key("styles.xml") {
        push_diagnostic(
            &mut diagnostics,
            "styles.xml",
            "ODP document styles were not projected",
        )?;
    }
    if root
        .elements()
        .any(|node| node.is(OFFICE_NS, "automatic-styles"))
    {
        push_diagnostic(
            &mut diagnostics,
            "content.xml/automatic-styles",
            "ODP automatic styles were not projected",
        )?;
    }
    for (page_index, page) in pages.into_iter().enumerate() {
        presentation.add_slide(0)?;
        Arc::make_mut(&mut presentation.slides[page_index])
            .slide
            .common_slide_data
            .shape_tree
            .children
            .clear();
        if let Some(name) = page.name {
            Arc::make_mut(&mut presentation.slides[page_index])
                .slide
                .common_slide_data
                .name = Some(name);
        }
        for (object_index, object) in page.objects.into_iter().enumerate() {
            let path = format!("slides/{page_index}/objects/{object_index}");
            match object {
                OdpObject::Text(rect, text) => {
                    presentation
                        .slide_mut(page_index)
                        .expect("new slide exists")
                        .add_textbox(rect.left, rect.top, rect.width, rect.height)?
                        .set_text(&text)?;
                }
                OdpObject::Shape(rect, text) => {
                    let mut slide = presentation
                        .slide_mut(page_index)
                        .expect("new slide exists");
                    let mut shape = slide.add_unstyled_shape(
                        "rect",
                        rect.left,
                        rect.top,
                        rect.width,
                        rect.height,
                    )?;
                    if !text.is_empty() {
                        shape.set_text(&text)?;
                    }
                }
                OdpObject::Table(rect, rows) => {
                    let columns = rows.iter().map(Vec::len).max().unwrap_or(0);
                    if rows.is_empty() || columns == 0 {
                        push_diagnostic(&mut diagnostics, &path, "empty ODP table was dropped")?;
                        continue;
                    }
                    let mut slide = presentation
                        .slide_mut(page_index)
                        .expect("new slide exists");
                    let mut shape = slide.add_table(
                        rows.len(),
                        columns,
                        rect.left,
                        rect.top,
                        rect.width,
                        rect.height,
                    )?;
                    let mut table = shape.table_mut().expect("new table remains typed");
                    for (row_index, row) in rows.iter().enumerate() {
                        for (column_index, text) in row.iter().enumerate() {
                            table
                                .cell_mut(row_index, column_index)
                                .expect("source-built cell is in range")
                                .set_text(text);
                        }
                    }
                }
                OdpObject::Image(rect, href) => {
                    let target = href.strip_prefix("./").unwrap_or(&href);
                    validate_entry_name(target)?;
                    let Some(image) = entries.get(target) else {
                        push_diagnostic(
                            &mut diagnostics,
                            &path,
                            "unresolved ODP image was dropped",
                        )?;
                        continue;
                    };
                    let filename = Path::new(target)
                        .file_name()
                        .and_then(|name| name.to_str())
                        .unwrap_or("image.png");
                    presentation.add_picture(
                        page_index,
                        image,
                        filename,
                        rect.left,
                        rect.top,
                        Some(rect.width),
                        Some(rect.height),
                    )?;
                }
                OdpObject::Unsupported(message) => {
                    push_diagnostic(&mut diagnostics, &path, &message)?;
                }
            }
        }
        if let Some(notes) = page.notes.filter(|notes| !notes.is_empty()) {
            attach_notes(&mut presentation, page_index, &notes)?;
        }
    }
    Ok(OdpReadResult {
        presentation,
        diagnostics,
    })
}

fn parse_pages(root: &XmlNode) -> Result<Vec<OdpPage>> {
    if !root.is(OFFICE_NS, "document-content") {
        return Err(odp_error(
            Some("content.xml"),
            0,
            "invalid ODP content root",
        ));
    }
    let body = unique_direct_child(root, OFFICE_NS, "body")?;
    let presentation = unique_direct_child(body, OFFICE_NS, "presentation")?;
    let pages = presentation
        .elements()
        .filter(|node| node.is(DRAW_NS, "page"))
        .map(parse_page)
        .collect::<Result<Vec<_>>>()?;
    if pages.len() > 10_000 {
        return Err(odp_error(
            Some("content.xml"),
            0,
            "ODP exceeds the slide limit",
        ));
    }
    Ok(pages)
}

fn unique_direct_child<'a>(node: &'a XmlNode, namespace: &str, local: &str) -> Result<&'a XmlNode> {
    let mut matches = node.elements().filter(|child| child.is(namespace, local));
    let child = matches.next().ok_or_else(|| {
        odp_error(
            Some("content.xml"),
            0,
            format!("required ODP {local} element is missing"),
        )
    })?;
    if matches.next().is_some() {
        return Err(odp_error(
            Some("content.xml"),
            0,
            format!("ODP {local} element is duplicated"),
        ));
    }
    Ok(child)
}

fn parse_page(page: &XmlNode) -> Result<OdpPage> {
    let mut output = OdpPage {
        name: page.attr(Some(DRAW_NS), "name").map(str::to_owned),
        ..OdpPage::default()
    };
    for child in page.elements() {
        if child.is(PRESENTATION_NS, "notes") {
            if output.notes.is_some() {
                return Err(odp_error(
                    Some("content.xml"),
                    0,
                    "ODP page has duplicate notes",
                ));
            }
            output.notes = Some(node_text(child));
            continue;
        }
        if child.is(DRAW_NS, "frame") {
            let rect = frame_rect(child)?;
            let direct = child.elements().collect::<Vec<_>>();
            let tables = direct
                .iter()
                .filter(|node| node.is(TABLE_NS, "table"))
                .copied()
                .collect::<Vec<_>>();
            let text_boxes = direct
                .iter()
                .filter(|node| node.is(DRAW_NS, "text-box"))
                .copied()
                .collect::<Vec<_>>();
            let objects = direct
                .iter()
                .filter(|node| {
                    node.is(DRAW_NS, "object")
                        || node.is(DRAW_NS, "object-ole")
                        || node.is(DRAW_NS, "plugin")
                })
                .copied()
                .collect::<Vec<_>>();
            let images = direct
                .iter()
                .filter(|node| node.is(DRAW_NS, "image"))
                .copied()
                .collect::<Vec<_>>();
            let primary_count = usize::from(!tables.is_empty())
                + usize::from(!text_boxes.is_empty())
                + usize::from(!objects.is_empty());
            if primary_count > 1
                || tables.len() > 1
                || text_boxes.len() > 1
                || objects.len() > 1
                || images.len() > 1
            {
                return Err(odp_error(
                    Some("content.xml"),
                    0,
                    "ODP frame has multiple payloads",
                ));
            }
            if let Some(table) = tables.first() {
                output
                    .objects
                    .push(OdpObject::Table(rect, table_rows(table)?));
            } else if !objects.is_empty() {
                output.objects.push(OdpObject::Unsupported(
                    "unsupported ODP embedded object was dropped".to_owned(),
                ));
            } else if let Some(image) = images.first() {
                if let Some(href) = image.attr(Some(XLINK_NS), "href") {
                    output.objects.push(OdpObject::Image(rect, href.to_owned()));
                } else {
                    output.objects.push(OdpObject::Unsupported(
                        "ODP image without an href was dropped".to_owned(),
                    ));
                }
            } else if !text_boxes.is_empty() {
                output.objects.push(OdpObject::Text(rect, node_text(child)));
            } else {
                output.objects.push(OdpObject::Unsupported(
                    "unsupported ODP frame was dropped".to_owned(),
                ));
            }
        } else if child.is(DRAW_NS, "custom-shape") {
            let geometries = child
                .elements()
                .filter(|node| node.is(DRAW_NS, "enhanced-geometry"))
                .collect::<Vec<_>>();
            if geometries.len() == 1
                && matches!(
                    geometries[0].attr(Some(DRAW_NS), "type"),
                    Some("rectangle" | "ooxml-rect")
                )
            {
                output
                    .objects
                    .push(OdpObject::Shape(frame_rect(child)?, node_text(child)));
            } else {
                output.objects.push(OdpObject::Unsupported(
                    "unsupported ODP custom shape was dropped".to_owned(),
                ));
            }
        } else if child.is(DRAW_NS, "g")
            || child.is(DRAW_NS, "connector")
            || child.is(PRESENTATION_NS, "animations")
        {
            output.objects.push(OdpObject::Unsupported(
                "unsupported ODP presentation content was dropped".to_owned(),
            ));
        } else if matches!(
            child.name.namespace.as_deref(),
            Some(DRAW_NS | PRESENTATION_NS)
        ) {
            output.objects.push(OdpObject::Unsupported(format!(
                "unsupported ODP {} content was dropped",
                child.name.local
            )));
        }
        if output.objects.len() > MAX_SLIDE_OBJECTS {
            return Err(odp_error(
                Some("content.xml"),
                0,
                "ODP page exceeds the object limit",
            ));
        }
    }
    Ok(output)
}

fn frame_rect(node: &XmlNode) -> Result<FrameRect> {
    Ok(FrameRect {
        left: parse_length(required_svg_length(node, "x")?)?,
        top: parse_length(required_svg_length(node, "y")?)?,
        width: parse_positive_length(required_svg_length(node, "width")?)?,
        height: parse_positive_length(required_svg_length(node, "height")?)?,
    })
}

fn required_svg_length<'a>(node: &'a XmlNode, local: &str) -> Result<&'a str> {
    node.attr(Some(SVG_NS), local).ok_or_else(|| {
        odp_error(
            Some("content.xml"),
            0,
            format!("ODP geometry is missing svg:{local}"),
        )
    })
}

fn parse_length(value: &str) -> Result<Emu> {
    let (number, unit) = value.trim().split_at(
        value
            .trim()
            .find(|character: char| character.is_ascii_alphabetic())
            .unwrap_or(value.trim().len()),
    );
    let number = number
        .parse::<f64>()
        .map_err(|_| odp_error(Some("content.xml"), 0, "invalid ODP length"))?;
    if !number.is_finite() {
        return Err(odp_error(Some("content.xml"), 0, "non-finite ODP length"));
    }
    let scale = match unit {
        "cm" => 360_000.0,
        "in" => 914_400.0,
        "pt" => 12_700.0,
        "mm" => 36_000.0,
        _ => Err(odp_error(
            Some("content.xml"),
            0,
            "unsupported ODP length unit",
        ))?,
    };
    let emu = number * scale;
    if !emu.is_finite() || emu < i64::MIN as f64 || emu > i64::MAX as f64 {
        return Err(odp_error(
            Some("content.xml"),
            0,
            "ODP length exceeds the supported range",
        ));
    }
    Ok(Emu(emu as i64))
}

fn parse_positive_length(value: &str) -> Result<Emu> {
    let value = parse_length(value)?;
    if value.0 <= 0 {
        return Err(odp_error(
            Some("content.xml"),
            0,
            "ODP extent must be positive",
        ));
    }
    Ok(value)
}

fn node_text(node: &XmlNode) -> String {
    let mut output = String::new();
    collect_node_text(node, &mut output);
    output.trim_matches('\n').to_owned()
}

fn collect_node_text(node: &XmlNode, output: &mut String) {
    let paragraph = node.is(TEXT_NS, "p") || node.is(TEXT_NS, "h");
    for child in &node.children {
        match child {
            XmlChild::Text(text) => output.push_str(text),
            XmlChild::Element(element) if element.is(TEXT_NS, "tab") => output.push('\t'),
            XmlChild::Element(element) if element.is(TEXT_NS, "line-break") => output.push('\u{b}'),
            XmlChild::Element(element) if element.is(TEXT_NS, "s") => {
                let count = element
                    .attr(Some(TEXT_NS), "c")
                    .and_then(|value| value.parse::<usize>().ok())
                    .unwrap_or(1)
                    .min(10_000);
                output.extend(std::iter::repeat_n(' ', count));
            }
            XmlChild::Element(element) => collect_node_text(element, output),
        }
    }
    if paragraph && !output.ends_with('\n') {
        output.push('\n');
    }
}

fn table_rows(table: &XmlNode) -> Result<Vec<Vec<String>>> {
    let rows = table
        .elements()
        .filter(|row| row.is(TABLE_NS, "table-row"))
        .map(|row| {
            if row.attr(Some(TABLE_NS), "number-rows-repeated").is_some() {
                return Err(odp_error(
                    Some("content.xml"),
                    0,
                    "repeated ODP table rows are unsupported",
                ));
            }
            row.elements()
                .filter(|cell| {
                    cell.is(TABLE_NS, "table-cell") || cell.is(TABLE_NS, "covered-table-cell")
                })
                .map(|cell| {
                    if cell
                        .attr(Some(TABLE_NS), "number-columns-repeated")
                        .is_some()
                    {
                        return Err(odp_error(
                            Some("content.xml"),
                            0,
                            "repeated ODP table cells are unsupported",
                        ));
                    }
                    Ok(node_text(cell))
                })
                .collect::<Result<Vec<_>>>()
        })
        .collect::<Result<Vec<_>>>()?;
    let cells = rows.iter().try_fold(0_usize, |count, row| {
        count
            .checked_add(row.len())
            .filter(|count| *count <= MAX_TABLE_CELLS)
    });
    if cells.is_none() {
        return Err(odp_error(
            Some("content.xml"),
            0,
            "ODP table exceeds the cell limit",
        ));
    }
    Ok(rows)
}

fn attach_notes(presentation: &mut Presentation, slide_index: usize, text: &str) -> Result<()> {
    let escaped = escape_xml(text);
    let xml = format!(
        r#"<p:notes xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{escaped}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#,
        A_NS = DRAWINGML_NS,
    );
    let notes = CT_NotesSlide::from_xml(xml.as_bytes()).map_err(|error| Error::MalformedPart {
        part_name: "ODP notes".to_owned(),
        message: error.to_string(),
    })?;
    let part_name = format!("/ppt/notesSlides/notesSlide{}.xml", slide_index + 1);
    let slide_part = presentation.slides[slide_index].part_name.clone();
    let relationships = presentation.package.get_or_create_part_rels(&slide_part);
    relationships.add(
        rel_types::NOTES_SLIDE,
        &relative_part_target(&slide_part, &part_name),
    );
    let mut notes_relationships = Relationships::new();
    notes_relationships.add(
        rel_types::SLIDE,
        &relative_part_target(&part_name, &slide_part),
    );
    presentation
        .package
        .part_rels
        .insert(part_name.clone(), notes_relationships);
    presentation
        .package
        .content_types
        .add_override(&part_name, content_types::NOTES_SLIDE);
    presentation.package.set_part(
        &part_name,
        notes.to_xml().map_err(|error| Error::MalformedPart {
            part_name: part_name.clone(),
            message: error.to_string(),
        })?,
    );
    Arc::make_mut(&mut presentation.slides[slide_index]).notes =
        Some(NotesRecord { part_name, notes });
    Ok(())
}

struct OdpWriter<'a> {
    presentation: &'a Presentation,
    diagnostics: Vec<OdpDiagnostic>,
    images: BTreeMap<String, Vec<u8>>,
}

impl<'a> OdpWriter<'a> {
    fn new(presentation: &'a Presentation) -> Self {
        Self {
            presentation,
            diagnostics: Vec::new(),
            images: BTreeMap::new(),
        }
    }

    fn write(mut self) -> Result<OdpWriteResult> {
        let content = self.content_xml()?;
        let manifest = self.manifest_xml();
        if content.len() as u64 > OdpLimits::DEFAULT.archive.max_part_uncompressed_bytes {
            return Err(odp_error(
                Some("content.xml"),
                0,
                "ODP output exceeds the part limit",
            ));
        }
        let limits = OdpLimits::DEFAULT.archive;
        if self.images.len().saturating_add(3) > limits.max_entries {
            return Err(odp_error(None, 0, "ODP output exceeds the entry limit"));
        }
        let mut total = (ODP_MIMETYPE.len() + content.len() + manifest.len()) as u64;
        for (name, bytes) in &self.images {
            if bytes.len() as u64 > limits.max_part_uncompressed_bytes {
                return Err(odp_error(
                    Some(name),
                    0,
                    "ODP output part exceeds the size limit",
                ));
            }
            total = total
                .checked_add(bytes.len() as u64)
                .filter(|total| *total <= limits.max_total_uncompressed_bytes)
                .ok_or_else(|| odp_error(Some(name), 0, "ODP output exceeds the total limit"))?;
        }
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut archive = ZipWriter::new(&mut cursor);
            write_entry(
                &mut archive,
                "mimetype",
                ODP_MIMETYPE,
                CompressionMethod::Stored,
            )?;
            write_entry(
                &mut archive,
                "content.xml",
                content.as_bytes(),
                CompressionMethod::Deflated,
            )?;
            for (name, bytes) in &self.images {
                write_entry(&mut archive, name, bytes, CompressionMethod::Deflated)?;
            }
            write_entry(
                &mut archive,
                "META-INF/manifest.xml",
                manifest.as_bytes(),
                CompressionMethod::Deflated,
            )?;
            archive
                .finish()
                .map_err(|error| odp_error(None, 0, format!("cannot finish ODP ZIP: {error}")))?;
        }
        if cursor.get_ref().len() as u64 > OdpLimits::DEFAULT.archive.max_total_uncompressed_bytes {
            return Err(odp_error(
                None,
                0,
                "ODP output exceeds the total size limit",
            ));
        }
        Ok(OdpWriteResult {
            bytes: cursor.into_inner(),
            diagnostics: self.diagnostics,
        })
    }

    fn content_xml(&mut self) -> Result<String> {
        let mut output = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE_NS}" xmlns:draw="{DRAW_NS}" xmlns:text="{TEXT_NS}" xmlns:table="{TABLE_NS}" xmlns:presentation="{PRESENTATION_NS}" xmlns:svg="{SVG_NS}" xmlns:xlink="{XLINK_NS}" office:version="1.3"><office:body><office:presentation>"#
        );
        for (slide_index, record) in self.presentation.slides.iter().enumerate() {
            if record.slide.common_slide_data.background.is_some() {
                push_diagnostic(
                    &mut self.diagnostics,
                    &format!("slides/{slide_index}/background"),
                    "slide background appearance was omitted",
                )?;
            }
            if record.slide.transition.is_some() {
                push_diagnostic(
                    &mut self.diagnostics,
                    &format!("slides/{slide_index}/transition"),
                    "slide transition was omitted",
                )?;
            }
            if record.slide.timing.is_some() {
                push_diagnostic(
                    &mut self.diagnostics,
                    &format!("slides/{slide_index}/timing"),
                    "slide animation timing was omitted",
                )?;
            }
            let name = record
                .slide
                .common_slide_data
                .name
                .as_deref()
                .unwrap_or("Slide");
            reserve_escaped_text(&output, name)?;
            write!(output, r#"<draw:page draw:name="{}">"#, escape_xml(name)).unwrap();
            for (shape_index, child) in record
                .slide
                .common_slide_data
                .shape_tree
                .children
                .iter()
                .enumerate()
            {
                self.write_shape(&mut output, slide_index, shape_index, child)?;
                ensure_xml_limit(&output)?;
            }
            if let Some(notes) = &record.notes {
                output.push_str("<presentation:notes><draw:frame svg:x=\"0cm\" svg:y=\"0cm\" svg:width=\"20cm\" svg:height=\"5cm\"><draw:text-box>");
                write_paragraphs(&mut output, &notes.notes.notes_text())?;
                output.push_str("</draw:text-box></draw:frame></presentation:notes>");
            }
            output.push_str("</draw:page>");
            ensure_xml_limit(&output)?;
        }
        output.push_str("</office:presentation></office:body></office:document-content>");
        ensure_xml_limit(&output)?;
        Ok(output)
    }

    fn write_shape(
        &mut self,
        output: &mut String,
        slide_index: usize,
        shape_index: usize,
        child: &ShapeTreeChild,
    ) -> Result<()> {
        let path = format!("slides/{slide_index}/shapes/{shape_index}");
        match child {
            ShapeTreeChild::Shape(shape) => {
                if shape
                    .shape_properties
                    .fill
                    .as_ref()
                    .is_some_and(|fill| !matches!(fill, Fill::NoFill(_)))
                    || shape.shape_properties.line.is_some()
                    || shape.shape_properties.effects.is_some()
                {
                    push_diagnostic(
                        &mut self.diagnostics,
                        &format!("{path}/appearance"),
                        "shape appearance was omitted",
                    )?;
                }
                let Some(rect) = shape
                    .shape_properties
                    .transform
                    .as_ref()
                    .and_then(transform_rect)
                else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "shape without bounds was dropped",
                    );
                };
                let text = shape_ref(child).text().unwrap_or_default();
                if shape.shape_properties.preset_geometry.is_some() {
                    write_custom_shape(output, rect, &text)?;
                } else {
                    write_frame_start(output, rect);
                    output.push_str("<draw:text-box>");
                    write_paragraphs(output, &text)?;
                    output.push_str("</draw:text-box></draw:frame>");
                }
            }
            ShapeTreeChild::GraphicFrame(frame) => {
                let Some(rect) = transform_rect(&frame.transform) else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "graphic frame without bounds was dropped",
                    );
                };
                if let Some(table) = shape_ref(child).table() {
                    write_frame_start(output, rect);
                    output.push_str("<table:table>");
                    for row in 0..table.row_count() {
                        output.push_str("<table:table-row>");
                        for column in 0..table.column_count() {
                            let text = table
                                .cell(row, column)
                                .map(|cell| cell.text())
                                .unwrap_or_default();
                            output.push_str("<table:table-cell office:value-type=\"string\">");
                            write_paragraphs(output, &text)?;
                            output.push_str("</table:table-cell>");
                        }
                        output.push_str("</table:table-row>");
                    }
                    output.push_str("</table:table></draw:frame>");
                } else {
                    push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "unsupported graphic frame was dropped",
                    )?;
                }
            }
            ShapeTreeChild::Picture(picture) => {
                if picture.media.is_some() {
                    push_diagnostic(
                        &mut self.diagnostics,
                        &format!("{path}/media"),
                        "media playback metadata was omitted",
                    )?;
                }
                let Some(rect) = picture
                    .shape_properties
                    .transform
                    .as_ref()
                    .and_then(transform_rect)
                else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "picture without bounds was dropped",
                    );
                };
                let relationship_id = picture
                    .blip_fill
                    .as_ref()
                    .and_then(|fill| fill.blip.as_ref())
                    .and_then(|blip| blip.embed.as_deref());
                let Some(relationship_id) = relationship_id else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "linked or unresolved picture was dropped",
                    );
                };
                let slide_part = &self.presentation.slides[slide_index].part_name;
                let Some(relationship) = self
                    .presentation
                    .package
                    .get_part_rels(slide_part)
                    .and_then(|relationships| relationships.get_by_id(relationship_id))
                else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "unresolved picture was dropped",
                    );
                };
                if relationship_is_external(relationship) {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "external picture was dropped",
                    );
                }
                let part = OpcPackage::resolve_rel_target(slide_part, &relationship.target);
                let Some(bytes) = self.presentation.package.get_part(&part) else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "missing picture was dropped",
                    );
                };
                let Some(format) = oxml_media::ImageFormat::sniff(bytes) else {
                    return push_diagnostic(
                        &mut self.diagnostics,
                        &path,
                        "picture with an unsupported signature was dropped",
                    );
                };
                let extension = format.extension();
                let name = format!(
                    "Pictures/image-{}-{}.{}",
                    slide_index + 1,
                    shape_index + 1,
                    extension
                );
                self.images.insert(name.clone(), bytes.to_vec());
                write_frame_start(output, rect);
                write!(
                    output,
                    r#"<draw:image xlink:href="{name}" xlink:type="simple"/></draw:frame>"#
                )
                .unwrap();
            }
            ShapeTreeChild::GroupShape(_)
            | ShapeTreeChild::Connector(_)
            | ShapeTreeChild::AlternateContent(_) => {
                push_diagnostic(
                    &mut self.diagnostics,
                    &path,
                    "unsupported shape was dropped",
                )?;
            }
        }
        Ok(())
    }

    fn manifest_xml(&self) -> String {
        let mut output = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>"#,
            std::str::from_utf8(ODP_MIMETYPE).expect("ASCII MIME type")
        );
        for name in self.images.keys() {
            let media_type = Path::new(name)
                .extension()
                .and_then(|value| value.to_str())
                .and_then(oxml_media::ImageFormat::from_extension)
                .map(oxml_media::ImageFormat::content_type)
                .unwrap_or("application/octet-stream");
            write!(
                output,
                r#"<manifest:file-entry manifest:full-path="{name}" manifest:media-type="{media_type}"/>"#
            )
            .unwrap();
        }
        output.push_str("</manifest:manifest>");
        output
    }
}

fn transform_rect(transform: &CT_Transform2D) -> Option<FrameRect> {
    let offset = transform.offset.as_ref()?;
    let extent = transform.extent.as_ref()?;
    (extent.cx.0 > 0 && extent.cy.0 > 0).then_some(FrameRect {
        left: offset.x,
        top: offset.y,
        width: extent.cx,
        height: extent.cy,
    })
}

fn write_frame_start(output: &mut String, rect: FrameRect) {
    write!(
        output,
        "<draw:frame svg:x=\"{:.6}in\" svg:y=\"{:.6}in\" svg:width=\"{:.6}in\" svg:height=\"{:.6}in\">",
        rect.left.to_inches(),
        rect.top.to_inches(),
        rect.width.to_inches(),
        rect.height.to_inches(),
    )
    .unwrap();
}

fn write_custom_shape(output: &mut String, rect: FrameRect, text: &str) -> Result<()> {
    write!(
        output,
        r#"<draw:custom-shape svg:x="{:.6}in" svg:y="{:.6}in" svg:width="{:.6}in" svg:height="{:.6}in">"#,
        rect.left.to_inches(),
        rect.top.to_inches(),
        rect.width.to_inches(),
        rect.height.to_inches(),
    )
    .expect("writing to String cannot fail");
    write_paragraphs(output, text)?;
    output.push_str("<draw:enhanced-geometry draw:type=\"rectangle\"/></draw:custom-shape>");
    ensure_xml_limit(output)
}

fn write_paragraphs(output: &mut String, text: &str) -> Result<()> {
    let text = text.trim_end_matches('\n');
    reserve_escaped_text(output, text)?;
    if text.is_empty() {
        output.push_str("<text:p/>");
        return ensure_xml_limit(output);
    }
    for line in text.split('\n') {
        write!(output, "<text:p>{}</text:p>", escape_xml(line))
            .expect("writing to String cannot fail");
    }
    ensure_xml_limit(output)
}

fn reserve_escaped_text(output: &str, text: &str) -> Result<()> {
    let worst_case = text
        .len()
        .checked_mul(24)
        .and_then(|bytes| bytes.checked_add(32))
        .ok_or_else(|| odp_error(Some("content.xml"), 0, "ODP output size overflowed"))?;
    output
        .len()
        .checked_add(worst_case)
        .filter(|size| *size <= OdpLimits::DEFAULT.archive.max_part_uncompressed_bytes as usize)
        .ok_or_else(|| odp_error(Some("content.xml"), 0, "ODP output exceeds the part limit"))?;
    Ok(())
}

fn ensure_xml_limit(output: &str) -> Result<()> {
    if output.len() as u64 > OdpLimits::DEFAULT.archive.max_part_uncompressed_bytes {
        return Err(odp_error(
            Some("content.xml"),
            0,
            "ODP output exceeds the part limit",
        ));
    }
    Ok(())
}

fn push_diagnostic(diagnostics: &mut Vec<OdpDiagnostic>, path: &str, message: &str) -> Result<()> {
    if diagnostics.len() >= MAX_DIAGNOSTICS {
        return Err(odp_error(None, 0, "ODP exceeds the diagnostic limit"));
    }
    diagnostics.push(OdpDiagnostic {
        path: path.to_owned(),
        message: message.to_owned(),
    });
    Ok(())
}

fn escape_xml(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            character if character == '\t' || character == '\n' || character == '\r' => {
                output.push(character)
            }
            character if character >= '\u{20}' && character != '\u{7f}' => output.push(character),
            _ => output.push('\u{fffd}'),
        }
    }
    output
}

fn write_entry(
    archive: &mut ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
    method: CompressionMethod,
) -> Result<()> {
    archive
        .start_file(
            name,
            SimpleFileOptions::DEFAULT
                .compression_method(method)
                .unix_permissions(0o644),
        )
        .map_err(|error| odp_error(Some(name), 0, format!("cannot create ODP entry: {error}")))?;
    archive
        .write_all(bytes)
        .map_err(|error| odp_error(Some(name), 0, format!("cannot write ODP entry: {error}")))
}
