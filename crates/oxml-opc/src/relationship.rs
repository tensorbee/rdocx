//! Parsing and writing of `.rels` relationship files.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};

use crate::error::{OpcError, Result};

const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/relationships";

/// Well-known OOXML relationship types.
pub mod rel_types {
    // Package-level relationships.
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const THUMBNAIL: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
    pub const DIGITAL_SIGNATURE_ORIGIN: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    pub const DIGITAL_SIGNATURE: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

    // Shared officeDocument relationships.
    pub const DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const AUDIO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
    pub const VIDEO: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
    pub const POWERPOINT_MEDIA: &str =
        "http://schemas.microsoft.com/office/2007/relationships/media";
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    pub const WEB_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    pub const FONT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const GLOSSARY_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const DIAGRAM_DATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
    pub const DIAGRAM_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";
    pub const DIAGRAM_QUICK_STYLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";
    pub const DIAGRAM_COLORS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";
    pub const DIAGRAM_DRAWING: &str =
        "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
    pub const PACKAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const CONTROL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
    pub const STRICT_OLE_OBJECT: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject";
    pub const STRICT_CONTROL: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
    pub const ACTIVEX_CONTROL_BINARY: &str =
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
    pub const VBA_PROJECT: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    pub const VBA_PROJECT_SIGNATURE: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
    pub const VBA_PROJECT_SIGNATURE_AGILE: &str =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile";

    // SpreadsheetML relationships.
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

    // PresentationML relationships.
    pub const SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    pub const SLIDE_LAYOUT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    pub const SLIDE_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    pub const NOTES_SLIDE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    pub const NOTES_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
    pub const PRES_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
    pub const VIEW_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
    pub const TABLE_STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
    pub const HANDOUT_MASTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";
    pub const POWERPOINT_COMMENTS: &str =
        "http://schemas.microsoft.com/office/2018/10/relationships/comments";
    pub const POWERPOINT_AUTHORS: &str =
        "http://schemas.microsoft.com/office/2018/10/relationships/authors";
}

/// A single relationship entry.
#[derive(Debug, Clone, PartialEq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

/// A collection of relationships parsed from a `.rels` file.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    pub items: Vec<Relationship>,
    next_id: u32,
    // The bytes read and the items they gave, shared between clones.
    source: Option<Arc<(Vec<u8>, Vec<Relationship>)>>,
}

impl Relationships {
    pub fn new() -> Self {
        Self {
            items: Vec::new(),
            next_id: 1,
            source: None,
        }
    }

    /// Parse from XML bytes.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut items = Vec::new();
        let mut ids = HashSet::new();
        let mut max_id: u32 = 0;
        let mut buf = Vec::new();
        let mut depth = 0usize;
        let mut owned_root = false;

        loop {
            let decoder = reader.decoder();
            match reader.read_resolved_event_into(&mut buf) {
                Ok((namespace, Event::Empty(ref e)))
                    if owned_root
                        && depth == 1
                        && matches!(namespace, ResolveResult::Bound(value) if namespace_uri(value.as_ref(), decoder)? == std::str::from_utf8(RELATIONSHIPS_NAMESPACE).expect("static namespace"))
                        && e.local_name().as_ref() == b"Relationship" =>
                {
                    let relationship = parse_relationship(e, reader.decoder())?;
                    if !ids.insert(relationship.id.clone()) {
                        return Err(OpcError::InvalidRelationship);
                    }
                    update_max_relationship_id(&relationship.id, &mut max_id);
                    items.push(relationship);
                }
                Ok((namespace, Event::Start(ref e)))
                    if owned_root
                        && depth == 1
                        && matches!(namespace, ResolveResult::Bound(value) if namespace_uri(value.as_ref(), decoder)? == std::str::from_utf8(RELATIONSHIPS_NAMESPACE).expect("static namespace"))
                        && e.local_name().as_ref() == b"Relationship" =>
                {
                    let relationship = parse_relationship(e, reader.decoder())?;
                    if !ids.insert(relationship.id.clone()) {
                        return Err(OpcError::InvalidRelationship);
                    }
                    update_max_relationship_id(&relationship.id, &mut max_id);
                    items.push(relationship);
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                }
                Ok((namespace, Event::Start(ref start))) => {
                    if depth == 0 {
                        owned_root = matches!(namespace, ResolveResult::Bound(value) if namespace_uri(value.as_ref(), decoder)? == std::str::from_utf8(RELATIONSHIPS_NAMESPACE).expect("static namespace"))
                            && start.local_name().as_ref() == b"Relationships";
                    }
                    depth += 1;
                }
                Ok((_, Event::End(_))) => {
                    depth = depth.checked_sub(1).ok_or(OpcError::InvalidRelationship)?
                }
                Ok((_, Event::Eof)) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }

        if depth != 0 {
            return Err(OpcError::InvalidRelationship);
        }
        Ok(Relationships {
            source: Some(Arc::new((xml.to_vec(), items.clone()))),
            items,
            next_id: next_relationship_number(max_id),
        })
    }

    /// Rebinds one imported scope, preserving opaque source around owned records.
    /// ID changes are explicit, and absent proposed records are removed atomically.
    pub fn rebind_for_import(
        &self,
        items: Vec<Relationship>,
        ids: &HashMap<String, String>,
    ) -> Result<Self> {
        self.validate_ids()?;
        let mut proposed = Self::new();
        proposed.items = items;
        proposed.validate_ids()?;
        for item in &proposed.items {
            if item.rel_type.is_empty()
                || item.target.is_empty()
                || item.rel_type.chars().any(char::is_control)
                || item.target.chars().any(char::is_control)
                || item
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| !matches!(mode, "External" | "Internal"))
            {
                return Err(OpcError::InvalidRelationship);
            }
        }
        let source = match &self.source {
            Some(source) => source.0.clone(),
            None => self.to_xml()?,
        };
        let parsed = relationship_xml(&source)?;
        let proposed_by_id = proposed
            .items
            .iter()
            .map(|item| (item.id.as_str(), item))
            .collect::<HashMap<_, _>>();
        let mut edits = Vec::new();
        let mut seen = HashSet::new();
        let mut old = HashSet::new();
        for record in &parsed.records {
            let id = record
                .attributes
                .iter()
                .find(|attribute| attribute.name == b"Id")
                .ok_or(OpcError::InvalidRelationship)?;
            if !old.insert(id.value.as_str()) {
                return Err(OpcError::InvalidRelationship);
            }
            let mapped = ids.get(&id.value).map(String::as_str).unwrap_or(&id.value);
            let Some(item) = proposed_by_id.get(mapped) else {
                edits.push((record.range.clone(), Vec::new()));
                continue;
            };
            if !seen.insert(item.id.as_str()) {
                return Err(OpcError::InvalidRelationship);
            }
            for (name, value) in [
                (b"Id".as_slice(), Some(item.id.as_str())),
                (b"Type", Some(item.rel_type.as_str())),
                (b"Target", Some(item.target.as_str())),
                (b"TargetMode", item.target_mode.as_deref()),
            ] {
                match (
                    record
                        .attributes
                        .iter()
                        .find(|attribute| attribute.name == name),
                    value,
                ) {
                    (Some(attribute), Some(value)) if attribute.value != value => edits.push((
                        attribute.value_range.clone(),
                        quick_xml::escape::escape(value).into_owned().into_bytes(),
                    )),
                    (Some(attribute), None) => {
                        edits.push((attribute.full_range.clone(), Vec::new()))
                    }
                    (None, Some(value)) => {
                        let at = record.head_end - if record.empty { 2 } else { 1 };
                        edits.push((
                            at..at,
                            format!(
                                " {}=\"{}\"",
                                std::str::from_utf8(name)
                                    .map_err(|_| OpcError::InvalidRelationship)?,
                                quick_xml::escape::escape(value)
                            )
                            .into_bytes(),
                        ));
                    }
                    _ => {}
                }
            }
        }
        let current_ids = self
            .items
            .iter()
            .map(|item| item.id.as_str())
            .collect::<HashSet<_>>();
        if ids
            .keys()
            .any(|id| !current_ids.contains(id.as_str()) && !old.contains(id.as_str()))
        {
            return Err(OpcError::InvalidRelationship);
        }
        // A changed public ID does not identify its historical opaque row.
        // An explicit retained-ID map proves that ownership. Otherwise refuse
        // a missing raw owner combined with newly emitted, unowned rows.
        let unclaimed_raw_owner = old
            .iter()
            .any(|id| !current_ids.contains(*id) && !ids.contains_key(*id));
        if unclaimed_raw_owner
            && proposed
                .items
                .iter()
                .any(|item| !seen.contains(item.id.as_str()))
        {
            return Err(OpcError::InvalidRelationship);
        }
        let prefix = parsed
            .name
            .rsplit_once(':')
            .map(|(prefix, _)| format!("{prefix}:"))
            .unwrap_or_default();
        let mut added = Vec::new();
        for item in &proposed.items {
            if !seen.contains(item.id.as_str()) {
                let mut start = BytesStart::new(format!("{prefix}Relationship"));
                start.push_attribute(("Id", item.id.as_str()));
                start.push_attribute(("Type", item.rel_type.as_str()));
                start.push_attribute(("Target", item.target.as_str()));
                if let Some(mode) = &item.target_mode {
                    start.push_attribute(("TargetMode", mode.as_str()));
                }
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Empty(start))?;
                added.extend(writer.into_inner());
            }
        }
        if parsed.empty {
            let replacement = [
                b">".as_slice(),
                &added,
                format!("</{}>", parsed.name).as_bytes(),
            ]
            .concat();
            edits.push((parsed.head_end - 2..parsed.head_end, replacement));
        } else if !added.is_empty() {
            edits.push((parsed.close_start..parsed.close_start, added));
        }
        let xml = apply_relationship_edits(&source, edits)?;
        let result = Self::from_xml(&xml)?;
        if result.items != proposed.items {
            return Err(OpcError::InvalidRelationship);
        }
        Ok(result)
    }

    /// Merges an imported scope without rebuilding its opaque XML.
    /// Conflicting root metadata is refused before either scope changes.
    pub fn merge_for_import(&self, other: &Self) -> Result<(Self, HashMap<String, String>)> {
        let left_xml = self
            .rebind_for_import(self.items.clone(), &HashMap::new())?
            .to_xml()?;
        let left = relationship_xml(&left_xml)?;
        let right_xml = other
            .rebind_for_import(other.items.clone(), &HashMap::new())?
            .to_xml()?;
        let right = relationship_xml(&right_xml)?;
        let mut root_attributes = Vec::new();
        for attribute in &right.attributes {
            if let Some(existing) = left.attributes.iter().find(|item| {
                if attribute.name == b"xmlns" || attribute.name.starts_with(b"xmlns:") {
                    item.name == attribute.name
                } else {
                    item.expanded_name == attribute.expanded_name
                }
            }) {
                if existing.value != attribute.value {
                    return Err(OpcError::InvalidRelationship);
                }
            } else {
                root_attributes.extend_from_slice(&right_xml[attribute.full_range.clone()]);
            }
        }
        let mut occupied = self.clone();
        let mut ids = HashMap::new();
        let mut imported = other.items.clone();
        for item in &mut imported {
            let old = item.id.clone();
            if occupied.get_by_id(&item.id).is_some() {
                item.id = occupied.allocate_id();
            }
            ids.insert(old, item.id.clone());
            occupied.items.push(item.clone());
        }
        let rebound = other.rebind_for_import(imported, &ids)?.to_xml()?;
        let rebound_root = relationship_xml(&rebound)?;
        let interior = if rebound_root.empty {
            Vec::new()
        } else {
            let mut inner = rebound[rebound_root.head_end..rebound_root.close_start].to_vec();
            let mut edits = Vec::new();
            for child in &rebound_root.children {
                let mut reader = quick_xml::Reader::from_reader(&rebound[child.clone()]);
                let start = match reader.read_event()? {
                    Event::Start(start) | Event::Empty(start) => start,
                    _ => return Err(OpcError::InvalidRelationship),
                };
                let empty = rebound[child.start + reader.buffer_position() as usize - 2] == b'/';
                let declared = start
                    .attributes()
                    .map(|attribute| attribute.map(|attribute| attribute.key.as_ref().to_vec()))
                    .collect::<std::result::Result<HashSet<_>, _>>()?;
                let mut closure = Vec::new();
                for attribute in rebound_root.attributes.iter().filter(|attribute| {
                    attribute.name == b"xmlns" || attribute.name.starts_with(b"xmlns:")
                }) {
                    if !declared.contains(&attribute.name) {
                        closure.extend_from_slice(&rebound[attribute.full_range.clone()]);
                    }
                }
                if !declared.contains(b"xmlns".as_slice())
                    && !rebound_root
                        .attributes
                        .iter()
                        .any(|attribute| attribute.name == b"xmlns")
                {
                    closure.extend_from_slice(b" xmlns=\"\"");
                }
                let at = child.start + reader.buffer_position() as usize
                    - if empty { 2 } else { 1 }
                    - rebound_root.head_end;
                edits.push((at..at, closure));
            }
            inner = apply_relationship_edits(&inner, edits)?;
            inner
        };
        let mut edits = Vec::new();
        if left.empty {
            let replacement = [
                root_attributes.as_slice(),
                b">",
                &interior,
                format!("</{}>", left.name).as_bytes(),
            ]
            .concat();
            edits.push((left.head_end - 2..left.head_end, replacement));
        } else {
            edits.push((left.head_end - 1..left.head_end - 1, root_attributes));
            edits.push((left.close_start..left.close_start, interior));
        }
        let xml = apply_relationship_edits(&left_xml, edits)?;
        let result = Self::from_xml(&xml)?;
        if result.items != occupied.items {
            return Err(OpcError::InvalidRelationship);
        }
        Ok((result, ids))
    }

    /// Serialize to XML bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.validate_ids()?;
        if let Some(source) = &self.source
            && source.1 == self.items
        {
            return Ok(source.0.clone());
        }
        let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut rels_start = BytesStart::new("Relationships");
        rels_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(rels_start))?;

        for rel in &self.items {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(("Id", rel.id.as_str()));
            elem.push_attribute(("Type", rel.rel_type.as_str()));
            elem.push_attribute(("Target", rel.target.as_str()));
            if let Some(ref mode) = rel.target_mode {
                elem.push_attribute(("TargetMode", mode.as_str()));
            }
            writer.write_event(Event::Empty(elem))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;

        Ok(writer.into_inner())
    }

    /// Compare retained metadata of two qualified records after assigned ids and
    /// targets have been matched by the caller. Opaque values retain the full
    /// decoded namespace context, because their QName semantics are unmodelled.
    pub fn record_metadata_equal_for_import(
        &self,
        id: &str,
        other: &Self,
        other_id: &str,
    ) -> Result<bool> {
        let left = self
            .rebind_for_import(self.items.clone(), &HashMap::new())?
            .to_xml()?;
        let right = other
            .rebind_for_import(other.items.clone(), &HashMap::new())?
            .to_xml()?;
        let a = relationship_xml(&left)?;
        let b = relationship_xml(&right)?;
        let record = |root: &RelationshipXmlRoot, id: &str| {
            root.records.iter().position(|record| {
                record.attributes.iter().any(|attribute| {
                    attribute.expanded_name.0.is_none()
                        && attribute.name == b"Id"
                        && attribute.value == id
                })
            })
        };
        let (Some(ai), Some(bi)) = (record(&a, id), record(&b, other_id)) else {
            return Err(OpcError::InvalidRelationship);
        };
        let a_record = &a.records[ai];
        let b_record = &b.records[bi];
        let metadata = |record: &RelationshipXmlRecord| {
            record
                .attributes
                .iter()
                .enumerate()
                .filter(|(_, attribute)| {
                    attribute.name != b"xmlns"
                        && !attribute.name.starts_with(b"xmlns:")
                        && (attribute.expanded_name.0.is_some()
                            || !matches!(
                                attribute.name.as_slice(),
                                b"Id" | b"Type" | b"Target" | b"TargetMode"
                            ))
                })
                .map(|(index, _)| index)
                .collect::<Vec<_>>()
        };
        let a_attributes = metadata(a_record);
        let b_attributes = metadata(b_record);
        let content = |xml: &[u8], record: &RelationshipXmlRecord| -> Vec<u8> {
            if record.empty {
                Vec::new()
            } else {
                let interior = &xml[record.head_end..record.range.end];
                interior[..interior.iter().rposition(|byte| *byte == b'<').unwrap_or(0)].to_vec()
            }
        };
        if a_attributes.is_empty()
            && b_attributes.is_empty()
            && content(&left, a_record).is_empty()
            && content(&right, b_record).is_empty()
        {
            return Ok(true);
        }
        // Namespace closure can add redundant declarations to a copied record.
        // Compare its complete effective authority, not declaration placement.
        let namespaces = |root: &RelationshipXmlRoot, record: &RelationshipXmlRecord| {
            root.attributes
                .iter()
                .chain(&record.attributes)
                .filter(|attribute| {
                    attribute.name == b"xmlns" || attribute.name.starts_with(b"xmlns:")
                })
                .map(|attribute| (attribute.name.clone(), attribute.value.clone()))
                .collect::<std::collections::BTreeMap<_, _>>()
        };
        if namespaces(&a, a_record) != namespaces(&b, b_record)
            || a_attributes.len() != b_attributes.len()
        {
            return Ok(false);
        }
        for (left_index, right_index) in a_attributes.iter().zip(b_attributes) {
            let left_attribute = &a_record.attributes[*left_index];
            let right_attribute = &b_record.attributes[right_index];
            if left_attribute.expanded_name != right_attribute.expanded_name
                || left[left_attribute.full_range.clone()]
                    != right[right_attribute.full_range.clone()]
            {
                return Ok(false);
            }
        }
        Ok(content(&left, a_record) == content(&right, b_record))
    }

    /// Find a relationship by its ID.
    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.items.iter().find(|r| r.id == id)
    }

    /// Find the first relationship matching a given type.
    pub fn get_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.items.iter().find(|r| r.rel_type == rel_type)
    }

    /// Find all relationships matching a given type.
    pub fn get_all_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.items
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .collect()
    }

    /// Add a new relationship and return its generated ID.
    pub fn add(&mut self, rel_type: &str, target: &str) -> String {
        let id = self.allocate_id();
        self.items.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
        id
    }

    /// Add an externally-targeted relationship (e.g. a hyperlink URL) and
    /// return its generated ID.
    pub fn add_external(&mut self, rel_type: &str, target: &str) -> String {
        let id = self.allocate_id();
        self.items.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: Some("External".to_string()),
        });
        id
    }

    /// Add a relationship with a specific ID.
    ///
    /// If a relationship with this ID already exists, it is replaced.
    /// The `next_id` counter is updated to avoid future collisions.
    pub fn add_with_id(&mut self, id: &str, rel_type: &str, target: &str) {
        self.items.retain(|r| r.id != id);
        self.items.push(Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
        if let Some(num) = id.strip_prefix("rId").and_then(|s| s.parse::<u32>().ok())
            && num >= self.next_id
        {
            self.next_id = next_relationship_number(num);
        }
    }

    fn allocate_id(&mut self) -> String {
        let mut candidate = self.next_id;
        let numeric_space = usize::try_from(u32::MAX).unwrap_or(usize::MAX);
        let attempts = self.items.len().saturating_add(1).min(numeric_space);
        for _ in 0..attempts {
            let id = format!("rId{candidate}");
            let next = next_relationship_number(candidate);
            if self.items.iter().all(|relationship| relationship.id != id) {
                self.next_id = next;
                return id;
            }
            candidate = next;
        }

        for ordinal in 1u128.. {
            let id = format!("rIdGenerated{ordinal}");
            if self.items.iter().all(|relationship| relationship.id != id) {
                return id;
            }
        }
        unreachable!("the generated relationship id space is unbounded")
    }

    /// Whether two relationship sets hold the same items, identifier counter
    /// and source.
    pub(crate) fn same_as(&self, other: &Relationships) -> bool {
        self.items == other.items && self.next_id == other.next_id && self.source == other.source
    }

    /// Whether these relationships still hold the items their source bytes
    /// were read into, with the identifier counter reading them sets, so
    /// that writing and reading them back gives them unchanged.
    pub(crate) fn is_as_read(&self) -> bool {
        let Some(source) = &self.source else {
            return false;
        };
        let mut maximum = 0;
        for item in &self.items {
            update_max_relationship_id(&item.id, &mut maximum);
        }
        source.1 == self.items && self.next_id == next_relationship_number(maximum)
    }

    pub(crate) fn validate_ids(&self) -> Result<()> {
        let mut ids = HashSet::with_capacity(self.items.len());
        if self
            .items
            .iter()
            .all(|relationship| ids.insert(relationship.id.as_str()))
        {
            Ok(())
        } else {
            Err(OpcError::InvalidRelationship)
        }
    }
}

struct RelationshipXmlAttribute {
    name: Vec<u8>,
    expanded_name: (Option<String>, Vec<u8>),
    value: String,
    value_range: std::ops::Range<usize>,
    full_range: std::ops::Range<usize>,
}

struct RelationshipXmlRecord {
    range: std::ops::Range<usize>,
    head_end: usize,
    empty: bool,
    attributes: Vec<RelationshipXmlAttribute>,
}

struct RelationshipXmlRoot {
    name: String,
    head_end: usize,
    close_start: usize,
    empty: bool,
    attributes: Vec<RelationshipXmlAttribute>,
    records: Vec<RelationshipXmlRecord>,
    children: Vec<std::ops::Range<usize>>,
}

fn namespace_uri(bytes: &[u8], decoder: Decoder) -> Result<String> {
    // NamespaceResolver retains the declaration's lexical attribute value.
    // Namespace identity uses XML attribute normalization, including references.
    let declaration = quick_xml::events::attributes::Attribute {
        key: quick_xml::name::QName(b"xmlns"),
        value: std::borrow::Cow::Borrowed(bytes),
    };
    Ok(declaration
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
        .into_owned())
}

fn relationship_xml(xml: &[u8]) -> Result<RelationshipXmlRoot> {
    // quick-xml consumes its initial BOM without counting it in buffer_position.
    // Parse the remaining slice explicitly and translate every splice offset.
    let origin = if xml.starts_with(b"\xef\xbb\xbf") {
        3
    } else {
        0
    };
    let mut reader = NsReader::from_reader(&xml[origin..]);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    let mut record = None;
    let mut child = None;
    loop {
        let begin = origin + reader.buffer_position() as usize;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        let qualified = matches!(namespace, ResolveResult::Bound(uri) if namespace_uri(uri.as_ref(), decoder)? == std::str::from_utf8(RELATIONSHIPS_NAMESPACE).expect("static namespace"));
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let end = origin + reader.buffer_position() as usize;
                let mut attributes = Vec::new();
                let mut expanded_attributes = HashSet::new();
                let raw_start: &[u8] = start.as_ref();
                for attribute in start.attributes() {
                    let attribute = attribute?;
                    let name_offset = (attribute.key.as_ref().as_ptr() as usize)
                        .checked_sub(raw_start.as_ptr() as usize)
                        .ok_or(OpcError::InvalidRelationship)?;
                    let value_offset = (attribute.value.as_ptr() as usize)
                        .checked_sub(raw_start.as_ptr() as usize)
                        .ok_or(OpcError::InvalidRelationship)?;
                    let value_start = begin + 1 + value_offset;
                    let value_end = value_start + attribute.value.len();
                    if value_end >= xml.len() {
                        return Err(OpcError::InvalidRelationship);
                    }
                    let mut full_start = begin + 1 + name_offset;
                    while full_start > begin && xml[full_start - 1].is_ascii_whitespace() {
                        full_start -= 1;
                    }
                    let declaration = attribute.key.as_ref() == b"xmlns"
                        || attribute.key.as_ref().starts_with(b"xmlns:");
                    let expanded_name = if declaration {
                        (
                            Some("http://www.w3.org/2000/xmlns/".to_owned()),
                            attribute.key.as_ref().to_vec(),
                        )
                    } else {
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        let namespace = match namespace {
                            ResolveResult::Bound(uri) => {
                                Some(namespace_uri(uri.as_ref(), decoder)?)
                            }
                            ResolveResult::Unbound => None,
                            ResolveResult::Unknown(_) => return Err(OpcError::InvalidRelationship),
                        };
                        (namespace, local.as_ref().to_vec())
                    };
                    if !expanded_attributes.insert(expanded_name.clone()) {
                        return Err(OpcError::InvalidRelationship);
                    }
                    attributes.push(RelationshipXmlAttribute {
                        name: attribute.key.as_ref().to_vec(),
                        expanded_name,
                        value: attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?
                            .into_owned(),
                        value_range: value_start..value_end,
                        full_range: full_start..value_end + 1,
                    });
                }
                if depth == 0 {
                    if root.is_some()
                        || !qualified
                        || start.local_name().as_ref() != b"Relationships"
                    {
                        return Err(OpcError::InvalidRelationship);
                    }
                    root = Some(RelationshipXmlRoot {
                        name: std::str::from_utf8(start.name().as_ref())
                            .map_err(|_| OpcError::InvalidRelationship)?
                            .to_owned(),
                        head_end: end,
                        close_start: end,
                        empty,
                        attributes,
                        records: Vec::new(),
                        children: Vec::new(),
                    });
                } else if depth == 1 {
                    child = Some(begin);
                    if qualified && start.local_name().as_ref() == b"Relationship" {
                        let current = RelationshipXmlRecord {
                            range: begin..end,
                            head_end: end,
                            empty,
                            attributes,
                        };
                        if empty {
                            root.as_mut()
                                .ok_or(OpcError::InvalidRelationship)?
                                .records
                                .push(current);
                        } else {
                            record = Some(current);
                        }
                    }
                    if empty {
                        root.as_mut()
                            .ok_or(OpcError::InvalidRelationship)?
                            .children
                            .push(begin..end);
                        child = None;
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(OpcError::InvalidRelationship)?;
                let current = root.as_mut().ok_or(OpcError::InvalidRelationship)?;
                if depth == 1 {
                    if let Some(mut record) = record.take() {
                        record.range.end = origin + reader.buffer_position() as usize;
                        current.records.push(record);
                    }
                    if let Some(begin) = child.take() {
                        current
                            .children
                            .push(begin..origin + reader.buffer_position() as usize);
                    }
                } else if depth == 0 {
                    current.close_start = begin;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if depth != 0 || record.is_some() || child.is_some() {
        return Err(OpcError::InvalidRelationship);
    }
    root.ok_or(OpcError::InvalidRelationship)
}

fn apply_relationship_edits(
    xml: &[u8],
    mut edits: Vec<(std::ops::Range<usize>, Vec<u8>)>,
) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| range.start);
    let mut end = 0;
    for (range, _) in &edits {
        if range.start < end || range.start > range.end || range.end > xml.len() {
            return Err(OpcError::InvalidRelationship);
        }
        end = range.end;
    }
    let mut result = xml.to_vec();
    for (range, bytes) in edits.into_iter().rev() {
        result.splice(range, bytes);
    }
    Ok(result)
}

fn parse_relationship(element: &BytesStart<'_>, decoder: Decoder) -> Result<Relationship> {
    let mut id = None;
    let mut rel_type = None;
    let mut target = None;
    let mut target_mode = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
            .into_owned();
        match attribute.key.as_ref() {
            b"Id" => id = Some(value),
            b"Type" => rel_type = Some(value),
            b"Target" => target = Some(value),
            b"TargetMode" => target_mode = Some(value),
            _ => {}
        }
    }
    match (id, rel_type, target) {
        (Some(id), Some(rel_type), Some(target)) => Ok(Relationship {
            id,
            rel_type,
            target,
            target_mode,
        }),
        _ => Err(OpcError::InvalidRelationship),
    }
}

fn update_max_relationship_id(id: &str, maximum: &mut u32) {
    if let Some(number) = id.strip_prefix("rId").and_then(|value| value.parse().ok()) {
        *maximum = (*maximum).max(number);
    }
}

fn next_relationship_number(current: u32) -> u32 {
    current.checked_add(1).unwrap_or(1).max(1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relationship_and_content_type_constants_are_unique_and_well_formed() {
        const PACKAGE_PREFIX: &str =
            "http://schemas.openxmlformats.org/package/2006/relationships/";
        const OFFICE_PREFIX: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

        let package_relationships = [rel_types::CORE_PROPERTIES, rel_types::THUMBNAIL];
        let office_relationships = [
            rel_types::DOCUMENT,
            rel_types::STYLES,
            rel_types::NUMBERING,
            rel_types::HEADER,
            rel_types::FOOTER,
            rel_types::IMAGE,
            rel_types::SETTINGS,
            rel_types::WEB_SETTINGS,
            rel_types::FONT_TABLE,
            rel_types::FONT,
            rel_types::THEME,
            rel_types::HYPERLINK,
            rel_types::FOOTNOTES,
            rel_types::ENDNOTES,
            rel_types::GLOSSARY_DOCUMENT,
            rel_types::CHART,
            rel_types::DIAGRAM_DATA,
            rel_types::DIAGRAM_LAYOUT,
            rel_types::DIAGRAM_QUICK_STYLE,
            rel_types::DIAGRAM_COLORS,
            rel_types::PACKAGE,
            rel_types::OLE_OBJECT,
            rel_types::CONTROL,
            rel_types::WORKSHEET,
            rel_types::SHARED_STRINGS,
            rel_types::EXTENDED_PROPERTIES,
            rel_types::CUSTOM_PROPERTIES,
            rel_types::SLIDE,
            rel_types::SLIDE_LAYOUT,
            rel_types::SLIDE_MASTER,
            rel_types::NOTES_SLIDE,
            rel_types::NOTES_MASTER,
            rel_types::PRES_PROPS,
            rel_types::VIEW_PROPS,
            rel_types::TABLE_STYLES,
            rel_types::HANDOUT_MASTER,
        ];
        let strict_relationships = [rel_types::STRICT_OLE_OBJECT, rel_types::STRICT_CONTROL];
        let microsoft_relationships = [
            rel_types::ACTIVEX_CONTROL_BINARY,
            rel_types::VBA_PROJECT,
            rel_types::VBA_PROJECT_SIGNATURE,
            rel_types::VBA_PROJECT_SIGNATURE_AGILE,
        ];

        let mut relationship_values = std::collections::HashSet::new();
        for value in package_relationships {
            assert!(value.starts_with(PACKAGE_PREFIX));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in office_relationships {
            assert!(value.starts_with(OFFICE_PREFIX));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in strict_relationships {
            assert!(value.starts_with("http://purl.oclc.org/ooxml/officeDocument/relationships/"));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in microsoft_relationships {
            assert!(value.starts_with("http://schemas.microsoft.com/office/2006/relationships/"));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }
        for value in [
            rel_types::POWERPOINT_COMMENTS,
            rel_types::POWERPOINT_AUTHORS,
            rel_types::DIAGRAM_DRAWING,
        ] {
            assert!(value.starts_with("http://schemas.microsoft.com/office/"));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(relationship_values.insert(value));
        }

        let content_type_values = [
            crate::content_types::RELATIONSHIPS,
            crate::content_types::XML,
            crate::content_types::CORE_PROPERTIES,
            crate::content_types::EXTENDED_PROPERTIES,
            crate::content_types::CUSTOM_PROPERTIES,
            crate::content_types::THEME,
            crate::content_types::CHART,
            crate::content_types::WORD_GLOSSARY,
            crate::content_types::WORD_DOCUMENT,
            crate::content_types::WORD_DOCUMENT_MACRO_ENABLED,
            crate::content_types::WORD_TEMPLATE,
            crate::content_types::WORD_TEMPLATE_MACRO_ENABLED,
            crate::content_types::PRESENTATION,
            crate::content_types::SLIDESHOW,
            crate::content_types::SLIDE,
            crate::content_types::SLIDE_LAYOUT,
            crate::content_types::SLIDE_MASTER,
            crate::content_types::NOTES_SLIDE,
            crate::content_types::NOTES_MASTER,
            crate::content_types::PRES_PROPS,
            crate::content_types::VIEW_PROPS,
            crate::content_types::TABLE_STYLES,
            crate::content_types::HANDOUT_MASTER,
            crate::content_types::POWERPOINT_COMMENTS,
            crate::content_types::POWERPOINT_AUTHORS,
            crate::content_types::WORKBOOK,
            crate::content_types::EMBEDDED_WORKBOOK,
            crate::content_types::WORKSHEET,
            crate::content_types::SHARED_STRINGS,
            crate::content_types::STYLES,
        ];

        let mut content_types = std::collections::HashSet::new();
        for value in content_type_values {
            let (kind, subtype) = value.split_once('/').expect("valid MIME type");
            assert_eq!(kind, "application");
            assert!(!subtype.is_empty());
            assert!(!subtype.contains('/'));
            assert!(!value.chars().any(char::is_whitespace));
            assert!(content_types.insert(value));
        }
    }

    #[test]
    fn round_trip_relationships() {
        let mut rels = Relationships::new();
        rels.add(rel_types::DOCUMENT, "word/document.xml");
        rels.add(rel_types::STYLES, "word/styles.xml");

        let xml = rels.to_xml().unwrap();
        let parsed = Relationships::from_xml(&xml).unwrap();

        assert_eq!(parsed.items.len(), 2);
        assert_eq!(parsed.items[0].id, "rId1");
        assert_eq!(parsed.items[0].target, "word/document.xml");
        assert_eq!(parsed.items[1].id, "rId2");
    }

    #[test]
    fn find_by_type() {
        let mut rels = Relationships::new();
        rels.add(rel_types::DOCUMENT, "word/document.xml");
        rels.add(rel_types::STYLES, "word/styles.xml");

        let doc = rels.get_by_type(rel_types::DOCUMENT).unwrap();
        assert_eq!(doc.target, "word/document.xml");
    }

    #[test]
    fn high_numeric_relationship_ids_roll_over_without_collision() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4294967294" Type="type-a" Target="a.xml"/>
</Relationships>"#;
        let mut relationships = Relationships::from_xml(xml).unwrap();
        assert_eq!(relationships.add("type-b", "b.xml"), "rId4294967295");
        assert_eq!(relationships.add("type-c", "c.xml"), "rId1");
        assert_eq!(
            relationships.add_external("type-d", "https://example.com"),
            "rId2"
        );
        let ids = relationships
            .items
            .iter()
            .map(|relationship| relationship.id.as_str())
            .collect::<std::collections::HashSet<_>>();
        assert_eq!(ids.len(), relationships.items.len());
        assert!(!ids.contains("rId0"));
    }

    #[test]
    fn parsed_u32_max_relationship_id_rolls_to_first_free_positive_id() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4294967295" Type="type-max" Target="max.xml"/>
  <Relationship Id="rId1" Type="type-one" Target="one.xml"/>
</Relationships>"#;
        let mut relationships = Relationships::from_xml(xml).unwrap();
        assert_eq!(relationships.add("type-two", "two.xml"), "rId2");
        relationships.add_with_id("rId4294967295", "type-max", "replacement.xml");
        assert_eq!(relationships.add("type-three", "three.xml"), "rId3");
    }

    #[test]
    fn aliased_and_expanded_relationship_entries_parse_but_foreign_decoys_do_not() {
        let xml = br#"<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:x="urn:foreign"><x:Relationship Id="rId9" Type="urn:wrong" Target="wrong.xml"/><pr:Relationship Id="rId2" Type="urn:first" Target="first.xml"></pr:Relationship><pr:Relationship Id="rId3" Type="urn:second" Target="second.xml"/></pr:Relationships>"#;
        let parsed = Relationships::from_xml(xml).expect("canonical relationships parse");
        assert_eq!(parsed.items.len(), 2);
        assert_eq!(parsed.items[0].id, "rId2");
        assert_eq!(parsed.items[1].id, "rId3");
    }

    #[test]
    fn duplicate_relationship_ids_are_rejected_on_parse_and_serialize() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:first" Target="first.xml"/><Relationship Id="rId1" Type="urn:second" Target="second.xml"/></Relationships>"#;
        assert!(matches!(
            Relationships::from_xml(xml),
            Err(OpcError::InvalidRelationship)
        ));

        let mut relationships = Relationships::new();
        relationships.items.push(Relationship {
            id: "producer".to_owned(),
            rel_type: "urn:first".to_owned(),
            target: "first.xml".to_owned(),
            target_mode: None,
        });
        relationships.items.push(Relationship {
            id: "producer".to_owned(),
            rel_type: "urn:second".to_owned(),
            target: "second.xml".to_owned(),
            target_mode: None,
        });
        assert!(matches!(
            relationships.to_xml(),
            Err(OpcError::InvalidRelationship)
        ));
    }

    #[test]
    fn encoded_relationship_values_are_decoded_before_validation_and_escape_once() {
        let duplicate = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:first" Target="first.xml"/><Relationship Id="rId&#49;" Type="urn:second" Target="second.xml"/></Relationships>"#;
        assert!(matches!(
            Relationships::from_xml(duplicate),
            Err(OpcError::InvalidRelationship)
        ));

        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId&#49;" Type="urn:test?a=1&amp;b=2" Target="folder/a&amp;b.xml"/></Relationships>"#;
        let parsed = Relationships::from_xml(xml).unwrap();
        assert_eq!(parsed.items[0].id, "rId1");
        assert_eq!(parsed.items[0].rel_type, "urn:test?a=1&b=2");
        assert_eq!(parsed.items[0].target, "folder/a&b.xml");
        let saved = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
        assert!(saved.contains("folder/a&amp;b.xml"));
        assert!(!saved.contains("&amp;amp;"));
        assert_eq!(
            Relationships::from_xml(saved.as_bytes()).unwrap().items,
            parsed.items
        );
    }

    #[test]
    fn unchanged_relationship_parts_serialize_byte_for_byte() {
        let xml = br#"<?xml version='1.0' encoding='UTF-8'?>
<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
 <r:Relationship Target="word/document.xml" Type="urn:document" Id="producer-id"></r:Relationship>
</r:Relationships>"#;
        let mut relationships = Relationships::from_xml(xml).unwrap();
        assert_eq!(relationships.to_xml().unwrap(), xml);

        relationships.items[0].target = "word/changed.xml".to_owned();
        let changed = relationships.to_xml().unwrap();
        assert_ne!(changed, xml);
        assert!(
            String::from_utf8(changed)
                .unwrap()
                .contains("word/changed.xml")
        );
    }

    #[test]
    fn fx216_relationship_projection_excludes_foreign_capsules() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:fx216:opaque" z:root="keep"><Relationship Id="owned" Type="urn:owned" Target="owned.xml" z:marker="record"><z:child/></Relationship><z:capsule><Relationship Id="foreign" Type="urn:foreign" Target="https://foreign.example" TargetMode="External"/></z:capsule></Relationships>"#;
        let relationships = Relationships::from_xml(xml).unwrap();
        assert_eq!(relationships.items.len(), 1);
        assert_eq!(relationships.items[0].id, "owned");
        assert_eq!(relationships.to_xml().unwrap(), xml);
    }

    #[test]
    fn fx216_scoped_relationship_rebind_preserves_opaque_source() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:fx216:opaque" z:root="keep"><Relationship Id="owned" Type="urn:owned" Target="owned.xml" z:marker="record"><z:child/></Relationship></Relationships>"#;
        let relationships = Relationships::from_xml(xml).unwrap();
        let mut items = relationships.items.clone();
        items[0].target = "copied.xml".into();
        let rebound = relationships
            .rebind_for_import(items, &HashMap::new())
            .unwrap();
        let changed = String::from_utf8(rebound.to_xml().unwrap()).unwrap();
        assert!(
            changed.contains("z:root=\"keep\""),
            "import needs a scoped preservation path"
        );
        assert!(changed.contains("z:marker=\"record\"") && changed.contains("<z:child/>"));
    }

    #[test]
    fn fx216_relationship_bom_rebind_and_either_side_merge_keep_original_coordinates() {
        for declaration in ["", "<?xml version='1.0'?>"] {
            let raw = format!(
                "{declaration}<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:z='urn:bom' z:root='keep'><Relationship Id='rId1' Type='urn:owned' Target='old.xml' z:record='keep'><z:child/></Relationship><!--tail--></Relationships>"
            );
            let xml = [b"\xef\xbb\xbf".as_slice(), raw.as_bytes()].concat();
            let original = Relationships::from_xml(&xml).unwrap();
            assert_eq!(original.items.len(), 1);
            assert_eq!(original.to_xml().unwrap(), xml);
            let mut items = original.items.clone();
            items[0].target = "new.xml".into();
            let changed = original.rebind_for_import(items, &HashMap::new()).unwrap();
            assert_eq!(
                changed.to_xml().unwrap(),
                [
                    b"\xef\xbb\xbf".as_slice(),
                    raw.replace("old.xml", "new.xml").as_bytes()
                ]
                .concat()
            );
            let plain = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='plain' Type='urn:plain' Target='plain.xml'/></Relationships>"#).unwrap();
            let (left_bom, _) = original.merge_for_import(&plain).unwrap();
            let left_bytes = left_bom.to_xml().unwrap();
            assert!(left_bytes.starts_with(b"\xef\xbb\xbf"));
            assert_eq!(left_bom.items.len(), 2);
            let (right_bom, _) = plain.merge_for_import(&original).unwrap();
            let right_bytes = right_bom.to_xml().unwrap();
            assert!(!right_bytes.starts_with(b"\xef\xbb\xbf"));
            assert_eq!(right_bom.items.len(), 2);
            for bytes in [&left_bytes, &right_bytes] {
                let text = std::str::from_utf8(bytes).unwrap();
                assert!(
                    text.contains("z:root='keep'")
                        && text.contains("z:record='keep'")
                        && text.contains("<z:child/>")
                        && text.contains("<!--tail-->")
                );
            }
            assert_eq!(original.to_xml().unwrap(), xml);
        }
    }

    #[test]
    fn fx216_relationship_encoded_namespace_owns_root_and_direct_records() {
        for xml in [
            br#"<q:Relationships xmlns:q='http://schemas.openxmlformats.org/package/2006/relationship&#115;'><q:Relationship Id='owned' Type='urn:owned' Target='old.xml'/></q:Relationships>"#.as_slice(),
            br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:q='http://schemas.openxmlformats.org/package/2006/relationship&#x73;'><q:Relationship Id='owned' Type='urn:owned' Target='old.xml'/></Relationships>"#.as_slice(),
        ] {
            let relationships = Relationships::from_xml(xml).unwrap();
            assert_eq!(relationships.items.len(), 1);
            assert_eq!(relationships.to_xml().unwrap(), xml);
            let mut items = relationships.items.clone();
            items[0].target = "copied.xml".into();
            let rebound = relationships.rebind_for_import(items, &HashMap::new()).unwrap();
            assert_eq!(rebound.items[0].target, "copied.xml");
            assert!(std::str::from_utf8(&rebound.to_xml().unwrap()).unwrap().contains("relationship&#"));
            assert_eq!(relationships.to_xml().unwrap(), xml);
        }
    }

    #[test]
    fn fx216_relationship_entity_equivalent_namespace_aliases_are_one_identity() {
        let left = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:z='urn:metadata' z:mark='left'/>"#).unwrap();
        let conflicting = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:y='urn:met&#97;data' y:mark='right'/>"#).unwrap();
        let before = left.to_xml().unwrap();
        assert!(left.merge_for_import(&conflicting).is_err());
        let equal = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:y='urn:met&#97;data' y:mark='left'><y:opaque/></Relationships>"#).unwrap();
        let (merged, _) = left.merge_for_import(&equal).unwrap();
        let raw = merged.to_xml().unwrap();
        let mut reader = quick_xml::Reader::from_reader(raw.as_slice());
        let start = match reader.read_event().unwrap() {
            Event::Start(start) => start,
            _ => panic!("root"),
        };
        assert_eq!(
            start
                .attributes()
                .map(|attribute| attribute.unwrap())
                .filter(|attribute| attribute.key.local_name().as_ref() == b"mark")
                .count(),
            1
        );
        let duplicate = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:z='urn:metadata' xmlns:y='urn:met&#97;data' z:mark='left' y:mark='left'/>"#).unwrap();
        assert!(
            duplicate
                .rebind_for_import(Vec::new(), &HashMap::new())
                .is_err()
        );
        let different = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:y='urn:different' y:mark='right'/>"#).unwrap();
        assert!(left.merge_for_import(&different).is_ok());
        assert_eq!(left.to_xml().unwrap(), before);
    }

    #[test]
    fn fx216_relationship_root_alias_conflict_refuses_and_equal_alias_is_single() {
        let left = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:z='urn:metadata' z:mark='left'/>"#).unwrap();
        let conflicting = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:y='urn:metadata' y:mark='right'/>"#).unwrap();
        let before = left.to_xml().unwrap();
        assert!(left.merge_for_import(&conflicting).is_err());
        let equal = Relationships::from_xml(br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:y='urn:metadata' y:mark='left'><y:opaque/></Relationships>"#).unwrap();
        let (merged, _) = left.merge_for_import(&equal).unwrap();
        let xml = merged.to_xml().unwrap();
        let mut reader = NsReader::from_reader(xml.as_slice());
        let mut buffer = Vec::new();
        let (_, event) = reader.read_resolved_event_into(&mut buffer).unwrap();
        let root = match event {
            Event::Start(root) => root,
            _ => panic!("merged root"),
        };
        let marks = root.attributes().map(|attribute| {
            let attribute = attribute.unwrap();
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            usize::from(matches!(namespace, ResolveResult::Bound(uri) if uri.as_ref() == b"urn:metadata") && local.as_ref() == b"mark")
        }).sum::<usize>();
        assert_eq!(marks, 1);
        assert!(std::str::from_utf8(&xml).unwrap().contains("<y:opaque"));
        assert_eq!(left.to_xml().unwrap(), before);
    }

    #[test]
    fn fx216_relationship_merge_keeps_both_scopes_and_rolls_back_root_conflicts() {
        let left = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:opaque" z:left='keep'><Relationship Id='rId1' Type='urn:a' Target='a'/><z:tail/></Relationships>"#;
        let right = br#"<q:Relationships xmlns:q="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:opaque" z:right='keep'><!--source--><q:Relationship Id='rId1' Type='urn:b' Target='https://b' TargetMode='External' z:record='keep'><z:child/></q:Relationship><capsule><q:Relationship Id='ghost' Type='urn:foreign' Target='unknown'/></capsule></q:Relationships>"#;
        let original = Relationships::from_xml(left).unwrap();
        let source = Relationships::from_xml(right).unwrap();
        let (merged, ids) = original.merge_for_import(&source).unwrap();
        assert_ne!(ids["rId1"], "rId1");
        assert_eq!(merged.items.len(), 2);
        assert_eq!(merged.items[1].target_mode.as_deref(), Some("External"));
        let raw = String::from_utf8(merged.to_xml().unwrap()).unwrap();
        for retained in [
            "z:left='keep'",
            "z:right='keep'",
            "<!--source-->",
            "z:record='keep'",
            "<z:child/>",
            "Id='ghost'",
        ] {
            assert!(raw.contains(retained), "{retained}");
        }
        assert!(raw.contains("xmlns=\"\""));
        let conflict = Relationships::from_xml(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:opaque" z:left='unequal'/>"#).unwrap();
        assert!(original.merge_for_import(&conflict).is_err());
        assert_eq!(original.to_xml().unwrap(), left);
        assert_eq!(source.to_xml().unwrap(), right);
    }

    #[test]
    fn fx216_relationship_rebind_removes_adds_and_rolls_back_invalid_schema() {
        let xml = br#"<?xml version='1.0'?><q:Relationships xmlns:q="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:fx216:opaque" z:root='keep'><q:Relationship Target='https://old.example' z:record='keep' Id='kept' TargetMode='External' Type='urn:owned'><z:child/></q:Relationship><!--between--><q:Relationship Id='removed' Type='urn:owned' Target='removed.xml'/><z:capsule><q:Relationship Id='foreign' Type='urn:metadata' Target='foreign.xml'/></z:capsule><z:tail/></q:Relationships>"#;
        let original = Relationships::from_xml(xml).unwrap();
        let items = vec![
            Relationship {
                id: "renamed".into(),
                rel_type: "urn:owned".into(),
                target: "copied.xml".into(),
                target_mode: None,
            },
            Relationship {
                id: "added".into(),
                rel_type: "urn:new".into(),
                target: "https://new.example/?a=1&b=2".into(),
                target_mode: Some("External".into()),
            },
        ];
        let map = HashMap::from([("kept".into(), "renamed".into())]);
        let changed = original.rebind_for_import(items.clone(), &map).unwrap();
        assert_eq!(changed.items, items);
        let raw = String::from_utf8(changed.to_xml().unwrap()).unwrap();
        assert!(
            raw.contains("z:root='keep'")
                && raw.contains("z:record='keep'")
                && raw.contains("<z:child/>")
                && raw.contains("<!--between-->")
                && raw.contains("<z:tail/>")
        );
        assert!(raw.contains("<z:capsule><q:Relationship Id='foreign' Type='urn:metadata' Target='foreign.xml'/></z:capsule>"));
        assert!(!raw.contains("Id='removed'"));
        assert!(!raw.contains("TargetMode='External'"));
        assert_eq!(changed.items[1].target, "https://new.example/?a=1&b=2");
        assert_eq!(original.to_xml().unwrap(), xml);
        assert!(
            original
                .rebind_for_import(vec![items[0].clone(), items[0].clone()], &map)
                .is_err()
        );
        let mut invalid = items.clone();
        invalid[0].target.clear();
        assert!(original.rebind_for_import(invalid, &map).is_err());
        let mut invalid = items;
        invalid[0].target_mode = Some("invalid".into());
        assert!(original.rebind_for_import(invalid, &map).is_err());
        assert!(
            original
                .rebind_for_import(
                    Vec::new(),
                    &HashMap::from([("foreign".into(), "guessed".into())])
                )
                .is_err()
        );
        assert_eq!(original.to_xml().unwrap(), xml);
    }

    #[test]
    fn fx216_relationship_rebind_maps_rows_added_after_parsing() {
        let xml = br#"<q:Relationships xmlns:q="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:opaque" z:root='keep'><q:Relationship Id='kept' Type='urn:owned' Target='old.xml' z:record='keep'><!--inside--></q:Relationship><!--between--><z:tail/></q:Relationships>"#;
        let mut original = Relationships::from_xml(xml).unwrap();
        original.add_with_id("draft", "urn:new", "new.xml");
        let before = original.to_xml().unwrap();
        let mut items = original.items.clone();
        items.last_mut().unwrap().id = "imported".into();
        let map = HashMap::from([("draft".into(), "imported".into())]);
        let changed = original.rebind_for_import(items.clone(), &map).unwrap();
        assert_eq!(changed.items, items);
        let saved = changed.to_xml().unwrap();
        let written = std::str::from_utf8(&saved).unwrap();
        for opaque in [
            "z:root='keep'",
            "z:record='keep'",
            "<!--inside-->",
            "<!--between-->",
            "<z:tail/>",
        ] {
            assert!(written.contains(opaque), "lost {opaque}");
        }
        assert_eq!(Relationships::from_xml(&saved).unwrap().items, items);
        assert!(
            original
                .rebind_for_import(
                    items.clone(),
                    &HashMap::from([("unknown".into(), "imported".into())])
                )
                .is_err()
        );
        let mut duplicate = original.clone();
        duplicate.items.push(duplicate.items[0].clone());
        assert!(duplicate.rebind_for_import(items, &map).is_err());
        assert_eq!(original.to_xml().unwrap(), before);
    }

    #[test]
    fn fx216_relationship_rebind_requires_explicit_renamed_raw_owner() {
        let xml = br#"<q:Relationships xmlns:q="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:opaque" z:root='keep'><q:Relationship Id='kept' Type='urn:owned' Target='old.xml' z:record='keep'><z:child keep='&#65;'/><!--inside--></q:Relationship><!--between--><z:tail/></q:Relationships>"#;
        let mut original = Relationships::from_xml(xml).unwrap();
        original.items[0].id = "draft".into();
        let before_items = original.items.clone();
        let before_source = original.source.clone();
        let before = original.to_xml().unwrap();
        let mut items = original.items.clone();
        items[0].id = "imported".into();
        items[0].target = "new.xml".into();

        // Neither the new public ID nor matching position/target proves which
        // retained XML row owns the opaque payload.
        let current_map = HashMap::from([("draft".into(), "imported".into())]);
        assert!(matches!(
            original.rebind_for_import(items.clone(), &current_map),
            Err(crate::OpcError::InvalidRelationship)
        ));
        assert!(matches!(
            original.rebind_for_import(original.items.clone(), &HashMap::new()),
            Err(crate::OpcError::InvalidRelationship)
        ));
        assert_eq!(original.items, before_items);
        assert_eq!(original.source, before_source);
        assert_eq!(original.to_xml().unwrap(), before);

        // Naming the retained source ID is a complete ownership route even
        // though the caller has already changed the public item ID.
        let retained_map = HashMap::from([("kept".into(), "imported".into())]);
        let changed = original
            .rebind_for_import(items.clone(), &retained_map)
            .unwrap();
        assert_eq!(changed.items, items);
        let saved = changed.to_xml().unwrap();
        let written = std::str::from_utf8(&saved).unwrap();
        for opaque in [
            "z:root='keep'",
            "z:record='keep'",
            "<z:child keep='&#65;'/>",
            "<!--inside-->",
            "<!--between-->",
            "<z:tail/>",
        ] {
            assert!(written.contains(opaque), "lost {opaque}");
        }
        assert!(written.contains("Id='imported'"));
        assert!(written.contains("Target='new.xml'"));
        assert_eq!(Relationships::from_xml(&saved).unwrap().items, items);
        assert_eq!(original.items, before_items);
        assert_eq!(original.source, before_source);
        assert_eq!(original.to_xml().unwrap(), before);
    }

    #[test]
    fn fx216_relationship_rebind_distinguishes_removal_from_ambiguous_addition() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:z="urn:opaque" z:root='keep'><Relationship Id='kept' Type='urn:owned' Target='old.xml' z:record='keep'><z:child/></Relationship><!--between--><z:tail/></Relationships>"#;
        let parsed = Relationships::from_xml(xml).unwrap();
        let mut ambiguous = parsed.clone();
        ambiguous.items.clear();
        ambiguous.add_with_id("draft", "urn:new", "new.xml");
        let before_items = ambiguous.items.clone();
        let before_source = ambiguous.source.clone();
        let before = ambiguous.to_xml().unwrap();
        let mut items = ambiguous.items.clone();
        items[0].id = "imported".into();
        let map = HashMap::from([("draft".into(), "imported".into())]);
        assert!(matches!(
            ambiguous.rebind_for_import(items.clone(), &map),
            Err(crate::OpcError::InvalidRelationship)
        ));
        assert_eq!(ambiguous.items, before_items);
        assert_eq!(ambiguous.source, before_source);
        assert_eq!(ambiguous.to_xml().unwrap(), before);

        // A removal is still supported. Staging it first makes a subsequent
        // new row unambiguous without inventing a raw-owner correspondence.
        let mut removed = parsed
            .rebind_for_import(Vec::new(), &HashMap::new())
            .unwrap();
        let raw = String::from_utf8(removed.to_xml().unwrap()).unwrap();
        assert!(!raw.contains("z:record") && !raw.contains("<z:child/>"));
        assert!(
            raw.contains("z:root='keep'")
                && raw.contains("<!--between-->")
                && raw.contains("<z:tail/>")
        );
        removed.add_with_id("draft", "urn:new", "new.xml");
        let changed = removed.rebind_for_import(items.clone(), &map).unwrap();
        assert_eq!(changed.items, items);
        let raw = String::from_utf8(changed.to_xml().unwrap()).unwrap();
        assert!(raw.contains("z:root='keep'") && raw.contains("<z:tail/>"));
        assert!(
            changed
                .rebind_for_import(
                    changed.items.clone(),
                    &HashMap::from([("unknown".into(), "imported".into())]),
                )
                .is_err()
        );
        assert_eq!(parsed.to_xml().unwrap(), xml);
    }

    #[test]
    fn fx216_import_record_metadata_ignores_redundant_closure_declarations() {
        let uri = "http://schemas.openxmlformats.org/package/2006/relationships";
        let source = Relationships::from_xml(format!(r#"<Relationships xmlns="{uri}" xmlns:z="urn:style-record"><Relationship Id="picture" Type="image" Target="media/image.png" z:keep='&#69;'><z:payload value="z"/></Relationship></Relationships>"#).as_bytes()).unwrap();
        let destination =
            Relationships::from_xml(format!(r#"<Relationships xmlns="{uri}"/>"#).as_bytes())
                .unwrap();
        let (merged, ids) = destination.merge_for_import(&source).unwrap();
        assert!(
            source
                .record_metadata_equal_for_import("picture", &merged, &ids["picture"])
                .unwrap()
        );
        let changed = Relationships::from_xml(format!(r#"<Relationships xmlns="{uri}" xmlns:z="urn:style-record"><Relationship Id="picture" Type="image" Target="media/image.png" z:keep='&#69;' xmlns:z="urn:foreign"><z:payload value="z"/></Relationship></Relationships>"#).as_bytes()).unwrap();
        assert!(
            !source
                .record_metadata_equal_for_import("picture", &changed, "picture")
                .unwrap()
        );
        let encoded = Relationships::from_xml(format!(r#"<Relationships xmlns="{uri}"><Relationship Id="picture" Type="image" Target="media/image.png" z:keep='&#69;' xmlns:z="urn:style-rec&#111;rd"><z:payload value="z"/></Relationship></Relationships>"#).as_bytes()).unwrap();
        assert!(
            source
                .record_metadata_equal_for_import("picture", &encoded, "picture")
                .unwrap()
        );
    }

    #[test]
    fn fx216_import_record_metadata_checks_opaque_namespace_authority() {
        let uri = "http://schemas.openxmlformats.org/package/2006/relationships";
        let fixture = |binding: &str, id: &str| {
            format!(
                r#"<Relationships xmlns="{uri}" xmlns:z="{binding}"><Relationship Id="{id}" Type="font" Target="font.bin" z:marker='&#65;'><z:payload prefix="z"/></Relationship><z:capsule><Relationship Id="ghost" Type="font" Target="hidden"/></z:capsule></Relationships>"#
            )
        };
        let source = Relationships::from_xml(fixture("urn:metadata", "source").as_bytes()).unwrap();
        let equivalent =
            Relationships::from_xml(fixture("urn:met&#97;data", "destination").as_bytes()).unwrap();
        let foreign =
            Relationships::from_xml(fixture("urn:foreign", "destination").as_bytes()).unwrap();
        assert!(
            source
                .record_metadata_equal_for_import("source", &equivalent, "destination")
                .unwrap()
        );
        assert!(
            !source
                .record_metadata_equal_for_import("source", &foreign, "destination")
                .unwrap()
        );
        assert!(
            source
                .record_metadata_equal_for_import("ghost", &source, "source")
                .is_err()
        );
        let plain = Relationships::from_xml(format!(r#"<x:Relationships xmlns:x="{uri}"><x:Relationship Id="other" Type="font" Target="other.bin"/></x:Relationships>"#).as_bytes()).unwrap();
        let other = Relationships::from_xml(format!(r#"<Relationships xmlns="{uri}"><Relationship Id="source" Type="font" Target="font.bin"></Relationship></Relationships>"#).as_bytes()).unwrap();
        assert!(
            plain
                .record_metadata_equal_for_import("other", &other, "source")
                .unwrap()
        );
    }
}
