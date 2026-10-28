//! Deterministic native document comparison and tracked-revision generation.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;

use base64::Engine;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, Writer, XmlVersion};
use rdocx_oxml::content_control::{CT_Sdt, SdtContent};
use rdocx_oxml::document::{BodyContent, CT_Document};
use rdocx_oxml::namespace::W_NS;
use rdocx_oxml::properties::CT_PPr;
use rdocx_oxml::shared::ST_PageOrientation;
use rdocx_oxml::table::{CT_Row, CT_Tbl, CT_TblPr, CT_Tc, CT_TrPr, CellContent};
use rdocx_oxml::text::{
    CT_P, CT_R, CT_Text, CommentRangeMarker, RunContent, declare_w14_on_part_root,
};
use sha2::{Digest, Sha256};

use crate::revision::validate_revision_timestamp;
use crate::{Document, Error, Result, StoryKind};

use oxml_opc::OpcPackage;
use oxml_opc::relationship::{Relationship, rel_types};

pub(crate) const COMMENT_COMPARISON_PART: &str = "/customXml/rdocxComparisonComments.xml";
pub(crate) const COMMENT_COMPARISON_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
pub(crate) const COMMENT_COMPARISON_OPEN: &str =
    "<rdocx:comparisonComments xmlns:rdocx=\"urn:rdocx:comparison:comments:1\">";
pub(crate) const COMMENT_COMPARISON_CLOSE: &str = "</rdocx:comparisonComments>";

fn comment_relationship(relationship: &Relationship) -> bool {
    matches!(
        relationship.rel_type.as_str(),
        rel_types::COMMENTS | crate::comments::COMMENTS_EXTENDED_REL_TYPE
    )
}

fn related_part_snapshot(
    document: &Document,
    name: &str,
    seen: &mut HashSet<String>,
) -> Result<serde_json::Value> {
    if !seen.insert(name.to_owned()) {
        return Ok(serde_json::Value::Null);
    }
    let xml = document
        .package
        .get_part(name)
        .ok_or_else(|| Error::Other(format!("missing comment relationship target {name}")))?;
    let relationships = document.package.get_part_rels(name);
    let mut children = Vec::new();
    if let Some(relationships) = relationships {
        for relationship in &relationships.items {
            if crate::document::relationship_is_internal(relationship) {
                let target = OpcPackage::resolve_rel_target(name, &relationship.target);
                let child = related_part_snapshot(document, &target, seen)?;
                if !child.is_null() {
                    children.push(child);
                }
            }
        }
    }
    let rels = relationships
        .map(|rels| rels.to_xml())
        .transpose()
        .map_err(|error| Error::Other(error.to_string()))?;
    Ok(serde_json::json!({
        "name": name,
        "content_type": document.package.content_types.content_type_for(name),
        "xml": base64::engine::general_purpose::STANDARD.encode(xml),
        "rels": rels.map(|raw| base64::engine::general_purpose::STANDARD.encode(raw)),
        "related": children,
    }))
}

fn comment_snapshot(document: &Document) -> Result<serde_json::Value> {
    let mut parts = Vec::new();
    let mut seen = HashSet::new();
    if let Some(relationships) = document.package.get_part_rels(&document.doc_part_name) {
        for relationship in &relationships.items {
            if !comment_relationship(relationship)
                || !crate::document::relationship_is_internal(relationship)
            {
                continue;
            }
            let name =
                OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
            let mut part = related_part_snapshot(document, &name, &mut seen)?;
            let object = part
                .as_object_mut()
                .ok_or_else(|| Error::Other("duplicate comment comparison part".to_owned()))?;
            object.extend(
                serde_json::json!({
                    "id": relationship.id,
                    "type": relationship.rel_type,
                    "target": relationship.target,
                })
                .as_object()
                .cloned()
                .unwrap_or_default(),
            );
            parts.push(part);
        }
    }
    Ok(serde_json::Value::Array(parts))
}

fn snapshot_part_index<'a>(
    part: &'a serde_json::Value,
    index: &mut HashMap<String, &'a serde_json::Value>,
) -> Result<()> {
    let name = part
        .get("name")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| Error::Other("comparison comment part has no name".to_owned()))?;
    index.insert(name.to_owned(), part);
    for child in part
        .get("related")
        .and_then(serde_json::Value::as_array)
        .ok_or_else(|| Error::Other("comparison comment part has no related list".to_owned()))?
    {
        snapshot_part_index(child, index)?;
    }
    Ok(())
}

fn snapshot_part_signature(
    name: &str,
    index: &HashMap<String, &serde_json::Value>,
    visited: &mut HashSet<String>,
) -> Result<String> {
    if !visited.insert(name.to_owned()) {
        return Ok("shared".to_owned());
    }
    let part = index
        .get(name)
        .ok_or_else(|| Error::Other(format!("comparison comment target {name} is absent")))?;
    let xml = part
        .get("xml")
        .and_then(serde_json::Value::as_str)
        .unwrap_or_default();
    let content_type = part.get("content_type").and_then(serde_json::Value::as_str);
    let mut rel_signatures = Vec::new();
    if let Some(raw) = part.get("rels").and_then(serde_json::Value::as_str) {
        let decoded = base64::engine::general_purpose::STANDARD
            .decode(raw)
            .map_err(|error| Error::Other(error.to_string()))?;
        let rels = oxml_opc::relationship::Relationships::from_xml(&decoded)
            .map_err(|error| Error::Other(error.to_string()))?;
        for relationship in rels.items {
            let target = if crate::document::relationship_is_internal(&relationship) {
                let child = OpcPackage::resolve_rel_target(name, &relationship.target);
                snapshot_part_signature(&child, index, visited)?
            } else {
                relationship.target.clone()
            };
            rel_signatures.push(format!(
                "{:?}:{:?}:{:?}:{target}",
                relationship.id, relationship.rel_type, relationship.target_mode
            ));
        }
    }
    rel_signatures.sort();
    visited.remove(name);
    Ok(format!("{content_type:?}:{xml}:{rel_signatures:?}"))
}

fn comment_snapshots_match(left: &serde_json::Value, right: &serde_json::Value) -> Result<bool> {
    let signature = |value: &serde_json::Value| -> Result<Vec<String>> {
        let roots = value
            .as_array()
            .ok_or_else(|| Error::Other("invalid comparison comment snapshot".to_owned()))?;
        let mut index = HashMap::new();
        for root in roots {
            snapshot_part_index(root, &mut index)?;
        }
        let mut signatures = Vec::new();
        for root in roots {
            let name = root
                .get("name")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| Error::Other("comment root has no name".to_owned()))?;
            signatures.push(format!(
                "{name}:{}",
                snapshot_part_signature(name, &index, &mut HashSet::new())?
            ));
        }
        signatures.sort();
        Ok(signatures)
    };
    Ok(signature(left)? == signature(right)?)
}

fn apply_related_part(document: &mut Document, part: &serde_json::Value) -> Result<()> {
    let field = |key: &str| -> Result<&str> {
        part.get(key)
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| Error::Other(format!("invalid comparison comment field {key}")))
    };
    let name = field("name")?;
    let xml = base64::engine::general_purpose::STANDARD
        .decode(field("xml")?)
        .map_err(|error| Error::Other(error.to_string()))?;
    document.package.set_part(name, xml);
    if let Some(content_type) = part.get("content_type").and_then(serde_json::Value::as_str) {
        document
            .package
            .content_types
            .add_override(name, content_type);
    }
    if let Some(encoded) = part.get("rels").and_then(serde_json::Value::as_str) {
        let xml = base64::engine::general_purpose::STANDARD
            .decode(encoded)
            .map_err(|error| Error::Other(error.to_string()))?;
        let rels = oxml_opc::relationship::Relationships::from_xml(&xml)
            .map_err(|error| Error::Other(error.to_string()))?;
        document.package.set_part_rels(name, rels);
    } else {
        document.package.remove_part_rels(name);
    }
    for child in part
        .get("related")
        .and_then(serde_json::Value::as_array)
        .ok_or_else(|| Error::Other("invalid related comment parts".to_owned()))?
    {
        apply_related_part(document, child)?;
    }
    Ok(())
}

fn comment_part_has_inbound(document: &Document, child: &str) -> bool {
    document
        .package
        .package_rels
        .items
        .iter()
        .any(|relationship| {
            crate::document::relationship_is_internal(relationship)
                && OpcPackage::resolve_rel_target("/", &relationship.target) == child
        })
        || document.package.part_rels.iter().any(|(owner, rels)| {
            rels.items.iter().any(|relationship| {
                crate::document::relationship_is_internal(relationship)
                    && OpcPackage::resolve_rel_target(owner, &relationship.target) == child
            })
        })
}

fn remove_orphan_comment_part(document: &mut Document, name: &str, seen: &mut HashSet<String>) {
    if !seen.insert(name.to_owned()) {
        return;
    }
    let children = document
        .package
        .remove_part_rels(name)
        .into_iter()
        .flat_map(|rels| rels.items)
        .filter(crate::document::relationship_is_internal)
        .map(|relationship| OpcPackage::resolve_rel_target(name, &relationship.target))
        .collect::<Vec<_>>();
    document.package.remove_part(name);
    document.package.content_types.remove_override(name);
    for child in children {
        if !comment_part_has_inbound(document, &child) {
            remove_orphan_comment_part(document, &child, seen);
        }
    }
}

fn extended_comment_snapshots_match(left: &serde_json::Value, right: &serde_json::Value) -> bool {
    let extended = |value: &serde_json::Value| {
        value
            .as_array()
            .into_iter()
            .flatten()
            .filter(|part| {
                part.get("type").and_then(serde_json::Value::as_str)
                    == Some(crate::comments::COMMENTS_EXTENDED_REL_TYPE)
            })
            .map(|part| {
                part.get("xml")
                    .and_then(serde_json::Value::as_str)
                    .unwrap_or_default()
                    .to_owned()
            })
            .collect::<Vec<_>>()
    };
    extended(left) == extended(right)
}

fn comment_package_links_match(left: &serde_json::Value, right: &serde_json::Value) -> bool {
    let links = |value: &serde_json::Value| {
        let mut parts = value
            .as_array()
            .into_iter()
            .flatten()
            .map(|part| {
                (
                    part.get("name").cloned().unwrap_or_default().to_string(),
                    part.get("rels").cloned().unwrap_or_default().to_string(),
                    part.get("related").cloned().unwrap_or_default().to_string(),
                )
            })
            .collect::<Vec<_>>();
        parts.sort();
        parts
    };
    links(left) == links(right)
}

fn remap_comment_targets(document: &Document, snapshot: &mut serde_json::Value) -> Result<()> {
    fn plan(
        document: &Document,
        part: &serde_json::Value,
        child: bool,
        names: &mut HashMap<String, String>,
        occupied: &mut HashSet<String>,
    ) -> Result<()> {
        let name = part
            .get("name")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| Error::Other("comment snapshot target has no name".to_owned()))?;
        let encoded = part
            .get("xml")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| Error::Other("comment snapshot target has no bytes".to_owned()))?;
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(encoded)
            .map_err(|error| Error::Other(error.to_string()))?;
        let desired_rels = part
            .get("rels")
            .and_then(serde_json::Value::as_str)
            .map(|value| base64::engine::general_purpose::STANDARD.decode(value))
            .transpose()
            .map_err(|error| Error::Other(error.to_string()))?;
        let existing_rels = document
            .package
            .get_part_rels(name)
            .map(|rels| rels.to_xml())
            .transpose()
            .map_err(|error| Error::Other(error.to_string()))?;
        let desired_content_type = part.get("content_type").and_then(serde_json::Value::as_str);
        let collision = document
            .package
            .get_part(name)
            .is_some_and(|old| old != bytes)
            || existing_rels != desired_rels
            || (document.package.contains_part(name)
                && document.package.content_types.content_type_for(name) != desired_content_type);
        let target = if child && collision {
            let (stem, extension) = name.rsplit_once('.').unwrap_or((name, ""));
            let mut ordinal = 1usize;
            loop {
                let candidate = if extension.is_empty() {
                    format!("{stem}-rdocx-comment-{ordinal}")
                } else {
                    format!("{stem}-rdocx-comment-{ordinal}.{extension}")
                };
                if !document.package.contains_part(&candidate)
                    && occupied.insert(candidate.to_ascii_lowercase())
                {
                    break candidate;
                }
                ordinal = ordinal
                    .checked_add(1)
                    .ok_or_else(|| Error::Other("comment target names exhausted".to_owned()))?;
            }
        } else {
            name.to_owned()
        };
        names.insert(name.to_owned(), target);
        for related in part
            .get("related")
            .and_then(serde_json::Value::as_array)
            .ok_or_else(|| Error::Other("comment snapshot lacks related parts".to_owned()))?
        {
            plan(document, related, true, names, occupied)?;
        }
        Ok(())
    }

    fn rewrite(part: &mut serde_json::Value, names: &HashMap<String, String>) -> Result<()> {
        let old_name = part
            .get("name")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| Error::Other("comment snapshot target has no name".to_owned()))?
            .to_owned();
        if let Some(encoded) = part.get("rels").and_then(serde_json::Value::as_str) {
            let bytes = base64::engine::general_purpose::STANDARD
                .decode(encoded)
                .map_err(|error| Error::Other(error.to_string()))?;
            let mut rels = oxml_opc::relationship::Relationships::from_xml(&bytes)
                .map_err(|error| Error::Other(error.to_string()))?;
            for relationship in &mut rels.items {
                if crate::document::relationship_is_internal(relationship) {
                    let target = OpcPackage::resolve_rel_target(&old_name, &relationship.target);
                    if let Some(mapped) = names.get(&target) {
                        relationship.target = mapped.clone();
                    }
                }
            }
            let xml = rels
                .to_xml()
                .map_err(|error| Error::Other(error.to_string()))?;
            part["rels"] =
                serde_json::Value::String(base64::engine::general_purpose::STANDARD.encode(xml));
        }
        part["name"] = serde_json::Value::String(names.get(&old_name).cloned().unwrap_or(old_name));
        for related in part
            .get_mut("related")
            .and_then(serde_json::Value::as_array_mut)
            .ok_or_else(|| Error::Other("comment snapshot lacks related parts".to_owned()))?
        {
            rewrite(related, names)?;
        }
        Ok(())
    }

    let roots = snapshot
        .as_array()
        .ok_or_else(|| Error::Other("invalid comparison comment snapshot".to_owned()))?;
    let mut occupied = document
        .package
        .parts
        .keys()
        .map(|name| name.to_ascii_lowercase())
        .collect::<HashSet<_>>();
    let mut names = HashMap::new();
    for root in roots {
        plan(document, root, false, &mut names, &mut occupied)?;
    }
    for root in snapshot
        .as_array_mut()
        .expect("snapshot roots were checked")
    {
        rewrite(root, &names)?;
    }
    Ok(())
}

pub(crate) fn apply_comment_snapshot(
    document: &mut Document,
    snapshot: &serde_json::Value,
) -> Result<()> {
    let parts = snapshot
        .as_array()
        .ok_or_else(|| Error::Other("invalid comparison comment snapshot".to_owned()))?;
    if let Some(relationships) = document.package.get_part_rels_mut(&document.doc_part_name) {
        let removed = relationships
            .items
            .iter()
            .filter(|item| comment_relationship(item))
            .map(|item| OpcPackage::resolve_rel_target(&document.doc_part_name, &item.target))
            .collect::<Vec<_>>();
        relationships
            .items
            .retain(|item| !comment_relationship(item));
        let mut children = Vec::new();
        for name in removed {
            if let Some(rels) = document.package.remove_part_rels(&name) {
                children.extend(
                    rels.items
                        .into_iter()
                        .filter(crate::document::relationship_is_internal)
                        .map(|relationship| {
                            OpcPackage::resolve_rel_target(&name, &relationship.target)
                        }),
                );
            }
            document.package.remove_part(&name);
            document.package.content_types.remove_override(&name);
        }
        let mut seen = HashSet::new();
        for child in children {
            if !comment_part_has_inbound(document, &child) {
                remove_orphan_comment_part(document, &child, &mut seen);
            }
        }
    }
    for part in parts {
        let field = |key: &str| -> Result<&str> {
            part.get(key)
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| Error::Other(format!("invalid comparison comment field {key}")))
        };
        apply_related_part(document, part)?;
        document
            .package
            .get_or_create_part_rels(&document.doc_part_name)
            .items
            .push(Relationship {
                id: field("id")?.to_owned(),
                rel_type: field("type")?.to_owned(),
                target: field("target")?.to_owned(),
                target_mode: None,
            });
    }
    document.comments = None;
    document.comments_part_name = None;
    document.comments_owned = false;
    document.comments_extended = None;
    document.comments_extended_part_name = None;
    document.comments_extended_owned = false;
    Ok(())
}

#[cfg(test)]
thread_local! {
    static FAIL_AFTER_COMPARISON_STAGING: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

type ControlPropertySignature<'a> = Option<(
    Option<rdocx_oxml::content_control::SdtType>,
    Option<&'a rdocx_oxml::content_control::CT_DataBinding>,
)>;

/// The content-control properties that name, protect or file a control
/// without changing what it holds, by their `w:sdtPr` element names.
const CONTROL_METADATA: [&str; 5] = ["tag", "alias", "lock", "placeholder", "docPartGallery"];

/// A comparison difference that cannot be represented as a content revision.
///
/// The redline keeps the original for every diagnostic. The message starts
/// with a stable prefix naming the difference: `formatting differs` for
/// formatting that cannot be revised, and `content-control <name> differs`
/// for a control's `tag`, `alias`, `lock`, `placeholder` or
/// `docPartGallery`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComparisonDiagnostic {
    pub location: String,
    pub message: String,
}

/// A Word story category that can be excluded from document comparison.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ComparisonStoryKind {
    Main,
    Header,
    Footer,
    Comment,
    TextBox,
    Footnote,
    Endnote,
}

/// The text boundary used when generating content revisions.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ComparisonGranularity {
    /// Preserve the legacy whole-run comparison behavior.
    #[default]
    Run,
    /// Compare maximal word, whitespace, and punctuation or symbol units.
    Word,
    /// Compare individual Unicode scalar values.
    Character,
}

/// Policy controls for native document comparison.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ComparisonOptions {
    pub granularity: ComparisonGranularity,
    pub ignore_formatting: bool,
    pub ignore_whitespace: bool,
    pub ignore_fields: bool,
    pub ignore_comments: bool,
    pub ignored_stories: Vec<ComparisonStoryKind>,
}

struct Metadata<'a> {
    author: &'a str,
    timestamp: &'a str,
    options: &'a ComparisonOptions,
    ids: IdAllocator,
    /// Whether an input holds a modeled revision, which the comparison
    /// keeps only in owners that both inputs hold unchanged.
    existing_revisions: bool,
}

struct IdAllocator {
    used: HashSet<i32>,
    next: i32,
}

impl ComparisonStoryKind {
    fn label(self) -> &'static str {
        match self {
            Self::Header => "header",
            Self::Footer => "footer",
            Self::Comment => "comments",
            Self::Footnote => "footnotes",
            Self::Endnote => "endnotes",
            Self::Main => "body",
            Self::TextBox => "text-box",
        }
    }

    fn root_local(self) -> &'static str {
        match self {
            Self::Header => "hdr",
            Self::Footer => "ftr",
            Self::Comment => "comments",
            Self::Footnote => "footnotes",
            Self::Endnote => "endnotes",
            Self::Main | Self::TextBox => "document",
        }
    }

    fn owner_local(self) -> Option<&'static str> {
        match self {
            Self::Comment => Some("comment"),
            Self::Footnote => Some("footnote"),
            Self::Endnote => Some("endnote"),
            Self::Main | Self::Header | Self::Footer | Self::TextBox => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StoryPart {
    kind: ComparisonStoryKind,
    part_name: String,
}

#[derive(Default)]
struct TextBoxMarkers {
    main: String,
    related: HashMap<String, String>,
}

impl IdAllocator {
    fn new(used: HashSet<i32>) -> Self {
        Self { used, next: 0 }
    }

    fn allocate(&mut self) -> Result<i32> {
        while self.used.contains(&self.next) {
            self.next = self
                .next
                .checked_add(1)
                .ok_or_else(|| Error::Other("comparison revision ids are exhausted".to_owned()))?;
        }
        let id = self.next;
        self.used.insert(id);
        self.next = self.next.saturating_add(1);
        Ok(id)
    }

    fn revision(
        &mut self,
        kind: &str,
        author: &str,
        timestamp: &str,
        inner: &str,
    ) -> Result<String> {
        let id = self.allocate()?;
        Ok(Self::revision_with_id(kind, author, timestamp, inner, id))
    }

    fn revision_with_id(kind: &str, author: &str, timestamp: &str, inner: &str, id: i32) -> String {
        let author = quick_xml::escape::escape(author);
        let timestamp = quick_xml::escape::escape(timestamp);
        format!(
            r#"<w:{kind} w:id="{id}" w:author="{author}" w:date="{timestamp}">{inner}</w:{kind}>"#
        )
    }

    fn marker(&mut self, kind: &str, author: &str, timestamp: &str) -> Result<String> {
        let id = self.allocate()?;
        Ok(Self::marker_with_id(kind, author, timestamp, id))
    }

    fn marker_with_id(kind: &str, author: &str, timestamp: &str, id: i32) -> String {
        let author = quick_xml::escape::escape(author);
        let timestamp = quick_xml::escape::escape(timestamp);
        format!(r#"<w:{kind} w:id="{id}" w:author="{author}" w:date="{timestamp}"/>"#)
    }
}

struct MovePairs {
    original: Vec<Option<i32>>,
    edited: Vec<Option<i32>>,
}

fn pair_moves(
    aligned: &[(Option<usize>, Option<usize>)],
    original: &[String],
    edited: &[String],
    ids: &mut IdAllocator,
) -> Result<MovePairs> {
    let mut pairs = MovePairs {
        original: vec![None; original.len()],
        edited: vec![None; edited.len()],
    };
    let unmatched_edited = aligned
        .iter()
        .filter_map(|(left, right)| left.is_none().then_some(*right).flatten())
        .collect::<Vec<_>>();
    for left in aligned
        .iter()
        .filter_map(|(left, right)| right.is_none().then_some(*left).flatten())
    {
        let Some(right) = unmatched_edited
            .iter()
            .copied()
            .find(|right| pairs.edited[*right].is_none() && original[left] == edited[*right])
        else {
            continue;
        };
        let id = ids.allocate()?;
        pairs.original[left] = Some(id);
        pairs.edited[right] = Some(id);
    }
    Ok(pairs)
}

impl Document {
    /// Compare every supported Word story with `edited` and record tracked changes.
    pub fn compare(
        &mut self,
        edited: &Document,
        author: &str,
        timestamp: &str,
    ) -> Result<Vec<ComparisonDiagnostic>> {
        self.compare_with_options(edited, author, timestamp, &ComparisonOptions::default())
    }

    /// Compare supported Word stories using an explicit granularity and ignore policy.
    pub fn compare_with_options(
        &mut self,
        edited: &Document,
        author: &str,
        timestamp: &str,
        options: &ComparisonOptions,
    ) -> Result<Vec<ComparisonDiagnostic>> {
        validate_revision_timestamp(timestamp)?;
        validate_comparison_options(options)?;
        let original = comparison_input(self)?;
        let mut edited = comparison_input(edited)?;
        if original.package.contains_part(COMMENT_COMPARISON_PART)
            || edited.package.contains_part(COMMENT_COMPARISON_PART)
        {
            return Err(Error::Other(
                "comparison requires resolved comment revisions".to_owned(),
            ));
        }
        let comments_ignored = story_ignored(options, ComparisonStoryKind::Comment);
        let original_comments = if comments_ignored {
            serde_json::Value::Null
        } else {
            comment_snapshot(&original)?
        };
        let edited_comments = if comments_ignored {
            serde_json::Value::Null
        } else {
            comment_snapshot(&edited)?
        };
        let comments_changed =
            !comments_ignored && !comment_snapshots_match(&original_comments, &edited_comments)?;
        let mut original_stories = story_parts_with_options(&original, options)?;
        let mut edited_stories = story_parts_with_options(&edited, options)?;
        // Existing revisions are kept where both inputs hold them unchanged,
        // and every other change near them is refused at its owner.
        let existing_revisions = contains_modeled_revisions(&original, &original_stories, options)?
            || contains_modeled_revisions(&edited, &edited_stories, options)?;
        if comments_changed
            && let Some((left, right)) = original_stories
                .iter()
                .find(|story| story.kind == ComparisonStoryKind::Comment)
                .zip(
                    edited_stories
                        .iter()
                        .find(|story| story.kind == ComparisonStoryKind::Comment),
                )
            && left == right
        {
            let original_source =
                std::str::from_utf8(story_xml(&original, left)?).map_err(utf8_error)?;
            let edited_source =
                std::str::from_utf8(story_xml(&edited, right)?).map_err(utf8_error)?;
            let root_shell = |source| -> Result<Vec<String>> {
                Ok(canonical_owned_story(source, "comment")?
                    .0
                    .into_iter()
                    .filter(|token| token != "owner")
                    .collect())
            };
            if root_shell(original_source)? != root_shell(edited_source)? {
                return Err(Error::Other(format!(
                    "comments story root shell changed in {}",
                    left.part_name
                )));
            }
        }
        let comment_stories = original_stories
            .iter()
            .find(|story| story.kind == ComparisonStoryKind::Comment)
            .zip(
                edited_stories
                    .iter()
                    .find(|story| story.kind == ComparisonStoryKind::Comment),
            );
        let compare_comment_story = |left: &StoryPart, right: &StoryPart| {
            compare_story_part(
                story_xml(&original, left).unwrap_or_default(),
                story_xml(&edited, right).unwrap_or_default(),
                left,
                &mut Metadata {
                    author,
                    timestamp,
                    options,
                    ids: IdAllocator::new(HashSet::new()),
                    existing_revisions,
                },
                &mut Vec::new(),
            )
        };
        let compatible_comment_story = comment_stories.is_some_and(|(left, right)| {
            left == right && compare_comment_story(left, right).is_ok()
        });
        let snapshot_comments = comments_changed
            && (!compatible_comment_story
                || !extended_comment_snapshots_match(&original_comments, &edited_comments)
                || !comment_package_links_match(&original_comments, &edited_comments));
        if snapshot_comments && existing_revisions {
            // The comment package revision replaces every comment part whole,
            // so it would rewrite the revisions a comment holds.
            for (document, stories) in [(&original, &original_stories), (&edited, &edited_stories)]
            {
                for story in stories
                    .iter()
                    .filter(|story| story.kind == ComparisonStoryKind::Comment)
                {
                    if modeled_revision_count_with_options(story_xml(document, story)?, options)?
                        > 0
                    {
                        let reason = match comment_stories {
                            Some((left, right)) if left == right => compare_comment_story(
                                left, right,
                            )
                            .err()
                            .map_or_else(
                                || "comments were added, removed, answered or resolved".to_owned(),
                                |error| error.to_string(),
                            ),
                            _ => "the comment parts differ".to_owned(),
                        };
                        return Err(Error::Other(format!(
                            "comparison cannot replace the comments of {} whole, because a comment holds an existing tracked revision: {reason}",
                            story.part_name
                        )));
                    }
                }
            }
        }
        if snapshot_comments {
            original_stories.retain(|story| story.kind != ComparisonStoryKind::Comment);
            edited_stories.retain(|story| story.kind != ComparisonStoryKind::Comment);
        }
        if original_stories != edited_stories {
            return Err(Error::Other(
                "document comparison requires identical related-story shells".to_owned(),
            ));
        }
        let (carried_links, mut charts) = remap_equivalent_story_relationships(
            &original,
            &mut edited,
            &original_stories,
            &edited_stories,
        )?;
        reject_cross_story_moves(&original, &edited, &original_stories, options)?;

        let original_xml = original
            .package
            .get_part(&original.doc_part_name)
            .ok_or_else(|| {
                Error::Other(format!("missing main story {}", original.doc_part_name))
            })?;
        let edited_xml = edited
            .package
            .get_part(&edited.doc_part_name)
            .ok_or_else(|| Error::Other(format!("missing main story {}", edited.doc_part_name)))?;
        if original_xml == edited_xml {
            let mut stories_unchanged = true;
            for story in &original_stories {
                if story_xml(&original, story)? != story_xml(&edited, story)? {
                    stories_unchanged = false;
                    break;
                }
            }
            if stories_unchanged && !snapshot_comments {
                return Ok(Vec::new());
            }
        }
        let text_box_markers =
            comparison_text_box_markers(&original, &edited, &original_stories, options)?;
        let mut used_ids = if story_ignored(options, ComparisonStoryKind::Main) {
            HashSet::new()
        } else {
            word_ids_with_options(original_xml, options)?
        };
        if !story_ignored(options, ComparisonStoryKind::Main) {
            used_ids.extend(word_ids_with_options(edited_xml, options)?);
        }
        for story in &original_stories {
            used_ids.extend(word_ids_with_options(
                story_xml(&original, story)?,
                options,
            )?);
            used_ids.extend(word_ids_with_options(story_xml(&edited, story)?, options)?);
        }
        let input_ids = used_ids.clone();
        let mut metadata = Metadata {
            author,
            timestamp,
            options,
            ids: IdAllocator::new(used_ids),
            existing_revisions,
        };
        let mut diagnostics = Vec::new();
        let tracked_body = if story_ignored(options, ComparisonStoryKind::Main) {
            extract_body_inner(original_xml)?.to_owned()
        } else if story_ignored(options, ComparisonStoryKind::TextBox) {
            let original_source = std::str::from_utf8(original_xml).map_err(utf8_error)?;
            let edited_source = std::str::from_utf8(edited_xml).map_err(utf8_error)?;
            let original_bindings = root_namespace_bindings(original_source, "document", &[])?;
            let edited_bindings = root_namespace_bindings(edited_source, "document", &[])?;
            compare_story_inner(
                extract_body_inner(original_xml)?,
                extract_body_inner(edited_xml)?,
                "body",
                "w",
                &original_bindings,
                &edited_bindings,
                &mut metadata,
                &mut diagnostics,
            )?
        } else {
            let original_body = extract_body_inner(original_xml)?;
            let edited_body = extract_body_inner(edited_xml)?;
            let original_spans = story_content_spans(original_body, "w")?;
            let edited_spans = story_content_spans(edited_body, "w")?;
            let original_source = std::str::from_utf8(original_xml).map_err(utf8_error)?;
            let edited_source = std::str::from_utf8(edited_xml).map_err(utf8_error)?;
            let original_bindings = root_namespace_bindings(original_source, "document", &[])?;
            let edited_bindings = root_namespace_bindings(edited_source, "document", &[])?;
            let original_document = story_document(original_body, &original_bindings)?;
            let edited_document = story_document(edited_body, &edited_bindings)?;
            if original_spans.len() != original_document.body.content.len()
                || edited_spans.len() != edited_document.body.content.len()
            {
                return Err(Error::Other(
                    "comparison could not correlate main-story owners".to_owned(),
                ));
            }
            let held_original = HeldRevisions::of(
                original_body,
                &original_spans,
                &original_bindings,
                &metadata,
            )?;
            let held_edited =
                HeldRevisions::of(edited_body, &edited_spans, &edited_bindings, &metadata)?;
            compare_body(
                &original_document,
                &edited_document,
                "body",
                Some((original_body, &original_spans)),
                [&held_original, &held_edited],
                &mut metadata,
                &mut diagnostics,
            )?
        };
        let mut tracked_xml = replace_body_inner(original_xml, &tracked_body)?.into_bytes();
        // Content from the edited document can use the `w14` its own root
        // declares, which the original root may not.
        declare_w14_on_part_root(&mut tracked_xml)?;
        let tracked_xml = crate::document::uniquify_drawing_ids_in_xml(&tracked_xml)?;
        let tracked_model_xml = close_drawing_namespaces(tracked_xml, true)?;
        let tracked = CT_Document::from_xml(&tracked_model_xml)?;
        tracked.to_xml()?;
        let mut candidate = original.clone_for_staging();
        candidate.document = tracked;
        candidate
            .package
            .set_part(&candidate.doc_part_name, tracked_model_xml);
        for story in &original_stories {
            let tracked_story = compare_story_part(
                story_xml(&original, story)?,
                story_xml(&edited, story)?,
                story,
                &mut metadata,
                &mut diagnostics,
            )?;
            candidate.package.set_part(&story.part_name, tracked_story);
        }
        if snapshot_comments {
            let revision_id = metadata.ids.allocate()?;
            let mut carried_comments = edited_comments.clone();
            remap_comment_targets(&candidate, &mut carried_comments)?;
            apply_comment_snapshot(&mut candidate, &carried_comments)?;
            let record = serde_json::to_vec(&serde_json::json!({
                "version": 1,
                "author": author,
                "timestamp": timestamp,
                "id": revision_id,
                "original": original_comments,
            }))
            .map_err(|error| Error::Other(error.to_string()))?;
            candidate.package.set_part(
                COMMENT_COMPARISON_PART,
                format!(
                    "{COMMENT_COMPARISON_OPEN}{}{COMMENT_COMPARISON_CLOSE}",
                    base64::engine::general_purpose::STANDARD.encode(record)
                )
                .into_bytes(),
            );
            candidate
                .package
                .content_types
                .add_override(COMMENT_COMPARISON_PART, "application/xml");
            candidate
                .package
                .get_or_create_part_rels(&candidate.doc_part_name)
                .add(
                    COMMENT_COMPARISON_REL,
                    "../customXml/rdocxComparisonComments.xml",
                );
        }
        // A hyperlink that only the edited side targets arrives with the
        // edited content that holds it.
        let mut referenced_by_owner = HashMap::<String, HashSet<String>>::new();
        let mut carried_image = false;
        for carried in carried_links {
            let CarriedRelationship {
                owner,
                relationship,
                media,
                chart,
            } = carried;
            if !referenced_by_owner.contains_key(&owner) {
                let ids = candidate
                    .package
                    .get_part(&owner)
                    .map(crate::document::xml_relationship_ids_in_order)
                    .transpose()?
                    .unwrap_or_default()
                    .into_iter()
                    .collect();
                referenced_by_owner.insert(owner.clone(), ids);
            }
            if referenced_by_owner
                .get(&owner)
                .is_some_and(|ids| ids.contains(&relationship.id))
            {
                if let Some(CarriedMedia {
                    part_name,
                    bytes,
                    content_type,
                }) = media
                {
                    carried_image = true;
                    candidate.package.set_part(&part_name, bytes);
                    candidate
                        .package
                        .content_types
                        .add_override(&part_name, &content_type);
                }
                if let Some(source_part) = chart {
                    crate::field::copy_fragment_part(
                        &edited.package,
                        &mut candidate,
                        &source_part,
                        &charts.part_map,
                        &mut charts.copied,
                    )?;
                }
                candidate
                    .package
                    .get_or_create_part_rels(&owner)
                    .items
                    .push(relationship);
            }
        }
        if carried_image {
            let mut bytes = std::io::Cursor::new(Vec::new());
            candidate.package.write_to(&mut bytes)?;
            candidate = Document::from_bytes(bytes.get_ref())?;
        }
        candidate = reopen_staged(candidate)?;
        #[cfg(test)]
        FAIL_AFTER_COMPARISON_STAGING.with(|fail| {
            if fail.replace(false) {
                return Err(Error::Other(
                    "injected staged comparison postcondition failure".to_owned(),
                ));
            }
            Ok(())
        })?;
        // The postconditions resolve only the revisions this comparison
        // wrote, so a kept revision stays in both resolved copies and must
        // match the input that holds it.
        let new_ids = metadata
            .ids
            .used
            .difference(&input_ids)
            .copied()
            .collect::<HashSet<_>>();
        if existing_revisions {
            refuse_changed_existing_revisions(
                &original,
                &candidate,
                &original_stories,
                options,
                &new_ids,
            )?;
        }
        let accepted_body = resolved_package(
            &candidate,
            |document| document.accept_revision_ids(&new_ids),
            &original_stories,
            options,
            &text_box_markers,
        )?;
        let mut edited_package =
            normalized_package(&edited, &edited_stories, options, &text_box_markers)?;
        let main_ignored = story_ignored(options, ComparisonStoryKind::Main);
        if main_ignored {
            edited_package.0 = accepted_body.0.clone();
        }
        if accepted_body != edited_package {
            return Err(comparison_postcondition_error(
                "acceptance",
                "edited",
                &accepted_body,
                &edited_package,
            ));
        }
        if snapshot_comments {
            let mut accepted = candidate.clone_for_staging();
            accepted.accept_revision_ids(&new_ids)?;
            if !comment_snapshots_match(&comment_snapshot(&accepted)?, &edited_comments)? {
                return Err(Error::Other(
                    "comparison acceptance does not reproduce edited comments".to_owned(),
                ));
            }
        }
        let rejected_package = resolved_package(
            &candidate,
            |document| document.reject_revision_ids(&new_ids),
            &original_stories,
            options,
            &text_box_markers,
        )?;
        let mut original_package =
            normalized_package(&original, &original_stories, options, &text_box_markers)?;
        if main_ignored {
            original_package.0 = rejected_package.0.clone();
        }
        if rejected_package != original_package {
            return Err(comparison_postcondition_error(
                "rejection",
                "original",
                &rejected_package,
                &original_package,
            ));
        }
        if snapshot_comments {
            let mut rejected = candidate.clone_for_staging();
            rejected.reject_revision_ids(&new_ids)?;
            if !comment_snapshots_match(&comment_snapshot(&rejected)?, &original_comments)? {
                return Err(Error::Other(
                    "comparison rejection does not reproduce original comments".to_owned(),
                ));
            }
        }

        self.commit_staged_mutation(candidate);
        Ok(diagnostics)
    }
}

fn story_ignored(options: &ComparisonOptions, kind: ComparisonStoryKind) -> bool {
    options.ignored_stories.contains(&kind)
        || (options.ignore_comments && kind == ComparisonStoryKind::Comment)
}

fn validate_comparison_options(options: &ComparisonOptions) -> Result<()> {
    let mut unique = HashSet::new();
    if options
        .ignored_stories
        .iter()
        .any(|kind| !unique.insert(*kind))
    {
        return Err(Error::Other(
            "comparison options contain a duplicate ignored story".to_owned(),
        ));
    }
    Ok(())
}

fn story_parts(document: &Document) -> Result<Vec<StoryPart>> {
    story_parts_with_options(document, &ComparisonOptions::default())
}

fn story_parts_with_options(
    document: &Document,
    options: &ComparisonOptions,
) -> Result<Vec<StoryPart>> {
    let relationships = document.package.get_part_rels(&document.doc_part_name);
    let mut stories = Vec::new();
    let mut seen = HashSet::new();
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
    for section in sections {
        for (kind, references) in [
            (ComparisonStoryKind::Header, section.header_refs.as_slice()),
            (ComparisonStoryKind::Footer, section.footer_refs.as_slice()),
        ] {
            if story_ignored(options, kind) {
                continue;
            }
            for reference in references {
                let relationships = relationships.ok_or_else(|| {
                    Error::Other(format!(
                        "{} reference {} has no relationship set",
                        kind.label(),
                        reference.rel_id
                    ))
                })?;
                let matching = relationships
                    .items
                    .iter()
                    .filter(|relationship| relationship.id == reference.rel_id)
                    .collect::<Vec<_>>();
                if matching.len() != 1 {
                    return Err(Error::Other(format!(
                        "{} reference {} has {} relationships",
                        kind.label(),
                        reference.rel_id,
                        matching.len()
                    )));
                }
                let relationship = matching[0];
                let expected_type = if kind == ComparisonStoryKind::Header {
                    rel_types::HEADER
                } else {
                    rel_types::FOOTER
                };
                if relationship.rel_type != expected_type
                    || !crate::document::relationship_is_internal(relationship)
                {
                    continue;
                }
                let part_name =
                    OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
                if document.package.get_part(&part_name).is_none() {
                    return Err(Error::Other(format!(
                        "{} relationship {} targets missing part {part_name}",
                        kind.label(),
                        reference.rel_id
                    )));
                }
                if seen.insert((kind, part_name.clone())) {
                    stories.push(StoryPart { kind, part_name });
                }
            }
        }
    }

    if let Some(relationships) = relationships {
        for (kind, relationship_type) in [
            (ComparisonStoryKind::Comment, rel_types::COMMENTS),
            (ComparisonStoryKind::Footnote, rel_types::FOOTNOTES),
            (ComparisonStoryKind::Endnote, rel_types::ENDNOTES),
        ] {
            if story_ignored(options, kind) {
                continue;
            }
            for relationship in &relationships.items {
                if relationship.rel_type != relationship_type {
                    continue;
                }
                if !crate::document::relationship_is_internal(relationship) {
                    continue;
                }
                let part_name =
                    OpcPackage::resolve_rel_target(&document.doc_part_name, &relationship.target);
                if document.package.get_part(&part_name).is_none() {
                    return Err(Error::Other(format!(
                        "{} relationship {} targets missing part {part_name}",
                        kind.label(),
                        relationship.id
                    )));
                }
                if !seen.insert((kind, part_name.clone())) {
                    return Err(Error::Other(format!(
                        "{} story part {part_name} is referenced more than once",
                        kind.label()
                    )));
                }
                stories.push(StoryPart { kind, part_name });
            }
        }
    }
    Ok(stories)
}

pub(crate) fn revision_story_parts(document: &Document) -> Result<Vec<(StoryKind, String)>> {
    story_parts(document).map(|stories| {
        stories
            .into_iter()
            .map(|story| {
                let kind = match story.kind {
                    ComparisonStoryKind::Main => StoryKind::Body,
                    ComparisonStoryKind::Header => StoryKind::Header,
                    ComparisonStoryKind::Footer => StoryKind::Footer,
                    ComparisonStoryKind::Comment => StoryKind::Comment,
                    ComparisonStoryKind::TextBox => StoryKind::TextBox,
                    ComparisonStoryKind::Footnote => StoryKind::Footnote,
                    ComparisonStoryKind::Endnote => StoryKind::Endnote,
                };
                (kind, story.part_name)
            })
            .collect()
    })
}

fn story_xml<'a>(document: &'a Document, story: &StoryPart) -> Result<&'a [u8]> {
    document.package.get_part(&story.part_name).ok_or_else(|| {
        Error::Other(format!(
            "missing {} story {}",
            story.kind.label(),
            story.part_name
        ))
    })
}

#[derive(Clone)]
struct CarriedMedia {
    part_name: String,
    bytes: Vec<u8>,
    content_type: String,
}

struct CarriedRelationship {
    owner: String,
    relationship: Relationship,
    media: Option<CarriedMedia>,
    chart: Option<String>,
}

const MAX_COMPARISON_CHART_PARTS: usize = 4096;
const MAX_COMPARISON_CHART_BYTES: usize = 256 * 1024 * 1024;

#[derive(Default)]
struct CarriedCharts {
    part_map: BTreeMap<String, String>,
    copied: HashSet<String>,
    bytes: usize,
}

/// Give equivalent edited image, chart and hyperlink relationships their original ids.
/// Changed image/chart payloads and new hyperlinks get distinct ids and are carried
/// into the redline when its tracked content references them.
fn remap_equivalent_story_relationships(
    original: &Document,
    edited: &mut Document,
    original_stories: &[StoryPart],
    edited_stories: &[StoryPart],
) -> Result<(Vec<CarriedRelationship>, CarriedCharts)> {
    let mut reserved_media = HashSet::new();
    let mut charts = CarriedCharts::default();
    let mut carried = remap_equivalent_owner_relationships(
        original,
        edited,
        &original.doc_part_name,
        &edited.doc_part_name.clone(),
        &mut reserved_media,
        &mut charts,
    )?;
    for (left, right) in original_stories.iter().zip(edited_stories) {
        carried.extend(remap_equivalent_owner_relationships(
            original,
            edited,
            &left.part_name,
            &right.part_name,
            &mut reserved_media,
            &mut charts,
        )?);
    }
    Ok((carried, charts))
}

fn remap_equivalent_owner_relationships(
    original: &Document,
    edited: &mut Document,
    original_owner: &str,
    edited_owner: &str,
    reserved_media: &mut HashSet<String>,
    charts: &mut CarriedCharts,
) -> Result<Vec<CarriedRelationship>> {
    let original_relationships = original
        .package
        .get_part_rels(original_owner)
        .cloned()
        .unwrap_or_default();
    let edited_relationships = edited
        .package
        .get_part_rels(edited_owner)
        .cloned()
        .unwrap_or_default();
    let mut used = HashSet::new();
    let mut remap = HashMap::new();
    let mut unmatched_links = Vec::new();
    let mut unmatched_images = Vec::new();
    let mut unmatched_charts = Vec::new();
    let mut original_charts = HashMap::new();
    for relationship in &original_relationships.items {
        if relationship.rel_type == rel_types::CHART {
            original_charts.insert(
                relationship.id.clone(),
                chart_relationship_signature(original, original_owner, relationship)?.0,
            );
        }
    }
    for right in &edited_relationships.items {
        if right.rel_type == rel_types::CHART {
            let (signature, parts) = chart_relationship_signature(edited, edited_owner, right)?;
            if let Some((index, left)) =
                original_relationships
                    .items
                    .iter()
                    .enumerate()
                    .find(|(index, left)| {
                        !used.contains(index) && original_charts.get(&left.id) == Some(&signature)
                    })
            {
                used.insert(index);
                if left.id != right.id {
                    remap.insert(right.id.clone(), left.id.clone());
                }
            } else {
                let source_part = OpcPackage::resolve_rel_target(edited_owner, &right.target)
                    .to_ascii_lowercase();
                reserve_comparison_chart_parts(original, edited, parts, reserved_media, charts)?;
                unmatched_charts.push((right.clone(), source_part));
            }
            continue;
        }
        if right.rel_type != rel_types::IMAGE && right.rel_type != rel_types::HYPERLINK {
            continue;
        }
        let right_payload = relationship_payload(edited, edited_owner, right);
        let Some((index, left)) =
            original_relationships
                .items
                .iter()
                .enumerate()
                .find(|(index, left)| {
                    !used.contains(index)
                        && left.rel_type == right.rel_type
                        && left.target_mode == right.target_mode
                        && relationship_payload(original, original_owner, left) == right_payload
                })
        else {
            if right.rel_type == rel_types::HYPERLINK {
                unmatched_links.push(right.clone());
            } else if crate::document::relationship_is_internal(right) {
                let target = OpcPackage::resolve_rel_target(edited_owner, &right.target);
                let bytes = edited.package.get_part(&target).ok_or_else(|| {
                    Error::Other(format!("comparison image target is missing: {target}"))
                })?;
                let content_type = edited
                    .package
                    .content_types
                    .content_type_for(&target)
                    .ok_or_else(|| {
                        Error::Other(format!("comparison image has no content type: {target}"))
                    })?;
                unmatched_images.push((
                    right.clone(),
                    target,
                    bytes.to_vec(),
                    content_type.to_owned(),
                ));
            }
            continue;
        };
        used.insert(index);
        if left.id != right.id {
            remap.insert(right.id.clone(), left.id.clone());
        }
    }
    let mut occupied = edited_relationships
        .items
        .iter()
        .chain(&original_relationships.items)
        .map(|relationship| relationship.id.clone())
        .chain(remap.values().cloned())
        .collect::<HashSet<_>>();
    // The redline allocates a carried hyperlink's id the way the original's
    // relationships do, which is the id staging keeps when it reopens it.
    let mut allocator = original_relationships.clone();
    let mut carried = Vec::with_capacity(unmatched_links.len() + unmatched_images.len());
    for mut link in unmatched_links {
        let fresh = allocator.add_external(&link.rel_type, &link.target);
        occupied.insert(fresh.clone());
        remap.insert(link.id.clone(), fresh.clone());
        link.id = fresh;
        carried.push(CarriedRelationship {
            owner: original_owner.to_owned(),
            relationship: link,
            media: None,
            chart: None,
        });
    }
    for (mut image, source_part, bytes, content_type) in unmatched_images {
        let existing = original_relationships
            .items
            .iter()
            .find(|relationship| {
                if relationship.rel_type != rel_types::IMAGE
                    || !crate::document::relationship_is_internal(relationship)
                {
                    return false;
                }
                let target = OpcPackage::resolve_rel_target(original_owner, &relationship.target);
                original.package.content_types.content_type_for(&target)
                    == Some(content_type.as_str())
                    && relationship_payload(original, original_owner, relationship).as_deref()
                        == Some(bytes.as_slice())
            })
            .map(|relationship| (relationship.target.clone(), None))
            .or_else(|| {
                carried.iter().find_map(|existing: &CarriedRelationship| {
                    let media = existing.media.as_ref()?;
                    (media.bytes == bytes && media.content_type == content_type)
                        .then(|| (existing.relationship.target.clone(), Some(media.clone())))
                })
            });
        if let Some((target, media)) = existing {
            let mut id_ordinal = 1;
            let fresh = loop {
                let candidate = format!("rdocxComparisonImage{id_ordinal}");
                if occupied.insert(candidate.clone()) {
                    break candidate;
                }
                id_ordinal += 1;
            };
            remap.insert(image.id.clone(), fresh.clone());
            image.id = fresh;
            image.target = target;
            carried.push(CarriedRelationship {
                owner: original_owner.to_owned(),
                relationship: image,
                media,
                chart: None,
            });
            continue;
        }
        let extension = source_part.rsplit('.').next().unwrap_or("bin");
        let owner_dir = original_owner.rsplit_once('/').map_or("", |(dir, _)| dir);
        let mut ordinal = 1;
        let part_name = loop {
            let part_name = format!("{owner_dir}/media/rdocxComparison{ordinal}.{extension}");
            if !original.package.contains_part(&part_name)
                && !edited.package.contains_part(&part_name)
                && reserved_media.insert(part_name.clone())
            {
                break part_name;
            }
            ordinal += 1;
        };
        let target = format!("media/rdocxComparison{ordinal}.{extension}");
        let mut id_ordinal = 1;
        let fresh = loop {
            let candidate = format!("rdocxComparisonImage{id_ordinal}");
            if occupied.insert(candidate.clone()) {
                break candidate;
            }
            id_ordinal += 1;
        };
        remap.insert(image.id.clone(), fresh.clone());
        image.id = fresh;
        image.target = target;
        carried.push(CarriedRelationship {
            owner: original_owner.to_owned(),
            relationship: image,
            media: Some(CarriedMedia {
                part_name,
                bytes,
                content_type,
            }),
            chart: None,
        });
    }
    for (mut chart, source_part) in unmatched_charts {
        let mut ordinal = 1;
        let fresh = loop {
            let id = format!("rdocxComparisonChart{ordinal}");
            if occupied.insert(id.clone()) {
                break id;
            }
            ordinal += 1;
        };
        remap.insert(chart.id.clone(), fresh.clone());
        chart.id = fresh;
        chart.target = charts.part_map[&source_part].clone();
        carried.push(CarriedRelationship {
            owner: original_owner.to_owned(),
            relationship: chart,
            media: None,
            chart: Some(source_part),
        });
    }
    if remap.is_empty() {
        return Ok(carried);
    }
    let mut collision_remap = HashMap::new();
    let remapped_sources = remap.keys().cloned().collect::<HashSet<_>>();
    for target in remap.values() {
        if remapped_sources.contains(target)
            || edited_relationships
                .items
                .iter()
                .all(|relationship| relationship.id != *target)
        {
            continue;
        }
        let mut ordinal = collision_remap.len() + 1;
        let replacement = loop {
            let candidate = format!("rdocxComparison{ordinal}");
            if occupied.insert(candidate.clone()) {
                break candidate;
            }
            ordinal += 1;
        };
        collision_remap.insert(target.clone(), replacement);
    }
    collision_remap.extend(remap);
    let source = edited
        .package
        .get_part(edited_owner)
        .ok_or_else(|| Error::Other(format!("missing comparison story {edited_owner}")))?;
    let updated = crate::document::remap_xml_relationship_ids(source, &collision_remap)?;
    if edited_owner == edited.doc_part_name {
        edited.document = CT_Document::from_xml(&updated)?;
    }
    edited.package.set_part(edited_owner, updated);
    if let Some(relationships) = edited.package.get_part_rels_mut(edited_owner) {
        for relationship in &mut relationships.items {
            if let Some(updated) = collision_remap.get(&relationship.id) {
                relationship.id.clone_from(updated);
            }
        }
        relationships.to_xml()?;
    }
    Ok(carried)
}

fn reserve_comparison_chart_parts(
    original: &Document,
    edited: &Document,
    parts: Vec<String>,
    reserved: &mut HashSet<String>,
    charts: &mut CarriedCharts,
) -> Result<()> {
    for source in parts {
        if charts.part_map.contains_key(&source) {
            continue;
        }
        if charts.part_map.len() >= MAX_COMPARISON_CHART_PARTS {
            return Err(Error::Other(
                "comparison aggregate chart resources exceed part limit".to_owned(),
            ));
        }
        let size = comparison_chart_part_size(edited, &source)?;
        charts.bytes = charts.bytes.checked_add(size).ok_or_else(|| {
            Error::Other("comparison aggregate chart resources overflow".to_owned())
        })?;
        if charts.bytes > MAX_COMPARISON_CHART_BYTES {
            return Err(Error::Other(
                "comparison aggregate chart resources exceed byte limit".to_owned(),
            ));
        }
        let (directory, filename) = source.rsplit_once('/').unwrap_or(("", &source));
        let mut ordinal = 1;
        let target = loop {
            let target = format!("{directory}/rdocxComparison{ordinal}-{filename}");
            if !original.package.contains_part(&target)
                && !edited.package.contains_part(&target)
                && reserved.insert(target.to_ascii_lowercase())
            {
                break target;
            }
            ordinal += 1;
        };
        charts.part_map.insert(source, target);
    }
    Ok(())
}

fn comparison_chart_part_size(document: &Document, name: &str) -> Result<usize> {
    let bytes = document
        .package
        .get_part(name)
        .ok_or_else(|| Error::Other(format!("comparison chart target is missing: {name}")))?;
    let relationship_bytes = document
        .package
        .get_part_rels(name)
        .map(|relationships| relationships.to_xml())
        .transpose()?
        .map_or(0, |xml| xml.len());
    bytes
        .len()
        .checked_add(relationship_bytes)
        .ok_or_else(|| Error::Other("comparison chart closure size overflow".to_owned()))
}

/// Hash a chart's complete owned graph independently of package part names.
/// Sorted relationship ids define stable node ordinals, including shared and cyclic edges.
fn chart_relationship_signature(
    document: &Document,
    owner: &str,
    relationship: &Relationship,
) -> Result<(String, Vec<String>)> {
    const MAX_CHART_DEPTH: usize = 128;
    if !crate::document::relationship_is_internal(relationship) {
        return Err(Error::Other(
            "comparison requires an internal chart relationship".to_owned(),
        ));
    }
    fn field(hash: &mut Sha256, bytes: &[u8]) {
        hash.update((bytes.len() as u64).to_le_bytes());
        hash.update(bytes);
    }
    let root = OpcPackage::resolve_rel_target(owner, &relationship.target).to_ascii_lowercase();
    let mut nodes = vec![root.clone()];
    let mut ordinals = HashMap::from([(root.clone(), 0usize)]);
    let mut pending = vec![(root, 0usize)];
    let mut visited = HashSet::new();
    let mut hash = Sha256::new();
    let mut total_bytes = 0usize;
    while let Some((name, depth)) = pending.pop() {
        if !visited.insert(name.clone()) {
            continue;
        }
        if depth > MAX_CHART_DEPTH {
            return Err(Error::Other(
                "comparison chart closure exceeds graph limit".to_owned(),
            ));
        }
        let content_type = document
            .package
            .content_types
            .content_type_for(&name)
            .ok_or_else(|| {
                Error::Other(format!("comparison chart part has no content type: {name}"))
            })?;
        if content_type.starts_with("application/vnd.openxmlformats-package.digital-signature") {
            return Err(Error::Other(
                "comparison cannot remap integrity-bound chart resources".to_owned(),
            ));
        }
        let bytes = document
            .package
            .get_part(&name)
            .ok_or_else(|| Error::Other(format!("comparison chart target is missing: {name}")))?;
        total_bytes = total_bytes
            .checked_add(comparison_chart_part_size(document, &name)?)
            .ok_or_else(|| Error::Other("comparison chart closure size overflow".to_owned()))?;
        if total_bytes > MAX_COMPARISON_CHART_BYTES {
            return Err(Error::Other(
                "comparison chart closure exceeds byte limit".to_owned(),
            ));
        }
        field(&mut hash, content_type.as_bytes());
        field(&mut hash, bytes);
        let mut edges = document
            .package
            .get_part_rels(&name)
            .map(|rels| rels.items.clone())
            .unwrap_or_default();
        edges.sort_by(|left, right| left.id.cmp(&right.id));
        field(&mut hash, &(edges.len() as u64).to_le_bytes());
        let mut ids = HashSet::new();
        let mut children = Vec::new();
        for edge in edges {
            if !ids.insert(edge.id.clone()) {
                return Err(Error::Other(format!(
                    "comparison chart relationship id is ambiguous: {}",
                    edge.id
                )));
            }
            field(&mut hash, edge.id.as_bytes());
            field(&mut hash, edge.rel_type.as_bytes());
            if crate::document::relationship_is_internal(&edge) {
                field(&mut hash, b"internal");
                let target =
                    OpcPackage::resolve_rel_target(&name, &edge.target).to_ascii_lowercase();
                let next = if let Some(ordinal) = ordinals.get(&target) {
                    *ordinal
                } else {
                    if nodes.len() >= MAX_COMPARISON_CHART_PARTS {
                        return Err(Error::Other(
                            "comparison chart closure exceeds graph limit".to_owned(),
                        ));
                    }
                    let ordinal = nodes.len();
                    ordinals.insert(target.clone(), ordinal);
                    nodes.push(target.clone());
                    ordinal
                };
                field(&mut hash, &(next as u64).to_le_bytes());
                children.push((target, depth + 1));
            } else {
                if edge.target_mode.as_deref() != Some("External") {
                    return Err(Error::Other(
                        "comparison chart relationship has an invalid target mode".to_owned(),
                    ));
                }
                field(&mut hash, b"external");
                field(&mut hash, edge.target.as_bytes());
            }
        }
        // Match the existing copier's sorted depth-first recursion, rather than
        // shortest-path depth. Shortcut edges must not bypass its stack bound.
        pending.extend(children.into_iter().rev());
    }
    Ok((
        hash.finalize()
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect(),
        nodes,
    ))
}

fn relationship_payload(
    document: &Document,
    owner: &str,
    relationship: &Relationship,
) -> Option<Vec<u8>> {
    if !crate::document::relationship_is_internal(relationship) {
        return Some(relationship.target.as_bytes().to_vec());
    }
    let target = OpcPackage::resolve_rel_target(owner, &relationship.target);
    document.package.get_part(&target).map(<[u8]>::to_vec)
}

fn contains_modeled_revisions(
    document: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
) -> Result<bool> {
    if !story_ignored(options, ComparisonStoryKind::Main)
        && modeled_revision_count_with_options(
            document
                .package
                .get_part(&document.doc_part_name)
                .ok_or_else(|| {
                    Error::Other(format!("missing main story {}", document.doc_part_name))
                })?,
            options,
        )? > 0
    {
        return Ok(true);
    }
    stories
        .iter()
        .map(|story| modeled_revision_count_with_options(story_xml(document, story)?, options))
        .try_fold(false, |found, count| count.map(|count| found || count > 0))
}

fn reject_cross_story_moves(
    original: &Document,
    edited: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
) -> Result<()> {
    let mut original_by_story = Vec::new();
    let mut edited_by_story = Vec::new();
    if !story_ignored(options, ComparisonStoryKind::Main) {
        original_by_story.push((
            "document".to_owned(),
            normalized_body_with_options(&original.document, options),
        ));
        edited_by_story.push((
            "document".to_owned(),
            normalized_body_with_options(&edited.document, options),
        ));
    }
    for story in stories {
        let identity = format!("{}:{}", story.kind.label(), story.part_name);
        let original_xml = story_xml(original, story)?;
        let edited_xml = story_xml(edited, story)?;
        let marker = if story_ignored(options, ComparisonStoryKind::TextBox) {
            let original_source = std::str::from_utf8(original_xml).map_err(utf8_error)?;
            let edited_source = std::str::from_utf8(edited_xml).map_err(utf8_error)?;
            let prefix = root_prefix(original_source, story.kind.root_local())?;
            Some(text_box_marker_local(
                original_source,
                edited_source,
                &prefix,
            )?)
        } else {
            None
        };
        original_by_story.push((
            identity.clone(),
            normalized_story_part(
                original,
                &story.part_name,
                original_xml,
                story.kind,
                options,
                marker.as_deref(),
                normalized_body_with_options,
            )?,
        ));
        edited_by_story.push((
            identity,
            normalized_story_part(
                edited,
                &story.part_name,
                edited_xml,
                story.kind,
                options,
                marker.as_deref(),
                normalized_body_with_options,
            )?,
        ));
    }

    let mut removed = HashMap::<String, Vec<String>>::new();
    let mut inserted = HashMap::<String, Vec<String>>::new();
    for ((identity, before), (_, after)) in original_by_story.iter().zip(&edited_by_story) {
        let before = signature_counts(before);
        let after = signature_counts(after);
        for (signature, count) in &before {
            let delta = count.saturating_sub(*after.get(signature).unwrap_or(&0));
            removed
                .entry(signature.clone())
                .or_default()
                .extend(std::iter::repeat_n(identity.clone(), delta));
        }
        for (signature, count) in &after {
            let delta = count.saturating_sub(*before.get(signature).unwrap_or(&0));
            inserted
                .entry(signature.clone())
                .or_default()
                .extend(std::iter::repeat_n(identity.clone(), delta));
        }
    }
    for (signature, sources) in removed {
        let Some(destinations) = inserted.get(&signature) else {
            continue;
        };
        if sources
            .iter()
            .any(|source| destinations.iter().any(|destination| source != destination))
        {
            return Err(Error::Other(
                "comparison cannot represent a move between Word stories".to_owned(),
            ));
        }
    }
    Ok(())
}

fn signature_counts(signatures: &[String]) -> HashMap<String, usize> {
    let mut counts = HashMap::new();
    for signature in signatures {
        *counts.entry(signature.clone()).or_default() += 1;
    }
    counts
}

fn compare_story_part(
    original: &[u8],
    edited: &[u8],
    story: &StoryPart,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<Vec<u8>> {
    let original = std::str::from_utf8(original).map_err(utf8_error)?;
    let edited = std::str::from_utf8(edited).map_err(utf8_error)?;
    let original_root = root_inner_range(original, story.kind.root_local())?;
    let edited_root = root_inner_range(edited, story.kind.root_local())?;
    let word_prefix = root_prefix(original, story.kind.root_local())?;
    let original_bindings = root_namespace_bindings(original, story.kind.root_local(), &[])?;
    let edited_bindings = root_namespace_bindings(edited, story.kind.root_local(), &[])?;
    let tracked = if let Some(owner_local) = story.kind.owner_local() {
        compare_owned_story(
            original,
            edited,
            original_root,
            edited_root,
            owner_local,
            story,
            &word_prefix,
            &original_bindings,
            &edited_bindings,
            metadata,
            diagnostics,
        )?
    } else {
        let tracked_inner = compare_story_inner(
            &original[original_root.clone()],
            &edited[edited_root],
            &format!("{}:{}", story.kind.label(), story.part_name),
            &word_prefix,
            &original_bindings,
            &edited_bindings,
            metadata,
            diagnostics,
        )?;
        let mut tracked = original.to_owned();
        tracked.replace_range(original_root, &tracked_inner);
        tracked
    };
    let mut tracked = tracked.into_bytes();
    declare_w14_on_part_root(&mut tracked)?;
    crate::revision::modeled_revision_count(&tracked)?;
    Ok(tracked)
}

#[allow(clippy::too_many_arguments)]
fn compare_owned_story(
    original: &str,
    edited: &str,
    original_root: Range<usize>,
    edited_root: Range<usize>,
    owner_local: &str,
    story: &StoryPart,
    word_prefix: &str,
    original_bindings: &[(String, String)],
    edited_bindings: &[(String, String)],
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_spans = direct_word_element_spans_prefix_aware(original, owner_local)?;
    let edited_spans = direct_word_element_spans_prefix_aware(edited, owner_local)?;
    if original_spans.len() != edited_spans.len() {
        return Err(Error::Other(format!(
            "{} story shell count changed in {}",
            story.kind.label(),
            story.part_name
        )));
    }
    let (original_skeleton, original_owners) = canonical_owned_story(original, owner_local)?;
    let (edited_skeleton, edited_owners) = canonical_owned_story(edited, owner_local)?;
    if original_owners.len() != original_spans.len() || edited_owners.len() != edited_spans.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate {} story owners in {}",
            story.kind.label(),
            story.part_name
        )));
    }
    if original_skeleton != edited_skeleton {
        return Err(Error::Other(format!(
            "{} story root shell changed in {}",
            story.kind.label(),
            story.part_name
        )));
    }
    let mut replacements = Vec::with_capacity(original_spans.len());
    for (index, (left, right)) in original_spans.iter().zip(&edited_spans).enumerate() {
        let left_xml = &original[left.clone()];
        let right_xml = &edited[right.clone()];
        if original_owners[index].first() != edited_owners[index].first() {
            return Err(Error::Other(format!(
                "{} owner shell changed at {}[{index}]",
                story.kind.label(),
                story.part_name
            )));
        }
        if matches!(
            story.kind,
            ComparisonStoryKind::Footnote | ComparisonStoryKind::Endnote
        ) && !normal_note_owner(left_xml)?
        {
            if original_owners[index] != edited_owners[index] {
                return Err(Error::Other(format!(
                    "{} separator shell changed at {}[{index}]",
                    story.kind.label(),
                    story.part_name
                )));
            }
            replacements.push(left_xml.to_owned());
            continue;
        }
        let left_inner = element_inner_range_any_prefix(left_xml, owner_local)?;
        let right_inner = element_inner_range_any_prefix(right_xml, owner_local)?;
        let left_bindings = root_namespace_bindings(left_xml, owner_local, original_bindings)?;
        let right_bindings = root_namespace_bindings(right_xml, owner_local, edited_bindings)?;
        let tracked_inner = compare_story_inner(
            &left_xml[left_inner.clone()],
            &right_xml[right_inner],
            &format!(
                "{}:{}/{}[{index}]",
                story.kind.label(),
                story.part_name,
                owner_local
            ),
            word_prefix,
            &left_bindings,
            &right_bindings,
            metadata,
            diagnostics,
        )?;
        let mut owner = left_xml.to_owned();
        owner.replace_range(left_inner, &tracked_inner);
        replacements.push(owner);
    }
    let mut tracked = original.to_owned();
    for (span, replacement) in original_spans.into_iter().zip(replacements).rev() {
        tracked.replace_range(span, &replacement);
    }
    if root_inner_range(&tracked, story.kind.root_local())?.start != original_root.start
        || edited_root.start == usize::MAX
    {
        return Err(Error::Other(
            "comparison story root moved unexpectedly".to_owned(),
        ));
    }
    Ok(tracked)
}

fn compare_story_inner(
    original: &str,
    edited: &str,
    location: &str,
    word_prefix: &str,
    original_bindings: &[(String, String)],
    edited_bindings: &[(String, String)],
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    compare_story_inner_impl(
        original,
        edited,
        location,
        word_prefix,
        original_bindings,
        edited_bindings,
        metadata,
        diagnostics,
        true,
    )
}

#[allow(clippy::too_many_arguments)]
fn compare_story_inner_impl(
    original: &str,
    edited: &str,
    location: &str,
    word_prefix: &str,
    original_bindings: &[(String, String)],
    edited_bindings: &[(String, String)],
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
    scan_text_boxes: bool,
) -> Result<String> {
    let original_text_boxes = if scan_text_boxes {
        text_box_spans(original, word_prefix)?
    } else {
        Vec::new()
    };
    let edited_text_boxes = if scan_text_boxes {
        text_box_spans(edited, word_prefix)?
    } else {
        Vec::new()
    };
    if !original_text_boxes.is_empty() || !edited_text_boxes.is_empty() {
        if text_box_host_skeletons(original, &original_text_boxes)?
            != text_box_host_skeletons(edited, &edited_text_boxes)?
        {
            return Err(Error::Other(format!(
                "comparison cannot change nested text-box host shells at {location}"
            )));
        }
        if story_ignored(metadata.options, ComparisonStoryKind::TextBox) {
            let marker = text_box_marker_local(original, edited, word_prefix)?;
            let (masked_original, preserved) =
                mask_text_box_subtrees(original, word_prefix, &marker)?;
            let (masked_edited, _) = mask_text_box_subtrees(edited, word_prefix, &marker)?;
            let mut tracked = compare_story_inner_impl(
                &masked_original,
                &masked_edited,
                location,
                word_prefix,
                original_bindings,
                edited_bindings,
                metadata,
                diagnostics,
                false,
            )?;
            restore_text_box_subtrees(&mut tracked, &preserved, word_prefix, &marker, location)?;
            return Ok(tracked);
        }
        if original_text_boxes.len() != edited_text_boxes.len() {
            return Err(Error::Other(format!(
                "comparison cannot change nested text-box hosts at {location}"
            )));
        }
        let text_box_coordinates = text_box_coordinates(original, &original_text_boxes)?;
        let mut tracked_boxes = Vec::with_capacity(original_text_boxes.len());
        for ((run_ordinal, box_ordinal), (left, right)) in text_box_coordinates
            .into_iter()
            .zip(original_text_boxes.iter().zip(&edited_text_boxes))
        {
            let left_box = &original[left.clone()];
            let right_box = &edited[right.clone()];
            let left_inner = element_inner_range_any_prefix(left_box, "txbxContent")?;
            let right_inner = element_inner_range_any_prefix(right_box, "txbxContent")?;
            let left_bindings =
                root_namespace_bindings(left_box, "txbxContent", original_bindings)?;
            let right_bindings =
                root_namespace_bindings(right_box, "txbxContent", edited_bindings)?;
            let tracked_inner = compare_story_inner(
                &left_box[left_inner.clone()],
                &right_box[right_inner],
                &format!("{location}/text-box[{run_ordinal}:{box_ordinal}]"),
                word_prefix,
                &left_bindings,
                &right_bindings,
                metadata,
                diagnostics,
            )?;
            let mut tracked_box = left_box.to_owned();
            tracked_box.replace_range(left_inner, &tracked_inner);
            tracked_boxes.push(tracked_box);
        }
        let tracked_hosts = replace_text_boxes_in_hosts(
            original,
            &original_text_boxes,
            &tracked_boxes,
            word_prefix,
        )?;
        let marker = text_box_marker_local(original, edited, word_prefix)?;
        let (masked_original, _) = mask_text_box_subtrees(original, word_prefix, &marker)?;
        let (masked_edited, _) = mask_text_box_subtrees(edited, word_prefix, &marker)?;
        let mut tracked = compare_story_inner_impl(
            &masked_original,
            &masked_edited,
            location,
            word_prefix,
            original_bindings,
            edited_bindings,
            metadata,
            diagnostics,
            false,
        )?;
        restore_text_box_subtrees(&mut tracked, &tracked_hosts, word_prefix, &marker, location)?;
        return Ok(tracked);
    }
    let original_spans = story_content_spans(original, word_prefix)?;
    let edited_spans = story_content_spans(edited, word_prefix)?;
    let original_document = story_document(original, original_bindings)?;
    let edited_document = story_document(edited, edited_bindings)?;
    if original_spans.len() != original_document.body.content.len()
        || edited_spans.len() != edited_document.body.content.len()
    {
        return Err(Error::Other(format!(
            "comparison could not correlate related-story owners at {location}"
        )));
    }
    let held_original = HeldRevisions::of(original, &original_spans, original_bindings, metadata)?;
    let held_edited = HeldRevisions::of(edited, &edited_spans, edited_bindings, metadata)?;
    compare_body(
        &original_document,
        &edited_document,
        location,
        Some((original, &original_spans)),
        [&held_original, &held_edited],
        metadata,
        diagnostics,
    )
    .map_err(|error| Error::Other(format!("comparison failed in {location}: {error}")))
}

fn story_content_spans(xml: &str, word_prefix: &str) -> Result<Vec<Range<usize>>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open_elements = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_section = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes()
        );
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let is_section = is_section && element.local_name().as_ref() == b"sectPr";
                open_elements.push((depth == 1 && !is_section).then_some(before));
                depth += 1;
            }
            Event::Empty(element)
                if depth == 1 && !(is_section && element.local_name().as_ref() == b"sectPr") =>
            {
                spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
            }
            Event::Empty(_) => {}
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open_elements.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn text_box_coordinates(xml: &str, boxes: &[Range<usize>]) -> Result<Vec<(usize, usize)>> {
    let mut coordinates = Vec::with_capacity(boxes.len());
    let mut previous_run = None;
    let mut run_ordinal = 0usize;
    let mut box_ordinal = 0usize;
    for text_box in boxes {
        let run = containing_run(xml, text_box.start)?;
        if previous_run.as_ref() != Some(&run) {
            run_ordinal += usize::from(previous_run.is_some());
            box_ordinal = 0;
            previous_run = Some(run);
        }
        coordinates.push((run_ordinal, box_ordinal));
        box_ordinal += 1;
    }
    Ok(coordinates)
}

fn text_box_marker_local(original: &str, edited: &str, word_prefix: &str) -> Result<String> {
    for ordinal in 0usize.. {
        let local = format!("textBoxHost{ordinal}");
        if !contains_private_text_box_marker(original, word_prefix, &local)?
            && !contains_private_text_box_marker(edited, word_prefix, &local)?
        {
            return Ok(local);
        }
    }
    unreachable!()
}

fn contains_private_text_box_marker(xml: &str, word_prefix: &str, local: &str) -> Result<bool> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let private = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == b"urn:rdocx:comparison:private"
        );
        match event {
            Event::Start(element) | Event::Empty(element)
                if private && element.local_name().as_ref() == local.as_bytes() =>
            {
                return Ok(true);
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buffer.clear();
    }
}

fn mask_text_box_subtrees(
    xml: &str,
    word_prefix: &str,
    marker_local: &str,
) -> Result<(String, Vec<String>)> {
    let boxes = text_box_spans(xml, word_prefix)?;
    let spans = text_box_host_spans(xml, &boxes)?;
    let preserved = spans
        .iter()
        .map(|span| xml[span.clone()].to_owned())
        .collect::<Vec<_>>();
    let mut masked = xml.to_owned();
    for (index, span) in spans.into_iter().enumerate().rev() {
        masked.replace_range(
            span,
            &format!(
                r#"<rdocxcmp:{marker_local} xmlns:rdocxcmp="urn:rdocx:comparison:private" rdocxcmp:index="{index}"/>"#
            ),
        );
    }
    Ok((masked, preserved))
}

fn restore_text_box_subtrees(
    xml: &mut String,
    preserved: &[String],
    word_prefix: &str,
    marker_local: &str,
    location: &str,
) -> Result<()> {
    let spans = private_text_box_host_spans(xml, word_prefix, marker_local)?;
    if spans.len() != preserved.len() {
        return Err(Error::Other(format!(
            "comparison lost an ignored text-box subtree at {location}"
        )));
    }
    for (index, span) in spans.into_iter().enumerate().rev() {
        xml.replace_range(span, &preserved[index]);
    }
    Ok(())
}

fn text_box_host_spans(xml: &str, boxes: &[Range<usize>]) -> Result<Vec<Range<usize>>> {
    let mut hosts = Vec::new();
    for text_box in boxes {
        let run = containing_run(xml, text_box.start)?;
        let relative = text_box.start - run.0;
        let children = direct_element_spans(&xml[run.0..run.1])?;
        let child = children
            .into_iter()
            .find(|child| child.start <= relative && relative < child.end)
            .ok_or_else(|| Error::Other("text-box subtree has no run-child host".to_owned()))?;
        let host = run.0 + child.start..run.0 + child.end;
        if hosts.last() != Some(&host) {
            hosts.push(host);
        }
    }
    Ok(hosts)
}

fn text_box_host_skeletons(xml: &str, boxes: &[Range<usize>]) -> Result<Vec<String>> {
    text_box_host_spans(xml, boxes)?
        .into_iter()
        .map(|host| {
            let host = &xml[host];
            let boxes = text_box_spans(host, "w")?;
            Ok(story_skeleton(host, &boxes))
        })
        .collect()
}

fn replace_text_boxes_in_hosts(
    xml: &str,
    boxes: &[Range<usize>],
    replacements: &[String],
    word_prefix: &str,
) -> Result<Vec<String>> {
    let mut replacement_index = 0usize;
    let mut hosts = Vec::new();
    for span in text_box_host_spans(xml, boxes)? {
        let mut host = xml[span].to_owned();
        let host_boxes = text_box_spans(&host, word_prefix)?;
        let count = host_boxes.len();
        for (box_span, replacement) in host_boxes
            .into_iter()
            .zip(&replacements[replacement_index..replacement_index + count])
            .rev()
        {
            host.replace_range(box_span, replacement);
        }
        replacement_index += count;
        hosts.push(host);
    }
    if replacement_index != replacements.len() {
        return Err(Error::Other(
            "comparison could not correlate nested text-box owners".to_owned(),
        ));
    }
    Ok(hosts)
}

fn direct_element_spans(xml: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(_) => {
                open.push((depth == 1).then_some(before));
                depth += 1;
            }
            Event::Empty(_) if depth == 1 => spans.push(before..after),
            Event::Empty(_) => {}
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open.pop() {
                    spans.push(start..after);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn private_text_box_host_spans(
    xml: &str,
    word_prefix: &str,
    marker_local: &str,
) -> Result<Vec<Range<usize>>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut stack = Vec::<Option<usize>>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_marker = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == b"urn:rdocx:comparison:private"
        );
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => stack.push(
                (is_marker && element.local_name().as_ref() == marker_local.as_bytes())
                    .then_some(before),
            ),
            Event::Empty(element)
                if is_marker && element.local_name().as_ref() == marker_local.as_bytes() =>
            {
                spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
            }
            Event::Empty(_) => {}
            Event::End(_) => {
                if let Some(Some(start)) = stack.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn without_text_box_subtrees(xml: &[u8]) -> Result<Vec<u8>> {
    let source = std::str::from_utf8(xml).map_err(utf8_error)?;
    let boxes = text_box_spans(source, "w")?;
    if boxes.is_empty() {
        return Ok(xml.to_vec());
    }
    let mut output = source.to_owned();
    for span in boxes.into_iter().rev() {
        output.replace_range(span, "");
    }
    Ok(output.into_bytes())
}

fn modeled_revision_count_with_options(xml: &[u8], options: &ComparisonOptions) -> Result<usize> {
    if story_ignored(options, ComparisonStoryKind::TextBox) {
        crate::revision::modeled_revision_count(&without_text_box_subtrees(xml)?)
    } else {
        crate::revision::modeled_revision_count(xml)
    }
}

fn word_ids_with_options(xml: &[u8], options: &ComparisonOptions) -> Result<HashSet<i32>> {
    if story_ignored(options, ComparisonStoryKind::TextBox) {
        word_ids(&without_text_box_subtrees(xml)?)
    } else {
        word_ids(xml)
    }
}

fn text_box_spans(xml: &str, word_prefix: &str) -> Result<Vec<Range<usize>>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut stack = Vec::<Option<usize>>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_text_box =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                stack.push(
                    (is_text_box && element.local_name().as_ref() == b"txbxContent")
                        .then_some(before),
                );
            }
            Event::Empty(element)
                if is_text_box && element.local_name().as_ref() == b"txbxContent" =>
            {
                spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
            }
            Event::End(_) => {
                if let Some(Some(start)) = stack.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn root_prefix(xml: &str, local: &str) -> Result<String> {
    let suffix = format!(":{local}");
    let suffix_at = xml
        .find(&suffix)
        .ok_or_else(|| Error::Other(format!("comparison {local} root has no prefix")))?;
    let name_start = xml[..suffix_at]
        .rfind('<')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("comparison {local} root has no start")))?;
    Ok(xml[name_start..suffix_at].to_owned())
}

fn element_inner_range_any_prefix(xml: &str, local: &str) -> Result<Range<usize>> {
    let open_end = xml
        .find('>')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("{local} source has no start tag")))?;
    let name_end = xml[..open_end]
        .find(local)
        .map(|at| at + local.len())
        .ok_or_else(|| Error::Other(format!("{local} source has no owner name")))?;
    let name_start = xml[..name_end]
        .rfind('<')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("{local} source has no owner start")))?;
    let close = format!("</{}>", &xml[name_start..name_end]);
    let close_start = xml
        .rfind(&close)
        .ok_or_else(|| Error::Other(format!("{local} source has no end tag")))?;
    Ok(open_end..close_start)
}

fn story_document(inner: &str, namespace_bindings: &[(String, String)]) -> Result<CT_Document> {
    let (xml, _) = story_wrapper(inner, namespace_bindings)?;
    let xml = close_drawing_namespaces(xml, false)?;
    CT_Document::from_xml(&xml).map_err(Into::into)
}

/// The owners of one side of a story that hold an existing modeled revision.
#[derive(Default)]
struct HeldRevisions {
    /// One flag per owner, in the order of the story's content spans.
    owners: Vec<bool>,
    /// Whether the story's own section properties hold one.
    section: bool,
}

impl HeldRevisions {
    /// Find the owners of story content `inner`, split by `spans`, that hold
    /// a modeled revision. A revision outside every owner is in the section
    /// properties, the only other modeled child of a story. All flags are
    /// clear when no input holds an existing revision.
    fn of(
        inner: &str,
        spans: &[Range<usize>],
        bindings: &[(String, String)],
        metadata: &Metadata<'_>,
    ) -> Result<Self> {
        if !metadata.existing_revisions {
            return Ok(Self::default());
        }
        let (xml, offset) = story_wrapper(inner, bindings)?;
        let mut held = Self {
            owners: vec![false; spans.len()],
            section: false,
        };
        for (range, _) in crate::revision::modeled_revision_spans(&xml)? {
            // The spans are in document order and do not overlap.
            let start = range.start.saturating_sub(offset);
            let index = spans.partition_point(|span| span.end <= start);
            match spans.get(index).filter(|span| span.contains(&start)) {
                Some(_) => held.owners[index] = true,
                None => held.section = true,
            }
        }
        Ok(held)
    }

    fn holds(&self, index: Option<usize>) -> bool {
        index.is_some_and(|index| self.owners.get(index).copied().unwrap_or(false))
    }
}

/// Wrap story content in a Word document and body that declare
/// `namespace_bindings`, and return the wrapper with the offset of `inner`.
fn story_wrapper(inner: &str, namespace_bindings: &[(String, String)]) -> Result<(Vec<u8>, usize)> {
    let mut writer = Writer::new(Vec::with_capacity(inner.len() + 256));
    let mut suffix = 0usize;
    let wrapper_prefix = loop {
        let candidate = if suffix == 0 {
            "rdocxcmp".to_owned()
        } else {
            format!("rdocxcmp{suffix}")
        };
        let binding_name = format!("xmlns:{candidate}");
        if namespace_bindings
            .iter()
            .all(|(name, _)| name != &binding_name)
        {
            break candidate;
        }
        suffix += 1;
    };
    let document_name = format!("{wrapper_prefix}:document");
    let body_name = format!("{wrapper_prefix}:body");
    let binding_name = format!("xmlns:{wrapper_prefix}");
    let mut document = BytesStart::new(document_name.as_str());
    document.push_attribute((binding_name.as_str(), W_NS));
    for (name, value) in namespace_bindings {
        document.push_attribute((name.as_str(), value.as_str()));
    }
    writer.write_event(Event::Start(document))?;
    writer.write_event(Event::Start(BytesStart::new(body_name.as_str())))?;
    let offset = writer.get_ref().len();
    writer.get_mut().extend_from_slice(inner.as_bytes());
    writer.write_event(Event::End(BytesEnd::new(body_name.as_str())))?;
    writer.write_event(Event::End(BytesEnd::new(document_name.as_str())))?;
    Ok((writer.into_inner(), offset))
}

fn comparison_input(document: &Document) -> Result<Document> {
    let mut candidate = document.clone_for_staging();
    if let Some(source) = candidate
        .package
        .get_part(&candidate.doc_part_name)
        .map(<[u8]>::to_vec)
    {
        let parsed_source = CT_Document::from_xml(&source)?;
        let scoped_source = close_drawing_namespaces(source, true)?;
        let scoped_document = CT_Document::from_xml(&scoped_source)?;
        if parsed_source == candidate.document {
            candidate.document = scoped_document;
            candidate
                .package
                .set_part(&candidate.doc_part_name, scoped_source);
        } else {
            retain_matching_drawing_namespaces(&mut candidate.document, &scoped_document);
        }
    }
    candidate.prepare_staged_package()?;
    let source = candidate
        .package
        .get_part(&candidate.doc_part_name)
        .ok_or_else(|| Error::Other(format!("missing main story {}", candidate.doc_part_name)))?;
    let source = close_drawing_namespaces(source.to_vec(), true)?;
    candidate.document = CT_Document::from_xml(&source)?;
    candidate.package.set_part(&candidate.doc_part_name, source);
    Ok(candidate)
}

fn retain_matching_drawing_namespaces(target: &mut CT_Document, scoped_source: &CT_Document) {
    let mut source_drawings = Vec::new();
    crate::document::visit_all_drawings(&scoped_source.body.content, &mut |drawing| {
        source_drawings.push((drawing_signature(drawing), drawing.clone()));
    });
    crate::document::visit_all_drawings_mut(&mut target.body.content, &mut |drawing| {
        let signature = drawing_signature(drawing);
        let Some(index) = source_drawings
            .iter()
            .position(|(candidate, _)| candidate == &signature)
        else {
            return;
        };
        let (_, source) = source_drawings.remove(index);
        if let (Some(target), Some(source)) = (&mut drawing.inline, source.inline) {
            target.raw_xml = source.raw_xml;
        } else if let (Some(target), Some(source)) = (&mut drawing.anchor, source.anchor) {
            target.raw_xml = source.raw_xml;
        }
    });
}

fn drawing_signature(drawing: &rdocx_oxml::drawing::CT_Drawing) -> rdocx_oxml::drawing::CT_Drawing {
    let mut signature = drawing.clone();
    if let Some(inline) = &mut signature.inline {
        inline.raw_xml = None;
    }
    if let Some(anchor) = &mut signature.anchor {
        anchor.raw_xml = None;
    }
    signature
}

fn close_drawing_namespaces(
    xml: Vec<u8>,
    preserve_root_namespace_ownership: bool,
) -> Result<Vec<u8>> {
    let root_scope = preserve_root_namespace_ownership
        .then(|| comparison_root_namespace_scope(&xml))
        .transpose()?;
    let mut reader = NsReader::from_reader(xml.as_slice());
    reader.config_mut().trim_text(false);
    let mut edits = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison drawing scan failed: {error}")))?
        {
            Event::Start(element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                if matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
                    && matches!(element.local_name().as_ref(), b"inline" | b"anchor")
                {
                    let name = element.name().as_ref().to_vec();
                    reader
                        .read_to_end_into(quick_xml::name::QName(&name), &mut Vec::new())
                        .map_err(|error| {
                            Error::Other(format!("comparison drawing capture failed: {error}"))
                        })?;
                    let end = reader.buffer_position() as usize;
                    let mut scope = crate::document::story_namespace_scope_at(&xml, start)?;
                    if let Some(root_scope) = &root_scope {
                        scope.retain(|prefix, namespace| root_scope.get(prefix) != Some(namespace));
                    }
                    let scope = required_external_bindings(&xml[start..end], &scope)?;
                    let replacement = crate::document::close_content_fragment_namespaces(
                        &xml[start..end],
                        &scope,
                    )?;
                    if replacement != xml[start..end] {
                        edits.push((start, end, replacement));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let mut output = xml;
    for (start, end, replacement) in edits.into_iter().rev() {
        output.splice(start..end, replacement);
    }
    Ok(output)
}

fn comparison_root_namespace_scope(
    xml: &[u8],
) -> Result<std::collections::BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("comparison root namespace scan failed: {error}"))
        })? {
            Event::Start(_) | Event::Empty(_) => {
                return crate::document::story_namespace_scope_at(xml, start);
            }
            Event::Eof => {
                return Err(Error::Other(
                    "comparison namespace source has no root element".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn required_external_bindings(
    xml: &[u8],
    external: &std::collections::BTreeMap<String, String>,
) -> Result<std::collections::BTreeMap<String, String>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut scopes = Vec::<HashSet<String>>::new();
    let mut required = std::collections::BTreeMap::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("comparison drawing binding scan failed: {error}"))
        })? {
            Event::Start(element) => {
                let declarations = declared_prefixes(&element)?;
                record_required_bindings(&element, &scopes, &declarations, external, &mut required);
                scopes.push(declarations);
            }
            Event::Empty(element) => {
                let declarations = declared_prefixes(&element)?;
                record_required_bindings(&element, &scopes, &declarations, external, &mut required);
            }
            Event::End(_) => {
                scopes.pop();
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(required)
}

fn declared_prefixes(element: &BytesStart<'_>) -> Result<HashSet<String>> {
    let mut declarations = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Other(format!("comparison drawing namespace failed: {error}"))
        })?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" {
            declarations.insert(String::new());
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            declarations.insert(std::str::from_utf8(prefix).map_err(utf8_error)?.to_owned());
        }
    }
    Ok(declarations)
}

fn record_required_bindings(
    element: &BytesStart<'_>,
    scopes: &[HashSet<String>],
    declarations: &HashSet<String>,
    external: &std::collections::BTreeMap<String, String>,
    required: &mut std::collections::BTreeMap<String, String>,
) {
    record_required_prefix(
        qualified_prefix(element.name().as_ref()).unwrap_or(""),
        scopes,
        declarations,
        external,
        required,
    );
    for attribute in element.attributes().filter_map(|attribute| attribute.ok()) {
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if let Some(prefix) = qualified_prefix(name) {
            record_required_prefix(prefix, scopes, declarations, external, required);
        }
    }
}

fn record_required_prefix(
    prefix: &str,
    scopes: &[HashSet<String>],
    declarations: &HashSet<String>,
    external: &std::collections::BTreeMap<String, String>,
    required: &mut std::collections::BTreeMap<String, String>,
) {
    let internally_bound =
        declarations.contains(prefix) || scopes.iter().rev().any(|scope| scope.contains(prefix));
    if !internally_bound && let Some(namespace) = external.get(prefix) {
        required
            .entry(prefix.to_owned())
            .or_insert_with(|| namespace.clone());
    }
}

fn qualified_prefix(name: &[u8]) -> Option<&str> {
    let separator = name.iter().position(|byte| *byte == b':')?;
    std::str::from_utf8(&name[..separator]).ok()
}

fn root_namespace_bindings(
    xml: &str,
    expected_local: &str,
    inherited: &[(String, String)],
) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison namespace parse failed: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != expected_local.as_bytes() {
                    return Err(Error::Other(format!(
                        "comparison namespace scope expected {expected_local} root"
                    )));
                }
                let mut bindings = inherited.to_vec();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Other(format!("comparison namespace attribute failed: {error}"))
                    })?;
                    let key = attribute.key.as_ref();
                    if key != b"xmlns" && !key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let name = std::str::from_utf8(key).map_err(utf8_error)?.to_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                        .map_err(|error| {
                            Error::Other(format!(
                                "comparison namespace value decode failed: {error}"
                            ))
                        })?
                        .into_owned();
                    bindings.retain(|(existing, _)| existing != &name);
                    bindings.push((name, value));
                }
                return Ok(bindings);
            }
            Event::Eof => {
                return Err(Error::Other(format!(
                    "comparison {expected_local} namespace scope is missing"
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn root_inner_range(xml: &str, expected_local: &str) -> Result<Range<usize>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if element.local_name().as_ref() != expected_local.as_bytes() || !is_word {
                        return Err(Error::Other(format!(
                            "comparison expected a Word {expected_local} root"
                        )));
                    }
                    start = Some(after);
                }
                depth += 1;
            }
            Event::Empty(element) if depth == 0 => {
                if element.local_name().as_ref() == expected_local.as_bytes() && is_word {
                    return Ok(after..after);
                }
                return Err(Error::Other(format!(
                    "comparison expected a Word {expected_local} root"
                )));
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok(start.unwrap_or(before)..before);
                }
            }
            Event::Eof => {
                return Err(Error::Other(format!(
                    "comparison XML has no closed {expected_local} root"
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn direct_word_element_spans_prefix_aware(xml: &str, local: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let target =
                    depth == 1 && element.local_name().as_ref() == local.as_bytes() && is_word;
                open.push(target.then_some(before));
                depth += 1;
            }
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == local.as_bytes() && is_word {
                    spans.push(before..after);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open.pop() {
                    spans.push(start..after);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn story_skeleton(xml: &str, spans: &[Range<usize>]) -> String {
    let mut skeleton = xml.to_owned();
    for span in spans.iter().rev() {
        skeleton.replace_range(span.clone(), "<owner/>");
    }
    skeleton
}

/// An owned story read so that two serializations of one tree compare equal:
/// the part with each owner as one placeholder, then each owner, whose first
/// token is its start tag.
///
/// Names resolve to their namespaces, attributes compare as a sorted set and
/// an empty element reads as a start and an end. The XML declaration,
/// comments, processing instructions, whitespace-only text, namespace
/// declarations and Markup Compatibility attributes are left out. They say
/// how the part is written, not what it holds, and the redline keeps the
/// original bytes.
fn canonical_owned_story(xml: &str, owner_local: &str) -> Result<(Vec<String>, Vec<Vec<String>>)> {
    let scan_error = |error: &dyn std::fmt::Display| {
        Error::Other(format!("comparison story shell scan failed: {error}"))
    };
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut skeleton = Vec::new();
    let mut owners = Vec::<Vec<String>>::new();
    let mut in_owner = false;
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| scan_error(&error))?;
        let tokens = match (in_owner, owners.last_mut()) {
            (true, Some(owner)) => owner,
            _ => &mut skeleton,
        };
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                tokens.push("end".to_owned());
                depth = depth.saturating_sub(1);
                in_owner &= depth > 1;
                buffer.clear();
                continue;
            }
            Event::Text(text) if !text.iter().all(u8::is_ascii_whitespace) => {
                tokens.push(format!("text {:?}", String::from_utf8_lossy(&text)));
                buffer.clear();
                continue;
            }
            Event::CData(text) => {
                tokens.push(format!("text {:?}", String::from_utf8_lossy(&text)));
                buffer.clear();
                continue;
            }
            Event::GeneralRef(reference) => {
                tokens.push(format!(
                    "reference {:?}",
                    String::from_utf8_lossy(&reference)
                ));
                buffer.clear();
                continue;
            }
            Event::Eof => break,
            _ => {
                buffer.clear();
                continue;
            }
        };
        let resolver = reader.resolver();
        let (namespace, local) = resolver.resolve_element(element.name());
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| scan_error(&error))?;
            let key = attribute.key.as_ref();
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                continue;
            }
            let (attribute_namespace, attribute_local) = resolver.resolve_attribute(attribute.key);
            if matches!(
                attribute_namespace,
                ResolveResult::Bound(Namespace(uri)) if uri == oxml_core::xml::MC_NS.as_bytes()
            ) {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .map_err(|error| scan_error(&error))?;
            attributes.push(format!(
                "{attribute_namespace:?} {:?} {value:?}",
                String::from_utf8_lossy(attribute_local.as_ref())
            ));
        }
        attributes.sort();
        let start = format!(
            "start {namespace:?} {:?} {attributes:?}",
            String::from_utf8_lossy(local.as_ref())
        );
        let is_owner = depth == 1
            && local.as_ref() == owner_local.as_bytes()
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let tokens = if is_owner {
            skeleton.push("owner".to_owned());
            in_owner = !empty;
            owners.push(Vec::new());
            owners.last_mut().expect("owner was just pushed")
        } else {
            tokens
        };
        tokens.push(start);
        if empty {
            tokens.push("end".to_owned());
        } else {
            depth += 1;
        }
        buffer.clear();
    }
    Ok((skeleton, owners))
}

fn normal_note_owner(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let (_, event) = reader
        .read_resolved_event_into(&mut buffer)
        .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
    let (Event::Start(element) | Event::Empty(element)) = event else {
        return Err(Error::Other(
            "comparison note has no owner element".to_owned(),
        ));
    };
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == b"type"
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
        {
            return Ok(attribute.value.as_ref().is_empty());
        }
    }
    Ok(true)
}

pub(crate) fn reopen_staged(candidate: Document) -> Result<Document> {
    candidate.prepare_and_reopen_staged()
}

type NormalizedPackage = (Vec<String>, Vec<(ComparisonStoryKind, String, Vec<String>)>);

fn comparison_postcondition_error(
    outcome: &str,
    expected: &str,
    actual: &NormalizedPackage,
    wanted: &NormalizedPackage,
) -> Error {
    Error::Other(format!(
        "comparison {outcome} does not reproduce the {expected} stories at {}",
        first_normalized_mismatch(actual, wanted)
    ))
}

fn first_normalized_mismatch(left: &NormalizedPackage, right: &NormalizedPackage) -> String {
    let main_count = left.0.len().min(right.0.len());
    for index in 0..main_count {
        if left.0[index] != right.0[index] {
            return format!("body story item[{index}]");
        }
    }
    if left.0.len() != right.0.len() {
        return format!("body story item[{main_count}]");
    }
    let story_count = left.1.len().min(right.1.len());
    for story_index in 0..story_count {
        let (left_kind, left_part, left_items) = &left.1[story_index];
        let (right_kind, right_part, right_items) = &right.1[story_index];
        if left_kind != right_kind || left_part != right_part {
            return format!("story inventory[{story_index}]");
        }
        let item_count = left_items.len().min(right_items.len());
        for item_index in 0..item_count {
            if left_items[item_index] != right_items[item_index] {
                return format!(
                    "{} story {} item[{item_index}]",
                    left_kind.label(),
                    left_part
                );
            }
        }
        if left_items.len() != right_items.len() {
            return format!(
                "{} story {} item[{item_count}]",
                left_kind.label(),
                left_part
            );
        }
    }
    format!("story inventory[{story_count}]")
}

fn comparison_text_box_markers(
    original: &Document,
    edited: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
) -> Result<TextBoxMarkers> {
    if !story_ignored(options, ComparisonStoryKind::TextBox) {
        return Ok(TextBoxMarkers::default());
    }
    let original_main = original
        .package
        .get_part(&original.doc_part_name)
        .ok_or_else(|| Error::Other(format!("missing main story {}", original.doc_part_name)))?;
    let edited_main = edited
        .package
        .get_part(&edited.doc_part_name)
        .ok_or_else(|| Error::Other(format!("missing main story {}", edited.doc_part_name)))?;
    let original_main = std::str::from_utf8(original_main).map_err(utf8_error)?;
    let edited_main = std::str::from_utf8(edited_main).map_err(utf8_error)?;
    let mut markers = TextBoxMarkers {
        main: text_box_marker_local(original_main, edited_main, "w")?,
        related: HashMap::new(),
    };
    for story in stories {
        let original_xml = std::str::from_utf8(story_xml(original, story)?).map_err(utf8_error)?;
        let edited_xml = std::str::from_utf8(story_xml(edited, story)?).map_err(utf8_error)?;
        let prefix = root_prefix(original_xml, story.kind.root_local())?;
        markers.related.insert(
            story.part_name.clone(),
            text_box_marker_local(original_xml, edited_xml, &prefix)?,
        );
    }
    Ok(markers)
}

/// Refuse a redline that does not keep each existing revision of the
/// original byte for byte, with its author, date and id, in every compared
/// story. The revisions this comparison wrote carry `new_ids`, which no id of
/// a compared story uses. Ignored stories and text boxes are not seeded and
/// keep their original bytes, so they are left out here.
fn refuse_changed_existing_revisions(
    original: &Document,
    candidate: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
    new_ids: &HashSet<i32>,
) -> Result<()> {
    let mut parts = Vec::new();
    if !story_ignored(options, ComparisonStoryKind::Main) {
        parts.push(original.doc_part_name.clone());
    }
    parts.extend(stories.iter().map(|story| story.part_name.clone()));
    for part in parts {
        let revisions = |document: &Document| -> Result<Vec<Vec<u8>>> {
            let xml = document
                .package
                .get_part(&part)
                .ok_or_else(|| Error::Other(format!("missing comparison story part {part}")))?;
            let xml = if story_ignored(options, ComparisonStoryKind::TextBox) {
                without_text_box_subtrees(xml)?
            } else {
                xml.to_vec()
            };
            Ok(crate::revision::modeled_revision_spans(&xml)?
                .into_iter()
                .filter(|(_, id)| !new_ids.contains(id))
                .map(|(range, _)| xml[range].to_vec())
                .collect())
        };
        if revisions(original)? != revisions(candidate)? {
            return Err(Error::Other(format!(
                "comparison would change an existing tracked revision in {part}"
            )));
        }
    }
    Ok(())
}

fn resolved_package(
    candidate: &Document,
    resolve: impl FnOnce(&mut Document) -> Result<usize>,
    stories: &[StoryPart],
    options: &ComparisonOptions,
    text_box_markers: &TextBoxMarkers,
) -> Result<NormalizedPackage> {
    let mut resolved = candidate.clone_for_staging();
    resolve(&mut resolved)?;
    normalized_package(&resolved, stories, options, text_box_markers)
}

fn normalized_package(
    document: &Document,
    stories: &[StoryPart],
    options: &ComparisonOptions,
    text_box_markers: &TextBoxMarkers,
) -> Result<NormalizedPackage> {
    let mut related = Vec::with_capacity(stories.len());
    for story in stories {
        related.push((
            story.kind,
            story.part_name.clone(),
            normalized_story_part(
                document,
                &story.part_name,
                story_xml(document, story)?,
                story.kind,
                options,
                text_box_markers
                    .related
                    .get(&story.part_name)
                    .map(String::as_str),
                postcondition_body,
            )?,
        ));
    }
    let source = document
        .package
        .get_part(&document.doc_part_name)
        .ok_or_else(|| Error::Other(format!("missing main story {}", document.doc_part_name)))?;
    let source = std::str::from_utf8(source).map_err(utf8_error)?;
    let main_source;
    let source = if story_ignored(options, ComparisonStoryKind::TextBox) {
        let (masked, _) = mask_text_box_subtrees(source, "w", &text_box_markers.main)?;
        main_source = masked;
        main_source.as_str()
    } else {
        source
    };
    let bindings = root_namespace_bindings(source, "document", &[])?;
    let mut main_document = story_document(extract_body_inner(source.as_bytes())?, &bindings)?;
    normalize_drawing_relationships(&mut main_document, document, &document.doc_part_name)?;
    let main = postcondition_body(&main_document, options);
    Ok((main, related))
}

fn normalized_story_part(
    document: &Document,
    physical_owner: &str,
    xml: &[u8],
    kind: ComparisonStoryKind,
    options: &ComparisonOptions,
    text_box_marker: Option<&str>,
    story_projection: fn(&CT_Document, &ComparisonOptions) -> Vec<String>,
) -> Result<Vec<String>> {
    let masked;
    let xml = if story_ignored(options, ComparisonStoryKind::TextBox) {
        let source = std::str::from_utf8(xml).map_err(utf8_error)?;
        let prefix = root_prefix(source, kind.root_local())?;
        let marker = text_box_marker
            .ok_or_else(|| Error::Other("missing shared comparison text-box marker".to_owned()))?;
        masked = mask_text_box_subtrees(source, &prefix, marker)?.0;
        &masked
    } else {
        std::str::from_utf8(xml).map_err(utf8_error)?
    };
    let root = root_inner_range(xml, kind.root_local())?;
    if let Some(owner_local) = kind.owner_local() {
        let mut normalized = Vec::new();
        for owner in direct_word_element_spans_prefix_aware(xml, owner_local)? {
            let owner_xml = &xml[owner];
            if matches!(
                kind,
                ComparisonStoryKind::Footnote | ComparisonStoryKind::Endnote
            ) && !normal_note_owner(owner_xml)?
            {
                continue;
            }
            let inner = element_inner_range_any_prefix(owner_xml, owner_local)?;
            let root_bindings = root_namespace_bindings(xml, kind.root_local(), &[])?;
            let owner_bindings = root_namespace_bindings(owner_xml, owner_local, &root_bindings)?;
            let mut model = story_document(&owner_xml[inner], &owner_bindings)?;
            normalize_drawing_relationships(&mut model, document, physical_owner)?;
            normalized.extend(story_projection(&model, options));
        }
        Ok(normalized)
    } else {
        let bindings = root_namespace_bindings(xml, kind.root_local(), &[])?;
        let mut model = story_document(&xml[root], &bindings)?;
        normalize_drawing_relationships(&mut model, document, physical_owner)?;
        Ok(story_projection(&model, options))
    }
}

fn normalize_drawing_relationships(
    model: &mut CT_Document,
    document: &Document,
    owner: &str,
) -> Result<()> {
    let mut result = Ok(());
    crate::document::visit_all_drawings_mut(&mut model.body.content, &mut |drawing| {
        if result.is_err() {
            return;
        }
        result = (|| {
            if let Some(inline) = &mut drawing.inline {
                inline.doc_pr_id = 0;
                normalize_relationship_id(&mut inline.embed_id, document, owner)?;
                if let Some(id) = &mut inline.chart_rel_id {
                    normalize_relationship_id(id, document, owner)?;
                }
            }
            if let Some(anchor) = &mut drawing.anchor {
                anchor.doc_pr_id = 0;
                normalize_relationship_id(&mut anchor.embed_id, document, owner)?;
                if let Some(id) = &mut anchor.chart_rel_id {
                    normalize_relationship_id(id, document, owner)?;
                }
            }
            Ok(())
        })();
    });
    result
}

fn normalize_relationship_id(id: &mut String, document: &Document, owner: &str) -> Result<()> {
    let Some(relationship) = document
        .package
        .get_part_rels(owner)
        .and_then(|relationships| relationships.get_by_id(id))
    else {
        return Ok(());
    };
    let signature = if relationship.rel_type == rel_types::CHART {
        format!(
            "internal:{}:{}",
            rel_types::CHART,
            chart_relationship_signature(document, owner, relationship)?.0
        )
    } else if crate::document::relationship_is_internal(relationship) {
        let target = OpcPackage::resolve_rel_target(owner, &relationship.target);
        let Some(bytes) = document.package.get_part(&target) else {
            return Ok(());
        };
        let digest = Sha256::digest(bytes);
        let mut hex = String::with_capacity(digest.len() * 2);
        for byte in digest {
            use std::fmt::Write as _;
            let _ = write!(hex, "{byte:02x}");
        }
        format!("internal:{}:{hex}", relationship.rel_type)
    } else {
        format!(
            "external:{}:{}:{}",
            relationship.rel_type,
            relationship.target_mode.as_deref().unwrap_or(""),
            relationship.target
        )
    };
    *id = signature;
    Ok(())
}

/// Compare the owners of one story and return the tracked story content.
///
/// `held` names the owners of each side that hold an existing modeled
/// revision. Such an owner is kept only when both sides hold it unchanged,
/// and then its original bytes are emitted. Any other change that would
/// revise, delete, insert or move it, or mark its paragraph mark, is refused.
fn compare_body(
    original: &CT_Document,
    edited: &CT_Document,
    location: &str,
    original_source: Option<(&str, &[Range<usize>])>,
    held: [&HeldRevisions; 2],
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let [held_original, held_edited] = held;
    let original_signatures = original
        .body
        .content
        .iter()
        .map(|content| body_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .body
        .content
        .iter()
        .map(|content| body_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    fn paragraphs(content: &[BodyContent]) -> Vec<Option<&CT_P>> {
        content
            .iter()
            .map(|content| match content {
                BodyContent::Paragraph(paragraph) => Some(paragraph.as_ref()),
                _ => None,
            })
            .collect()
    }
    let original_paragraphs = paragraphs(&original.body.content);
    let edited_paragraphs = paragraphs(&edited.body.content);
    let (aligned, carried) = replace_changed_paragraph_runs(
        pair_gaps_by_kind(
            align(&original_signatures, &edited_signatures),
            &original_signatures,
            &edited_signatures,
        ),
        &original_paragraphs,
        &edited_paragraphs,
        &original_signatures,
        &edited_signatures,
        metadata.options,
    )?;
    refuse_uncarried_field_owners(
        &aligned,
        &original_paragraphs,
        &edited_paragraphs,
        &carried,
        location,
    )?;
    let moves = pair_moves(
        &aligned,
        &original_signatures,
        &edited_signatures,
        &mut metadata.ids,
    )?;
    let trailing_paragraph_insert_start = aligned
        .iter()
        .enumerate()
        .rev()
        .take_while(|(_, (left, right))| {
            left.is_none()
                && right
                    .and_then(|index| edited.body.content.get(index))
                    .is_some_and(|content| matches!(content, BodyContent::Paragraph(_)))
        })
        .map(|(position, _)| position)
        .last();
    let trailing_paragraph_delete_start = aligned
        .iter()
        .enumerate()
        .rev()
        .take_while(|(_, (left, right))| {
            right.is_none()
                && left.is_some_and(|index| {
                    moves.original[index].is_none()
                        && matches!(original.body.content[index], BodyContent::Paragraph(_))
                })
        })
        .map(|(position, _)| position)
        .last();
    // A final paragraph change marks the paragraph mark of an owner already
    // emitted, which must not be one that holds an existing revision.
    let refuse_held_mark = |position: Option<usize>| -> Result<()> {
        let Some((left, right)) = position.map(|position| aligned[position]) else {
            return Ok(());
        };
        if held_original.holds(left) || held_edited.holds(right) {
            return Err(Error::Other(format!(
                "comparison cannot mark the paragraph mark of {location}/paragraph[{}] for the final paragraph change after it, because that paragraph holds an existing tracked revision",
                left.or(right).unwrap_or_default()
            )));
        }
        Ok(())
    };
    let mut output: Vec<(EmittedOwner, String)> = Vec::new();
    let mut trailing_merged = None;
    for (position, (original_index, edited_index)) in aligned.iter().copied().enumerate() {
        let item = |(left, right): (Option<usize>, Option<usize>)| {
            right
                .and_then(|index| edited.body.content.get(index))
                .or_else(|| left.and_then(|index| original.body.content.get(index)))
        };
        let next_takes_mark = next_kept_owner(&aligned, position, |pair| {
            matches!(item(pair), Some(BodyContent::Table(_)))
        })
        .is_some_and(|pair| {
            matches!(
                item(pair),
                Some(BodyContent::Paragraph(_) | BodyContent::Table(_))
            )
        });
        match (original_index, edited_index) {
            (Some(left), Some(right)) => {
                let original_content = &original.body.content[left];
                let edited_content = &edited.body.content[right];
                let compared = if original_content == edited_content {
                    if let Some((source, spans)) = original_source {
                        source
                            .get(spans[left].clone())
                            .ok_or_else(|| {
                                Error::Other(format!(
                                    "comparison source span is invalid at {location}[{left}]"
                                ))
                            })?
                            .to_owned()
                    } else {
                        body_content_xml(original_content)?
                    }
                } else {
                    let held_reason = match (
                        held_original.holds(Some(left)),
                        held_edited.holds(Some(right)),
                    ) {
                        (true, true) => Some("the edited document changes it"),
                        (true, false) => Some(
                            "the edited document does not hold the revision the original holds there",
                        ),
                        (false, true) => Some(
                            "the edited document holds a tracked revision there that the original does not",
                        ),
                        (false, false) => None,
                    };
                    if let Some(reason) = held_reason {
                        return Err(Error::Other(format!(
                            "comparison cannot change {}, which holds an existing tracked revision: {reason}",
                            body_location(location, original_content, left)
                        )));
                    }
                    let content_source =
                        original_source.and_then(|(source, spans)| source.get(spans[left].clone()));
                    compare_body_content(
                        original_content,
                        edited_content,
                        content_source,
                        &body_location(location, edited_content, right),
                        metadata,
                        diagnostics,
                    )?
                };
                output.push((
                    EmittedOwner::of(original_content, edited_content, None),
                    compared,
                ));
            }
            (Some(left), None) => {
                let content = &original.body.content[left];
                if held_original.holds(Some(left)) {
                    let place = body_location(location, content, left);
                    return Err(Error::Other(if moves.original[left].is_some() {
                        format!(
                            "comparison cannot move {place}, which holds an existing tracked revision"
                        )
                    } else if carried.original.contains(&left) {
                        format!(
                            "comparison cannot replace {place} whole, because it holds an existing tracked revision"
                        )
                    } else {
                        format!(
                            "comparison cannot delete {place}, which holds an existing tracked revision"
                        )
                    }));
                }
                if let Some(id) = moves.original[left] {
                    if let BodyContent::Paragraph(paragraph) = content
                        && !next_takes_mark
                    {
                        refuse_held_mark(output.len().checked_sub(1))?;
                        let merged =
                            mark_previous_paragraph_with_id(&mut output, "moveFrom", id, metadata)?;
                        output.push(final_paragraph(
                            paragraph,
                            "moveFrom",
                            Some(id),
                            merged,
                            metadata,
                        )?);
                    } else {
                        output.push((
                            EmittedOwner::of(content, content, None),
                            moved_body_content(content, "moveFrom", id, metadata)?,
                        ));
                    }
                } else if let BodyContent::Paragraph(paragraph) = content
                    && trailing_paragraph_delete_start.is_some_and(|start| position >= start)
                {
                    if trailing_paragraph_delete_start == Some(position) {
                        refuse_held_mark(previous_mark_position(&output, "del"))?;
                        trailing_merged = mark_previous_paragraph(&mut output, "del", metadata)?;
                    }
                    if next_takes_mark {
                        output.push((
                            EmittedOwner::of(content, content, Some("del")),
                            deleted_body_content(content, metadata)?,
                        ));
                    } else {
                        output.push(final_paragraph(
                            paragraph,
                            "del",
                            None,
                            trailing_merged,
                            metadata,
                        )?);
                    }
                } else if let BodyContent::Paragraph(paragraph) = content
                    && !next_takes_mark
                {
                    refuse_held_mark(previous_mark_position(&output, "del"))?;
                    let merged = mark_previous_paragraph(&mut output, "del", metadata)?;
                    output.push(final_paragraph(paragraph, "del", None, merged, metadata)?);
                } else {
                    output.push((
                        EmittedOwner::of(content, content, Some("del")),
                        deleted_body_content(content, metadata)?,
                    ));
                }
            }
            (None, Some(right)) => {
                let content = &edited.body.content[right];
                if held_edited.holds(Some(right)) {
                    let action = if moves.edited[right].is_some() {
                        "move"
                    } else {
                        "insert"
                    };
                    return Err(Error::Other(format!(
                        "comparison cannot {action} edited {}, which holds a tracked revision that the original does not hold there",
                        body_location(location, content, right)
                    )));
                }
                if let Some(id) = moves.edited[right] {
                    if let BodyContent::Paragraph(paragraph) = content
                        && !next_takes_mark
                    {
                        refuse_held_mark(output.len().checked_sub(1))?;
                        let merged =
                            mark_previous_paragraph_with_id(&mut output, "moveTo", id, metadata)?;
                        output.push(final_paragraph(
                            paragraph,
                            "moveTo",
                            Some(id),
                            merged,
                            metadata,
                        )?);
                    } else {
                        output.push((
                            EmittedOwner::of(content, content, None),
                            moved_body_content(content, "moveTo", id, metadata)?,
                        ));
                    }
                } else if let BodyContent::Paragraph(paragraph) = content
                    && trailing_paragraph_insert_start.is_some_and(|start| position >= start)
                {
                    if trailing_paragraph_insert_start == Some(position) {
                        refuse_held_mark(previous_mark_position(&output, "ins"))?;
                        trailing_merged = mark_previous_paragraph(&mut output, "ins", metadata)?;
                    }
                    if next_takes_mark {
                        output.push((
                            EmittedOwner::of(content, content, None),
                            inserted_body_content(content, metadata)?,
                        ));
                    } else {
                        output.push(final_paragraph(
                            paragraph,
                            "ins",
                            None,
                            trailing_merged,
                            metadata,
                        )?);
                    }
                } else if let BodyContent::Paragraph(paragraph) = content
                    && !next_takes_mark
                {
                    refuse_held_mark(previous_mark_position(&output, "ins"))?;
                    let merged = mark_previous_paragraph(&mut output, "ins", metadata)?;
                    output.push(final_paragraph(paragraph, "ins", None, merged, metadata)?);
                } else {
                    output.push((
                        EmittedOwner::of(content, content, Some("ins")),
                        inserted_body_content(content, metadata)?,
                    ));
                }
            }
            (None, None) => unreachable!(),
        }
    }
    let mut output = if let Some((source, spans)) = original_source {
        let section_start = if original.body.sect_pr.is_some() {
            direct_word_child_start(source, "w", b"sectPr")?
        } else {
            None
        };
        interleave_story_source(source, spans, &aligned, output, section_start)?
    } else {
        output.into_iter().map(|(_, xml)| xml).collect::<String>()
    };
    if (held_original.section || held_edited.section)
        && original.body.sect_pr != edited.body.sect_pr
    {
        return Err(Error::Other(format!(
            "comparison cannot change the section properties of {location}, which hold an existing tracked revision"
        )));
    }
    let section_xml = section_properties_xml(
        original.body.sect_pr.as_ref(),
        edited.body.sect_pr.as_ref(),
        "section",
        metadata,
        diagnostics,
    )?;
    output.push_str(&section_xml);
    Ok(output)
}

fn interleave_story_source(
    source: &str,
    spans: &[Range<usize>],
    aligned: &[(Option<usize>, Option<usize>)],
    output: Vec<(EmittedOwner<'_>, String)>,
    section_start: Option<usize>,
) -> Result<String> {
    if spans.len() != aligned.iter().filter(|(left, _)| left.is_some()).count()
        || aligned.len() != output.len()
    {
        return Err(Error::Other(
            "comparison source spans do not match related-story owners".to_owned(),
        ));
    }
    let content_end = section_start.unwrap_or(source.len());
    let mut tracked = String::new();
    if let Some(first) = spans.first() {
        tracked.push_str(&source[..first.start]);
    } else {
        tracked.push_str(&source[..content_end]);
    }
    let mut next_original = 0usize;
    for ((left, _), (_, xml)) in aligned.iter().zip(output) {
        if let Some(index) = left {
            if *index != next_original {
                return Err(Error::Other(
                    "comparison related-story owner order is not monotonic".to_owned(),
                ));
            }
            if *index > 0 {
                tracked.push_str(&source[spans[index - 1].end..spans[*index].start]);
            }
            next_original += 1;
        }
        tracked.push_str(&xml);
    }
    if let Some(last) = spans.last() {
        if last.end > content_end {
            return Err(Error::Other(
                "comparison section properties overlap body content".to_owned(),
            ));
        }
        tracked.push_str(&source[last.end..content_end]);
    }
    Ok(tracked)
}

fn direct_word_child_start(
    xml: &str,
    word_prefix: &str,
    local_name: &[u8],
) -> Result<Option<usize>> {
    let open = format!(
        r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:{word_prefix}="{W_NS}">"#
    );
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes()
        );
        match event {
            Event::Start(element) => {
                if depth == 1 && is_word && element.local_name().as_ref() == local_name {
                    return Ok(Some(before.saturating_sub(offset)));
                }
                depth += 1;
            }
            Event::Empty(element)
                if depth == 1 && is_word && element.local_name().as_ref() == local_name =>
            {
                return Ok(Some(before.saturating_sub(offset)));
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(None),
            _ => {}
        }
        buffer.clear();
    }
}

fn moved_body_content(
    content: &BodyContent,
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => moved_paragraph(paragraph, kind, id, metadata),
        BodyContent::Table(table) => moved_table(table, kind, id, metadata),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Err(Error::Other(
            "comparison cannot move a content-control or opaque body node".to_owned(),
        )),
    }
}

fn moved_paragraph(
    paragraph: &CT_P,
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<String> {
    refuse_moved_markers(paragraph)?;
    let marker = IdAllocator::marker_with_id(kind, metadata.author, metadata.timestamp, id);
    let mut properties = paragraph.properties.clone().unwrap_or_default();
    properties.rpr = Some(properties.rpr.take().unwrap_or_default());
    let properties = marked_paragraph_properties_xml(&property_xml(&properties)?, &marker)?;
    let mut output = format!("<w:p>{properties}");
    output.push_str(&wrapped_paragraph_children(paragraph, false, |content| {
        Ok(IdAllocator::revision_with_id(
            kind,
            metadata.author,
            metadata.timestamp,
            content,
            id,
        ))
    })?);
    output.push_str("</w:p>");
    Ok(output)
}

/// Both ends of a move would hold the paragraph's bookmarks and comment
/// ranges, and each may appear once.
fn refuse_moved_markers(paragraph: &CT_P) -> Result<()> {
    if paragraph.bookmark_markers.is_empty() && paragraph.comment_ranges.is_empty() {
        return Ok(());
    }
    Err(Error::Other(
        "comparison cannot move a paragraph that holds a bookmark or comment range".to_owned(),
    ))
}

fn moved_table(table: &CT_Tbl, kind: &str, id: i32, metadata: &Metadata<'_>) -> Result<String> {
    let source = table_xml(table)?;
    let replacements = table
        .rows
        .iter()
        .map(|row| moved_row(row, kind, id, metadata))
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(&source, "tr", &replacements)
}

fn moved_row(row: &CT_Row, kind: &str, id: i32, metadata: &Metadata<'_>) -> Result<String> {
    let marker = IdAllocator::marker_with_id(kind, metadata.author, metadata.timestamp, id);
    let mut xml = row_xml(row)?;
    let properties = direct_word_element_spans(&xml, "trPr")?;
    if let Some(span) = properties.first() {
        let updated = append_word_child(&xml[span.clone()], "trPr", &marker)?;
        xml.replace_range(span.clone(), &updated);
        Ok(xml)
    } else {
        let open = xml
            .find('>')
            .ok_or_else(|| Error::Other("row XML has no start".to_owned()))?
            + 1;
        Ok(format!(
            "{}<w:trPr>{marker}</w:trPr>{}",
            &xml[..open],
            &xml[open..]
        ))
    }
}

/// What a later final paragraph change can do with an owner already emitted.
#[derive(Clone, Copy)]
enum EmittedOwner<'a> {
    /// A paragraph whose mark can carry the change, with its properties after
    /// acceptance and after rejection.
    Paragraph {
        accepted: Option<&'a CT_PPr>,
        rejected: Option<&'a CT_PPr>,
    },
    /// A table whose every row carries this `w:ins` or `w:del` marker. A change
    /// of the same kind crosses it, because resolving that kind removes the
    /// table and leaves the paragraphs on either side adjacent.
    MarkedTable(&'static str),
    /// Source whitespace between owners, which every change crosses.
    Whitespace,
    Other,
}

impl<'a> EmittedOwner<'a> {
    /// The owner emitted for an original and an edited item, the same item
    /// when only one side has it.
    fn of(
        original: &'a BodyContent,
        edited: &'a BodyContent,
        marked_kind: Option<&'static str>,
    ) -> Self {
        match (original, edited, marked_kind) {
            (BodyContent::Paragraph(original), BodyContent::Paragraph(edited), _) => {
                Self::paragraph(original, edited)
            }
            (_, BodyContent::Table(_), Some(kind)) => Self::MarkedTable(kind),
            _ => Self::Other,
        }
    }

    fn of_control(
        original: &'a SdtContent,
        edited: &'a SdtContent,
        marked_kind: Option<&'static str>,
    ) -> Self {
        match (original, edited, marked_kind) {
            (SdtContent::Paragraph(original), SdtContent::Paragraph(edited), _) => {
                Self::paragraph(original, edited)
            }
            (_, SdtContent::Table(_), Some(kind)) => Self::MarkedTable(kind),
            _ => Self::Other,
        }
    }

    fn paragraph(original: &'a CT_P, edited: &'a CT_P) -> Self {
        Self::Paragraph {
            accepted: edited.properties.as_ref(),
            rejected: original.properties.as_ref(),
        }
    }
}

/// The aligned owner after an inserted, deleted or moved one at `position`
/// that the resolution removing its paragraph mark keeps. A table on the same
/// side is crossed, because that resolution removes every row of it.
///
/// When that owner is a paragraph or a table, the paragraph at `position`
/// carries its own change: removing its mark merges it into that paragraph,
/// or, before a table, removes the paragraph, which has no content left.
/// Otherwise its mark ends the story, and [`final_paragraph`] applies.
fn next_kept_owner(
    aligned: &[(Option<usize>, Option<usize>)],
    position: usize,
    is_table: impl Fn((Option<usize>, Option<usize>)) -> bool,
) -> Option<(Option<usize>, Option<usize>)> {
    let (left, right) = aligned[position];
    aligned[position + 1..]
        .iter()
        .copied()
        .find(|&(next_left, next_right)| {
            next_left.is_some() != left.is_some()
                || next_right.is_some() != right.is_some()
                || !is_table((next_left, next_right))
        })
}

/// The position in `output` of the owner whose paragraph mark
/// [`mark_previous_paragraph`] marks for a final `kind` change.
fn previous_mark_position(output: &[(EmittedOwner<'_>, String)], kind: &str) -> Option<usize> {
    output.iter().rposition(|(owner, _)| match owner {
        EmittedOwner::MarkedTable(marked) => *marked != kind,
        EmittedOwner::Whitespace => false,
        EmittedOwner::Paragraph { .. } | EmittedOwner::Other => true,
    })
}

/// Mark the paragraph before a final inserted or deleted paragraph, and
/// return its properties on the side that removes that mark, which
/// [`final_paragraph`] gives the paragraph the two merge into.
fn mark_previous_paragraph<'a>(
    output: &mut [(EmittedOwner<'a>, String)],
    kind: &str,
    metadata: &mut Metadata<'_>,
) -> Result<Option<&'a CT_PPr>> {
    let Some((EmittedOwner::Paragraph { accepted, rejected }, paragraph)) =
        previous_mark_position(output, kind).map(|position| &mut output[position])
    else {
        return Err(Error::Other(
            "comparison needs an adjacent paragraph for a final paragraph change".to_owned(),
        ));
    };
    let marker = metadata
        .ids
        .marker(kind, metadata.author, metadata.timestamp)?;
    *paragraph = marked_paragraph_xml(paragraph, &marker)?;
    Ok(if kind == "del" { *accepted } else { *rejected })
}

fn marked_paragraph_xml(paragraph: &str, marker: &str) -> Result<String> {
    let paragraph_properties = direct_word_element_spans(paragraph, "pPr")?;
    if let Some(properties_span) = paragraph_properties.first() {
        let updated = marked_paragraph_properties_xml(&paragraph[properties_span.clone()], marker)?;
        let mut marked = paragraph.to_owned();
        marked.replace_range(properties_span.clone(), &updated);
        return Ok(marked);
    }
    let open = paragraph
        .find('>')
        .ok_or_else(|| Error::Other("paragraph XML has no start".to_owned()))?
        + 1;
    let properties = format!("<w:pPr><w:rPr>{marker}</w:rPr></w:pPr>");
    if paragraph.as_bytes().get(open.saturating_sub(2)..open) == Some(b"/>") {
        let name_end = paragraph[1..]
            .find(|character: char| {
                character.is_ascii_whitespace() || matches!(character, '/' | '>')
            })
            .map(|offset| offset + 1)
            .ok_or_else(|| Error::Other("paragraph XML has no qualified name".to_owned()))?;
        let name = &paragraph[1..name_end];
        return Ok(format!("{}>{properties}</{name}>", &paragraph[..open - 2]));
    }
    Ok(format!(
        "{}{properties}{}",
        &paragraph[..open],
        &paragraph[open..]
    ))
}

/// Put a paragraph-mark revision marker into serialized `w:pPr` in schema
/// order: first in `w:rPr`, and a new `w:rPr` before `w:sectPr` and
/// `w:pPrChange`.
fn marked_paragraph_properties_xml(properties: &str, marker: &str) -> Result<String> {
    if let Some(run_span) = direct_word_element_spans(properties, "rPr")?.first() {
        let run_properties = &properties[run_span.clone()];
        let updated_run = with_mark_revision(run_properties, marker)?;
        let mut updated = properties.to_owned();
        updated.replace_range(run_span.clone(), &updated_run);
        return Ok(updated);
    }
    let run_properties = format!("<w:rPr>{marker}</w:rPr>");
    let later = direct_word_element_spans(properties, "sectPr")?
        .into_iter()
        .chain(direct_word_element_spans(properties, "pPrChange")?)
        .map(|span| span.start)
        .min();
    if let Some(start) = later {
        let mut updated = properties.to_owned();
        updated.insert_str(start, &run_properties);
        return Ok(updated);
    }
    append_word_child(properties, "pPr", &run_properties)
}

/// [`mark_previous_paragraph`] for a final moved paragraph, whose move
/// pair shares `id`. The paragraph before must be the last owner emitted.
fn mark_previous_paragraph_with_id<'a>(
    output: &mut [(EmittedOwner<'a>, String)],
    kind: &str,
    id: i32,
    metadata: &Metadata<'_>,
) -> Result<Option<&'a CT_PPr>> {
    let Some((EmittedOwner::Paragraph { accepted, rejected }, paragraph)) = output.last_mut()
    else {
        return Err(Error::Other(
            "comparison needs an adjacent paragraph for a final paragraph move".to_owned(),
        ));
    };
    let marker = IdAllocator::marker_with_id(kind, metadata.author, metadata.timestamp, id);
    *paragraph = marked_paragraph_xml(paragraph, &marker)?;
    Ok(if kind == "moveFrom" {
        *accepted
    } else {
        *rejected
    })
}

/// A final inserted, deleted or moved paragraph. Its own mark stays as the
/// story terminator, and the mark of the paragraph before it carries the
/// change. A move wraps the runs with the `move_id` of its pair.
///
/// Removing that mark merges the paragraph before into this one, and the
/// merged paragraph keeps these properties. `merged` holds the properties the
/// paragraph before has on that side, so this paragraph carries the accepted
/// properties with a `w:pPrChange` and a mark `w:rPrChange` that hold the
/// rejected ones, and both resolutions keep each paragraph's properties.
fn final_paragraph<'a>(
    paragraph: &'a CT_P,
    kind: &str,
    move_id: Option<i32>,
    merged: Option<&'a CT_PPr>,
    metadata: &mut Metadata<'_>,
) -> Result<(EmittedOwner<'a>, String)> {
    let own = paragraph.properties.as_ref();
    let (accepted, rejected) = if matches!(kind, "del" | "moveFrom") {
        (merged, own)
    } else {
        (own, merged)
    };
    let mut output = String::from("<w:p>");
    output.push_str(&final_paragraph_properties_xml(
        own, accepted, rejected, metadata,
    )?);
    output.push_str(&wrapped_paragraph_children(
        paragraph,
        kind == "del",
        |content| match move_id {
            Some(id) => Ok(IdAllocator::revision_with_id(
                kind,
                metadata.author,
                metadata.timestamp,
                content,
                id,
            )),
            None => metadata
                .ids
                .revision(kind, metadata.author, metadata.timestamp, content),
        },
    )?);
    output.push_str("</w:p>");
    Ok((EmittedOwner::Paragraph { accepted, rejected }, output))
}

fn compare_body_content(
    original: &BodyContent,
    edited: &BodyContent,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    match (original, edited) {
        (BodyContent::Paragraph(left), BodyContent::Paragraph(right)) => compare_paragraph(
            left,
            right,
            original_source,
            location,
            metadata,
            diagnostics,
        ),
        (BodyContent::Table(left), BodyContent::Table(right)) => compare_table(
            left,
            right,
            original_source,
            location,
            metadata,
            diagnostics,
        ),
        (BodyContent::ContentControl(left), BodyContent::ContentControl(right)) => {
            if let Some(source) = original_source {
                compare_control_from_xml(left, right, location, metadata, diagnostics, source)
            } else {
                compare_control(left, right, location, metadata, diagnostics)
            }
        }
        _ if body_signature(original) == body_signature(edited) => body_content_xml(original),
        _ => Err(Error::Other(format!(
            "comparison cannot replace unlike body structures at {location}"
        ))),
    }
}

fn deleted_body_content(content: &BodyContent, metadata: &mut Metadata<'_>) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => deleted_paragraph(paragraph, metadata),
        BodyContent::Table(table) => marked_table(table, "del", metadata),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Err(Error::Other(
            "comparison cannot delete an unmatched content-control or opaque body node".to_owned(),
        )),
    }
}

fn inserted_body_content(content: &BodyContent, metadata: &mut Metadata<'_>) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => inserted_paragraph(paragraph, metadata),
        BodyContent::Table(table) => marked_table(table, "ins", metadata),
        BodyContent::ContentControl(_) | BodyContent::RawXml(_) => Err(Error::Other(
            "comparison cannot insert an unmatched content-control or opaque body node".to_owned(),
        )),
    }
}

fn compare_paragraph(
    original: &CT_P,
    edited: &CT_P,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if original.bookmark_markers == edited.bookmark_markers
        && paragraph_boundaries_differ(original, edited, metadata.options)
    {
        return Ok(format!(
            "{}{}",
            deleted_paragraph(original, metadata)?,
            inserted_paragraph(edited, metadata)?
        ));
    }
    let detached = detach_field_spans(original, original_source)?;
    let original_source = detached.as_deref().or(original_source);
    if uses_attributed_run_path(metadata.options) {
        return compare_granular_paragraph(
            original,
            edited,
            original_source,
            location,
            metadata,
            diagnostics,
        );
    }
    if !original.hyperlinks.is_empty()
        || !edited.hyperlinks.is_empty()
        || !original.comment_ranges.is_empty()
        || !edited.comment_ranges.is_empty()
        || !original.bookmark_markers.is_empty()
        || !edited.bookmark_markers.is_empty()
        || !original.content_controls.is_empty()
        || !edited.content_controls.is_empty()
        || original
            .extra_xml
            .iter()
            .any(|(position, raw)| !CT_P::raw_is_root_attributes(*position, raw))
        || edited
            .extra_xml
            .iter()
            .any(|(position, raw)| !CT_P::raw_is_root_attributes(*position, raw))
    {
        return compare_complex_paragraph(
            original,
            edited,
            original_source,
            location,
            metadata,
            diagnostics,
        );
    }

    let (mut output, paragraph_close) = paragraph_shell(original_source)?;
    output.push_str(&paragraph_properties_xml(
        original,
        edited,
        location,
        metadata,
        diagnostics,
    )?);
    let original_signatures = original
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let aligned = align(&original_signatures, &edited_signatures);
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original.runs,
            &edited.runs,
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let original_run_spans = original_source
        .map(|source| modeled_paragraph_run_spans(original, source))
        .transpose()?
        .unwrap_or_default();
    if original_source.is_some() && original_run_spans.len() != original.runs.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate paragraph run owners at {location}"
        )));
    }
    for (left, right) in aligned {
        match (left, right) {
            (Some(i), Some(j)) => {
                if original_signatures[i] == edited_signatures[j] {
                    if original.runs[i] == edited.runs[j]
                        && let Some(source) = original_source
                    {
                        output.push_str(&source[original_run_spans[i].clone()]);
                    } else {
                        output.push_str(&compared_run_xml(
                            &original.runs[i],
                            &edited.runs[j],
                            &format!("{location}/run[{j}]"),
                            metadata,
                            diagnostics,
                        )?);
                    }
                } else {
                    let deleted = deleted_run_xml(&original.runs[i])?;
                    output.push_str(&metadata.ids.revision(
                        "del",
                        metadata.author,
                        metadata.timestamp,
                        &deleted,
                    )?);
                    let inserted = policy_inserted_run_xml(
                        &edited.runs[j],
                        Some(&original.runs[i]),
                        metadata.options,
                    )?;
                    output.push_str(&metadata.ids.revision(
                        "ins",
                        metadata.author,
                        metadata.timestamp,
                        &inserted,
                    )?);
                }
            }
            (Some(i), None) => {
                if run_is_ignored(&original.runs[i], metadata.options) {
                    if let Some(source) = original_source {
                        output.push_str(&source[original_run_spans[i].clone()]);
                    } else {
                        output.push_str(&paragraph_owned_run_xml(&original.runs[i])?);
                    }
                    continue;
                }
                let run = deleted_run_xml(&original.runs[i])?;
                output.push_str(&metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &run,
                )?);
            }
            (None, Some(j)) => {
                if run_is_ignored(&edited.runs[j], metadata.options) {
                    continue;
                }
                let run = paragraph_owned_run_xml(&edited.runs[j])?;
                output.push_str(&metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &run,
                )?);
            }
            (None, None) => unreachable!(),
        }
    }
    output.push_str(&paragraph_close);
    Ok(output)
}

fn paragraph_shell(source: Option<&str>) -> Result<(String, String)> {
    let Some(source) = source else {
        return Ok(("<w:p>".to_owned(), "</w:p>".to_owned()));
    };
    let open_end = source
        .find('>')
        .ok_or_else(|| Error::Other("paragraph XML has no start tag".to_owned()))?;
    let name_end = source[1..]
        .find(|character: char| character.is_ascii_whitespace() || matches!(character, '/' | '>'))
        .map(|offset| offset + 1)
        .ok_or_else(|| Error::Other("paragraph XML has no qualified name".to_owned()))?;
    let name = &source[1..name_end];
    let mut opening = source[..=open_end].to_owned();
    if opening.ends_with("/>") {
        opening.truncate(opening.len() - 2);
        opening.push('>');
    }
    Ok((opening, format!("</{name}>")))
}

fn uses_attributed_run_path(options: &ComparisonOptions) -> bool {
    options.granularity != ComparisonGranularity::Run
        || options.ignore_formatting
        || options.ignore_whitespace
        || options.ignore_fields
        || options.ignore_comments
        || story_ignored(options, ComparisonStoryKind::TextBox)
}

#[derive(Clone)]
struct AttributedRunUnit {
    run: CT_R,
    ignored: bool,
    owner: usize,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum GranularAction {
    Preserve,
    Equal,
    Replace,
    Delete,
    Insert,
    Drop,
}

fn compare_granular_paragraph(
    original: &CT_P,
    edited: &CT_P,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let boundary_error = || {
        Error::Other(format!(
            "comparison cannot revise paragraph boundary structures at {location}"
        ))
    };
    if paragraph_boundaries_differ(original, edited, metadata.options) {
        return Err(boundary_error());
    }
    let original_boundaries = shell_run_boundaries(original, metadata.options);
    let edited_boundaries = shell_run_boundaries(edited, metadata.options);

    let original_run_signatures = original
        .runs
        .iter()
        .map(attributed_run_signature)
        .collect::<Vec<_>>();
    let edited_run_signatures = edited
        .runs
        .iter()
        .map(attributed_run_signature)
        .collect::<Vec<_>>();
    // Runs that all match stay whole even when a control differs, so a
    // control's metadata or content never moves run-indexed markers.
    if original_run_signatures == edited_run_signatures && original_boundaries == edited_boundaries
    {
        let properties =
            paragraph_properties_xml(original, edited, location, metadata, diagnostics)?;
        let spans = original_source
            .map(|source| modeled_paragraph_run_spans(original, source))
            .transpose()?
            .unwrap_or_default();
        if original_source.is_some() && spans.len() != original.runs.len() {
            return Err(Error::Other(format!(
                "comparison could not correlate granular run owners at {location}"
            )));
        }
        let replacements = original
            .runs
            .iter()
            .zip(&edited.runs)
            .enumerate()
            .map(|(index, (left, right))| {
                if left == right
                    && let Some(source) = original_source
                {
                    Ok(source[spans[index].clone()].to_owned())
                } else {
                    compared_run_xml(
                        left,
                        right,
                        &format!("{location}/run[{index}]"),
                        metadata,
                        diagnostics,
                    )
                }
            })
            .collect::<Result<Vec<_>>>()?;
        let output = replace_paragraph_properties_and_runs(
            original,
            original_source,
            &properties,
            &replacements,
        )?;
        if original.content_controls == edited.content_controls {
            return Ok(output);
        }
        return compare_granular_controls(
            original,
            edited,
            &output,
            location,
            metadata,
            diagnostics,
        );
    }

    let original_units = attributed_run_units(&original.runs, metadata.options);
    let edited_units = attributed_run_units(&edited.runs, metadata.options);
    let original_signatures = original_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();
    let edited_signatures = edited_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();

    let properties = paragraph_properties_xml(original, edited, location, metadata, diagnostics)?;
    let Some(cuts) = shell_unit_cuts(
        &original_units,
        &edited_units,
        original_boundaries.into_iter().zip(edited_boundaries),
    ) else {
        return Err(boundary_error());
    };
    let (aligned, segments) = align_between_shells(&original_signatures, &edited_signatures, &cuts);
    // The original run each shell segment starts at, before which the
    // redline copies the original bytes, shell tags included.
    let segment_runs = std::iter::once(0)
        .chain(cuts.iter().map(|&(cut, _)| {
            original_units
                .get(cut)
                .map_or(original.runs.len(), |unit| unit.owner)
        }))
        .collect::<Vec<_>>();
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &edited_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let (aligned, segments) = grouped_text_alignment(
        &aligned,
        &segments,
        &original_units,
        &edited_units,
        &original_signatures,
        &edited_signatures,
        metadata.options,
    );
    let grouped = coalesced_granular_alignment(
        &aligned,
        &original_units,
        &edited_units,
        &original_signatures,
        &edited_signatures,
    );
    let mut grouped_alignment = Vec::with_capacity(grouped.len());
    let mut grouped_segment_runs = Vec::with_capacity(grouped.len());
    let mut replacements = Vec::with_capacity(grouped.len());
    for (action, members) in grouped {
        let first = aligned[members.start];
        grouped_alignment.push(first);
        grouped_segment_runs.push(segment_runs[segments[members.start]]);
        let left_indices = members
            .clone()
            .filter_map(|index| aligned[index].0)
            .collect::<Vec<_>>();
        let right_indices = members
            .filter_map(|index| aligned[index].1)
            .collect::<Vec<_>>();
        let left = merge_unit_runs(&original_units, &left_indices);
        let right = merge_unit_runs(&edited_units, &right_indices);
        replacements.push(match action {
            GranularAction::Preserve => paragraph_owned_run_xml(left.as_ref().unwrap())?,
            GranularAction::Equal => compared_run_xml(
                left.as_ref().unwrap(),
                right.as_ref().unwrap(),
                &format!("{location}/run-unit[{}]", right_indices[0]),
                metadata,
                diagnostics,
            )?,
            GranularAction::Replace => {
                let left = left.as_ref().unwrap();
                let deleted = metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted_run_xml(left)?,
                )?;
                let inserted = metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &policy_inserted_run_xml(
                        right.as_ref().unwrap(),
                        Some(left),
                        metadata.options,
                    )?,
                )?;
                format!("{deleted}{inserted}")
            }
            GranularAction::Delete => metadata.ids.revision(
                "del",
                metadata.author,
                metadata.timestamp,
                &deleted_run_xml(left.as_ref().unwrap())?,
            )?,
            GranularAction::Insert => metadata.ids.revision(
                "ins",
                metadata.author,
                metadata.timestamp,
                &paragraph_owned_run_xml(right.as_ref().unwrap())?,
            )?,
            GranularAction::Drop => String::new(),
        });
    }
    interleave_granular_paragraph(
        original,
        edited,
        original_source,
        &original_units,
        &edited_units,
        &grouped_alignment,
        &grouped_segment_runs,
        &replacements,
        &properties,
        location,
        metadata,
        diagnostics,
    )
}

/// The hyperlink owners without their place in the paragraph, which
/// [`shell_unit_cuts`] follows through the unit alignment.
fn hyperlink_shells(paragraph: &CT_P, options: &ComparisonOptions) -> Vec<String> {
    paragraph
        .hyperlinks
        .iter()
        .map(|link| {
            format!(
                "{:?}:{:?}:{:?}:{:?}:{:?}:{:?}:{:?}",
                link.rel_id,
                link.anchor,
                link.tooltip,
                link.doc_location,
                link.extra_attributes,
                hyperlink_raw_boundaries(paragraph, link, options),
                link.preserved_raw_before,
            )
        })
        .collect()
}

/// The run boundaries of the shells that the attributed path keeps from the
/// original: where each hyperlink starts and ends and where each inline
/// control or opaque child sits.
fn shell_run_boundaries(paragraph: &CT_P, options: &ComparisonOptions) -> Vec<usize> {
    paragraph
        .hyperlinks
        .iter()
        .flat_map(|link| [link.run_start, link.run_end])
        .chain(paragraph.content_controls.iter().map(|(at, ..)| *at))
        .chain(paragraph.extra_xml.iter().filter_map(|(at, raw)| {
            (options.granularity != ComparisonGranularity::Run
                && !CT_P::raw_is_root_attributes(*at, raw))
            .then_some(*at)
        }))
        .chain(
            paragraph
                .bookmark_markers
                .iter()
                .map(|marker| marker.run_index()),
        )
        .chain(
            paragraph
                .comment_ranges
                .iter()
                .filter(|_| !options.ignore_comments)
                .map(|marker| match marker {
                    CommentRangeMarker::Start { run_index, .. }
                    | CommentRangeMarker::End { run_index, .. } => *run_index,
                }),
        )
        .collect()
}

/// The `(original, edited)` unit indices where the shell boundaries fall, in
/// document order, for `(original, edited)` run boundaries.
///
/// The units between two consecutive boundaries form a segment, and the
/// redline copies the original bytes before the first run of a segment,
/// shell tags included, before anything of that segment, so each boundary
/// moves with the words inserted or deleted around it. `None` when the
/// shells are not in the same order on both sides, or when the edited side
/// writes a unit between two boundaries that fall between the same two
/// original runs, because the original bytes there are not split. Ignorable
/// and empty units are left out, as in the accept and reject postconditions.
fn shell_unit_cuts(
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    boundaries: impl IntoIterator<Item = (usize, usize)>,
) -> Option<Vec<(usize, usize)>> {
    let mut boundaries = boundaries.into_iter().collect::<Vec<_>>();
    boundaries.sort_unstable();
    if boundaries.windows(2).any(|pair| pair[0].1 > pair[1].1) {
        return None;
    }
    let cuts = boundaries
        .into_iter()
        .map(|(original, edited)| {
            (
                original_units.partition_point(|unit| unit.owner < original),
                edited_units.partition_point(|unit| unit.owner < edited),
            )
        })
        .collect::<Vec<_>>();
    let mut start = (0, 0);
    cuts.iter()
        .all(|&end| {
            let writable = start.0 < end.0
                || edited_units[start.1..end.1]
                    .iter()
                    .all(|unit| unit_is_ignorable(unit) || unit_is_empty(unit));
            start = end;
            writable
        })
        .then_some(cuts)
}

/// The unit alignment made segment by segment between the shell boundaries
/// at `cuts`, so no unit is matched across a hyperlink or an inline control,
/// and the segment of each pair.
#[allow(clippy::type_complexity)]
fn align_between_shells(
    original: &[String],
    edited: &[String],
    cuts: &[(usize, usize)],
) -> (Vec<(Option<usize>, Option<usize>)>, Vec<usize>) {
    let mut aligned = Vec::with_capacity(original.len().max(edited.len()));
    let mut segments = Vec::with_capacity(aligned.capacity());
    let mut start = (0, 0);
    for (segment, &end) in cuts
        .iter()
        .chain(std::iter::once(&(original.len(), edited.len())))
        .enumerate()
    {
        for (left, right) in align(&original[start.0..end.0], &edited[start.1..end.1]) {
            aligned.push((
                left.map(|index| index + start.0),
                right.map(|index| index + start.1),
            ));
            segments.push(segment);
        }
        start = end;
    }
    (aligned, segments)
}

fn hyperlink_raw_boundaries(
    paragraph: &CT_P,
    link: &rdocx_oxml::text::HyperlinkSpan,
    options: &ComparisonOptions,
) -> Vec<(usize, usize, Vec<u8>)> {
    let logical_start = policy_run_boundary(paragraph, link.run_start, options);
    link.extra_xml
        .iter()
        .map(|(boundary, revisions_before, raw)| {
            let absolute = link.run_start.saturating_add(*boundary);
            (
                policy_run_boundary(paragraph, absolute, options).saturating_sub(logical_start),
                *revisions_before,
                raw.clone(),
            )
        })
        .collect()
}

fn policy_run_boundary(
    paragraph: &CT_P,
    run_boundary: usize,
    options: &ComparisonOptions,
) -> usize {
    attributed_run_units(
        &paragraph.runs[..run_boundary.min(paragraph.runs.len())],
        options,
    )
    .iter()
    .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
    .count()
}

fn unit_is_ignorable(unit: &AttributedRunUnit) -> bool {
    unit.ignored && semantic_run_raw(&unit.run).is_empty() && unit.run.alt_drawings.is_empty()
}

fn unit_is_empty(unit: &AttributedRunUnit) -> bool {
    unit.run.content.is_empty()
        && semantic_run_raw(&unit.run).is_empty()
        && unit.run.alt_drawings.is_empty()
        && unit.run.properties.is_none()
}

fn granular_action(
    pair: (Option<usize>, Option<usize>),
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    original_signatures: &[String],
    edited_signatures: &[String],
) -> GranularAction {
    match pair {
        (Some(i), Some(j))
            if original_units[i].ignored
                && edited_units[j].ignored
                && original_signatures[i] == edited_signatures[j] =>
        {
            GranularAction::Preserve
        }
        (Some(i), Some(j)) if original_signatures[i] == edited_signatures[j] => {
            GranularAction::Equal
        }
        (Some(_), Some(_)) => GranularAction::Replace,
        (Some(i), None) if unit_is_ignorable(&original_units[i]) => GranularAction::Preserve,
        (None, Some(j)) if unit_is_ignorable(&edited_units[j]) => GranularAction::Drop,
        (Some(_), None) => GranularAction::Delete,
        (None, Some(_)) => GranularAction::Insert,
        (None, None) => unreachable!(),
    }
}

fn coalesced_granular_alignment(
    aligned: &[(Option<usize>, Option<usize>)],
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    original_signatures: &[String],
    edited_signatures: &[String],
) -> Vec<(GranularAction, Range<usize>)> {
    let mut groups = Vec::new();
    let mut start = 0usize;
    while start < aligned.len() {
        let action = granular_action(
            aligned[start],
            original_units,
            edited_units,
            original_signatures,
            edited_signatures,
        );
        let owners = (
            aligned[start].0.map(|index| original_units[index].owner),
            aligned[start].1.map(|index| edited_units[index].owner),
        );
        let mut end = start + 1;
        if matches!(
            action,
            GranularAction::Replace | GranularAction::Delete | GranularAction::Insert
        ) {
            while end < aligned.len()
                && granular_action(
                    aligned[end],
                    original_units,
                    edited_units,
                    original_signatures,
                    edited_signatures,
                ) == action
                && (
                    aligned[end].0.map(|index| original_units[index].owner),
                    aligned[end].1.map(|index| edited_units[index].owner),
                ) == owners
            {
                end += 1;
            }
        }
        groups.push((action, start..end));
        start = end;
    }
    groups
}

/// The text of a unit that holds nothing but text, which grouping may move
/// into a deletion or an insertion.
fn unit_plain_text(unit: &AttributedRunUnit) -> Option<&str> {
    match unit.run.content.as_slice() {
        [RunContent::Text(text)]
            if !unit.ignored
                && semantic_run_raw(&unit.run).is_empty()
                && unit.run.alt_drawings.is_empty() =>
        {
            Some(&text.text)
        }
        _ => None,
    }
}

/// Count word starts while carrying state across adjacent attributed units.
/// Character fragments from one word must not count as separate words.
fn text_words(text: Option<&str>, in_word: &mut bool) -> usize {
    let mut words = 0;
    if let Some(text) = text {
        for character in text.chars() {
            let word = character.is_alphanumeric() || character == '_';
            words += usize::from(word && !*in_word);
            *in_word = word;
        }
    }
    words
}

/// The word and character alignment with each stretch of changes written as
/// all its deleted units, then all its inserted units, as Word shows a
/// rewritten passage, and the segment of each pair.
///
/// A stretch is consecutive changed text units. It grows over an unchanged
/// stretch between two changes when that stretch is only whitespace, then
/// over one that has fewer words than the deletion and the insertion on each
/// side of it, in the spirit of diff-match-patch `diff_cleanupSemantic`, so a
/// one-word anchor between one-word replacements stays. Anything but plain
/// text ends a stretch: a field, a drawing, a preserved child, an ignored
/// unit, an unchanged unit whose formatting changed, and a shell segment
/// boundary, where a hyperlink, an inline control, a bookmark or a comment
/// range sits. Run granularity keeps its alignment.
fn grouped_text_alignment(
    aligned: &[(Option<usize>, Option<usize>)],
    segments: &[usize],
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    original_signatures: &[String],
    edited_signatures: &[String],
    options: &ComparisonOptions,
) -> (Alignment, Vec<usize>) {
    /// Consecutive pairs of one kind in one segment: `changed` is `None`
    /// for a pair grouping must not move.
    struct Block {
        pairs: Range<usize>,
        changed: Option<bool>,
        deleted_words: usize,
        inserted_words: usize,
        whitespace: bool,
        deleted_in_word: bool,
        inserted_in_word: bool,
        original_owner: Option<usize>,
        edited_owner: Option<usize>,
    }
    if options.granularity == ComparisonGranularity::Run {
        return (aligned.to_vec(), segments.to_vec());
    }
    let same_owner_formatting =
        |left: Option<usize>, right: Option<usize>, units: &[AttributedRunUnit]| {
            left.zip(right).is_none_or(|(left, right)| {
                units[left].owner == units[right].owner
                    || units[left].run.properties == units[right].run.properties
            })
        };
    let mut blocks = Vec::<Block>::new();
    for (index, &(left, right)) in aligned.iter().enumerate() {
        let left_text = left.map(|index| unit_plain_text(&original_units[index]));
        let right_text = right.map(|index| unit_plain_text(&edited_units[index]));
        let plain = left_text.is_none_or(|text| text.is_some())
            && right_text.is_none_or(|text| text.is_some());
        let (left_text, right_text) = (left_text.flatten(), right_text.flatten());
        let same_formatting = match (left, right) {
            (Some(i), Some(j)) => {
                original_units[i].run.properties == edited_units[j].run.properties
            }
            _ => true,
        };
        let changed = match granular_action(
            (left, right),
            original_units,
            edited_units,
            original_signatures,
            edited_signatures,
        ) {
            // A replacement that ignores formatting writes its insertion in
            // the formatting of its deletion, which an insertion of its own
            // would not.
            GranularAction::Replace if plain && (same_formatting || !options.ignore_formatting) => {
                Some(true)
            }
            GranularAction::Delete | GranularAction::Insert if plain => Some(true),
            GranularAction::Equal if plain && same_formatting => Some(false),
            _ => None,
        };
        let inserted_text = right_text.filter(|_| changed == Some(true));
        let whitespace = left_text.is_some_and(|text| text.trim().is_empty());
        match blocks.last_mut() {
            Some(block)
                if changed.is_some()
                    && block.changed == changed
                    && segments[block.pairs.start] == segments[index]
                    && same_owner_formatting(block.original_owner, left, original_units)
                    && same_owner_formatting(block.edited_owner, right, edited_units) =>
            {
                block.pairs.end = index + 1;
                block.deleted_words += text_words(left_text, &mut block.deleted_in_word);
                block.inserted_words += text_words(inserted_text, &mut block.inserted_in_word);
                block.whitespace &= whitespace;
                block.original_owner = left.or(block.original_owner);
                block.edited_owner = right.or(block.edited_owner);
            }
            _ => {
                let mut deleted_in_word = false;
                let mut inserted_in_word = false;
                blocks.push(Block {
                    pairs: index..index + 1,
                    changed,
                    deleted_words: text_words(left_text, &mut deleted_in_word),
                    inserted_words: text_words(inserted_text, &mut inserted_in_word),
                    whitespace,
                    deleted_in_word,
                    inserted_in_word,
                    original_owner: left,
                    edited_owner: right,
                });
            }
        }
    }
    let between_changes = |block: usize| {
        let segment = segments[blocks[block].pairs.start];
        block > 0
            && block + 1 < blocks.len()
            && blocks[block].changed == Some(false)
            && [&blocks[block - 1], &blocks[block + 1]].iter().all(|side| {
                side.changed == Some(true)
                    && segments[side.pairs.start] == segment
                    && same_owner_formatting(
                        side.original_owner,
                        blocks[block].original_owner,
                        original_units,
                    )
                    && same_owner_formatting(
                        side.edited_owner,
                        blocks[block].edited_owner,
                        edited_units,
                    )
            })
    };
    let mut folded = (0..blocks.len())
        .map(|block| between_changes(block) && blocks[block].whitespace)
        .collect::<Vec<_>>();
    // The words deleted and inserted by the change, grown over whitespace,
    // that holds each change block.
    let mut stretch = vec![0usize; blocks.len()];
    let mut stretch_words = vec![(0usize, 0usize); blocks.len()];
    for block in 0..blocks.len() {
        if blocks[block].changed == Some(true) {
            stretch[block] = if block >= 2 && folded[block - 1] {
                stretch[block - 2]
            } else {
                block
            };
            stretch_words[stretch[block]].0 += blocks[block].deleted_words;
            stretch_words[stretch[block]].1 += blocks[block].inserted_words;
        }
    }
    for block in 0..blocks.len() {
        if between_changes(block) && !folded[block] {
            let (before, after) = (
                stretch_words[stretch[block - 1]],
                stretch_words[stretch[block + 1]],
            );
            folded[block] =
                blocks[block].deleted_words < before.0.min(before.1).min(after.0).min(after.1);
        }
    }
    let mut grouped = Vec::with_capacity(aligned.len());
    let mut grouped_segments = Vec::with_capacity(aligned.len());
    let mut block = 0;
    while block < blocks.len() {
        let mut end = block + 1;
        if blocks[block].changed == Some(true) {
            while end + 1 < blocks.len() && folded[end] {
                end += 2;
            }
        }
        let pairs = blocks[block].pairs.start..blocks[end - 1].pairs.end;
        if blocks[block].changed == Some(true) && pairs.len() > 1 {
            let deleted = pairs.clone().filter_map(|index| aligned[index].0);
            let inserted = pairs.clone().filter_map(|index| aligned[index].1);
            grouped.extend(deleted.map(|index| (Some(index), None)));
            grouped.extend(inserted.map(|index| (None, Some(index))));
            grouped_segments.resize(grouped.len(), segments[pairs.start]);
        } else {
            grouped.extend_from_slice(&aligned[pairs.clone()]);
            grouped_segments.extend_from_slice(&segments[pairs]);
        }
        block = end;
    }
    (grouped, grouped_segments)
}

fn merge_unit_runs(units: &[AttributedRunUnit], indices: &[usize]) -> Option<CT_R> {
    let mut merged = units.get(*indices.first()?)?.run.clone();
    let property_boundary = usize::from(merged.properties.is_some());
    for &index in &indices[1..] {
        let next = &units[index].run;
        let content_offset = merged.content.len();
        merged.content.extend(next.content.iter().cloned());
        merged
            .alt_drawings
            .extend(next.alt_drawings.iter().cloned());
        for (raw, &encoded) in next.extra_xml.iter().zip(&next.extra_xml_positions) {
            if CT_R::raw_child_is_root_attributes(encoded) {
                continue;
            }
            let boundary = CT_R::raw_child_position(encoded);
            let mut rebased = encoded;
            CT_R::set_raw_child_position(&mut rebased, boundary + content_offset);
            merged.extra_xml.push(raw.clone());
            merged.extra_xml_positions.push(rebased);
        }
        if property_boundary == 0 && next.properties.is_some() {
            merged.properties = next.properties.clone();
        }
    }
    Some(merged)
}

/// The paragraph source with `properties` in place of its own `w:pPr`. A
/// self-closing paragraph, such as an empty one, is opened first, so the
/// properties and runs go inside it.
fn paragraph_with_properties(
    paragraph: &CT_P,
    original_source: Option<&str>,
    properties: &str,
) -> Result<String> {
    let mut source = original_source
        .map(str::to_owned)
        .map_or_else(|| paragraph_xml(paragraph), Ok)?;
    // The shell opens a self-closing start tag, which is then the whole source.
    let (opening, closing) = paragraph_shell(Some(&source))?;
    if !source.starts_with(&opening) {
        source = format!("{opening}{closing}");
    }
    let property_spans = direct_word_element_spans(&source, "pPr")?;
    match (property_spans.first(), properties.is_empty()) {
        (Some(span), false) => source.replace_range(span.clone(), properties),
        (Some(span), true) => source.replace_range(span.clone(), ""),
        (None, false) => source.insert_str(opening.len(), properties),
        (None, true) => {}
    }
    Ok(source)
}

fn replace_paragraph_properties_and_runs(
    paragraph: &CT_P,
    original_source: Option<&str>,
    properties: &str,
    runs: &[String],
) -> Result<String> {
    let source = paragraph_with_properties(paragraph, original_source, properties)?;
    replace_paragraph_run_elements(paragraph, &source, runs)
}

#[allow(clippy::too_many_arguments)]
fn interleave_granular_paragraph(
    original: &CT_P,
    edited: &CT_P,
    original_source: Option<&str>,
    original_units: &[AttributedRunUnit],
    edited_units: &[AttributedRunUnit],
    aligned: &[(Option<usize>, Option<usize>)],
    segment_runs: &[usize],
    replacements: &[String],
    properties: &str,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let source = paragraph_with_properties(original, original_source, properties)?;
    let spans = modeled_paragraph_run_spans(original, &source)?;
    if spans.len() != original.runs.len()
        || aligned.len() != replacements.len()
        || aligned.len() != segment_runs.len()
    {
        return Err(Error::Other(format!(
            "comparison could not correlate granular run owners at {location}"
        )));
    }
    let run_start = |run: usize| {
        spans
            .get(run)
            .map_or_else(|| paragraph_close_start(&source), |span| Ok(span.start))
    };
    let insertion_boundary = run_start(0)?;
    let mut output = source[..insertion_boundary].to_owned();
    let mut cursor = insertion_boundary;
    let mut consumed_owner = None;
    let mut original_owner_units = vec![0usize; original.runs.len()];
    let mut edited_owner_units = vec![0usize; edited.runs.len()];
    for unit in original_units {
        original_owner_units[unit.owner] += 1;
    }
    for unit in edited_units {
        edited_owner_units[unit.owner] += 1;
    }
    for (((left, right), replacement), &segment_run) in
        aligned.iter().zip(replacements).zip(segment_runs)
    {
        // Text inserted at the start of a shell segment comes after the
        // shell tags that open it.
        let segment_start = run_start(segment_run)?;
        if segment_start > cursor {
            output.push_str(&source[cursor..segment_start]);
            cursor = segment_start;
        }
        let mut exact_run = None;
        if let Some(unit) = left.map(|index| &original_units[index])
            && consumed_owner != Some(unit.owner)
        {
            let span = &spans[unit.owner];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
            consumed_owner = Some(unit.owner);
            if let Some(right_owner) = right.map(|index| edited_units[index].owner)
                && original.runs[unit.owner] == edited.runs[right_owner]
                && original_owner_units[unit.owner] == 1
                && edited_owner_units[right_owner] == 1
            {
                exact_run = Some(&source[span.clone()]);
            }
        }
        output.push_str(exact_run.unwrap_or(replacement));
    }
    output.push_str(&source[cursor..]);
    compare_granular_controls(original, edited, &output, location, metadata, diagnostics)
}

/// Compare the inline controls of a paragraph whose revised runs are in
/// `output`, which still holds the original controls.
fn compare_granular_controls(
    original: &CT_P,
    edited: &CT_P,
    output: &str,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let control_spans = direct_word_element_spans(output, "sdt")?;
    if control_spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate granular content controls at {location}"
        )));
    }
    let mut control_replacements = Vec::with_capacity(control_spans.len());
    for (index, ((_, _, _, left), (_, _, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        control_replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[control_spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(output, "sdt", &control_replacements)
}

fn attributed_run_units(runs: &[CT_R], options: &ComparisonOptions) -> Vec<AttributedRunUnit> {
    let mut units = Vec::new();
    for (owner, run) in runs.iter().enumerate() {
        let unit_start = units.len();
        let mut content_units = Vec::<Range<usize>>::with_capacity(run.content.len());
        for content in &run.content {
            let start = units.len() - unit_start;
            let fragments = match content {
                RunContent::Text(text) => granular_text(text, options)
                    .into_iter()
                    .map(RunContent::Text)
                    .collect::<Vec<_>>(),
                RunContent::DeletedText(text) => granular_text(text, options)
                    .into_iter()
                    .map(RunContent::DeletedText)
                    .collect::<Vec<_>>(),
                content => vec![content.clone()],
            };
            for fragment in fragments {
                let ignored = ignored_run_content(&fragment, options);
                units.push(AttributedRunUnit {
                    run: CT_R {
                        properties: run.properties.clone(),
                        content: vec![fragment],
                        extra_xml: Vec::new(),
                        extra_xml_positions: Vec::new(),
                        alt_drawings: Vec::new(),
                    },
                    ignored,
                    owner,
                });
            }
            content_units.push(start..units.len() - unit_start);
        }
        if units.len() == unit_start {
            units.push(AttributedRunUnit {
                run: CT_R {
                    properties: run.properties.clone(),
                    content: Vec::new(),
                    extra_xml: Vec::new(),
                    extra_xml_positions: Vec::new(),
                    alt_drawings: run.alt_drawings.clone(),
                },
                ignored: false,
                owner,
            });
        } else {
            units[unit_start].run.alt_drawings = run.alt_drawings.clone();
        }
        attribute_raw_children(run, &content_units, &mut units[unit_start..]);
        let property_boundary = usize::from(run.properties.is_some());
        let mut inserted = 0usize;
        for (raw, &encoded) in run.extra_xml.iter().zip(&run.extra_xml_positions) {
            if !is_text_box_host_marker(raw) {
                continue;
            }
            let boundary = CT_R::raw_child_position(encoded);
            let offset = if boundary <= property_boundary || content_units.is_empty() {
                0
            } else {
                let content_index = (boundary - property_boundary - 1).min(content_units.len() - 1);
                content_units[content_index].end
            };
            let mut position = encoded;
            CT_R::set_raw_child_position(&mut position, property_boundary);
            units.insert(
                unit_start + offset + inserted,
                AttributedRunUnit {
                    run: CT_R {
                        properties: run.properties.clone(),
                        content: Vec::new(),
                        extra_xml: vec![raw.clone()],
                        extra_xml_positions: vec![position],
                        alt_drawings: Vec::new(),
                    },
                    ignored: false,
                    owner,
                },
            );
            inserted += 1;
        }
    }
    units
}

fn attribute_raw_children(
    run: &CT_R,
    content_units: &[Range<usize>],
    units: &mut [AttributedRunUnit],
) {
    let property_boundary = usize::from(run.properties.is_some());
    for (raw, &encoded) in run.extra_xml.iter().zip(&run.extra_xml_positions) {
        if is_text_box_host_marker(raw) {
            continue;
        }
        let boundary = CT_R::raw_child_position(encoded);
        let (unit_index, position) = if boundary <= property_boundary || content_units.is_empty() {
            (0, boundary.min(property_boundary))
        } else {
            let content_index = (boundary - property_boundary - 1).min(content_units.len() - 1);
            let range = &content_units[content_index];
            (range.end.saturating_sub(1), property_boundary + 1)
        };
        units[unit_index].run.extra_xml.push(raw.clone());
        let mut attributed = encoded;
        CT_R::set_raw_child_position(&mut attributed, position);
        units[unit_index].run.extra_xml_positions.push(attributed);
    }
}

fn is_text_box_host_marker(raw: &[u8]) -> bool {
    raw.windows(b"urn:rdocx:comparison:private".len())
        .any(|window| window == b"urn:rdocx:comparison:private")
        && raw
            .windows(b"textBoxHost".len())
            .any(|window| window == b"textBoxHost")
}

fn granular_text(text: &CT_Text, options: &ComparisonOptions) -> Vec<CT_Text> {
    let fragments = match options.granularity {
        ComparisonGranularity::Run if options.ignore_whitespace => whitespace_fragments(&text.text),
        ComparisonGranularity::Run => vec![text.text.clone()],
        ComparisonGranularity::Character => text.text.chars().map(String::from).collect(),
        ComparisonGranularity::Word => word_fragments(&text.text),
    };
    if fragments.is_empty() {
        return vec![text.clone()];
    }
    fragments
        .into_iter()
        .map(|value| CT_Text {
            // A unit is written back as its own `w:t`, where edge whitespace
            // needs the flag to survive. Whitespace in a unit then reads as
            // whitespace whatever the source flag was.
            preserve_space: text.preserve_space || has_edge_whitespace(&value),
            text: value,
        })
        .collect()
}

/// The whole-run signature that agrees with the attributed units.
///
/// It reads the space flag the way `granular_text` writes it on every unit,
/// so a run that differs only by the flag stays whole instead of being
/// rewritten one unit per run, which would move run-indexed bookmark ends.
fn attributed_run_signature(run: &CT_R) -> String {
    let mut run = run.clone();
    for content in &mut run.content {
        if let RunContent::Text(text) | RunContent::DeletedText(text) = content {
            text.preserve_space |= has_edge_whitespace(&text.text);
        }
    }
    run_signature(&run)
}

fn whitespace_fragments(text: &str) -> Vec<String> {
    let mut output = Vec::new();
    let mut current = String::new();
    let mut whitespace = None;
    for character in text.chars() {
        let next = character.is_whitespace();
        if whitespace.is_some_and(|value| value != next) {
            output.push(std::mem::take(&mut current));
        }
        current.push(character);
        whitespace = Some(next);
    }
    if !current.is_empty() {
        output.push(current);
    }
    output
}

fn word_fragments(text: &str) -> Vec<String> {
    fn class(character: char) -> u8 {
        if character.is_alphanumeric() || character == '_' {
            0
        } else if character.is_whitespace() {
            1
        } else {
            2
        }
    }
    let mut output = Vec::new();
    let mut current = String::new();
    let mut current_class = None;
    for character in text.chars() {
        let next_class = class(character);
        if current_class.is_some_and(|value| value != next_class) {
            output.push(std::mem::take(&mut current));
        }
        current.push(character);
        current_class = Some(next_class);
    }
    if !current.is_empty() {
        output.push(current);
    }
    output
}

fn ignored_run_content(content: &RunContent, options: &ComparisonOptions) -> bool {
    (options.ignore_fields && matches!(content, RunContent::Field(_)))
        || (options.ignore_comments && matches!(content, RunContent::CommentReference { .. }))
        || (options.ignore_whitespace
            && matches!(
                content,
                RunContent::Text(text) | RunContent::DeletedText(text)
                    if text.text.chars().all(char::is_whitespace)
            ))
}

fn run_is_ignored(run: &CT_R, options: &ComparisonOptions) -> bool {
    !run.content.is_empty()
        && run
            .content
            .iter()
            .all(|content| ignored_run_content(content, options))
        && semantic_run_raw(run).is_empty()
}

fn attributed_unit_signature(unit: &AttributedRunUnit) -> String {
    if unit.ignored {
        let kind = match unit.run.content.first() {
            Some(RunContent::Field(_)) => "ignored:field".to_owned(),
            Some(RunContent::CommentReference { .. }) => "ignored:comment".to_owned(),
            _ => "ignored:whitespace".to_owned(),
        };
        let raw = semantic_run_raw(&unit.run);
        format!("{kind}:{raw:?}:{:?}", unit.run.alt_drawings)
    } else {
        match unit.run.content.first() {
            Some(RunContent::CommentReference { id, .. }) if unit.run.content.len() == 1 => {
                let raw = semantic_run_raw(&unit.run);
                format!(
                    "comment:{id}:{:?}:{raw:?}:{:?}",
                    unit.run.properties, unit.run.alt_drawings
                )
            }
            _ => run_signature(&unit.run),
        }
    }
}

fn compare_complex_paragraph(
    original: &CT_P,
    edited: &CT_P,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if paragraph_signature_with_options(original, metadata.options)
        == paragraph_signature_with_options(edited, metadata.options)
    {
        if !metadata.options.ignore_formatting
            && paragraph_formatting(original) != paragraph_formatting(edited)
        {
            formatting_diagnostic(diagnostics, location.to_owned());
        }
        for (index, (left, right)) in original.runs.iter().zip(&edited.runs).enumerate() {
            if !metadata.options.ignore_formatting && left.properties != right.properties {
                formatting_diagnostic(diagnostics, format!("{location}/run[{index}]"));
            }
        }
        for (index, ((_, _, _, left), (_, _, _, right))) in original
            .content_controls
            .iter()
            .zip(&edited.content_controls)
            .enumerate()
        {
            compare_control(
                left,
                right,
                &format!("{location}/content-control[{index}]"),
                metadata,
                diagnostics,
            )?;
        }
        return original_source
            .map(str::to_owned)
            .map_or_else(|| paragraph_xml(original), Ok);
    }
    if paragraph_boundaries_differ(original, edited, metadata.options)
        || (!metadata.options.ignore_formatting
            && paragraph_properties_differ(
                original.properties.as_ref(),
                edited.properties.as_ref(),
            ))
    {
        return Err(Error::Other(format!(
            "comparison cannot revise paragraph boundary structures at {location}"
        )));
    }

    let source = original_source
        .map(str::to_owned)
        .map_or_else(|| paragraph_xml(original), Ok)?;
    let run_spans = modeled_paragraph_run_spans(original, &source)?;
    if run_spans.len() != original.runs.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate complex paragraph runs at {location}"
        )));
    }
    let original_signatures = original
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .runs
        .iter()
        .map(|run| run_signature_with_options(run, metadata.options))
        .collect::<Vec<_>>();
    let aligned = align(&original_signatures, &edited_signatures);
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original.runs,
            &edited.runs,
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let mut output = String::new();
    let mut cursor = 0usize;
    for (left, right) in aligned {
        if let Some(index) = left {
            let span = &run_spans[index];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
        }
        match (left, right) {
            (Some(i), Some(j)) if original_signatures[i] == edited_signatures[j] => {
                if original.runs[i] == edited.runs[j] {
                    output.push_str(&source[run_spans[i].clone()]);
                } else {
                    output.push_str(&compared_run_xml(
                        &original.runs[i],
                        &edited.runs[j],
                        &format!("{location}/run[{j}]"),
                        metadata,
                        diagnostics,
                    )?);
                }
            }
            (Some(i), Some(j)) => {
                let deleted = deleted_run_xml(&original.runs[i])?;
                output.push_str(&metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted,
                )?);
                let inserted = policy_inserted_run_xml(
                    &edited.runs[j],
                    Some(&original.runs[i]),
                    metadata.options,
                )?;
                output.push_str(&metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &inserted,
                )?);
            }
            (Some(i), None) => {
                if run_is_ignored(&original.runs[i], metadata.options) {
                    output.push_str(&source[run_spans[i].clone()]);
                    continue;
                }
                let deleted = deleted_run_xml(&original.runs[i])?;
                output.push_str(&metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted,
                )?);
            }
            (None, Some(j)) => {
                if run_is_ignored(&edited.runs[j], metadata.options) {
                    continue;
                }
                let inserted = paragraph_owned_run_xml(&edited.runs[j])?;
                output.push_str(&metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &inserted,
                )?);
            }
            (None, None) => unreachable!(),
        }
    }
    output.push_str(&source[cursor..]);
    let spans = direct_word_element_spans(&output, "sdt")?;
    if spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate paragraph content controls at {location}"
        )));
    }
    let mut replacements = Vec::with_capacity(spans.len());
    for (index, ((_, _, _, left), (_, _, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &replacements)
}

fn paragraph_control_boundaries(paragraph: &CT_P) -> Vec<(usize, usize, usize)> {
    paragraph
        .content_controls
        .iter()
        .map(|(at, raw_before, markers_before, _)| (*at, *raw_before, *markers_before))
        .collect()
}

/// Whether two paragraphs differ in an inline structure that revising the
/// paragraph in place cannot express: a hyperlink, comment range, bookmark,
/// preserved raw child or inline content control.
///
/// The word and character paths follow a hyperlink or control through the
/// text alignment, so only the shells and their slots count there.
fn paragraph_boundaries_differ(
    original: &CT_P,
    edited: &CT_P,
    options: &ComparisonOptions,
) -> bool {
    if (!options.ignore_comments && original.comment_ranges != edited.comment_ranges)
        || original.bookmark_markers != edited.bookmark_markers
    {
        return true;
    }
    if uses_attributed_run_path(options) {
        let control_slots = |paragraph: &CT_P| {
            paragraph
                .content_controls
                .iter()
                .map(|(_, raw_before, markers_before, _)| (*raw_before, *markers_before))
                .collect::<Vec<_>>()
        };
        return hyperlink_shells(original, options) != hyperlink_shells(edited, options)
            || control_slots(original) != control_slots(edited);
    }
    original.hyperlinks != edited.hyperlinks
        || original
            .extra_xml
            .iter()
            .filter(|(position, raw)| !CT_P::raw_is_root_attributes(*position, raw))
            .ne(edited
                .extra_xml
                .iter()
                .filter(|(position, raw)| !CT_P::raw_is_root_attributes(*position, raw)))
        || paragraph_control_boundaries(original) != paragraph_control_boundaries(edited)
}

fn paragraph_properties_xml(
    original: &CT_P,
    edited: &CT_P,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_section = original
        .properties
        .as_ref()
        .and_then(|properties| properties.sect_pr.as_ref());
    let edited_section = edited
        .properties
        .as_ref()
        .and_then(|properties| properties.sect_pr.as_ref());
    let tracked_section = section_properties_xml(
        original_section,
        edited_section,
        &format!("{location}/section"),
        metadata,
        diagnostics,
    )?;
    if metadata.options.ignore_formatting {
        return original
            .properties
            .as_ref()
            .map(property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }

    let original_modeled = modeled_paragraph_properties(original.properties.as_ref());
    let edited_modeled = modeled_paragraph_properties(edited.properties.as_ref());
    let changed = original_modeled != edited_modeled;
    let mut current = if changed {
        edited.properties.clone().unwrap_or_default()
    } else {
        original.properties.clone().unwrap_or_default()
    };
    current.sect_pr = None;
    current.change = None;
    let (
        original_numbering_xml,
        original_numbering_positions,
        original_numbering_position,
        original_revision_xml,
        original_revision_positions,
    ) = original
        .properties
        .as_ref()
        .map(|properties| {
            (
                properties.numbering_revision_xml.clone(),
                properties.numbering_revision_xml_positions.clone(),
                properties.numbering_revision_position,
                properties.revision_xml.clone(),
                properties.revision_xml_positions.clone(),
            )
        })
        .unwrap_or_default();
    let edited_properties = edited.properties.as_ref().cloned().unwrap_or_default();
    if edited_properties.numbering_revision_xml != original_numbering_xml
        || edited_properties.numbering_revision_xml_positions != original_numbering_positions
        || edited_properties.numbering_revision_position != original_numbering_position
        || edited_properties.revision_xml != original_revision_xml
        || edited_properties.revision_xml_positions != original_revision_positions
    {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    current.numbering_revision_xml = original_numbering_xml;
    current.numbering_revision_xml_positions = original_numbering_positions;
    current.numbering_revision_position = original_numbering_position;
    current.revision_xml = original_revision_xml;
    current.revision_xml_positions = original_revision_positions;
    let needs_owner = changed || !tracked_section.is_empty() || original.properties.is_some();
    if !needs_owner {
        return Ok(String::new());
    }
    let mut current_xml = property_xml(&current)?;
    if current_xml.is_empty() {
        current_xml.push_str("<w:pPr></w:pPr>");
    }
    // `w:pPrChange` holds only the base properties, so the mark records its
    // own change in a `w:rPrChange`, as Word writes it.
    let (original_base, original_mark) = split_paragraph_mark(original_modeled);
    let (edited_base, edited_mark) = split_paragraph_mark(edited_modeled);
    if original_mark != edited_mark {
        let previous = run_property_xml(original_mark.as_ref())?;
        let change =
            metadata
                .ids
                .revision("rPrChange", metadata.author, metadata.timestamp, &previous)?;
        if let Some(span) = direct_word_element_spans(&current_xml, "rPr")?
            .into_iter()
            .next()
        {
            let updated = append_word_child(&current_xml[span.clone()], "rPr", &change)?;
            current_xml.replace_range(span, &updated);
        } else {
            current_xml = inject_before_close(
                &current_xml,
                "</w:pPr>",
                &format!("<w:rPr>{change}</w:rPr>"),
            )?;
        }
    }
    if !tracked_section.is_empty() {
        current_xml = inject_before_close(&current_xml, "</w:pPr>", &tracked_section)?;
    }
    if original_base != edited_base {
        let previous = original_base
            .as_ref()
            .map(property_xml)
            .transpose()?
            .filter(|xml| !xml.is_empty())
            .unwrap_or_else(|| "<w:pPr/>".to_owned());
        let change =
            metadata
                .ids
                .revision("pPrChange", metadata.author, metadata.timestamp, &previous)?;
        current_xml = inject_before_close(&current_xml, "</w:pPr>", &change)?;
    }
    Ok(current_xml)
}

/// Modeled paragraph properties without the paragraph mark, and the mark.
fn split_paragraph_mark(
    properties: Option<CT_PPr>,
) -> (Option<CT_PPr>, Option<rdocx_oxml::properties::CT_RPr>) {
    match properties {
        Some(mut properties) => {
            let mark = properties.rpr.take();
            (nonempty_paragraph_properties(properties), mark)
        }
        None => (None, None),
    }
}

/// The `w:pPr` of a [`final_paragraph`]: the `own` properties when both
/// resolutions agree, otherwise the `accepted` ones with the section break of
/// `own`, a `w:pPrChange` holding the rejected paragraph properties, and a
/// mark `w:rPrChange` holding the rejected mark formatting.
///
/// Rejecting a `w:pPrChange` keeps a mark `w:rPr` only when it holds a
/// revision, so a mark with formatting also carries the `w:rPrChange` when
/// only the paragraph properties differ.
fn final_paragraph_properties_xml(
    own: Option<&CT_PPr>,
    accepted: Option<&CT_PPr>,
    rejected: Option<&CT_PPr>,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let accepted_modeled = modeled_paragraph_properties(accepted);
    let rejected_modeled = modeled_paragraph_properties(rejected);
    if metadata.options.ignore_formatting || accepted_modeled == rejected_modeled {
        return own
            .map(property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }
    let split = |properties: Option<CT_PPr>| match properties {
        Some(mut properties) => {
            let mark = properties.rpr.take();
            (nonempty_paragraph_properties(properties), mark)
        }
        None => (None, None),
    };
    let (accepted_base, accepted_mark) = split(accepted_modeled);
    let (rejected_base, rejected_mark) = split(rejected_modeled);
    let mark_changed = accepted_mark != rejected_mark;
    let mut current = accepted.cloned().unwrap_or_default();
    current.sect_pr = own.and_then(|properties| properties.sect_pr.clone());
    let mut xml = property_xml(&current)?;
    if xml.is_empty() {
        xml.push_str("<w:pPr></w:pPr>");
    }
    if mark_changed {
        let previous = run_property_xml(rejected_mark.as_ref())?;
        let change =
            metadata
                .ids
                .revision("rPrChange", metadata.author, metadata.timestamp, &previous)?;
        if let Some(span) = direct_word_element_spans(&xml, "rPr")?.into_iter().next() {
            let updated = append_word_child(&xml[span.clone()], "rPr", &change)?;
            xml.replace_range(span, &updated);
        } else {
            let run_properties = format!("<w:rPr>{change}</w:rPr>");
            match direct_word_element_spans(&xml, "sectPr")?.first() {
                Some(section) => xml.insert_str(section.start, &run_properties),
                None => xml = append_word_child(&xml, "pPr", &run_properties)?,
            }
        }
    }
    if accepted_base != rejected_base {
        let previous = rejected_base
            .as_ref()
            .map(property_xml)
            .transpose()?
            .filter(|previous| !previous.is_empty())
            .unwrap_or_else(|| "<w:pPr/>".to_owned());
        let change =
            metadata
                .ids
                .revision("pPrChange", metadata.author, metadata.timestamp, &previous)?;
        xml = inject_before_close(&xml, "</w:pPr>", &change)?;
    }
    Ok(xml)
}

fn modeled_paragraph_properties(properties: Option<&CT_PPr>) -> Option<CT_PPr> {
    properties.cloned().and_then(|mut properties| {
        properties.sect_pr = None;
        properties.num_ilvl_raw = None;
        properties.num_id_raw = None;
        properties.numbering_revision = None;
        properties.numbering_revision_xml.clear();
        properties.numbering_revision_xml_positions.clear();
        properties.numbering_revision_position = None;
        properties.change = None;
        properties.revision_xml.clear();
        properties.revision_xml_positions.clear();
        nonempty_paragraph_properties(properties)
    })
}

fn nonempty_paragraph_properties(mut properties: CT_PPr) -> Option<CT_PPr> {
    if properties
        .rpr
        .as_ref()
        .is_some_and(|mark| *mark == rdocx_oxml::properties::CT_RPr::default())
    {
        properties.rpr = None;
    }
    (properties != CT_PPr::default()).then_some(properties)
}

fn paragraph_properties_differ(original: Option<&CT_PPr>, edited: Option<&CT_PPr>) -> bool {
    let modeled = |properties: Option<&CT_PPr>| {
        properties.cloned().and_then(|mut properties| {
            if let Some(section) = properties.sect_pr.as_mut() {
                clear_default_orientation(section);
            }
            nonempty_paragraph_properties(properties)
        })
    };
    modeled(original) != modeled(edited)
}

/// Portrait is the schema default, which the section writer omits.
fn clear_default_orientation(section: &mut rdocx_oxml::document::CT_SectPr) {
    if section.orientation == Some(ST_PageOrientation::Portrait) {
        section.orientation = None;
    }
}

fn section_properties_xml(
    original: Option<&rdocx_oxml::document::CT_SectPr>,
    edited: Option<&rdocx_oxml::document::CT_SectPr>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let (Some(original), Some(edited)) = (original, edited) else {
        return match (original, edited) {
            (None, None) => Ok(String::new()),
            _ => Err(Error::Other(format!(
                "comparison cannot create or remove a section shell at {location}"
            ))),
        };
    };
    if original.header_refs != edited.header_refs || original.footer_refs != edited.footer_refs {
        return Err(Error::Other(format!(
            "comparison cannot change related-story references at {location}"
        )));
    }
    if metadata.options.ignore_formatting {
        return section_property_xml(original);
    }
    let mut original_modeled = original.clone();
    original_modeled.change = None;
    let mut edited_modeled = edited.clone();
    edited_modeled.change = None;
    clear_default_orientation(&mut original_modeled);
    clear_default_orientation(&mut edited_modeled);
    if original_modeled == edited_modeled {
        return section_property_xml(original);
    }
    if original.extra_xml != edited.extra_xml {
        formatting_diagnostic(diagnostics, location.to_owned());
        edited_modeled.extra_xml = original.extra_xml.clone();
    }
    let previous = section_property_xml(&original_modeled)?;
    let change = metadata.ids.revision(
        "sectPrChange",
        metadata.author,
        metadata.timestamp,
        &previous,
    )?;
    let current = section_property_xml(&edited_modeled)?;
    inject_before_close(&current, "</w:sectPr>", &change)
}

fn deleted_paragraph(paragraph: &CT_P, metadata: &mut Metadata<'_>) -> Result<String> {
    let mut output = String::from("<w:p>");
    output.push_str(&paragraph_mark_properties(
        paragraph.properties.as_ref(),
        "del",
        metadata,
    )?);
    output.push_str(&wrapped_paragraph_children(paragraph, true, |content| {
        metadata
            .ids
            .revision("del", metadata.author, metadata.timestamp, content)
    })?);
    output.push_str("</w:p>");
    Ok(output)
}

fn inserted_paragraph(paragraph: &CT_P, metadata: &mut Metadata<'_>) -> Result<String> {
    let mut output = String::from("<w:p>");
    output.push_str(&paragraph_mark_properties(
        paragraph.properties.as_ref(),
        "ins",
        metadata,
    )?);
    output.push_str(&wrapped_paragraph_children(paragraph, false, |content| {
        metadata
            .ids
            .revision("ins", metadata.author, metadata.timestamp, content)
    })?);
    output.push_str("</w:p>");
    Ok(output)
}

/// The children of a whole deleted, inserted or moved paragraph after its
/// properties, each group inside the wrapper that `wrap` writes.
///
/// Simple fields, bookmarks and comment ranges go inside the wrappers with
/// the runs, so resolving the revision removes or keeps them with the text.
/// A range or proofing marker shares the wrapper of the content after it,
/// or of the content before it at the end of the paragraph, so a paragraph
/// of runs alone keeps one wrapper per run.
///
/// A revision wrapper may not hold a hyperlink, so a hyperlink stays outside
/// and its own children go inside wrappers, as Word writes a tracked
/// hyperlink run. Resolving the revision keeps the hyperlink when it keeps
/// the text, and drops the hyperlink it leaves empty.
fn wrapped_paragraph_children(
    paragraph: &CT_P,
    deleted: bool,
    mut wrap: impl FnMut(&str) -> Result<String>,
) -> Result<String> {
    let mut children = paragraph.clone();
    children.properties = None;
    if deleted {
        children.runs = children.runs.iter().map(deleted_run).collect();
    }
    wrapped_children(&paragraph_xml(&children)?, deleted, false, &mut wrap)
}

/// Expanded Word names in a serialized owner. The writer supplies canonical
/// `w` names, while preserved children can shadow that binding or use aliases.
fn word_element_starts(xml: &str) -> Result<HashSet<usize>> {
    let opening = format!(r#"<scope xmlns:w="{W_NS}">"#);
    let scoped = format!("{opening}{xml}</scope>");
    let mut reader = NsReader::from_reader(scoped.as_bytes());
    let mut starts = HashSet::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        if matches!(event, Event::Start(_) | Event::Empty(_))
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
            && let Some(start) = before.checked_sub(opening.len())
        {
            starts.insert(start);
        }
        if matches!(event, Event::Eof) {
            return Ok(starts);
        }
        buffer.clear();
    }
}

/// The children of the element `xml`, a paragraph or a hyperlink, as
/// [`wrapped_paragraph_children`] writes them. Only a direct hyperlink is
/// admitted, bounding calls to the paragraph and one hyperlink owner. Captured
/// nested hyperlinks stay unsupported rather than growing a copied-subtree stack.
fn wrapped_children(
    xml: &str,
    deleted: bool,
    inside_hyperlink: bool,
    wrap: &mut impl FnMut(&str) -> Result<String>,
) -> Result<String> {
    let mut output = String::new();
    let mut groups: Vec<(String, bool)> = Vec::new();
    let word_starts = word_element_starts(xml)?;
    let mut cursor = xml.find('>').map_or(xml.len(), |at| at + 1);
    for span in direct_element_spans(xml)? {
        if cursor < span.start {
            groups.push((xml[cursor..span.start].to_owned(), false));
        }
        cursor = span.end;
        let word = word_starts.contains(&span.start);
        let child = &xml[span];
        let local = if word { element_local_name(child) } else { "" };
        if local == "hyperlink" {
            if inside_hyperlink {
                return Err(Error::Other(
                    "comparison cannot track nested hyperlinks in a whole paragraph".to_owned(),
                ));
            }
            let open_end = child
                .find('>')
                .ok_or_else(|| Error::Other("hyperlink XML has no start".to_owned()))?
                + 1;
            for (group, _) in groups.drain(..) {
                output.push_str(&wrap(&group)?);
            }
            // Resolution drops a hyperlink only when it removes its content.
            let inner = wrapped_children(child, deleted, true, wrap)?;
            if inner.is_empty() {
                return Err(Error::Other(
                    "comparison cannot track a whole paragraph's empty hyperlink".to_owned(),
                ));
            }
            output.push_str(&child[..open_end]);
            output.push_str(&inner);
            output.push_str(&child[child.rfind("</").unwrap_or(child.len())..]);
            continue;
        }
        let marker = matches!(
            local,
            "bookmarkStart" | "bookmarkEnd" | "commentRangeStart" | "commentRangeEnd" | "proofErr"
        );
        // Runs already hold deleted text, as `deleted_run_xml` writes them.
        // Field results, hyperlink runs and other owners are renamed whole.
        let child = match (deleted, local) {
            (true, "") => child.to_owned(),
            (true, "r") => renamed_word_elements(child, "w:instrText", "w:delInstrText"),
            (true, _) => deleted_text_xml(child),
            (false, _) => child.to_owned(),
        };
        match groups.last_mut() {
            Some((group, has_content)) if !*has_content => {
                group.push_str(&child);
                *has_content = !marker;
            }
            _ => groups.push((child, !marker)),
        }
    }
    let close = xml.rfind("</").unwrap_or(xml.len());
    if cursor < close {
        groups.push((xml[cursor..close].to_owned(), false));
    }
    if groups.len() > 1
        && let Some((markers, false)) = groups.pop_if(|(_, has_content)| !*has_content)
        && let Some((group, _)) = groups.last_mut()
    {
        group.push_str(&markers);
    }
    for (group, _) in &groups {
        output.push_str(&wrap(group)?);
    }
    Ok(output)
}

fn paragraph_mark_properties(
    properties: Option<&CT_PPr>,
    kind: &str,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let marker = metadata
        .ids
        .marker(kind, metadata.author, metadata.timestamp)?;
    let mut properties = properties.cloned().unwrap_or_default();
    properties.rpr = Some(properties.rpr.take().unwrap_or_default());
    marked_paragraph_properties_xml(&property_xml(&properties)?, &marker)
}

fn compare_table(
    original: &CT_Tbl,
    edited: &CT_Tbl,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if original.grid != edited.grid {
        let deleted = marked_table(original, "del", metadata)?;
        let inserted = marked_table(edited, "ins", metadata)?;
        return Ok(format!("{deleted}{inserted}"));
    }
    if original.extra_xml != edited.extra_xml
        || table_control_boundaries(original) != table_control_boundaries(edited)
    {
        return Err(Error::Other(format!(
            "comparison cannot revise table boundary structures at {location}"
        )));
    }
    let original_signatures = original
        .rows
        .iter()
        .map(|row| row_signature_with_options(row, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited
        .rows
        .iter()
        .map(|row| row_signature_with_options(row, metadata.options))
        .collect::<Vec<_>>();
    let mut source = original_source
        .map(str::to_owned)
        .map_or_else(|| table_xml(original), Ok)?;
    let properties = table_properties_xml(
        original.properties.as_ref(),
        edited.properties.as_ref(),
        location,
        metadata,
        diagnostics,
    )?;
    let property_spans = direct_word_element_spans(&source, "tblPr")?;
    match (property_spans.first(), properties.is_empty()) {
        (Some(span), false) => source.replace_range(span.clone(), &properties),
        (Some(span), true) => source.replace_range(span.clone(), ""),
        (None, false) => {
            let open = source
                .find('>')
                .ok_or_else(|| Error::Other("table XML has no start".to_owned()))?
                + 1;
            source = format!("{}{}{}", &source[..open], properties, &source[open..]);
        }
        (None, true) => {}
    }
    let row_spans = direct_word_element_spans(&source, "tr")?;
    if row_spans.len() != original.rows.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate direct table rows at {location}"
        )));
    }
    let aligned = align(&original_signatures, &edited_signatures);
    let mut output = String::new();
    let mut cursor = 0usize;
    for (alignment_index, &(left, right)) in aligned.iter().enumerate() {
        if let Some(index) = left {
            let span = &row_spans[index];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
        } else {
            let next_row_start = aligned[alignment_index + 1..]
                .iter()
                .find_map(|(next_left, _)| next_left.map(|index| row_spans[index].start))
                .unwrap_or_else(|| source.rfind("</w:tbl>").unwrap_or(source.len()));
            output.push_str(&source[cursor..next_row_start]);
            cursor = next_row_start;
        }
        match (left, right) {
            (Some(i), Some(j)) => {
                output.push_str(&compare_row(
                    &original.rows[i],
                    &edited.rows[j],
                    Some(&source[row_spans[i].clone()]),
                    &format!("{location}/row[{j}]"),
                    metadata,
                    diagnostics,
                )?);
            }
            (Some(i), None) => output.push_str(&marked_row(&original.rows[i], "del", metadata)?),
            (None, Some(j)) => output.push_str(&marked_row(&edited.rows[j], "ins", metadata)?),
            (None, None) => unreachable!(),
        }
    }
    output.push_str(&source[cursor..]);
    let spans = direct_word_element_spans(&output, "sdt")?;
    if spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate table content controls at {location}"
        )));
    }
    let mut replacements = Vec::with_capacity(spans.len());
    for (index, ((_, _, left), (_, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &replacements)
}

fn table_properties_xml(
    original: Option<&CT_TblPr>,
    edited: Option<&CT_TblPr>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if metadata.options.ignore_formatting {
        return original
            .map(table_property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }
    let original_modeled = modeled_table_properties(original);
    let edited_modeled = modeled_table_properties(edited);
    let unmodeled = |properties: &CT_TblPr| {
        (
            properties.extra_xml.clone(),
            properties.revision_xml.clone(),
        )
    };
    if original.map(unmodeled).unwrap_or_default() != edited.map(unmodeled).unwrap_or_default() {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    if original_modeled == edited_modeled {
        return original
            .map(table_property_xml)
            .transpose()
            .map(Option::unwrap_or_default);
    }
    let mut current = edited.cloned().unwrap_or_default();
    current.change = None;
    let (original_extra_xml, original_revision_xml) = original.map(unmodeled).unwrap_or_default();
    current.extra_xml = original_extra_xml;
    current.revision_xml = original_revision_xml;
    let previous = original_modeled
        .as_ref()
        .map(table_property_xml)
        .transpose()?
        .unwrap_or_else(|| "<w:tblPr/>".to_owned());
    let change = metadata.ids.revision(
        "tblPrChange",
        metadata.author,
        metadata.timestamp,
        &previous,
    )?;
    let current = table_property_xml(&current)?;
    inject_before_close(&current, "</w:tblPr>", &change)
}

fn modeled_table_properties(properties: Option<&CT_TblPr>) -> Option<CT_TblPr> {
    properties.cloned().map(|mut properties| {
        properties.change = None;
        properties.revision_xml.clear();
        properties.extra_xml.clear();
        properties
    })
}

fn table_property_xml(properties: &CT_TblPr) -> Result<String> {
    let mut bytes = Vec::new();
    properties.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn table_control_boundaries(table: &CT_Tbl) -> Vec<(usize, usize)> {
    table
        .content_controls
        .iter()
        .map(|(at, raw_before, _)| (*at, *raw_before))
        .collect()
}

fn marked_table(table: &CT_Tbl, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    let source = table_xml(table)?;
    let replacements = table
        .rows
        .iter()
        .map(|row| marked_row(row, kind, metadata))
        .collect::<Result<Vec<_>>>()?;
    let output = replace_direct_word_elements(&source, "tr", &replacements)?;
    let control_spans = direct_word_element_spans(&output, "sdt")?;
    if control_spans.len() != table.content_controls.len() {
        return Err(Error::Other(
            "comparison could not correlate table content controls while marking rows".to_owned(),
        ));
    }
    let replacements = table
        .content_controls
        .iter()
        .zip(&control_spans)
        .map(|((_, _, control), span)| {
            mark_control_owned_rows(control, kind, metadata, &output[span.clone()])
        })
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(&output, "sdt", &replacements)
}

fn mark_control_owned_rows(
    control: &CT_Sdt,
    kind: &str,
    metadata: &mut Metadata<'_>,
    original_xml: &str,
) -> Result<String> {
    let mut content = String::new();
    for child in &control.content {
        content.push_str(&match child {
            SdtContent::Paragraph(paragraph) => paragraph_xml(paragraph)?,
            SdtContent::Table(table) => marked_table(table, kind, metadata)?,
            SdtContent::Row(row) => marked_row(row, kind, metadata)?,
            SdtContent::Cell(cell) => cell_xml(cell)?,
            SdtContent::Run(run) => paragraph_owned_run_xml(run)?,
            SdtContent::ContentControl(nested) => {
                let nested_xml = control_xml(nested)?;
                mark_control_owned_rows(nested, kind, metadata, &nested_xml)?
            }
            SdtContent::RawXml(raw) => String::from_utf8(raw.clone()).map_err(utf8_error)?,
        });
    }
    replace_element_inner(original_xml, "w:sdtContent", &content)
}

fn marked_row(row: &CT_Row, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    marked_row_xml(&row_xml(row)?, kind, metadata)
}

/// Mark a serialized row as Word marks an inserted or deleted row: the row
/// marker in `w:trPr`, then every cell paragraph mark and every cell run,
/// nested tables included. Without the cell marks Word merges an adjacent
/// removed paragraph mark into the first cell and leaves the row unresolved.
fn marked_row_xml(row: &str, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    let marker = metadata
        .ids
        .marker(kind, metadata.author, metadata.timestamp)?;
    let mut xml = row.to_owned();
    let properties = direct_word_element_spans(&xml, "trPr")?;
    if let Some(properties_span) = properties.first() {
        let updated = append_word_child(&xml[properties_span.clone()], "trPr", &marker)?;
        xml.replace_range(properties_span.clone(), &updated);
    } else {
        // `w:trPr` follows an optional `w:tblPrEx`.
        let start = match direct_word_element_spans(&xml, "tblPrEx")?.first() {
            Some(exception) => exception.end,
            None => {
                xml.find('>')
                    .ok_or_else(|| Error::Other("row XML has no start".to_owned()))?
                    + 1
            }
        };
        xml.insert_str(start, &format!("<w:trPr>{marker}</w:trPr>"));
    }
    let cells = direct_word_element_spans(&xml, "tc")?
        .into_iter()
        .map(|span| marked_cell_xml(&xml[span], kind, metadata))
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(&xml, "tc", &cells)
}

fn marked_cell_xml(cell: &str, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    let paragraphs = direct_word_element_spans(cell, "p")?
        .into_iter()
        .map(|span| {
            let marker = metadata
                .ids
                .marker(kind, metadata.author, metadata.timestamp)?;
            let paragraph = marked_paragraph_xml(&cell[span], &marker)?;
            let paragraph = marked_direct_runs_xml(&paragraph, kind, metadata)?;
            let hyperlinks = direct_word_element_spans(&paragraph, "hyperlink")?
                .into_iter()
                .map(|span| marked_direct_runs_xml(&paragraph[span], kind, metadata))
                .collect::<Result<Vec<_>>>()?;
            replace_direct_word_elements(&paragraph, "hyperlink", &hyperlinks)
        })
        .collect::<Result<Vec<_>>>()?;
    let cell = replace_direct_word_elements(cell, "p", &paragraphs)?;
    let tables = direct_word_element_spans(&cell, "tbl")?
        .into_iter()
        .map(|span| {
            let table = &cell[span];
            let rows = direct_word_element_spans(table, "tr")?
                .into_iter()
                .map(|span| marked_row_xml(&table[span], kind, metadata))
                .collect::<Result<Vec<_>>>()?;
            replace_direct_word_elements(table, "tr", &rows)
        })
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(&cell, "tbl", &tables)
}

/// Wrap every direct `w:r` of `owner` in a revision, with deleted text as
/// `w:delText`.
fn marked_direct_runs_xml(owner: &str, kind: &str, metadata: &mut Metadata<'_>) -> Result<String> {
    let runs = direct_word_element_spans(owner, "r")?
        .into_iter()
        .map(|span| {
            let mut run = owner[span].to_owned();
            if kind == "del" {
                for text in direct_word_element_spans(&run, "t")?.into_iter().rev() {
                    let deleted = run[text.clone()]
                        .replacen("<w:t", "<w:delText", 1)
                        .replace("</w:t>", "</w:delText>");
                    run.replace_range(text, &deleted);
                }
            }
            metadata
                .ids
                .revision(kind, metadata.author, metadata.timestamp, &run)
        })
        .collect::<Result<Vec<_>>>()?;
    replace_direct_word_elements(owner, "r", &runs)
}

fn compare_row(
    original: &CT_Row,
    edited: &CT_Row,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if original.cells.len() != edited.cells.len()
        || row_raw_children(original) != row_raw_children(edited)
        || row_control_boundaries(original) != row_control_boundaries(edited)
    {
        return Err(Error::Other(format!(
            "comparison cannot revise row boundary structures at {location}"
        )));
    }
    if !metadata.options.ignore_formatting && row_formatting(original) != row_formatting(edited) {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    let mut output = original_source
        .map(str::to_owned)
        .map_or_else(|| row_xml(original), Ok)?;
    let cell_spans = direct_word_element_spans(&output, "tc")?;
    if cell_spans.len() != original.cells.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate row cells at {location}"
        )));
    }
    let mut cell_replacements = Vec::with_capacity(cell_spans.len());
    for (index, (left, right)) in original.cells.iter().zip(&edited.cells).enumerate() {
        cell_replacements.push(compare_cell_from_xml(
            left,
            right,
            &format!("{location}/cell[{index}]"),
            metadata,
            diagnostics,
            &output[cell_spans[index].clone()],
        )?);
    }
    output = replace_direct_word_elements(&output, "tc", &cell_replacements)?;

    let control_spans = direct_word_element_spans(&output, "sdt")?;
    if control_spans.len() != original.content_controls.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate row content controls at {location}"
        )));
    }
    let mut control_replacements = Vec::with_capacity(control_spans.len());
    for (index, ((_, _, left), (_, _, right))) in original
        .content_controls
        .iter()
        .zip(&edited.content_controls)
        .enumerate()
    {
        control_replacements.push(compare_control_from_xml(
            left,
            right,
            &format!("{location}/content-control[{index}]"),
            metadata,
            diagnostics,
            &output[control_spans[index].clone()],
        )?);
    }
    replace_direct_word_elements(&output, "sdt", &control_replacements)
}

/// Return the raw children of a row without the record of its start-tag
/// attributes, which carry producer identities and no content.
fn row_raw_children(row: &CT_Row) -> Vec<&(usize, Vec<u8>)> {
    row.extra_xml
        .iter()
        .filter(|(position, raw)| !CT_Row::raw_is_root_attributes(*position, raw))
        .collect()
}

fn row_control_boundaries(row: &CT_Row) -> Vec<(usize, usize)> {
    row.content_controls
        .iter()
        .map(|(at, raw_before, _)| (*at, *raw_before))
        .collect()
}

fn compare_cell_from_xml(
    original: &CT_Tc,
    edited: &CT_Tc,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
    original_xml: &str,
) -> Result<String> {
    if original.content.len() != edited.content.len() || original.extra_xml != edited.extra_xml {
        return Err(Error::Other(format!(
            "comparison cannot revise cell boundary structures at {location}"
        )));
    }
    if !metadata.options.ignore_formatting && original.properties != edited.properties {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    let children = cell_child_spans(original, original_xml, location)?;
    let mut replacements = Vec::with_capacity(children.len());
    for (index, ((left, right), (_, span))) in original
        .content
        .iter()
        .zip(&edited.content)
        .zip(&children)
        .enumerate()
    {
        let child_location = format!("{location}/content[{index}]");
        let source = &original_xml[span.clone()];
        replacements.push(match (left, right) {
            (CellContent::Paragraph(left), CellContent::Paragraph(right)) => compare_paragraph(
                left,
                right,
                Some(source),
                &child_location,
                metadata,
                diagnostics,
            )?,
            (CellContent::Table(left), CellContent::Table(right)) => compare_table(
                left,
                right,
                Some(source),
                &child_location,
                metadata,
                diagnostics,
            )?,
            (CellContent::ContentControl(left), CellContent::ContentControl(right)) => {
                compare_control_from_xml(
                    left,
                    right,
                    &child_location,
                    metadata,
                    diagnostics,
                    source,
                )?
            }
            _ => {
                return Err(Error::Other(format!(
                    "comparison cannot revise incompatible cell children at {child_location}"
                )));
            }
        });
    }
    replace_ranges(original_xml, &children, &replacements)
}

fn compare_control(
    original: &CT_Sdt,
    edited: &CT_Sdt,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_xml = control_xml(original)?;
    compare_control_from_xml(
        original,
        edited,
        location,
        metadata,
        diagnostics,
        &original_xml,
    )
}

fn compare_control_from_xml(
    original: &CT_Sdt,
    edited: &CT_Sdt,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
    original_xml: &str,
) -> Result<String> {
    if control_property_signature(original) != control_property_signature(edited) {
        return Err(Error::Other(format!(
            "comparison cannot revise content-control properties at {location}"
        )));
    }
    control_metadata_diagnostics(original, edited, location, diagnostics)?;
    let original_content = modeled_control_content(original);
    let edited_content = modeled_control_content(edited);
    let direct_run_or_raw =
        |content: &&SdtContent| matches!(content, SdtContent::Run(_) | SdtContent::RawXml(_));
    if uses_attributed_run_path(metadata.options)
        && original_content.iter().all(direct_run_or_raw)
        && edited_content.iter().all(direct_run_or_raw)
        && (!original_content.is_empty() || !edited_content.is_empty())
        && inline_control_raw_boundaries(&original_content, metadata.options)
            == inline_control_raw_boundaries(&edited_content, metadata.options)
    {
        let original_runs = original_content
            .iter()
            .filter_map(|content| match content {
                SdtContent::Run(run) => Some(run.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        let edited_runs = edited_content
            .iter()
            .filter_map(|content| match content {
                SdtContent::Run(run) => Some(run.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        let content_spans = direct_word_element_spans(original_xml, "sdtContent")?;
        let content_span = content_spans.first().ok_or_else(|| {
            Error::Other(format!(
                "comparison could not find inline-control content at {location}"
            ))
        })?;
        let content_source = &original_xml[content_span.clone()];
        let inner = element_inner_range_any_prefix(content_source, "sdtContent")?;
        // Raw children remain physical shell boundaries after the runs are
        // extracted. Logical boundary equality above does not preserve them
        // in the run-only alignment by itself.
        let raw_runs = |content: &[&SdtContent]| {
            let mut runs = 0;
            content
                .iter()
                .filter_map(|child| match child {
                    SdtContent::Run(_) => {
                        runs += 1;
                        None
                    }
                    SdtContent::RawXml(_) => Some(runs),
                    _ => unreachable!(),
                })
                .collect::<Vec<_>>()
        };
        let boundaries = raw_runs(&original_content)
            .into_iter()
            .zip(raw_runs(&edited_content))
            .collect::<Vec<_>>();
        let compared = compare_control_runs_with_options(
            &original_runs,
            &edited_runs,
            Some(&content_source[inner]),
            &boundaries,
            location,
            metadata,
            diagnostics,
        )?;
        return replace_element_inner(original_xml, "w:sdtContent", &compared);
    }
    let whitespace_slots = control_whitespace_slots(original)?;
    let original_signatures = original_content
        .iter()
        .map(|content| control_content_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    let edited_signatures = edited_content
        .iter()
        .map(|content| control_content_signature_with_options(content, metadata.options))
        .collect::<Vec<_>>();
    fn paragraphs<'a>(content: &[&'a SdtContent]) -> Vec<Option<&'a CT_P>> {
        content
            .iter()
            .map(|content| match content {
                SdtContent::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .collect()
    }
    let original_paragraphs = paragraphs(&original_content);
    let edited_paragraphs = paragraphs(&edited_content);
    let (aligned, carried) = replace_changed_paragraph_runs(
        pair_gaps_by_kind(
            align(&original_signatures, &edited_signatures),
            &original_signatures,
            &edited_signatures,
        ),
        &original_paragraphs,
        &edited_paragraphs,
        &original_signatures,
        &edited_signatures,
        metadata.options,
    )?;
    refuse_uncarried_field_owners(
        &aligned,
        &original_paragraphs,
        &edited_paragraphs,
        &carried,
        location,
    )?;
    let trailing_paragraph_insert_start = aligned
        .iter()
        .enumerate()
        .rev()
        .take_while(|(_, (left, right))| {
            left.is_none()
                && right
                    .and_then(|index| edited_content.get(index))
                    .is_some_and(|content| matches!(content, SdtContent::Paragraph(_)))
        })
        .map(|(position, _)| position)
        .last();
    let trailing_paragraph_delete_start = aligned
        .iter()
        .enumerate()
        .rev()
        .take_while(|(_, (left, right))| {
            right.is_none()
                && left
                    .and_then(|index| original_content.get(index))
                    .is_some_and(|content| matches!(content, SdtContent::Paragraph(_)))
        })
        .map(|(position, _)| position)
        .last();
    let content_spans = direct_word_element_spans(original_xml, "sdtContent")?;
    let content_span = content_spans.first().ok_or_else(|| {
        Error::Other(format!(
            "comparison could not find content-control content at {location}"
        ))
    })?;
    let content_source = &original_xml[content_span.clone()];
    let content_inner = element_inner_range_any_prefix(content_source, "sdtContent")?;
    let content_inner = &content_source[content_inner];
    let original_spans = story_content_spans(content_inner, "w")?;
    if original_spans.len() != original_content.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate content-control owners at {location}"
        )));
    }
    let mut content: Vec<(EmittedOwner, String)> = Vec::new();
    let mut trailing_merged = None;
    let mut whitespace_emitted = vec![false; whitespace_slots.len()];
    for (position, (left, right)) in aligned.iter().copied().enumerate() {
        let whitespace_boundary = left.or_else(|| {
            let right = right?;
            let (Some(next_left), None) = aligned.get(position + 1).copied()? else {
                return None;
            };
            (matches!(original_content[next_left], SdtContent::Paragraph(_))
                && matches!(edited_content[right], SdtContent::Table(_)))
            .then_some(next_left)
        });
        if let Some(index) = whitespace_boundary
            && !whitespace_emitted[index]
            && !whitespace_slots[index].is_empty()
        {
            content.push((EmittedOwner::Whitespace, whitespace_slots[index].clone()));
            whitespace_emitted[index] = true;
        }
        let item = |(left, right): (Option<usize>, Option<usize>)| {
            right
                .and_then(|index| edited_content.get(index).copied())
                .or_else(|| left.and_then(|index| original_content.get(index).copied()))
        };
        let next_takes_mark = next_kept_owner(&aligned, position, |pair| {
            matches!(item(pair), Some(SdtContent::Table(_)))
        })
        .is_some_and(|pair| {
            matches!(
                item(pair),
                Some(SdtContent::Paragraph(_) | SdtContent::Table(_))
            )
        });
        let child_location = format!("{location}/content[{position}]");
        match (left, right) {
            (Some(i), Some(j)) => content.push((
                EmittedOwner::of_control(original_content[i], edited_content[j], None),
                compare_control_content(
                    original_content[i],
                    edited_content[j],
                    Some(&content_inner[original_spans[i].clone()]),
                    &child_location,
                    metadata,
                    diagnostics,
                )?,
            )),
            (Some(i), None) => {
                if let SdtContent::Paragraph(paragraph) = original_content[i]
                    && trailing_paragraph_delete_start.is_some_and(|start| position >= start)
                {
                    if trailing_paragraph_delete_start == Some(position) {
                        trailing_merged = mark_previous_paragraph(&mut content, "del", metadata)?;
                    }
                    if next_takes_mark {
                        content.push((
                            EmittedOwner::paragraph(paragraph, paragraph),
                            marked_control_content(original_content[i], "del", metadata)?,
                        ));
                    } else {
                        content.push(final_paragraph(
                            paragraph,
                            "del",
                            None,
                            trailing_merged,
                            metadata,
                        )?);
                    }
                } else if let SdtContent::Paragraph(paragraph) = original_content[i]
                    && !next_takes_mark
                {
                    let merged = mark_previous_paragraph(&mut content, "del", metadata)?;
                    content.push(final_paragraph(paragraph, "del", None, merged, metadata)?);
                } else {
                    content.push((
                        EmittedOwner::of_control(
                            original_content[i],
                            original_content[i],
                            Some("del"),
                        ),
                        marked_control_content(original_content[i], "del", metadata)?,
                    ));
                }
            }
            (None, Some(j)) => {
                if let SdtContent::Paragraph(paragraph) = edited_content[j]
                    && trailing_paragraph_insert_start.is_some_and(|start| position >= start)
                {
                    if trailing_paragraph_insert_start == Some(position) {
                        trailing_merged = mark_previous_paragraph(&mut content, "ins", metadata)?;
                    }
                    if next_takes_mark {
                        content.push((
                            EmittedOwner::paragraph(paragraph, paragraph),
                            marked_control_content(edited_content[j], "ins", metadata)?,
                        ));
                    } else {
                        content.push(final_paragraph(
                            paragraph,
                            "ins",
                            None,
                            trailing_merged,
                            metadata,
                        )?);
                    }
                } else if let SdtContent::Paragraph(paragraph) = edited_content[j]
                    && !next_takes_mark
                {
                    let merged = mark_previous_paragraph(&mut content, "ins", metadata)?;
                    content.push(final_paragraph(paragraph, "ins", None, merged, metadata)?);
                } else {
                    content.push((
                        EmittedOwner::of_control(edited_content[j], edited_content[j], Some("ins")),
                        marked_control_content(edited_content[j], "ins", metadata)?,
                    ));
                }
            }
            (None, None) => unreachable!(),
        }
    }
    if let Some(trailing) = whitespace_slots.last()
        && !trailing.is_empty()
    {
        content.push((EmittedOwner::Whitespace, trailing.clone()));
    }
    let content = content.into_iter().map(|(_, xml)| xml).collect::<String>();
    replace_element_inner(original_xml, "w:sdtContent", &content)
}

fn inline_control_raw_boundaries(
    content: &[&SdtContent],
    options: &ComparisonOptions,
) -> Vec<(usize, Vec<u8>)> {
    let mut runs = Vec::new();
    let mut raw = Vec::new();
    for child in content {
        match child {
            SdtContent::Run(run) => runs.push((*run).clone()),
            SdtContent::RawXml(bytes) => {
                let boundary = attributed_run_units(&runs, options)
                    .into_iter()
                    .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
                    .count();
                raw.push((boundary, bytes.clone()));
            }
            _ => unreachable!(),
        }
    }
    raw
}

fn compare_control_content(
    original: &SdtContent,
    edited: &SdtContent,
    original_source: Option<&str>,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    match (original, edited) {
        (SdtContent::Paragraph(left), SdtContent::Paragraph(right)) => compare_paragraph(
            left,
            right,
            original_source,
            location,
            metadata,
            diagnostics,
        ),
        (SdtContent::Table(left), SdtContent::Table(right)) => compare_table(
            left,
            right,
            original_source,
            location,
            metadata,
            diagnostics,
        ),
        (SdtContent::ContentControl(left), SdtContent::ContentControl(right)) => {
            if let Some(source) = original_source {
                compare_control_from_xml(left, right, location, metadata, diagnostics, source)
            } else {
                compare_control(left, right, location, metadata, diagnostics)
            }
        }
        (SdtContent::Row(left), SdtContent::Row(right)) => compare_row(
            left,
            right,
            original_source,
            location,
            metadata,
            diagnostics,
        ),
        (SdtContent::Cell(left), SdtContent::Cell(right)) => compare_cell_from_xml(
            left,
            right,
            location,
            metadata,
            diagnostics,
            &original_source
                .map(str::to_owned)
                .map_or_else(|| cell_xml(left), Ok)?,
        ),
        (SdtContent::Run(left), SdtContent::Run(right)) => {
            if uses_attributed_run_path(metadata.options) {
                compare_control_runs_with_options(
                    std::slice::from_ref(left),
                    std::slice::from_ref(right),
                    None,
                    &[],
                    location,
                    metadata,
                    diagnostics,
                )
            } else if run_signature(left) == run_signature(right) {
                if !metadata.options.ignore_formatting && left.properties != right.properties {
                    formatting_diagnostic(diagnostics, location.to_owned());
                }
                if left == right
                    && let Some(source) = original_source
                {
                    Ok(source.to_owned())
                } else {
                    run_xml(left)
                }
            } else {
                let deleted = metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted_run_xml(left)?,
                )?;
                let inserted = metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &run_xml(right)?,
                )?;
                Ok(format!("{deleted}{inserted}"))
            }
        }
        (SdtContent::RawXml(left), SdtContent::RawXml(right)) if left == right => {
            String::from_utf8(left.clone()).map_err(utf8_error)
        }
        _ => Err(Error::Other(format!(
            "comparison cannot revise incompatible content-control children at {location}"
        ))),
    }
}

fn compare_control_runs_with_options(
    original: &[CT_R],
    edited: &[CT_R],
    original_source: Option<&str>,
    raw_boundaries: &[(usize, usize)],
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    let original_units = attributed_run_units(original, metadata.options);
    let edited_units = attributed_run_units(edited, metadata.options);
    let original_signatures = original_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();
    let edited_signatures = edited_units
        .iter()
        .map(attributed_unit_signature)
        .collect::<Vec<_>>();
    let cuts = if metadata.options.granularity == ComparisonGranularity::Run {
        Vec::new()
    } else {
        shell_unit_cuts(
            &original_units,
            &edited_units,
            raw_boundaries.iter().copied(),
        )
        .ok_or_else(|| {
            Error::Other(format!(
                "comparison cannot revise inline-control raw boundaries at {location}"
            ))
        })?
    };
    let (aligned, segments) = align_between_shells(&original_signatures, &edited_signatures, &cuts);
    let segment_runs = std::iter::once(0)
        .chain(cuts.iter().map(|&(cut, _)| {
            original_units
                .get(cut)
                .map_or(original.len(), |unit| unit.owner)
        }))
        .collect::<Vec<_>>();
    if !metadata.options.ignore_fields {
        validate_field_alignment(
            &aligned,
            &original_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &edited_units
                .iter()
                .map(|unit| unit.run.clone())
                .collect::<Vec<_>>(),
            &original_signatures,
            &edited_signatures,
            location,
        )?;
    }
    let (aligned, segments) = grouped_text_alignment(
        &aligned,
        &segments,
        &original_units,
        &edited_units,
        &original_signatures,
        &edited_signatures,
        metadata.options,
    );
    let grouped = coalesced_granular_alignment(
        &aligned,
        &original_units,
        &edited_units,
        &original_signatures,
        &edited_signatures,
    );
    let mut grouped_alignment = Vec::with_capacity(grouped.len());
    let mut grouped_segment_runs = Vec::with_capacity(grouped.len());
    let mut replacements = Vec::with_capacity(grouped.len());
    for (action, members) in grouped {
        grouped_alignment.push(aligned[members.start]);
        grouped_segment_runs.push(segment_runs[segments[members.start]]);
        let left_indices = members
            .clone()
            .filter_map(|index| aligned[index].0)
            .collect::<Vec<_>>();
        let right_indices = members
            .filter_map(|index| aligned[index].1)
            .collect::<Vec<_>>();
        let left = merge_unit_runs(&original_units, &left_indices);
        let right = merge_unit_runs(&edited_units, &right_indices);
        replacements.push(match action {
            GranularAction::Preserve => paragraph_owned_run_xml(left.as_ref().unwrap())?,
            GranularAction::Equal => compared_run_xml(
                left.as_ref().unwrap(),
                right.as_ref().unwrap(),
                location,
                metadata,
                diagnostics,
            )?,
            GranularAction::Replace => {
                let left = left.as_ref().unwrap();
                let deleted = metadata.ids.revision(
                    "del",
                    metadata.author,
                    metadata.timestamp,
                    &deleted_run_xml(left)?,
                )?;
                let inserted = metadata.ids.revision(
                    "ins",
                    metadata.author,
                    metadata.timestamp,
                    &policy_inserted_run_xml(
                        right.as_ref().unwrap(),
                        Some(left),
                        metadata.options,
                    )?,
                )?;
                format!("{deleted}{inserted}")
            }
            GranularAction::Delete => metadata.ids.revision(
                "del",
                metadata.author,
                metadata.timestamp,
                &deleted_run_xml(left.as_ref().unwrap())?,
            )?,
            GranularAction::Insert => metadata.ids.revision(
                "ins",
                metadata.author,
                metadata.timestamp,
                &paragraph_owned_run_xml(right.as_ref().unwrap())?,
            )?,
            GranularAction::Drop => String::new(),
        });
    }
    let Some(source) = original_source else {
        return Ok(replacements.concat());
    };
    let open = "<w:inlineControl>";
    let wrapped = format!("{open}{source}</w:inlineControl>");
    let spans = direct_word_element_spans(&wrapped, "r")?
        .into_iter()
        .map(|span| span.start - open.len()..span.end - open.len())
        .collect::<Vec<_>>();
    if spans.len() != original.len() {
        return Err(Error::Other(format!(
            "comparison could not correlate inline-control run owners at {location}"
        )));
    }
    let insertion_boundary = spans.first().map_or(source.len(), |span| span.start);
    let mut output = source[..insertion_boundary].to_owned();
    let mut cursor = insertion_boundary;
    let mut consumed_owner = None;
    for (((left, _), replacement), &segment_run) in grouped_alignment
        .iter()
        .zip(&replacements)
        .zip(&grouped_segment_runs)
    {
        // An insertion can be the first action after a raw child, with no
        // original unit to consume. Copy that shell before its new text.
        let segment_start = spans
            .get(segment_run)
            .map_or(source.len(), |span| span.start);
        if segment_start > cursor {
            output.push_str(&source[cursor..segment_start]);
            cursor = segment_start;
        }
        if let Some(unit) = left.map(|index| &original_units[index])
            && consumed_owner != Some(unit.owner)
        {
            let span = &spans[unit.owner];
            output.push_str(&source[cursor..span.start]);
            cursor = span.end;
            consumed_owner = Some(unit.owner);
        }
        output.push_str(replacement);
    }
    output.push_str(&source[cursor..]);
    Ok(output)
}

fn marked_control_content(
    content: &SdtContent,
    kind: &str,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    match content {
        SdtContent::Paragraph(paragraph) if kind == "del" => deleted_paragraph(paragraph, metadata),
        SdtContent::Paragraph(paragraph) => inserted_paragraph(paragraph, metadata),
        SdtContent::Table(table) => marked_table(table, kind, metadata),
        SdtContent::Row(row) => marked_row(row, kind, metadata),
        SdtContent::Run(run) if kind == "del" => metadata.ids.revision(
            kind,
            metadata.author,
            metadata.timestamp,
            &deleted_run_xml(run)?,
        ),
        SdtContent::Run(run) => {
            metadata
                .ids
                .revision(kind, metadata.author, metadata.timestamp, &run_xml(run)?)
        }
        SdtContent::RawXml(raw) => String::from_utf8(raw.clone()).map_err(utf8_error),
        SdtContent::Cell(_) | SdtContent::ContentControl(_) => Err(Error::Other(
            "comparison cannot add or remove an unmatched cell or nested content control"
                .to_owned(),
        )),
    }
}

fn compared_run_xml(
    original: &CT_R,
    edited: &CT_R,
    location: &str,
    metadata: &mut Metadata<'_>,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<String> {
    if metadata.options.ignore_fields && (run_is_field(original) || run_is_field(edited)) {
        return paragraph_owned_run_xml(original);
    }
    if metadata.options.ignore_formatting {
        return paragraph_owned_run_xml(original);
    }
    if let (Some(original_field), Some(edited_field)) = (run_field(original), run_field(edited)) {
        return compare_field_xml(original_field, edited_field, metadata);
    }
    let original_properties = modeled_run_properties(original);
    let edited_properties = modeled_run_properties(edited);
    if original_properties == edited_properties {
        if original.properties != edited.properties {
            formatting_diagnostic(diagnostics, location.to_owned());
        }
        return paragraph_owned_run_xml(original);
    }
    let previous = run_property_xml(original_properties.as_ref())?;
    let change =
        metadata
            .ids
            .revision("rPrChange", metadata.author, metadata.timestamp, &previous)?;
    if unmodeled_run_properties_differ(original, edited) {
        formatting_diagnostic(diagnostics, location.to_owned());
    }
    let mut current_run = edited.clone();
    preserve_unmodeled_run_properties(&mut current_run, original);
    let mut current = run_xml(&current_run)?;
    let properties = direct_word_element_spans(&current, "rPr")?;
    if let Some(span) = properties.first() {
        let updated = append_word_child(&current[span.clone()], "rPr", &change)?;
        current.replace_range(span.clone(), &updated);
    } else {
        let open = current
            .find('>')
            .ok_or_else(|| Error::Other("run XML has no start".to_owned()))?
            + 1;
        current = format!(
            "{}<w:rPr>{change}</w:rPr>{}",
            &current[..open],
            &current[open..]
        );
    }
    Ok(current)
}

fn compare_field_xml(
    original: &rdocx_oxml::text::Field,
    edited: &rdocx_oxml::text::Field,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let original_xml = field_xml(original)?;
    let edited_xml = field_xml(edited)?;
    if original == edited {
        return Ok(original_xml);
    }
    let same_instruction = original.effective_instruction() == edited.effective_instruction();
    if same_instruction && original.is_complex() == edited.is_complex() {
        if original.is_complex() {
            let (before, old_result, after) = complex_field_result(&original_xml)?;
            let (_, new_result, _) = complex_field_result(&edited_xml)?;
            return Ok(format!(
                "{before}{}{after}",
                tracked_field_result(&old_result, &new_result, metadata)?
            ));
        }
        let original_inner = element_inner_range(&original_xml, "fldSimple")?;
        let edited_inner = element_inner_range(&edited_xml, "fldSimple")?;
        let old_result = original_xml[original_inner.clone()].to_owned();
        let new_result = edited_xml[edited_inner].to_owned();
        let mut tracked = original_xml;
        tracked.replace_range(
            original_inner,
            &tracked_field_result(&old_result, &new_result, metadata)?,
        );
        return Ok(tracked);
    }

    let deleted = metadata.ids.revision(
        "del",
        metadata.author,
        metadata.timestamp,
        &deleted_text_xml(&original_xml),
    )?;
    let inserted =
        metadata
            .ids
            .revision("ins", metadata.author, metadata.timestamp, &edited_xml)?;
    Ok(format!("{deleted}{inserted}"))
}

fn element_inner_range(xml: &str, local: &str) -> Result<Range<usize>> {
    let start = xml
        .find('>')
        .map(|at| at + 1)
        .ok_or_else(|| Error::Other(format!("{local} source has no start tag")))?;
    let end = xml
        .rfind(&format!("</w:{local}>"))
        .ok_or_else(|| Error::Other(format!("{local} source has no end tag")))?;
    Ok(start..end)
}

fn field_xml(field: &rdocx_oxml::text::Field) -> Result<String> {
    paragraph_owned_run_xml(&CT_R {
        properties: None,
        content: vec![RunContent::Field(field.clone())],
        extra_xml: Vec::new(),
        extra_xml_positions: Vec::new(),
        alt_drawings: Vec::new(),
    })
}

fn tracked_field_result(
    original: &str,
    edited: &str,
    metadata: &mut Metadata<'_>,
) -> Result<String> {
    let deleted = metadata.ids.revision(
        "del",
        metadata.author,
        metadata.timestamp,
        &deleted_text_xml(original),
    )?;
    let inserted = metadata
        .ids
        .revision("ins", metadata.author, metadata.timestamp, edited)?;
    Ok(format!("{deleted}{inserted}"))
}

/// Rename each `w:t` element to `w:delText` and each `w:instrText` to
/// `w:delInstrText`, and no other element whose name starts the same way,
/// such as `w:tab`. Word refuses to open a deletion that holds `w:instrText`.
fn deleted_text_xml(xml: &str) -> String {
    renamed_word_elements(
        &renamed_word_elements(xml, "w:t", "w:delText"),
        "w:instrText",
        "w:delInstrText",
    )
}

/// Rename each `from` element to `to`, and no other element whose name
/// starts the same way.
fn renamed_word_elements(xml: &str, from: &str, to: &str) -> String {
    let mut output = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(at) = rest.find(from) {
        let (before, after) = rest.split_at(at);
        let after = &after[from.len()..];
        let is_tag = (before.ends_with('<') || before.ends_with("</"))
            && after.starts_with(|next: char| matches!(next, '>' | '/') || next.is_whitespace());
        output.push_str(before);
        output.push_str(if is_tag { to } else { from });
        rest = after;
    }
    output.push_str(rest);
    output
}

fn complex_field_result(xml: &str) -> Result<(String, String, String)> {
    let separate = field_character(xml, 0, "separate")
        .ok_or_else(|| Error::Other("complex field source has no separate boundary".to_owned()))?;
    // A producer may pack a whole field in one run, as Google Docs writes page
    // fields. Ending a run after `separate` and starting one at `end` reads it
    // as the same field written one run per part.
    let xml = split_field_run(xml, separate, true)?;
    let (_, result_start) = containing_run(&xml, separate)?;
    let end_marker = field_character(&xml, result_start, "end")
        .ok_or_else(|| Error::Other("complex field source has no end boundary".to_owned()))?;
    let xml = split_field_run(&xml, end_marker, false)?;
    // The split may have moved the `end` character further along.
    let end_marker = field_character(&xml, result_start, "end")
        .ok_or_else(|| Error::Other("complex field source has no end boundary".to_owned()))?;
    let (result_end, _) = containing_run(&xml, end_marker)?;
    Ok((
        xml[..result_start].to_owned(),
        xml[result_start..result_end].to_owned(),
        xml[result_end..].to_owned(),
    ))
}

fn field_character(xml: &str, from: usize, kind: &str) -> Option<usize> {
    let tail = &xml[from..];
    tail.find(&format!("fldCharType=\"{kind}\""))
        .or_else(|| tail.find(&format!("fldCharType='{kind}'")))
        .map(|offset| from + offset)
}

/// Split the run holding the field character at `marker` so that the
/// character ends its run (`after`) or starts it.
///
/// Both runs repeat the original start tag and run properties. A run with no
/// content on that side of the character is returned unchanged.
fn split_field_run(xml: &str, marker: usize, after: bool) -> Result<String> {
    let (run_start, run_end) = containing_run(xml, marker)?;
    let run = &xml[run_start..run_end];
    let children = direct_element_spans(run)?;
    let character = children
        .iter()
        .position(|child| child.contains(&(marker - run_start)))
        .ok_or_else(|| Error::Other("complex field character is not a run child".to_owned()))?;
    let is_properties = |child: &Range<usize>| {
        let mut reader = Reader::from_reader(run[child.clone()].as_bytes());
        matches!(
            reader.read_event(),
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"rPr"
        )
    };
    let first_content = usize::from(children.first().is_some_and(is_properties));
    let split = if after { character + 1 } else { character };
    if split <= first_content || split >= children.len() {
        return Ok(xml.to_owned());
    }
    let head = &run[..children[first_content].start];
    let close = run
        .rfind("</")
        .map(|at| &run[at..])
        .ok_or_else(|| Error::Other("complex field run has no end tag".to_owned()))?;
    let at = run_start + children[split].start;
    Ok(format!("{}{close}{head}{}", &xml[..at], &xml[at..]))
}

fn containing_run(xml: &str, at: usize) -> Result<(usize, usize)> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<(usize, bool)>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                stack.push((before, element.local_name().as_ref() == b"r"));
            }
            Event::End(_) => {
                let Some((start, run)) = stack.pop() else {
                    return Err(Error::Other("complex field XML is unbalanced".to_owned()));
                };
                if run && start <= at && at < after {
                    return Ok((start, after));
                }
            }
            Event::Empty(_) => {}
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Err(Error::Other(
        "complex field boundary has no owning run".to_owned(),
    ))
}

fn modeled_run_properties(run: &CT_R) -> Option<rdocx_oxml::properties::CT_RPr> {
    run.properties.clone().map(|mut properties| {
        properties.revision_markers.clear();
        properties.change = None;
        properties.revision_xml.clear();
        properties.revision_xml_positions.clear();
        properties.language_extra_attributes.clear();
        properties
    })
}

fn unmodeled_run_properties_differ(original: &CT_R, edited: &CT_R) -> bool {
    let original = original.properties.as_ref();
    let edited = edited.properties.as_ref();
    original
        .map(|properties| properties.language_extra_attributes.as_slice())
        .unwrap_or_default()
        != edited
            .map(|properties| properties.language_extra_attributes.as_slice())
            .unwrap_or_default()
        || original
            .map(|properties| properties.revision_xml.as_slice())
            .unwrap_or_default()
            != edited
                .map(|properties| properties.revision_xml.as_slice())
                .unwrap_or_default()
        || original
            .map(|properties| properties.revision_xml_positions.as_slice())
            .unwrap_or_default()
            != edited
                .map(|properties| properties.revision_xml_positions.as_slice())
                .unwrap_or_default()
}

fn preserve_unmodeled_run_properties(current: &mut CT_R, original: &CT_R) {
    let current_properties = current
        .properties
        .get_or_insert_with(rdocx_oxml::properties::CT_RPr::default);
    if let Some(original_properties) = original.properties.as_ref() {
        current_properties.language_extra_attributes =
            original_properties.language_extra_attributes.clone();
        current_properties.revision_xml = original_properties.revision_xml.clone();
        current_properties.revision_xml_positions =
            original_properties.revision_xml_positions.clone();
    } else {
        current_properties.language_extra_attributes.clear();
        current_properties.revision_xml.clear();
        current_properties.revision_xml_positions.clear();
    }
}

fn run_property_xml(properties: Option<&rdocx_oxml::properties::CT_RPr>) -> Result<String> {
    let mut bytes = Vec::new();
    if let Some(properties) = properties {
        properties.to_xml(&mut Writer::new(&mut bytes))?;
    }
    if bytes.is_empty() {
        Ok("<w:rPr/>".to_owned())
    } else {
        String::from_utf8(bytes).map_err(utf8_error)
    }
}

fn formatting_diagnostic(diagnostics: &mut Vec<ComparisonDiagnostic>, location: String) {
    diagnostics.push(ComparisonDiagnostic {
        location,
        message: "formatting differs and the original formatting was retained".to_owned(),
    });
}

fn body_location(location: &str, content: &BodyContent, index: usize) -> String {
    match content {
        BodyContent::Paragraph(_) => format!("{location}/paragraph[{index}]"),
        BodyContent::Table(_) => format!("{location}/table[{index}]"),
        BodyContent::ContentControl(_) => format!("{location}/content-control[{index}]"),
        BodyContent::RawXml(_) => format!("{location}/raw[{index}]"),
    }
}

fn align(original: &[String], edited: &[String]) -> Vec<(Option<usize>, Option<usize>)> {
    let mut lcs = vec![vec![0usize; edited.len() + 1]; original.len() + 1];
    for i in (0..original.len()).rev() {
        for j in (0..edited.len()).rev() {
            lcs[i][j] = if original[i] == edited[j] {
                1 + lcs[i + 1][j + 1]
            } else {
                lcs[i + 1][j].max(lcs[i][j + 1])
            };
        }
    }
    let mut matches = Vec::new();
    let (mut i, mut j) = (0usize, 0usize);
    while i < original.len() && j < edited.len() {
        if original[i] == edited[j] {
            matches.push((i, j));
            i += 1;
            j += 1;
        } else if lcs[i + 1][j] >= lcs[i][j + 1] {
            i += 1;
        } else {
            j += 1;
        }
    }
    let mut result = Vec::new();
    let (mut left, mut right) = (0usize, 0usize);
    for (matched_left, matched_right) in matches
        .into_iter()
        .chain(std::iter::once((original.len(), edited.len())))
    {
        let paired = (matched_left - left).min(matched_right - right);
        for offset in 0..paired {
            result.push((Some(left + offset), Some(right + offset)));
        }
        for index in left + paired..matched_left {
            result.push((Some(index), None));
        }
        for index in right + paired..matched_right {
            result.push((None, Some(index)));
        }
        if matched_left < original.len() {
            result.push((Some(matched_left), Some(matched_right)));
        }
        left = matched_left.saturating_add(1);
        right = matched_right.saturating_add(1);
    }
    result
}

/// Pair the owners left over in each gap between matched owners by kind, in
/// order: a paragraph with a paragraph, a table with a table and a control
/// with a control, as the prefix of their signatures names the kind. The
/// others stay deleted, then inserted, before the next pair. [`align`] pairs
/// them by position, which would delete and insert a changed table whole
/// beside a deleted paragraph.
fn pair_gaps_by_kind(aligned: Alignment, original: &[String], edited: &[String]) -> Alignment {
    fn kind(signature: &str) -> &str {
        signature
            .split_once(':')
            .map_or(signature, |(kind, _)| kind)
    }
    fn flush(
        gap: &mut (Vec<usize>, Vec<usize>),
        original: &[String],
        edited: &[String],
        result: &mut Alignment,
    ) {
        let (left, right) = std::mem::take(gap);
        let mut lcs = vec![vec![0usize; right.len() + 1]; left.len() + 1];
        for i in (0..left.len()).rev() {
            for j in (0..right.len()).rev() {
                lcs[i][j] = if kind(&original[left[i]]) == kind(&edited[right[j]]) {
                    1 + lcs[i + 1][j + 1]
                } else {
                    lcs[i + 1][j].max(lcs[i][j + 1])
                };
            }
        }
        let mut pairs = Vec::new();
        let (mut i, mut j) = (0usize, 0usize);
        while i < left.len() && j < right.len() {
            if kind(&original[left[i]]) == kind(&edited[right[j]]) {
                pairs.push((i, j));
                i += 1;
                j += 1;
            } else if lcs[i + 1][j] >= lcs[i][j + 1] {
                i += 1;
            } else {
                j += 1;
            }
        }
        let (mut deleted, mut inserted) = (0usize, 0usize);
        for (i, j) in pairs
            .into_iter()
            .chain(std::iter::once((left.len(), right.len())))
        {
            result.extend(left[deleted..i].iter().map(|&index| (Some(index), None)));
            result.extend(right[inserted..j].iter().map(|&index| (None, Some(index))));
            if i < left.len() {
                result.push((Some(left[i]), Some(right[j])));
            }
            (deleted, inserted) = (i + 1, j + 1);
        }
    }
    let mut result = Vec::with_capacity(aligned.len());
    let mut gap = (Vec::new(), Vec::new());
    for pair in aligned {
        match pair {
            (Some(i), Some(j)) if original[i] == edited[j] => {
                flush(&mut gap, original, edited, &mut result);
                result.push(pair);
            }
            (left, right) => {
                gap.0.extend(left);
                gap.1.extend(right);
            }
        }
    }
    flush(&mut gap, original, edited, &mut result);
    result
}

/// Replace each changed run of paragraphs that revising in place cannot
/// express, and name the paragraphs whose hyperlinks and simple fields a
/// whole-paragraph revision may carry.
///
/// A run is the consecutive changed entries of paragraphs only, between
/// unchanged owners, tables or controls. When one of its pairs differs in
/// its inline boundary structures, or gains or loses a modeled field, or
/// when it holds a hyperlink or simple field, the run becomes all its
/// original paragraphs deleted, then all its edited paragraphs inserted. The
/// run grows over its neighbouring paragraphs until each side holds every
/// complex field it begins or ends, so a table of contents whose end sits
/// in an unchanged paragraph is still deleted and inserted whole. Word then
/// removes the whole field on each side, with the hyperlinks and simple
/// fields inside it.
///
/// A run keeps its pairs, and their refusal, when a side cannot be closed
/// that way, when one of its hyperlinks or simple fields lies outside such a
/// field, or when both sides share a bookmark or comment range, which the
/// replacement would hold twice.
fn replace_changed_paragraph_runs(
    aligned: Vec<(Option<usize>, Option<usize>)>,
    original: &[Option<&CT_P>],
    edited: &[Option<&CT_P>],
    original_signatures: &[String],
    edited_signatures: &[String],
    options: &ComparisonOptions,
) -> Result<(Alignment, CarriedParagraphs)> {
    fn markers<'a>(paragraphs: impl Iterator<Item = &'a CT_P>) -> HashSet<String> {
        let mut markers = HashSet::new();
        for paragraph in paragraphs {
            for marker in &paragraph.bookmark_markers {
                markers.extend(marker.id().map(|id| format!("bookmark id {id}")));
                markers.extend(marker.name().map(|name| format!("bookmark {name}")));
            }
            for marker in &paragraph.comment_ranges {
                let (CommentRangeMarker::Start { id, .. } | CommentRangeMarker::End { id, .. }) =
                    marker;
                markers.insert(format!("comment {id}"));
            }
        }
        markers
    }
    let changed = |(left, right): &(Option<usize>, Option<usize>)| match (left, right) {
        (Some(i), Some(j)) => original_signatures[*i] != edited_signatures[*j],
        _ => true,
    };
    // In-place revision refuses a modeled field it cannot pair with one on
    // the other side.
    let fields = |paragraph: &CT_P| {
        paragraph
            .runs
            .iter()
            .filter(|run| run_is_field(run))
            .count()
    };
    let paragraphs_only = |(left, right): &(Option<usize>, Option<usize>)| {
        left.is_none_or(|i| original[i].is_some()) && right.is_none_or(|j| edited[j].is_some())
    };
    let sides = |range: &Range<usize>| {
        let entries = &aligned[range.clone()];
        (
            entries
                .iter()
                .filter_map(|(left, _)| left.and_then(|i| original[i]))
                .collect::<Vec<_>>(),
            entries
                .iter()
                .filter_map(|(_, right)| right.and_then(|j| edited[j]))
                .collect::<Vec<_>>(),
        )
    };
    // Grow a run until both sides hold their complex fields whole.
    let close = |mut range: Range<usize>| -> Result<Option<Range<usize>>> {
        loop {
            let (originals, edits) = sides(&range);
            let balances = [field_balance(&originals)?, field_balance(&edits)?];
            if balances.contains(&FieldBalance::Closes) {
                if range.start == 0 || !paragraphs_only(&aligned[range.start - 1]) {
                    return Ok(None);
                }
                range.start -= 1;
            } else if balances.contains(&FieldBalance::Opens) {
                if range.end == aligned.len() || !paragraphs_only(&aligned[range.end]) {
                    return Ok(None);
                }
                range.end += 1;
            } else if balances.contains(&FieldBalance::Uncarried)
                || !markers(originals.into_iter()).is_disjoint(&markers(edits.into_iter()))
            {
                return Ok(None);
            } else {
                return Ok(Some(range));
            }
        }
    };
    let mut accepted: Vec<Range<usize>> = Vec::new();
    let mut start = 0;
    for run in aligned.chunk_by(|first, second| {
        (changed(first), paragraphs_only(first)) == (changed(second), paragraphs_only(second))
    }) {
        let range = start..start + run.len();
        start = range.end;
        if !changed(&run[0]) || !paragraphs_only(&run[0]) {
            continue;
        }
        let inexpressible = run.iter().any(|(left, right)| {
            matches!(
                (left.and_then(|i| original[i]), right.and_then(|j| edited[j])),
                (Some(left), Some(right)) if paragraph_boundaries_differ(left, right, options)
                    || (!options.ignore_fields && fields(left) != fields(right))
            )
        });
        let (originals, edits) = sides(&range);
        let mut owners = false;
        for paragraph in originals.iter().chain(&edits) {
            owners |= paragraph_field_events(paragraph)?.contains(&FieldEvent::Owner);
        }
        if !inexpressible && !owners {
            continue;
        }
        let Some(mut range) = close(range)? else {
            continue;
        };
        // A grown run that reaches an accepted one replaces both together.
        while let Some(last) = accepted.last()
            && last.end > range.start
        {
            let merged = last.start.min(range.start)..last.end.max(range.end);
            match close(merged)? {
                Some(merged) => {
                    accepted.pop();
                    range = merged;
                }
                None => break,
            }
        }
        if accepted.last().is_none_or(|last| last.end <= range.start) {
            accepted.push(range);
        }
    }
    let mut replaced = Vec::with_capacity(aligned.len());
    let mut carried = CarriedParagraphs::default();
    let mut ranges = accepted.into_iter().peekable();
    let mut index = 0;
    while index < aligned.len() {
        let Some(range) = ranges.next_if(|range| range.start == index) else {
            replaced.push(aligned[index]);
            index += 1;
            continue;
        };
        let entries = &aligned[range.clone()];
        for (left, _) in entries {
            if let Some(i) = *left {
                carried.original.insert(i);
                replaced.push((Some(i), None));
            }
        }
        for (_, right) in entries {
            if let Some(j) = *right {
                carried.edited.insert(j);
                replaced.push((None, Some(j)));
            }
        }
        index = range.end;
    }
    Ok((replaced, carried))
}

/// Aligned original and edited owner indices.
type Alignment = Vec<(Option<usize>, Option<usize>)>;

/// The paragraphs, by index on each side, that a replaced run deletes or
/// inserts whole inside the complex fields that enclose their hyperlinks
/// and simple fields.
#[derive(Default)]
struct CarriedParagraphs {
    original: HashSet<usize>,
    edited: HashSet<usize>,
}

/// Refuse a whole deleted, inserted or moved paragraph that holds a simple
/// field outside a complex field carried with it.
///
/// A simple field may not sit inside a revision wrapper. Its own children
/// could, as a hyperlink's do, but rdocx reads a simple field as one field
/// run and resolves no revision inside it.
fn refuse_uncarried_field_owners(
    aligned: &[(Option<usize>, Option<usize>)],
    original: &[Option<&CT_P>],
    edited: &[Option<&CT_P>],
    carried: &CarriedParagraphs,
    location: &str,
) -> Result<()> {
    for (left, right) in aligned {
        let paragraph = match (left, right) {
            (Some(i), None) if !carried.original.contains(i) => original[*i],
            (None, Some(j)) if !carried.edited.contains(j) => edited[*j],
            _ => None,
        };
        if let Some(paragraph) = paragraph
            && paragraph_wraps_simple_field(paragraph)?
        {
            return Err(Error::Other(format!(
                "comparison cannot track a whole paragraph's simple field outside a field it replaces at {location}"
            )));
        }
    }
    Ok(())
}

/// Whether a whole-paragraph revision would wrap a simple field: one that is
/// a child of the paragraph or of a hyperlink there, whose children the
/// wrappers go inside.
fn paragraph_wraps_simple_field(paragraph: &CT_P) -> Result<bool> {
    let xml = paragraph_xml(paragraph)?;
    let word_starts = word_element_starts(&xml)?;
    let mut reader = Reader::from_reader(xml.as_bytes());
    // For each open element, whether the wrappers go inside it.
    let mut open: Vec<bool> = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let word = word_starts.contains(&(reader.buffer_position() as usize));
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let (element, empty) = match &event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                open.pop();
                buffer.clear();
                continue;
            }
            Event::Eof => return Ok(false),
            _ => {
                buffer.clear();
                continue;
            }
        };
        let local = element.local_name();
        let parent_wrapped = open.last().copied();
        if word && parent_wrapped == Some(true) && local.as_ref() == b"fldSimple" {
            return Ok(true);
        }
        if !empty {
            open.push(
                parent_wrapped
                    .is_none_or(|wrapped| wrapped && word && local.as_ref() == b"hyperlink"),
            );
        }
        buffer.clear();
    }
}

/// A complex field character or a field owner of a paragraph.
#[derive(Clone, Copy, PartialEq)]
enum FieldEvent {
    Begin,
    End,
    /// A hyperlink or simple field, which a revision wrapper may not hold.
    Owner,
}

/// The complex field characters and field owners of a paragraph in
/// document order.
fn paragraph_field_events(paragraph: &CT_P) -> Result<Vec<FieldEvent>> {
    let xml = paragraph_xml(paragraph)?;
    let mut reader = Reader::from_reader(xml.as_bytes());
    let mut events = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let (element, empty) = match &event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                buffer.clear();
                continue;
            }
            Event::Eof => break,
            _ => {
                buffer.clear();
                continue;
            }
        };
        match element.local_name().as_ref() {
            b"hyperlink" | b"fldSimple" if depth == 1 => events.push(FieldEvent::Owner),
            b"fldChar" => {
                let kind = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.local_name().as_ref() == b"fldCharType")
                    .map(|attribute| attribute.value.into_owned());
                match kind.as_deref() {
                    Some(b"begin") => events.push(FieldEvent::Begin),
                    Some(b"end") => events.push(FieldEvent::End),
                    _ => {}
                }
            }
            _ => {}
        }
        if !empty {
            depth += 1;
        }
        buffer.clear();
    }
    Ok(events)
}

#[derive(Clone, Copy, PartialEq)]
enum FieldBalance {
    /// Every complex field begun is ended, and every owner lies inside one.
    Whole,
    /// A complex field begun here ends after the paragraphs.
    Opens,
    /// A complex field ended here begins before the paragraphs.
    Closes,
    /// The fields are whole, but an owner lies outside all of them.
    Uncarried,
}

fn field_balance(paragraphs: &[&CT_P]) -> Result<FieldBalance> {
    let mut depth = 0usize;
    let mut uncarried = false;
    for paragraph in paragraphs {
        for event in paragraph_field_events(paragraph)? {
            match event {
                FieldEvent::Begin => depth += 1,
                FieldEvent::End if depth == 0 => return Ok(FieldBalance::Closes),
                FieldEvent::End => depth -= 1,
                FieldEvent::Owner => uncarried |= depth == 0,
            }
        }
    }
    Ok(if depth > 0 {
        FieldBalance::Opens
    } else if uncarried {
        FieldBalance::Uncarried
    } else {
        FieldBalance::Whole
    })
}

fn body_signature(content: &BodyContent) -> String {
    match content {
        BodyContent::Paragraph(paragraph) => format!("p:{}", paragraph_signature(paragraph)),
        BodyContent::Table(table) => format!("t:{}", table_signature(table)),
        BodyContent::ContentControl(control) => format!("s:{}", control_signature(control)),
        BodyContent::RawXml(raw) => format!("x:{raw:?}"),
    }
}

fn paragraph_signature(paragraph: &CT_P) -> String {
    let numbering = paragraph_numbering(paragraph);
    let extra_xml = paragraph
        .extra_xml
        .iter()
        .filter(|(position, raw)| !CT_P::raw_is_root_attributes(*position, raw))
        .collect::<Vec<_>>();
    format!(
        "{numbering:?}:{:?}:{:?}:{:?}:{:?}:{:?}:{:?}",
        paragraph.runs.iter().map(run_signature).collect::<Vec<_>>(),
        paragraph.hyperlinks,
        paragraph.comment_ranges,
        paragraph.bookmark_markers,
        extra_xml,
        paragraph
            .content_controls
            .iter()
            .map(|(at, raw_before, markers_before, control)| (
                at,
                raw_before,
                markers_before,
                control_signature(control),
            ))
            .collect::<Vec<_>>(),
    )
}

fn paragraph_numbering(paragraph: &CT_P) -> Option<(Option<u32>, Option<u32>)> {
    paragraph.properties.as_ref().and_then(|properties| {
        (properties.num_id.is_some() || properties.num_ilvl.is_some())
            .then_some((properties.num_id, properties.num_ilvl))
    })
}

fn run_signature(run: &CT_R) -> String {
    let raw = semantic_run_raw(run);
    format!(
        "{:?}:{raw:?}",
        run.content
            .iter()
            .map(run_content_signature)
            .collect::<Vec<_>>()
    )
}

fn semantic_run_raw(run: &CT_R) -> Vec<(&[u8], usize)> {
    if run.extra_xml.len() != run.extra_xml_positions.len() {
        return run
            .extra_xml
            .iter()
            .enumerate()
            .filter(|(_, raw)| !rdocx_oxml::text::is_root_attribute_record(raw))
            .map(|(position, raw)| (raw.as_slice(), position))
            .collect();
    }
    run.extra_xml
        .iter()
        .zip(&run.extra_xml_positions)
        .filter(|(_, position)| !CT_R::raw_child_is_root_attributes(**position))
        .map(|(raw, position)| (raw.as_slice(), *position))
        .collect()
}

fn run_content_signature(content: &RunContent) -> String {
    match content {
        RunContent::Field(_) => "field-owner".to_owned(),
        RunContent::Drawing(drawing) => format!("Drawing({:?})", drawing_signature(drawing)),
        RunContent::Text(text) => format!("Text({:?})", text_signature(text)),
        RunContent::DeletedText(text) => format!("DeletedText({:?})", text_signature(text)),
        content => format!("{content:?}"),
    }
}

/// The text and whether its `xml:space="preserve"` changes how it reads.
///
/// The flag only protects whitespace at an edge of the text. Producers write
/// it on every `w:t` or only where needed, so elsewhere it is serialization.
fn text_signature(text: &CT_Text) -> (&str, bool) {
    (
        &text.text,
        text.preserve_space && has_edge_whitespace(&text.text),
    )
}

/// Whether XML whitespace starts or ends the text.
fn has_edge_whitespace(text: &str) -> bool {
    let whitespace = |character: char| matches!(character, ' ' | '\t' | '\n' | '\r');
    text.starts_with(whitespace) || text.ends_with(whitespace)
}

fn table_signature(table: &CT_Tbl) -> String {
    format!(
        "{:?}:{:?}:{:?}:{:?}",
        table.grid,
        table.rows.iter().map(row_signature).collect::<Vec<_>>(),
        table.extra_xml,
        table
            .content_controls
            .iter()
            .map(|(at, raw_before, control)| (at, raw_before, control_signature(control)))
            .collect::<Vec<_>>()
    )
}

fn row_signature(row: &CT_Row) -> String {
    format!(
        "{:?}:{:?}:{:?}",
        row.cells.iter().map(cell_signature).collect::<Vec<_>>(),
        row_raw_children(row),
        row.content_controls
            .iter()
            .map(|(at, raw_before, control)| (at, raw_before, control_signature(control)))
            .collect::<Vec<_>>()
    )
}

fn cell_signature(cell: &rdocx_oxml::table::CT_Tc) -> String {
    format!(
        "{:?}:{:?}",
        cell.content
            .iter()
            .map(|content| match content {
                CellContent::Paragraph(paragraph) => paragraph_signature(paragraph),
                CellContent::Table(table) => table_signature(table),
                CellContent::ContentControl(control) => control_signature(control),
            })
            .collect::<Vec<_>>(),
        cell.extra_xml
    )
}

fn control_signature(control: &CT_Sdt) -> String {
    let properties = control_property_signature(control);
    format!(
        "{:?}:{:?}",
        properties,
        control
            .content
            .iter()
            .filter_map(|content| match content {
                SdtContent::Paragraph(paragraph) => Some(paragraph_signature(paragraph)),
                SdtContent::Table(table) => Some(table_signature(table)),
                SdtContent::Row(row) => Some(row_signature(row)),
                SdtContent::Cell(cell) => Some(cell_signature(cell)),
                SdtContent::Run(run) => Some(run_signature(run)),
                SdtContent::ContentControl(control) => Some(control_signature(control)),
                SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace) => None,
                SdtContent::RawXml(raw) => Some(format!("{raw:?}")),
            })
            .collect::<Vec<_>>()
    )
}

/// The content-control properties that decide what a control holds, its type
/// and data binding, which alignment, refusal and the accept and reject
/// postconditions compare.
///
/// `w:id` is left out because producers renumber it on save, and the
/// [`CONTROL_METADATA`] because a difference there is reported as a
/// diagnostic. A pair that differs only by those keeps the original's
/// `w:sdtPr`. A `w:sdtPr` with neither property reads like no `w:sdtPr`.
fn control_property_signature(control: &CT_Sdt) -> ControlPropertySignature<'_> {
    control
        .properties
        .as_ref()
        .map(|properties| (properties.control_type, properties.data_binding.as_ref()))
        .filter(|signature| !matches!(signature, (None, None)))
}

/// Report each [`CONTROL_METADATA`] property that differs between the two
/// controls. The redline keeps the original `w:sdtPr`.
fn control_metadata_diagnostics(
    original: &CT_Sdt,
    edited: &CT_Sdt,
    location: &str,
    diagnostics: &mut Vec<ComparisonDiagnostic>,
) -> Result<()> {
    if original.properties == edited.properties {
        return Ok(());
    }
    let edited_values = control_metadata(edited)?;
    for ((name, original_value), edited_value) in CONTROL_METADATA
        .iter()
        .zip(control_metadata(original)?)
        .zip(edited_values)
    {
        if original_value != edited_value {
            diagnostics.push(ComparisonDiagnostic {
                location: location.to_owned(),
                message: format!(
                    "content-control {name} differs and the original {name} was retained"
                ),
            });
        }
    }
    Ok(())
}

/// The [`CONTROL_METADATA`] values of a control, in that order.
///
/// The lock, placeholder and gallery are kept as raw `w:sdtPr` children, so
/// they are read from the serialized properties by local name.
fn control_metadata(control: &CT_Sdt) -> Result<[Option<String>; 5]> {
    let Some(properties) = &control.properties else {
        return Ok(Default::default());
    };
    let mut values = [
        properties.tag.clone(),
        properties.alias.clone(),
        None,
        None,
        None,
    ];
    let mut shell = control.clone();
    shell.content.clear();
    let xml = control_xml(&shell)?;
    let mut reader = Reader::from_str(&xml);
    let mut path = Vec::<Vec<u8>>::new();
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::Other(format!("comparison content-control scan failed: {error}"))
        })?;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                path.pop();
                buffer.clear();
                continue;
            }
            Event::Eof => break,
            _ => {
                buffer.clear();
                continue;
            }
        };
        let local = element.local_name().as_ref().to_vec();
        let parents = path.iter().map(Vec::as_slice).collect::<Vec<_>>();
        let slot = match (parents.as_slice(), local.as_slice()) {
            ([b"sdt", b"sdtPr"], b"lock") => Some(2),
            ([b"sdt", b"sdtPr", b"placeholder"], b"docPart") => Some(3),
            ([b"sdt", b"sdtPr", b"docPartObj" | b"docPartList"], b"docPartGallery") => Some(4),
            _ => None,
        };
        if let Some(slot) = slot {
            let mut value = String::new();
            for attribute in element.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::Other(format!(
                        "comparison content-control attribute failed: {error}"
                    ))
                })?;
                if attribute.key.local_name().as_ref() == b"val" {
                    value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                        .map_err(|error| {
                            Error::Other(format!(
                                "comparison content-control value failed: {error}"
                            ))
                        })?
                        .into_owned();
                }
            }
            values[slot] = Some(value);
        }
        if !empty {
            path.push(local);
        }
        buffer.clear();
    }
    Ok(values)
}

fn modeled_control_content(control: &CT_Sdt) -> Vec<&SdtContent> {
    control
        .content
        .iter()
        .filter(|content| {
            !matches!(content, SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace))
        })
        .collect()
}

fn control_whitespace_slots(control: &CT_Sdt) -> Result<Vec<String>> {
    let modeled_count = modeled_control_content(control).len();
    let mut slots = vec![String::new(); modeled_count + 1];
    let mut modeled_before = 0usize;
    for content in &control.content {
        match content {
            SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace) => {
                slots[modeled_before].push_str(std::str::from_utf8(raw).map_err(utf8_error)?);
            }
            _ => modeled_before += 1,
        }
    }
    Ok(slots)
}

fn normalized_body(document: &CT_Document) -> Vec<String> {
    document.body.content.iter().map(body_signature).collect()
}

/// The accept and reject postcondition projection of a story: each item's
/// policy signature followed by the modeled properties of every paragraph it
/// owns. The signature leaves formatting out because alignment matches changed
/// paragraphs through it, so without the properties a resolution that loses
/// a paragraph's alignment or mark formatting would pass.
fn postcondition_body(document: &CT_Document, options: &ComparisonOptions) -> Vec<String> {
    let mut items = normalized_body_with_options(document, options);
    if options.ignore_formatting {
        return items;
    }
    for (item, content) in items.iter_mut().zip(&document.body.content) {
        let mut properties = Vec::new();
        match content {
            BodyContent::Paragraph(paragraph) => {
                properties.push(modeled_paragraph_properties(paragraph.properties.as_ref()));
            }
            BodyContent::Table(table) => table_paragraph_properties(table, &mut properties),
            BodyContent::ContentControl(control) => {
                control_paragraph_properties(control, &mut properties);
            }
            BodyContent::RawXml(_) => {}
        }
        item.push_str(&format!(":{properties:?}"));
    }
    items
}

fn table_paragraph_properties(table: &CT_Tbl, properties: &mut Vec<Option<CT_PPr>>) {
    for cell in table.rows.iter().flat_map(|row| &row.cells) {
        cell_paragraph_properties(cell, properties);
    }
}

fn cell_paragraph_properties(cell: &CT_Tc, properties: &mut Vec<Option<CT_PPr>>) {
    for content in &cell.content {
        match content {
            CellContent::Paragraph(paragraph) => {
                properties.push(modeled_paragraph_properties(paragraph.properties.as_ref()));
            }
            CellContent::Table(table) => table_paragraph_properties(table, properties),
            CellContent::ContentControl(control) => {
                control_paragraph_properties(control, properties);
            }
        }
    }
}

fn control_paragraph_properties(control: &CT_Sdt, properties: &mut Vec<Option<CT_PPr>>) {
    for content in &control.content {
        match content {
            SdtContent::Paragraph(paragraph) => {
                properties.push(modeled_paragraph_properties(paragraph.properties.as_ref()));
            }
            SdtContent::Table(table) => table_paragraph_properties(table, properties),
            SdtContent::Row(row) => {
                for cell in &row.cells {
                    cell_paragraph_properties(cell, properties);
                }
            }
            SdtContent::Cell(cell) => cell_paragraph_properties(cell, properties),
            SdtContent::ContentControl(control) => {
                control_paragraph_properties(control, properties);
            }
            SdtContent::Run(_) | SdtContent::RawXml(_) => {}
        }
    }
}

fn normalized_body_with_options(
    document: &CT_Document,
    options: &ComparisonOptions,
) -> Vec<String> {
    if options == &ComparisonOptions::default() {
        return normalized_body(document);
    }
    document
        .body
        .content
        .iter()
        .map(|content| body_signature_with_options(content, options))
        .collect()
}

fn body_signature_with_options(content: &BodyContent, options: &ComparisonOptions) -> String {
    match content {
        BodyContent::Paragraph(paragraph) => {
            format!("p:{}", paragraph_signature_with_options(paragraph, options))
        }
        BodyContent::Table(table) => {
            format!("t:{}", table_signature_with_options(table, options))
        }
        BodyContent::ContentControl(control) => {
            format!("s:{}", control_signature_with_options(control, options))
        }
        BodyContent::RawXml(raw) => format!("x:{raw:?}"),
    }
}

fn paragraph_signature_with_options(paragraph: &CT_P, options: &ComparisonOptions) -> String {
    let numbering = (!options.ignore_formatting)
        .then(|| paragraph_numbering(paragraph))
        .flatten();
    let hyperlinks = paragraph
        .hyperlinks
        .iter()
        .map(|link| {
            (
                &link.rel_id,
                &link.anchor,
                &link.tooltip,
                &link.doc_location,
                &link.extra_attributes,
                hyperlink_raw_boundaries(paragraph, link, options),
                link.preserved_raw_before,
                policy_run_boundary(paragraph, link.run_start, options),
                policy_run_boundary(paragraph, link.run_end, options),
            )
        })
        .collect::<Vec<_>>();
    format!(
        "{numbering:?}:{:?}:{hyperlinks:?}",
        paragraph_tokens(paragraph, options)
    )
}

/// One token of a paragraph signature.
#[derive(PartialEq, Eq)]
enum ParagraphToken {
    /// A compared unit, or a whole run on the whole-run path.
    Unit(String),
    /// A unit the options ignore, such as whitespace, a field or a comment
    /// reference.
    Ignored(String),
    /// A preserved raw child, comment marker, inline control or hyperlink edge.
    Boundary(String),
}

/// Paragraph content as tokens in document order.
///
/// Raw children, comment markers and inline controls come out at each run
/// boundary in the order the serializer writes them, and each run follows
/// as its compared units. Splitting or merging runs therefore cannot move
/// a marker, and a run without content cannot reorder its neighbours.
/// Ignored units are kept only where they touch a boundary token, so a
/// marker on either side of ignored content still differs.
fn paragraph_tokens(paragraph: &CT_P, options: &ComparisonOptions) -> Vec<String> {
    let attributed = uses_attributed_run_path(options);
    let units = if attributed {
        attributed_run_units(&paragraph.runs, options)
    } else {
        Vec::new()
    };
    let mut next_unit = 0;
    let mut tokens = Vec::new();
    let mut open_hyperlink = None;
    for boundary in 0..=paragraph.runs.len() {
        let inside = paragraph
            .hyperlinks
            .iter()
            .rposition(|link| link.run_start <= boundary && boundary < link.run_end);
        if let Some(index) = open_hyperlink
            && open_hyperlink != inside
        {
            tokens.push(ParagraphToken::Boundary(format!("hyperlink-end:{index}")));
            open_hyperlink = None;
        }
        push_boundary_tokens(paragraph, boundary, options, &mut tokens);
        for (index, _) in paragraph.hyperlinks.iter().enumerate().filter(|(_, link)| {
            link.run_start == boundary
                && link.run_end == boundary
                && link.preserved_raw_before.is_none()
        }) {
            tokens.push(ParagraphToken::Boundary(format!("hyperlink-empty:{index}")));
        }
        let Some(run) = paragraph.runs.get(boundary) else {
            break;
        };
        if let Some(index) = inside
            && open_hyperlink != inside
        {
            tokens.push(ParagraphToken::Boundary(format!("hyperlink-start:{index}")));
            open_hyperlink = inside;
        }
        if !attributed {
            let signature = run_signature_with_options(run, options);
            if !signature.is_empty() {
                tokens.push(ParagraphToken::Unit(signature));
            }
            continue;
        }
        while let Some(unit) = units.get(next_unit).filter(|unit| unit.owner == boundary) {
            next_unit += 1;
            if unit_is_empty(unit) {
                continue;
            }
            let signature = attributed_unit_signature(unit);
            tokens.push(if unit_is_ignorable(unit) {
                ParagraphToken::Ignored(signature)
            } else {
                ParagraphToken::Unit(signature)
            });
        }
    }
    tokens
        .dedup_by(|next, previous| matches!(next, ParagraphToken::Ignored(_)) && next == previous);
    let ignored = |index: usize| matches!(tokens.get(index), Some(ParagraphToken::Ignored(_)));
    let boundary = |index: usize| matches!(tokens.get(index), Some(ParagraphToken::Boundary(_)));
    let mut keep = vec![true; tokens.len()];
    let mut start = 0;
    while start < tokens.len() {
        if !ignored(start) {
            start += 1;
            continue;
        }
        let end = (start..tokens.len())
            .find(|index| !ignored(*index))
            .unwrap_or(tokens.len());
        let touches_boundary = (start > 0 && boundary(start - 1)) || boundary(end);
        keep[start..end].fill(touches_boundary);
        start = end;
    }
    tokens
        .into_iter()
        .zip(keep)
        .filter(|(_, keep)| *keep)
        .map(|(token, _)| match token {
            ParagraphToken::Unit(signature) => format!("unit:{signature}"),
            ParagraphToken::Ignored(signature) | ParagraphToken::Boundary(signature) => signature,
        })
        .collect()
}

/// Push the raw children, comment markers and inline controls at one run
/// boundary in the order `CT_P` serialization writes them.
fn push_boundary_tokens(
    paragraph: &CT_P,
    boundary: usize,
    options: &ComparisonOptions,
    tokens: &mut Vec<ParagraphToken>,
) {
    let extras = paragraph
        .extra_xml
        .iter()
        .filter(|(position, raw)| {
            *position == boundary && !CT_P::raw_is_root_attributes(*position, raw)
        })
        .map(|(_, raw)| raw)
        .collect::<Vec<_>>();
    for raw_index in 0..=extras.len() {
        let markers = paragraph
            .comment_ranges
            .iter()
            .filter_map(|marker| {
                let (start, id, run_index, raw_before, has_child_content) = match *marker {
                    CommentRangeMarker::Start {
                        id,
                        run_index,
                        raw_before,
                        has_child_content,
                    } => (true, id, run_index, raw_before, has_child_content),
                    CommentRangeMarker::End {
                        id,
                        run_index,
                        raw_before,
                        has_child_content,
                    } => (false, id, run_index, raw_before, has_child_content),
                };
                (run_index == boundary && raw_before.min(extras.len()) == raw_index)
                    .then(|| format!("comment-range:{start}:{id}:{has_child_content}"))
            })
            .collect::<Vec<_>>();
        for marker_index in 0..=markers.len() {
            for (_, _, _, control) in
                paragraph
                    .content_controls
                    .iter()
                    .filter(|(at, raw_before, markers_before, _)| {
                        *at == boundary
                            && (*raw_before).min(extras.len()) == raw_index
                            && (*markers_before).min(markers.len()) == marker_index
                    })
            {
                tokens.push(ParagraphToken::Boundary(format!(
                    "control:{}",
                    control_signature_with_options(control, options)
                )));
            }
            if let Some(marker) = markers.get(marker_index)
                && !options.ignore_comments
            {
                tokens.push(ParagraphToken::Boundary(marker.clone()));
            }
        }
        if let Some(raw) = extras.get(raw_index) {
            tokens.push(ParagraphToken::Boundary(format!("raw:{raw:?}")));
        }
    }
}

fn run_signature_with_options(run: &CT_R, options: &ComparisonOptions) -> String {
    if options == &ComparisonOptions::default() {
        return run_signature(run);
    }
    let content = run
        .content
        .iter()
        .filter_map(|content| {
            if options.ignore_comments && matches!(content, RunContent::CommentReference { .. }) {
                return None;
            }
            if options.ignore_fields && matches!(content, RunContent::Field(_)) {
                return Some("ignored:field".to_owned());
            }
            match content {
                RunContent::Text(text) | RunContent::DeletedText(text)
                    if options.ignore_whitespace =>
                {
                    let visible = text
                        .text
                        .chars()
                        .filter(|character| !character.is_whitespace())
                        .collect::<String>();
                    (!visible.is_empty()).then_some(format!("text:{visible:?}"))
                }
                _ => Some(run_content_signature(content)),
            }
        })
        .collect::<Vec<_>>();
    let raw = semantic_run_raw(run);
    if content.is_empty() && raw.is_empty() {
        String::new()
    } else {
        format!("{content:?}:{raw:?}")
    }
}

fn table_signature_with_options(table: &CT_Tbl, options: &ComparisonOptions) -> String {
    format!(
        "{:?}:{:?}:{:?}:{:?}",
        table.grid,
        table
            .rows
            .iter()
            .map(|row| row_signature_with_options(row, options))
            .collect::<Vec<_>>(),
        table.extra_xml,
        table
            .content_controls
            .iter()
            .map(|(at, raw_before, control)| (
                at,
                raw_before,
                control_signature_with_options(control, options),
            ))
            .collect::<Vec<_>>()
    )
}

fn row_signature_with_options(row: &CT_Row, options: &ComparisonOptions) -> String {
    format!(
        "{:?}:{:?}:{:?}",
        row.cells
            .iter()
            .map(|cell| cell_signature_with_options(cell, options))
            .collect::<Vec<_>>(),
        row_raw_children(row),
        row.content_controls
            .iter()
            .map(|(at, raw_before, control)| (
                at,
                raw_before,
                control_signature_with_options(control, options),
            ))
            .collect::<Vec<_>>()
    )
}

fn cell_signature_with_options(cell: &CT_Tc, options: &ComparisonOptions) -> String {
    format!(
        "{:?}:{:?}",
        cell.content
            .iter()
            .map(|content| match content {
                CellContent::Paragraph(paragraph) => {
                    paragraph_signature_with_options(paragraph, options)
                }
                CellContent::Table(table) => table_signature_with_options(table, options),
                CellContent::ContentControl(control) => {
                    control_signature_with_options(control, options)
                }
            })
            .collect::<Vec<_>>(),
        cell.extra_xml
    )
}

fn control_signature_with_options(control: &CT_Sdt, options: &ComparisonOptions) -> String {
    let mut content = Vec::new();
    for child in &control.content {
        match child {
            SdtContent::RawXml(raw) if raw.iter().all(u8::is_ascii_whitespace) => {}
            SdtContent::Run(run) if uses_attributed_run_path(options) => {
                content.extend(
                    attributed_run_units(std::slice::from_ref(run), options)
                        .into_iter()
                        .filter(|unit| !unit_is_ignorable(unit) && !unit_is_empty(unit))
                        .map(|unit| format!("u:{}", attributed_unit_signature(&unit))),
                );
            }
            child => content.push(control_content_signature_with_options(child, options)),
        }
    }
    format!("{:?}:{:?}", control_property_signature(control), content)
}

fn control_content_signature_with_options(
    content: &SdtContent,
    options: &ComparisonOptions,
) -> String {
    match content {
        SdtContent::Paragraph(paragraph) => {
            format!("p:{}", paragraph_signature_with_options(paragraph, options))
        }
        SdtContent::Table(table) => {
            format!("t:{}", table_signature_with_options(table, options))
        }
        SdtContent::Row(row) => format!("r:{}", row_signature_with_options(row, options)),
        SdtContent::Cell(cell) => format!("c:{}", cell_signature_with_options(cell, options)),
        SdtContent::Run(run) => format!("u:{}", run_signature_with_options(run, options)),
        SdtContent::ContentControl(control) => {
            format!("s:{}", control_signature_with_options(control, options))
        }
        SdtContent::RawXml(raw) => format!("x:{raw:?}"),
    }
}

fn paragraph_formatting(paragraph: &CT_P) -> Option<CT_PPr> {
    paragraph.properties.clone().map(|mut properties| {
        properties.num_id = None;
        properties.num_ilvl = None;
        properties.num_id_raw = None;
        properties.num_ilvl_raw = None;
        properties.numbering_revision = None;
        properties.numbering_revision_xml.clear();
        properties.numbering_revision_xml_positions.clear();
        properties.numbering_revision_position = None;
        properties.change = None;
        properties.revision_xml.clear();
        if let Some(section) = properties.sect_pr.as_mut() {
            clear_default_orientation(section);
        }
        properties
    })
}

fn row_formatting(row: &CT_Row) -> Option<CT_TrPr> {
    row.properties.clone().map(|mut properties| {
        properties.revision_markers.clear();
        properties.revision_xml.clear();
        properties
    })
}

fn body_content_xml(content: &BodyContent) -> Result<String> {
    match content {
        BodyContent::Paragraph(paragraph) => paragraph_xml(paragraph),
        BodyContent::Table(table) => table_xml(table),
        BodyContent::ContentControl(control) => control_xml(control),
        BodyContent::RawXml(raw) => String::from_utf8(raw.as_ref().clone()).map_err(utf8_error),
    }
}

/// The paragraph source with each run that the reader split around a field
/// written as separate physical runs, one per modeled run.
fn detach_field_spans(paragraph: &CT_P, source: Option<&str>) -> Result<Option<String>> {
    let spans = paragraph.detached_field_spans()?;
    if spans.is_empty() {
        return Ok(None);
    }
    let mut source =
        source.map_or_else(|| paragraph_xml(paragraph), |source| Ok(source.to_owned()))?;
    // Each span starts at a physical run of the paragraph, in order.
    let runs = paragraph_run_spans(&source)?;
    let mut edits = Vec::with_capacity(spans.len());
    let mut cursor = 0;
    for (raw, detached) in spans {
        let raw = std::str::from_utf8(raw).map_err(utf8_error)?;
        let start = runs
            .iter()
            .map(|run| run.start)
            .find(|start| *start >= cursor && source[*start..].starts_with(raw))
            .ok_or_else(|| {
                Error::Other(
                    "comparison could not find a split field run in its paragraph".to_owned(),
                )
            })?;
        cursor = start + raw.len();
        edits.push((
            start..cursor,
            String::from_utf8(detached).map_err(utf8_error)?,
        ));
    }
    for (range, detached) in edits.into_iter().rev() {
        source.replace_range(range, &detached);
    }
    Ok(Some(source))
}

fn paragraph_xml(paragraph: &CT_P) -> Result<String> {
    let mut bytes = Vec::new();
    paragraph.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn property_xml(properties: &CT_PPr) -> Result<String> {
    let mut bytes = Vec::new();
    properties.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn run_xml(run: &CT_R) -> Result<String> {
    let mut bytes = Vec::new();
    run.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn paragraph_owned_run_xml(run: &CT_R) -> Result<String> {
    if !run_is_field(run) {
        return run_xml(run);
    }
    let mut paragraph = CT_P::new();
    paragraph.runs.push(run.clone());
    let xml = paragraph_xml(&paragraph)?;
    let start = xml
        .find('>')
        .ok_or_else(|| Error::Other("serialized paragraph has no start".to_owned()))?
        + 1;
    let end = xml
        .rfind("</w:p>")
        .ok_or_else(|| Error::Other("serialized paragraph has no end".to_owned()))?;
    Ok(xml[start..end].to_owned())
}

fn policy_inserted_run_xml(
    edited: &CT_R,
    original: Option<&CT_R>,
    options: &ComparisonOptions,
) -> Result<String> {
    if !options.ignore_formatting {
        return paragraph_owned_run_xml(edited);
    }
    let mut inserted = edited.clone();
    if let Some(original) = original {
        inserted.properties = original.properties.clone();
    }
    paragraph_owned_run_xml(&inserted)
}

fn run_is_field(run: &CT_R) -> bool {
    matches!(run.content.as_slice(), [RunContent::Field(_)])
}

fn run_field(run: &CT_R) -> Option<&rdocx_oxml::text::Field> {
    match run.content.as_slice() {
        [RunContent::Field(field)] => Some(field),
        _ => None,
    }
}

fn validate_field_alignment(
    aligned: &[(Option<usize>, Option<usize>)],
    original: &[CT_R],
    edited: &[CT_R],
    original_signatures: &[String],
    edited_signatures: &[String],
    location: &str,
) -> Result<()> {
    for (left, right) in aligned {
        let left_field = left.is_some_and(|index| run_is_field(&original[index]));
        let right_field = right.is_some_and(|index| run_is_field(&edited[index]));
        let unchanged_pair = match (*left, *right) {
            (Some(i), Some(j)) => original_signatures[i] == edited_signatures[j],
            _ => false,
        };
        if (left_field || right_field) && !unchanged_pair {
            return Err(Error::Other(format!(
                "comparison cannot revise a modeled field at {location}"
            )));
        }
    }
    Ok(())
}

/// A run as deleted content. The field instruction of a field that spans
/// paragraphs is a preserved child, so it is renamed in the written run.
fn deleted_run_xml(run: &CT_R) -> Result<String> {
    Ok(renamed_word_elements(
        &run_xml(&deleted_run(run))?,
        "w:instrText",
        "w:delInstrText",
    ))
}

fn deleted_run(run: &CT_R) -> CT_R {
    let mut deleted = run.clone();
    for content in &mut deleted.content {
        if let RunContent::Text(text) = content {
            *content = RunContent::DeletedText(text.clone());
        }
    }
    deleted
}

fn table_xml(table: &CT_Tbl) -> Result<String> {
    let mut bytes = Vec::new();
    table.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn row_xml(row: &CT_Row) -> Result<String> {
    let mut bytes = Vec::new();
    row.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn cell_xml(cell: &CT_Tc) -> Result<String> {
    let mut bytes = Vec::new();
    cell.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn control_xml(control: &CT_Sdt) -> Result<String> {
    let document = CT_Document {
        body: rdocx_oxml::document::CT_Body {
            content: vec![BodyContent::ContentControl(Box::new(control.clone()))],
            sect_pr: None,
        },
        ..CT_Document::new()
    };
    let xml = document.to_xml()?;
    extract_body_inner(&xml).map(str::to_owned)
}

fn section_property_xml(section: &rdocx_oxml::document::CT_SectPr) -> Result<String> {
    let mut bytes = Vec::new();
    section.to_xml(&mut Writer::new(&mut bytes))?;
    String::from_utf8(bytes).map_err(utf8_error)
}

fn replace_body_inner(document: &[u8], body: &str) -> Result<String> {
    let xml = std::str::from_utf8(document).map_err(utf8_error)?;
    let start = xml
        .find("<w:body")
        .and_then(|at| xml[at..].find('>').map(|offset| at + offset + 1))
        .ok_or_else(|| Error::Other("serialized document has no w:body start".to_owned()))?;
    let end = xml
        .rfind("</w:body>")
        .ok_or_else(|| Error::Other("serialized document has no w:body end".to_owned()))?;
    Ok(format!("{}{}{}", &xml[..start], body, &xml[end..]))
}

fn extract_body_inner(document: &[u8]) -> Result<&str> {
    let xml = std::str::from_utf8(document).map_err(utf8_error)?;
    let start = xml
        .find("<w:body")
        .and_then(|at| xml[at..].find('>').map(|offset| at + offset + 1))
        .ok_or_else(|| Error::Other("serialized document has no w:body start".to_owned()))?;
    let end = xml
        .rfind("</w:body>")
        .ok_or_else(|| Error::Other("serialized document has no w:body end".to_owned()))?;
    Ok(&xml[start..end])
}

fn replace_element_inner(xml: &str, name: &str, inner: &str) -> Result<String> {
    let opening = format!("<{name}");
    let closing = format!("</{name}>");
    let start = xml
        .find(&opening)
        .and_then(|at| xml[at..].find('>').map(|offset| at + offset + 1))
        .ok_or_else(|| Error::Other(format!("serialized XML has no {name} start")))?;
    let end = xml
        .rfind(&closing)
        .ok_or_else(|| Error::Other(format!("serialized XML has no {name} end")))?;
    Ok(format!("{}{}{}", &xml[..start], inner, &xml[end..]))
}

fn inject_before_close(xml: &str, closing: &str, addition: &str) -> Result<String> {
    let at = xml
        .rfind(closing)
        .ok_or_else(|| Error::Other(format!("serialized XML has no {closing}")))?;
    Ok(format!("{}{}{}", &xml[..at], addition, &xml[at..]))
}

fn append_word_child(xml: &str, local: &str, addition: &str) -> Result<String> {
    let closing = format!("</w:{local}>");
    if xml.ends_with(&closing) {
        return inject_before_close(xml, &closing, addition);
    }
    let Some(slash) = xml.rfind("/>") else {
        return Err(Error::Other(format!(
            "serialized w:{local} has no closing boundary"
        )));
    };
    Ok(format!(
        "{}>{addition}</w:{local}>{}",
        &xml[..slash],
        &xml[slash + 2..]
    ))
}

/// Add a revision marker to the `w:rPr` of a paragraph mark in schema
/// order: `w:ins`, `w:del`, `w:moveFrom` and `w:moveTo` come first, in that
/// order, before any formatting.
fn with_mark_revision(run_properties: &str, marker: &str) -> Result<String> {
    const ORDER: [&str; 4] = ["ins", "del", "moveFrom", "moveTo"];
    let rank = |xml: &str| {
        ORDER
            .iter()
            .position(|local| *local == element_local_name(xml))
    };
    let children = direct_element_spans(run_properties)?;
    let (Some(marker_rank), Some(open)) = (rank(marker), run_properties.find('>')) else {
        return append_word_child(run_properties, "rPr", marker);
    };
    if children.is_empty() {
        return append_word_child(run_properties, "rPr", marker);
    }
    let at = children
        .iter()
        .take_while(|span| {
            rank(&run_properties[(*span).clone()]).is_some_and(|rank| rank <= marker_rank)
        })
        .last()
        .map_or(open + 1, |span| span.end);
    let mut marked = run_properties.to_owned();
    marked.insert_str(at, marker);
    Ok(marked)
}

/// The local name of the element that `xml` starts with.
fn element_local_name(xml: &str) -> &str {
    let name_end = xml
        .find(|character: char| character.is_ascii_whitespace() || matches!(character, '/' | '>'))
        .unwrap_or(xml.len());
    xml.get(1..name_end)
        .and_then(|name| name.rsplit(':').next())
        .unwrap_or_default()
}

fn direct_word_element_spans(xml: &str, local: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let expected = format!("w:{local}");
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let target = depth == 1 && element.name().as_ref() == expected.as_bytes();
                open.push(target.then_some(before));
                depth += 1;
            }
            Event::Empty(element) => {
                if depth == 1 && element.name().as_ref() == expected.as_bytes() {
                    spans.push(before..after);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some(start)) = open.pop() {
                    spans.push(start..after);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn paragraph_close_start(xml: &str) -> Result<usize> {
    xml.rfind("</w:p>")
        .ok_or_else(|| Error::Other("serialized paragraph has no end".to_owned()))
}

fn paragraph_run_spans(xml: &str) -> Result<Vec<Range<usize>>> {
    let open = format!(r#"<rdocxcmp:root xmlns:rdocxcmp="urn:rdocx-compare" xmlns:w="{W_NS}">"#);
    let wrapped = format!("{open}{xml}</rdocxcmp:root>");
    let offset = open.len();
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut stack = Vec::<(Vec<u8>, Option<usize>)>::new();
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let is_word =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes());
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                let parent = stack.last().map(|(local, _)| local.as_slice());
                let target = is_word
                    && matches!(local.as_slice(), b"r" | b"fldSimple")
                    && parent.is_some_and(|value| value == b"p" || value == b"hyperlink");
                stack.push((local, target.then_some(before)));
            }
            Event::Empty(element) => {
                let local = element.local_name();
                let parent = stack.last().map(|(local, _)| local.as_slice());
                if is_word
                    && matches!(local.as_ref(), b"r" | b"fldSimple")
                    && parent.is_some_and(|value| value == b"p" || value == b"hyperlink")
                {
                    spans.push(before.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::End(_) => {
                if let Some((_, Some(start))) = stack.pop() {
                    spans.push(start.saturating_sub(offset)..after.saturating_sub(offset));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    spans.sort_by_key(|span| span.start);
    Ok(spans)
}

fn replace_paragraph_run_elements(
    paragraph: &CT_P,
    xml: &str,
    replacements: &[String],
) -> Result<String> {
    let spans = modeled_paragraph_run_spans(paragraph, xml)?;
    if spans.len() != replacements.len() {
        return Err(Error::Other(format!(
            "comparison expected {} paragraph run owners, found {}",
            replacements.len(),
            spans.len()
        )));
    }
    let mut output = xml.to_owned();
    for (span, replacement) in spans.into_iter().zip(replacements).rev() {
        output.replace_range(span, replacement);
    }
    Ok(output)
}

fn modeled_paragraph_run_spans(paragraph: &CT_P, xml: &str) -> Result<Vec<Range<usize>>> {
    let physical = paragraph_run_spans(xml)?;
    let mut projected = Vec::with_capacity(paragraph.runs.len());
    let mut cursor = 0usize;
    let mut previous_field_owner = None;
    for (modeled_index, run) in paragraph.runs.iter().enumerate() {
        let (physical_count, field_owner) = match run.content.as_slice() {
            [RunContent::Field(field)] if field.is_complex() => {
                if let Some(source) = field.detached_source()? {
                    let source = String::from_utf8(source).map_err(utf8_error)?;
                    let count = paragraph_run_spans(&format!("<w:p>{source}</w:p>"))?.len();
                    if count == 0 || cursor + count > physical.len() {
                        return Err(Error::Other(format!(
                            "comparison could not correlate split field run {modeled_index}"
                        )));
                    }
                    projected.push(physical[cursor].start..physical[cursor + count - 1].end);
                    cursor += count;
                    previous_field_owner = None;
                    continue;
                }
                let field_owner = field.source_owner_id().ok_or_else(|| {
                    Error::Other("parsed complex field has no source owner".to_owned())
                })?;
                if previous_field_owner == Some(field_owner) {
                    projected.push(cursor..cursor);
                    continue;
                }
                let (source, _) = field.source_replacement()?.ok_or_else(|| {
                    Error::Other("parsed complex field has no source ownership".to_owned())
                })?;
                let source = std::str::from_utf8(source).map_err(|error| {
                    Error::Other(format!("complex field source is not UTF-8: {error}"))
                })?;
                (
                    paragraph_run_spans(&format!("<w:p>{source}</w:p>"))?.len(),
                    Some(field_owner),
                )
            }
            _ => (1, None),
        };
        if physical_count == 0 || cursor + physical_count > physical.len() {
            return Err(Error::Other(format!(
                "comparison could not correlate paragraph run owner {modeled_index}: \
                 {physical_count} physical runs from cursor {cursor}, {} available",
                physical.len()
            )));
        }
        projected.push(physical[cursor].start..physical[cursor + physical_count - 1].end);
        cursor += physical_count;
        previous_field_owner = field_owner;
    }
    if cursor != physical.len() {
        return Err(Error::Other(format!(
            "comparison correlated {cursor} of {} physical paragraph runs",
            physical.len()
        )));
    }
    Ok(projected)
}

fn replace_direct_word_elements(xml: &str, local: &str, replacements: &[String]) -> Result<String> {
    let spans = direct_word_element_spans(xml, local)?;
    if spans.len() != replacements.len() {
        return Err(Error::Other(format!(
            "comparison expected {} direct w:{local} elements, found {}",
            replacements.len(),
            spans.len()
        )));
    }
    let mut output = xml.to_owned();
    for (span, replacement) in spans.into_iter().zip(replacements).rev() {
        output.replace_range(span, replacement);
    }
    Ok(output)
}

fn cell_child_spans(
    cell: &CT_Tc,
    xml: &str,
    location: &str,
) -> Result<Vec<(String, Range<usize>)>> {
    let candidates = direct_modeled_child_spans(xml)?;
    let mut spans = Vec::with_capacity(cell.content.len());
    let mut candidate_index = 0usize;
    for (index, child) in cell.content.iter().enumerate() {
        for (_, raw) in cell.extra_xml.iter().filter(|(at, _)| *at == index) {
            let raw = std::str::from_utf8(raw).map_err(utf8_error)?;
            if candidates
                .get(candidate_index)
                .is_some_and(|(_, span)| &xml[span.clone()] == raw)
            {
                candidate_index += 1;
            }
        }
        let expected_name = match child {
            CellContent::Paragraph(_) => "w:p",
            CellContent::Table(_) => "w:tbl",
            CellContent::ContentControl(_) => "w:sdt",
        };
        let (name, span) = candidates.get(candidate_index).ok_or_else(|| {
            Error::Other(format!(
                "comparison could not correlate serialized cell content at {location}"
            ))
        })?;
        if name != expected_name {
            return Err(Error::Other(format!(
                "comparison found {name} instead of {expected_name} at {location}"
            )));
        }
        spans.push((name.clone(), span.clone()));
        candidate_index += 1;
    }
    Ok(spans)
}

fn direct_modeled_child_spans(xml: &str) -> Result<Vec<(String, Range<usize>)>> {
    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let mut open: Vec<Option<(String, usize)>> = Vec::new();
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let name = std::str::from_utf8(element.name().as_ref())
                    .map_err(utf8_error)?
                    .to_owned();
                let modeled = depth == 1 && matches!(name.as_str(), "w:p" | "w:tbl" | "w:sdt");
                open.push(modeled.then_some((name, before)));
                depth += 1;
            }
            Event::Empty(element) => {
                let name = std::str::from_utf8(element.name().as_ref())
                    .map_err(utf8_error)?
                    .to_owned();
                if depth == 1 && matches!(name.as_str(), "w:p" | "w:tbl" | "w:sdt") {
                    spans.push((name, before..after));
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some(Some((name, start))) = open.pop() {
                    spans.push((name, start..after));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(spans)
}

fn replace_ranges(
    xml: &str,
    spans: &[(String, Range<usize>)],
    replacements: &[String],
) -> Result<String> {
    if spans.len() != replacements.len() {
        return Err(Error::Other(
            "comparison replacement count does not match modeled children".to_owned(),
        ));
    }
    let mut output = xml.to_owned();
    for ((_, span), replacement) in spans.iter().zip(replacements).rev() {
        output.replace_range(span.clone(), replacement);
    }
    Ok(output)
}

fn word_ids(xml: &[u8]) -> Result<HashSet<i32>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut ids = HashSet::new();
    let mut buffer = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Other(format!("comparison XML scan failed: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Other(error.to_string()))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if local.as_ref() == b"id"
                        && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == W_NS.as_bytes())
                        && let Ok(value) = std::str::from_utf8(&attribute.value)
                        && let Ok(id) = value.parse::<i32>()
                    {
                        ids.insert(id);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(ids)
}

fn utf8_error(error: impl std::fmt::Display) -> Error {
    Error::Other(format!("comparison XML is not UTF-8: {error}"))
}

#[cfg(test)]
mod tests {
    use super::{
        ComparisonGranularity, ComparisonOptions, FAIL_AFTER_COMPARISON_STAGING, align,
        attributed_run_units, attributed_unit_signature, canonical_owned_story,
        comparison_postcondition_error, complex_field_result, deleted_text_xml,
        grouped_text_alignment, normalized_body_with_options, pair_gaps_by_kind,
        postcondition_body, story_document, text_words, word_fragments,
    };
    use crate::Document;
    use crate::Length;
    use oxml_opc::OpcPackage;
    use oxml_opc::relationship::rel_types;
    use rdocx_oxml::document::BodyContent;
    use rdocx_oxml::namespace::W_NS;
    use rdocx_oxml::text::{BreakType, CT_R, CT_Text, RunContent};

    #[test]
    fn grouping_counts_joined_character_fragments_as_words() {
        let mut in_word = false;
        assert_eq!(text_words(Some("ins"), &mut in_word), 1);
        assert_eq!(text_words(Some("pec"), &mut in_word), 0);
        assert_eq!(text_words(Some("tion_123"), &mut in_word), 0);
        assert_eq!(text_words(None, &mut in_word), 0);
        assert_eq!(text_words(Some(" next"), &mut in_word), 1);
    }

    #[test]
    fn grouping_scan_preserves_unit_coverage_segments_and_determinism() {
        for granularity in [
            ComparisonGranularity::Run,
            ComparisonGranularity::Word,
            ComparisonGranularity::Character,
        ] {
            let options = ComparisonOptions {
                granularity,
                ..Default::default()
            };
            let left = attributed_run_units(
                &[CT_R::new(&"red green and blue yellow ".repeat(12))],
                &options,
            );
            let right = attributed_run_units(
                &[CT_R::new(&"cyan magenta and black white ".repeat(12))],
                &options,
            );
            let left_signatures = left
                .iter()
                .map(attributed_unit_signature)
                .collect::<Vec<_>>();
            let right_signatures = right
                .iter()
                .map(attributed_unit_signature)
                .collect::<Vec<_>>();
            let aligned = align(&left_signatures, &right_signatures);
            let segments = (0..aligned.len())
                .map(|index| index / 7)
                .collect::<Vec<_>>();
            let grouped = grouped_text_alignment(
                &aligned,
                &segments,
                &left,
                &right,
                &left_signatures,
                &right_signatures,
                &options,
            );
            assert_eq!(
                grouped,
                grouped_text_alignment(
                    &aligned,
                    &segments,
                    &left,
                    &right,
                    &left_signatures,
                    &right_signatures,
                    &options
                )
            );
            if granularity == ComparisonGranularity::Run {
                assert_eq!(grouped, (aligned.clone(), segments.clone()));
            }
            assert!(grouped.0.len() <= left.len() + right.len());
            assert_eq!(
                grouped
                    .0
                    .iter()
                    .filter_map(|pair| pair.0)
                    .collect::<Vec<_>>(),
                (0..left.len()).collect::<Vec<_>>()
            );
            assert_eq!(
                grouped
                    .0
                    .iter()
                    .filter_map(|pair| pair.1)
                    .collect::<Vec<_>>(),
                (0..right.len()).collect::<Vec<_>>()
            );
            let mut left_segments = vec![0; left.len()];
            let mut right_segments = vec![0; right.len()];
            for ((left, right), &segment) in aligned.iter().zip(&segments) {
                if let Some(left) = left {
                    left_segments[*left] = segment;
                }
                if let Some(right) = right {
                    right_segments[*right] = segment;
                }
            }
            for ((left, right), &segment) in grouped.0.iter().zip(&grouped.1) {
                if let Some(left) = left {
                    assert_eq!(left_segments[*left], segment);
                }
                if let Some(right) = right {
                    assert_eq!(right_segments[*right], segment);
                }
            }
        }
    }

    #[test]
    fn mixed_owner_gap_pairing_is_deterministic_and_preserves_every_owner() {
        let original = (0..48)
            .map(|index| format!("{}:old-{index}", ["p", "tbl", "sdt"][index % 3]))
            .collect::<Vec<_>>();
        let edited = (0..48)
            .filter(|index| index % 7 != 0)
            .map(|index| format!("{}:new-{index}", ["p", "tbl", "sdt"][index % 3]))
            .collect::<Vec<_>>();
        let paired = pair_gaps_by_kind(align(&original, &edited), &original, &edited);
        assert_eq!(
            paired,
            pair_gaps_by_kind(align(&original, &edited), &original, &edited)
        );
        assert_eq!(
            paired
                .iter()
                .filter_map(|(left, _)| *left)
                .collect::<Vec<_>>(),
            (0..original.len()).collect::<Vec<_>>()
        );
        assert_eq!(
            paired
                .iter()
                .filter_map(|(_, right)| *right)
                .collect::<Vec<_>>(),
            (0..edited.len()).collect::<Vec<_>>()
        );
        for (left, right) in paired {
            if let (Some(left), Some(right)) = (left, right) {
                assert_eq!(
                    original[left].split_once(':').unwrap().0,
                    edited[right].split_once(':').unwrap().0
                );
            }
        }
    }

    #[test]
    fn comparison_defaults_preserve_run_granularity() {
        let mut legacy = Document::new();
        legacy.add_paragraph("before");
        let mut explicit = legacy.clone_for_staging();
        let mut edited = Document::new();
        edited.add_paragraph("after");
        let legacy_diagnostics = legacy
            .compare(&edited, "Ada", "2026-09-04T09:00:00Z")
            .unwrap();
        let explicit_diagnostics = explicit
            .compare_with_options(
                &edited,
                "Ada",
                "2026-09-04T09:00:00Z",
                &ComparisonOptions::default(),
            )
            .unwrap();
        assert_eq!(legacy_diagnostics, explicit_diagnostics);
        assert_eq!(legacy.to_bytes().unwrap(), explicit.to_bytes().unwrap());
        assert_eq!(
            ComparisonOptions::default().granularity,
            ComparisonGranularity::Run
        );
    }

    #[test]
    fn synthetic_story_document_preserves_colliding_producer_prefix() {
        let document = story_document(
            r#"<rdocxcmp:p><rdocxcmp:r><rdocxcmp:t>foreign</rdocxcmp:t></rdocxcmp:r></rdocxcmp:p><w:p><w:r><w:t>word</w:t></w:r></w:p>"#,
            &[
                ("xmlns:rdocxcmp".to_owned(), "urn:producer".to_owned()),
                ("xmlns:w".to_owned(), W_NS.to_owned()),
            ],
        )
        .unwrap();

        assert!(matches!(document.body.content[0], BodyContent::RawXml(_)));
        assert!(matches!(
            document.body.content[1],
            BodyContent::Paragraph(_)
        ));
    }

    #[test]
    fn postcondition_projection_compares_paragraph_properties() {
        let body = |properties: &str| {
            let paragraph = format!("<w:p>{properties}<w:r><w:t>a</w:t></w:r></w:p>");
            story_document(
                &format!(
                    "{paragraph}<w:sdt><w:sdtContent>{paragraph}</w:sdtContent></w:sdt>\
                     <w:tbl><w:tr><w:tc>{paragraph}</w:tc></w:tr></w:tbl>"
                ),
                &[("xmlns:w".to_owned(), W_NS.to_owned())],
            )
            .unwrap()
        };
        let plain = body("");
        let default = ComparisonOptions::default();
        let ignored = ComparisonOptions {
            ignore_formatting: true,
            ..Default::default()
        };
        assert_eq!(
            postcondition_body(&plain, &default),
            postcondition_body(&body("<w:pPr><w:rPr/></w:pPr>"), &default)
        );
        for properties in [
            r#"<w:pPr><w:jc w:val="center"/></w:pPr>"#,
            "<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>",
        ] {
            let formatted = body(properties);
            // Alignment matches a paragraph whose properties changed.
            assert_eq!(
                normalized_body_with_options(&plain, &default),
                normalized_body_with_options(&formatted, &default)
            );
            let (plain_items, formatted_items) = (
                postcondition_body(&plain, &default),
                postcondition_body(&formatted, &default),
            );
            assert_eq!(plain_items.len(), 3);
            for (plain_item, formatted_item) in plain_items.iter().zip(&formatted_items) {
                assert_ne!(plain_item, formatted_item, "{properties}");
            }
            assert_eq!(
                postcondition_body(&plain, &ignored),
                postcondition_body(&formatted, &ignored)
            );
        }
    }

    #[test]
    fn word_and_character_granularity_split_only_text_content() {
        assert_eq!(
            word_fragments("élan_東京  \t—!?"),
            ["élan_東京", "  \t", "—!?"].map(str::to_owned)
        );
        let run = CT_R {
            properties: None,
            content: vec![
                RunContent::Text(CT_Text::new("A😀")),
                RunContent::Tab,
                RunContent::Break(BreakType::Page),
            ],
            extra_xml: vec![b"<x:raw xmlns:x=\"urn:f235\"/>".to_vec()],
            extra_xml_positions: vec![3],
            alt_drawings: Vec::new(),
        };
        let units = attributed_run_units(
            &[run],
            &ComparisonOptions {
                granularity: ComparisonGranularity::Character,
                ..Default::default()
            },
        );
        assert_eq!(units.len(), 4);
        assert!(matches!(&units[0].run.content[..], [RunContent::Text(text)] if text.text == "A"));
        assert!(matches!(&units[1].run.content[..], [RunContent::Text(text)] if text.text == "😀"));
        assert!(matches!(&units[2].run.content[..], [RunContent::Tab]));
        assert!(matches!(
            &units[3].run.content[..],
            [RunContent::Break(BreakType::Page)]
        ));
        assert_eq!(
            units
                .iter()
                .map(|unit| unit.run.extra_xml.len())
                .sum::<usize>(),
            1
        );
    }

    #[test]
    fn a_packed_field_result_is_read_as_one_run_per_part() {
        for w in ["w", "q"] {
            let shell = format!(r#"<{w}:r {w}:rsidR="00AB12CD"><{w}:rPr><{w}:b/></{w}:rPr>"#);
            let close = format!("</{w}:r>");
            let character = |kind: &str| format!(r#"<{w}:fldChar {w}:fldCharType="{kind}"/>"#);
            let (begin, separate, end) =
                (character("begin"), character("separate"), character("end"));
            let code = format!("<{w}:instrText>PAGE</{w}:instrText>");
            let result = format!("<{w}:t>1</{w}:t>");
            let expected = (
                format!("{shell}{begin}{code}{separate}{close}"),
                format!("{shell}{result}{close}"),
                format!("{shell}{end}{close}"),
            );
            for field in [
                format!("{shell}{begin}{code}{separate}{result}{end}{close}"),
                format!("{shell}{begin}{code}{separate}{close}{shell}{result}{end}{close}"),
                format!("{}{}{}", expected.0, expected.1, expected.2),
            ] {
                assert_eq!(complex_field_result(&field).unwrap(), expected, "{field}");
            }
            let uncached = format!("{shell}{begin}{code}{separate}{end}{close}");
            assert_eq!(
                complex_field_result(&uncached).unwrap(),
                (expected.0, String::new(), expected.2)
            );
        }
    }

    #[test]
    fn only_text_elements_become_deleted_text() {
        assert_eq!(
            deleted_text_xml(
                r#"<w:r><w:t>a</w:t><w:tab/><w:t xml:space="preserve"> b </w:t><w:t/><w:br w:type="page"/></w:r><w:ffData><w:textInput><w:maxLength w:val="4"/></w:textInput></w:ffData>"#
            ),
            r#"<w:r><w:delText>a</w:delText><w:tab/><w:delText xml:space="preserve"> b </w:delText><w:delText/><w:br w:type="page"/></w:r><w:ffData><w:textInput><w:maxLength w:val="4"/></w:textInput></w:ffData>"#
        );
        // Word refuses to open a deletion that holds `w:instrText`.
        assert_eq!(
            deleted_text_xml(
                r#"<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> TOC </w:instrText><w:instrTextual/></w:r>"#
            ),
            r#"<w:r><w:fldChar w:fldCharType="begin"/><w:delInstrText xml:space="preserve"> TOC </w:delInstrText><w:instrTextual/></w:r>"#
        );
    }

    #[test]
    fn an_owned_story_shell_reads_the_same_in_every_serialization() {
        let compatibility = oxml_core::xml::MC_NS;
        let word = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:mc="{compatibility}" xmlns:w="{W_NS}" mc:Ignorable="w14"><w:comment w:id="0" w:author="Ada"><w:p/></w:comment>
</w:comments>"#
        );
        let other = |author: &str| {
            format!(
                r#"<q:comments xmlns:q="{W_NS}"><!-- kept --><q:comment q:author="{author}" q:id="0"><q:p></q:p></q:comment></q:comments>"#
            )
        };
        let canonical = |xml: &str| canonical_owned_story(xml, "comment").unwrap();
        assert_eq!(canonical(&word), canonical(&other("Ada")));
        let (skeleton, owners) = canonical(&word);
        assert_eq!(skeleton.iter().filter(|token| *token == "owner").count(), 1);
        assert_eq!(owners.len(), 1);
        assert_ne!(canonical(&word), canonical(&other("Bob")));
    }

    #[test]
    fn staged_comparison_postcondition_failure_preserves_bytes_and_layout_cache() {
        let mut original = Document::new();
        original.add_paragraph("original");
        let mut edited = Document::new();
        edited.add_paragraph("edited");
        let before = original.to_bytes().unwrap();
        let layout = original.layout().unwrap();
        FAIL_AFTER_COMPARISON_STAGING.with(|fail| fail.set(true));
        let error = original
            .compare_with_options(
                &edited,
                "Ada",
                "2026-09-04T09:00:00Z",
                &ComparisonOptions::default(),
            )
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("staged comparison postcondition")
        );
        assert_eq!(original.to_bytes().unwrap(), before);
        assert!(std::sync::Arc::ptr_eq(&layout, &original.layout().unwrap()));
    }

    #[test]
    fn comparison_postcondition_diagnostic_names_one_item_without_model_bytes() {
        let actual = (
            vec!["Drawing(CT_Drawing { huge model })".to_owned()],
            Vec::new(),
        );
        let wanted = (vec!["different".to_owned()], Vec::new());
        let message =
            comparison_postcondition_error("acceptance", "edited", &actual, &wanted).to_string();

        assert_eq!(
            message,
            "comparison acceptance does not reproduce the edited stories at body story item[0]"
        );
        assert!(!message.contains("CT_Drawing"));
        assert!(message.len() < 1_000);
    }

    #[test]
    fn changed_image_payload_comparison_resolves_both_versions() {
        let original_png: &[u8] = &[
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x9e, 0xdd, 0x22,
            0x71, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
        ];
        let replacement_png: &[u8] = &[
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
        ];
        let mut original = Document::new();
        original.add_paragraph("unchanged");
        original.add_picture(
            original_png,
            "first.png",
            Length::pt(12.0),
            Length::pt(12.0),
        );
        let mut source = Document::from_bytes(&original.to_bytes().unwrap()).unwrap();
        let owner = source.doc_part_name.clone();
        let image_id = source
            .package
            .get_part_rels(&owner)
            .unwrap()
            .items
            .iter()
            .find(|relationship| relationship.rel_type == rel_types::IMAGE)
            .unwrap()
            .id
            .clone();
        let settings = source
            .package
            .get_part("/word/settings.xml")
            .unwrap()
            .to_vec();
        let mut edited = Document::from_bytes(&source.to_bytes().unwrap()).unwrap();
        edited.replace_image(&image_id, replacement_png).unwrap();
        assert_eq!(
            edited.image_data(&image_id).as_deref(),
            Some(replacement_png)
        );

        let mut redline = Document::from_bytes(&source.to_bytes().unwrap()).unwrap();
        assert!(
            redline
                .compare(&edited, "Reviewer", "2026-09-27T12:00:00Z")
                .unwrap()
                .is_empty()
        );
        let redline = redline.to_bytes().unwrap();
        let mut accepted = Document::from_bytes(&redline).unwrap();
        let mut rejected = Document::from_bytes(&redline).unwrap();
        accepted.accept_all().unwrap();
        rejected.reject_all().unwrap();
        for document in [&accepted, &rejected] {
            assert_eq!(
                document.package.get_part("/word/settings.xml"),
                Some(settings.as_slice())
            );
        }
        let accepted = Document::from_bytes(&accepted.to_bytes().unwrap()).unwrap();
        let rejected = Document::from_bytes(&rejected.to_bytes().unwrap()).unwrap();
        assert_eq!(
            accepted.image_data("rdocxComparisonImage1").as_deref(),
            Some(replacement_png)
        );
        assert_eq!(
            rejected.image_data(&image_id).as_deref(),
            Some(original_png)
        );
        assert_eq!(
            accepted.package.get_part("/word/settings.xml"),
            Some(settings.as_slice())
        );
        assert_eq!(
            rejected.package.get_part("/word/settings.xml"),
            Some(settings.as_slice())
        );
    }
    fn fx245_document(value: f64, closing: &str) -> Document {
        let mut doc = Document::new();
        doc.add_paragraph("Quarterly volume by region.");
        doc.add_chart(
            oxml_chart::ChartKind::Bar,
            Length::inches(6.0),
            Length::inches(2.6),
            &oxml_chart::ChartData {
                categories: vec!["Q1".into(), "Q2".into(), "Q3".into()],
                series: vec![
                    ("North".into(), vec![12.0, value, 15.0]),
                    ("South".into(), vec![9.0, 11.0, 10.0]),
                    ("West".into(), vec![7.0, 8.0, 13.0]),
                ],
                ..oxml_chart::ChartData::default()
            },
        )
        .unwrap();
        doc.add_paragraph(closing);
        Document::from_bytes(&doc.to_bytes().unwrap()).unwrap()
    }

    // Follow the saved graph independently of comparison normalization.
    fn fx245_chart_and_workbook(doc: &Document) -> (Vec<u8>, Vec<u8>) {
        let drawing = doc
            .document
            .body
            .content
            .iter()
            .find_map(|item| {
                let BodyContent::Paragraph(p) = item else {
                    return None;
                };
                p.runs.iter().flat_map(|r| &r.content).find_map(|content| {
                    let RunContent::Drawing(drawing) = content else {
                        return None;
                    };
                    drawing
                        .inline
                        .as_ref()
                        .and_then(|inline| inline.chart_rel_id.clone())
                })
            })
            .expect("one retained chart drawing");
        let edge = doc
            .package
            .get_part_rels(&doc.doc_part_name)
            .unwrap()
            .get_by_id(&drawing)
            .unwrap();
        assert_eq!(edge.rel_type, rel_types::CHART);
        let chart_name = OpcPackage::resolve_rel_target(&doc.doc_part_name, &edge.target);
        let chart = doc.package.get_part(&chart_name).unwrap().to_vec();
        let workbook = doc
            .package
            .get_part_rels(&chart_name)
            .unwrap()
            .items
            .iter()
            .find(|edge| edge.rel_type == rel_types::PACKAGE)
            .unwrap();
        let workbook_name = OpcPackage::resolve_rel_target(&chart_name, &workbook.target);
        (
            chart,
            doc.package.get_part(&workbook_name).unwrap().to_vec(),
        )
    }

    fn fx245_check(value: f64, closing: &str, granularity: ComparisonGranularity) {
        let mut original = fx245_document(14.0, "End of report.");
        let mut edited = fx245_document(value, closing);
        let before = original.to_bytes().unwrap();
        let edited_before = edited.to_bytes().unwrap();
        let original_resources = fx245_chart_and_workbook(&original);
        let edited_resources = fx245_chart_and_workbook(&edited);
        original
            .compare_with_options(
                &edited,
                "Reviewer",
                "2026-01-01T00:00:00Z",
                &ComparisonOptions {
                    granularity,
                    ..ComparisonOptions::default()
                },
            )
            .unwrap();
        if value != 14.0 {
            assert!(
                !original.revisions().is_empty(),
                "chart-only changes need a revision"
            );
        }
        assert_eq!(edited.to_bytes().unwrap(), edited_before);
        let redline = original.to_bytes().unwrap();
        let mut accepted = Document::from_bytes(&redline).unwrap();
        let mut rejected = Document::from_bytes(&redline).unwrap();
        accepted.accept_all().unwrap();
        rejected.reject_all().unwrap();
        let accepted = Document::from_bytes(&accepted.to_bytes().unwrap()).unwrap();
        let rejected = Document::from_bytes(&rejected.to_bytes().unwrap()).unwrap();
        assert_eq!(fx245_chart_and_workbook(&accepted), edited_resources);
        assert_eq!(fx245_chart_and_workbook(&rejected), original_resources);
        assert_eq!(accepted.text(), edited.text());
        assert_eq!(
            rejected.text(),
            Document::from_bytes(&before).unwrap().text()
        );
    }

    #[test]
    fn fx245_chart_data_changes_resolve_both_versions() {
        for granularity in [ComparisonGranularity::Word, ComparisonGranularity::Run] {
            fx245_check(18.0, "End of report.", granularity);
        }
    }

    #[test]
    fn fx245_chart_and_adjacent_text_changes_compare_together() {
        fx245_check(18.0, "End of the report.", ComparisonGranularity::Word);
    }

    #[test]
    fn fx245_unchanged_chart_does_not_create_a_revision() {
        fx245_check(14.0, "End of the report.", ComparisonGranularity::Word);
        let mut original = fx245_document(14.0, "End of report.");
        let before = original.to_bytes().unwrap();
        let edited = Document::from_bytes(&before).unwrap();
        original
            .compare(&edited, "Reviewer", "2026-01-01T00:00:00Z")
            .unwrap();
        assert!(original.revisions().is_empty());
        assert_eq!(original.to_bytes().unwrap(), before);
    }

    fn fx245_drawn_chart_names(doc: &Document) -> Vec<String> {
        doc.document
            .body
            .content
            .iter()
            .filter_map(|item| {
                let BodyContent::Paragraph(p) = item else {
                    return None;
                };
                p.runs.iter().flat_map(|r| &r.content).find_map(|content| {
                    let RunContent::Drawing(drawing) = content else {
                        return None;
                    };
                    let id = drawing.inline.as_ref()?.chart_rel_id.as_ref()?;
                    let edge = doc
                        .package
                        .get_part_rels(&doc.doc_part_name)?
                        .get_by_id(id)?;
                    Some(OpcPackage::resolve_rel_target(
                        &doc.doc_part_name,
                        &edge.target,
                    ))
                })
            })
            .collect()
    }

    fn fx245_chart_name(doc: &Document) -> String {
        fx245_drawn_chart_names(doc)
            .into_iter()
            .next()
            .expect("one surviving chart")
    }

    fn fx245_resolve_pair(mut original: Document, mut edited: Document) -> (Document, Document) {
        let original_resources = fx245_chart_and_workbook(&original);
        let edited_resources = fx245_chart_and_workbook(&edited);
        let edited_before = edited.to_bytes().unwrap();
        original
            .compare(&edited, "Reviewer", "2026-01-01T00:00:00Z")
            .unwrap();
        assert!(!original.revisions().is_empty());
        assert_eq!(edited.to_bytes().unwrap(), edited_before);
        let redline = original.to_bytes().unwrap();
        let mut accepted = Document::from_bytes(&redline).unwrap();
        let mut rejected = Document::from_bytes(&redline).unwrap();
        accepted.accept_all().unwrap();
        rejected.reject_all().unwrap();
        let accepted = Document::from_bytes(&accepted.to_bytes().unwrap()).unwrap();
        let rejected = Document::from_bytes(&rejected.to_bytes().unwrap()).unwrap();
        assert_eq!(fx245_chart_and_workbook(&accepted), edited_resources);
        assert_eq!(fx245_chart_and_workbook(&rejected), original_resources);
        (accepted, rejected)
    }

    #[test]
    fn fx245_workbook_only_changes_remain_owned() {
        use std::io::{Cursor, Read, Write};
        let mut original = fx245_document(14.0, "End of report.");
        let mut edited = Document::from_bytes(&original.to_bytes().unwrap()).unwrap();
        let chart = fx245_chart_name(&edited);
        let workbook = edited
            .package
            .get_part_rels(&chart)
            .unwrap()
            .items
            .iter()
            .find(|edge| edge.rel_type == rel_types::PACKAGE)
            .unwrap();
        let name = OpcPackage::resolve_rel_target(&chart, &workbook.target);
        let mut archive =
            zip::ZipArchive::new(Cursor::new(edited.package.get_part(&name).unwrap())).unwrap();
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let mut changed = false;
        for index in 0..archive.len() {
            let mut member = archive.by_index(index).unwrap();
            let mut bytes = Vec::new();
            member.read_to_end(&mut bytes).unwrap();
            let name = member.name().unwrap().into_owned();
            if name == "xl/worksheets/sheet1.xml" {
                let xml = String::from_utf8(bytes).unwrap();
                assert!(xml.contains("<v>14</v>"));
                bytes = xml.replace("<v>14</v>", "<v>18</v>").into_bytes();
                changed = true;
            }
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(&bytes).unwrap();
        }
        assert!(changed);
        edited
            .package
            .set_part(&name, writer.finish().unwrap().into_inner());
        assert_eq!(
            fx245_chart_and_workbook(&original).0,
            fx245_chart_and_workbook(&edited).0
        );
        assert_ne!(
            fx245_chart_and_workbook(&original).1,
            fx245_chart_and_workbook(&edited).1
        );
        fx245_resolve_pair(original, edited);
    }

    #[test]
    fn fx245_shared_chart_cycles_and_opaque_sidecars_survive_collision_remapping() {
        let mut original = fx245_document(14.0, "End of report.");
        original
            .add_chart(
                oxml_chart::ChartKind::Bar,
                Length::inches(6.0),
                Length::inches(2.6),
                &oxml_chart::ChartData {
                    categories: vec!["Q1".into(), "Q2".into(), "Q3".into()],
                    series: vec![("North".into(), vec![12.0, 14.0, 15.0])],
                    ..oxml_chart::ChartData::default()
                },
            )
            .unwrap();
        let chart = fx245_chart_name(&original);
        let owner = original.doc_part_name.clone();
        original
            .package
            .get_part_rels_mut(&owner)
            .unwrap()
            .items
            .iter_mut()
            .filter(|edge| edge.rel_type == rel_types::CHART)
            .nth(1)
            .unwrap()
            .target = chart.clone();
        let sidecar = "/word/sidecar.xml";
        let opaque = br#"<x:opaque xmlns:x="urn:producer"><!--keep--><?keep exact?><x:v a='one &amp; two'/></x:opaque>"#;
        original.package.set_part(sidecar, opaque.to_vec());
        original
            .package
            .content_types
            .add_override(sidecar, "application/xml");
        original
            .package
            .get_or_create_part_rels(&chart)
            .add_with_id("sidecar", "urn:sidecar", sidecar);
        original
            .package
            .get_or_create_part_rels(sidecar)
            .add_with_id("back", "urn:cycle", &chart.to_ascii_uppercase());
        original
            .package
            .get_or_create_part_rels(sidecar)
            .add_external("urn:external", "https://example.invalid/no-fetch");
        let collision = "/word/charts/rdocxComparison1-chart1.xml";
        original
            .package
            .set_part(collision, b"unrelated collision".to_vec());
        original
            .package
            .content_types
            .add_override(collision, "application/xml");
        original.identifiers.preserve_fragment_part(collision);
        let bytes = original.to_bytes().unwrap();
        let mut edited = Document::from_bytes(&bytes).unwrap();
        let xml = String::from_utf8(edited.package.get_part(&chart).unwrap().to_vec()).unwrap();
        assert!(xml.contains("<c:v>14</c:v>"));
        edited.package.set_part(
            &chart,
            xml.replace("<c:v>14</c:v>", "<c:v>18</c:v>").into_bytes(),
        );
        let original_parts = original.package.parts.clone();
        let mut redline = Document::from_bytes(&original.to_bytes().unwrap()).unwrap();
        redline
            .compare(&edited, "Reviewer", "2026-01-01T00:00:00Z")
            .unwrap();
        let redline = Document::from_bytes(&redline.to_bytes().unwrap()).unwrap();
        let added = redline
            .package
            .parts
            .iter()
            .filter(|(name, _)| !original_parts.contains_key(*name))
            .collect::<Vec<_>>();
        assert_eq!(
            added.len(),
            3,
            "one chart, workbook and sidecar copy despite repeated drawing roots"
        );
        assert_eq!(
            added.iter().map(|(_, bytes)| bytes.len()).sum::<usize>(),
            fx245_chart_and_workbook(&edited).0.len()
                + fx245_chart_and_workbook(&edited).1.len()
                + opaque.len()
        );
        let (accepted, rejected) = fx245_resolve_pair(original, edited);
        for doc in [&accepted, &rejected] {
            let ids = doc
                .document
                .body
                .content
                .iter()
                .filter_map(|item| {
                    let BodyContent::Paragraph(p) = item else {
                        return None;
                    };
                    p.runs.iter().flat_map(|r| &r.content).find_map(|content| {
                        let RunContent::Drawing(drawing) = content else {
                            return None;
                        };
                        drawing
                            .inline
                            .as_ref()
                            .and_then(|inline| inline.chart_rel_id.clone())
                    })
                })
                .collect::<Vec<_>>();
            assert_eq!(ids.len(), 2, "both originally shared drawings survive");
            let roots = fx245_drawn_chart_names(doc);
            assert_eq!(roots[0], roots[1], "physical chart sharing survives");
            let expected_chart = fx245_chart_and_workbook(doc).0;
            for id in ids {
                let edge = doc
                    .package
                    .get_part_rels(&doc.doc_part_name)
                    .unwrap()
                    .get_by_id(&id)
                    .unwrap();
                let name = OpcPackage::resolve_rel_target(&doc.doc_part_name, &edge.target);
                assert_eq!(doc.package.get_part(&name).unwrap(), expected_chart);
            }
            assert_eq!(
                doc.package.get_part(collision),
                Some(b"unrelated collision".as_slice())
            );
            let root = fx245_chart_name(doc);
            let rels = doc.package.get_part_rels(&root).unwrap();
            let edge = rels.get_by_id("sidecar").unwrap();
            let copied = OpcPackage::resolve_rel_target(&root, &edge.target);
            assert_eq!(doc.package.get_part(&copied), Some(opaque.as_slice()));
            let back = doc
                .package
                .get_part_rels(&copied)
                .unwrap()
                .get_by_id("back")
                .unwrap();
            assert_eq!(
                OpcPackage::resolve_rel_target(&copied, &back.target).to_ascii_lowercase(),
                root.to_ascii_lowercase()
            );
            assert!(
                doc.package
                    .get_part_rels(&copied)
                    .unwrap()
                    .items
                    .iter()
                    .any(|edge| edge.target == "https://example.invalid/no-fetch"
                        && edge.target_mode.as_deref() == Some("External"))
            );
        }
    }

    #[test]
    fn fx245_invalid_chart_closure_refuses_atomically() {
        for invalid in [
            "missing",
            "duplicate",
            "depth",
            "shortcut-depth",
            "external",
        ] {
            let mut original = fx245_document(14.0, "End of report.");
            let mut edited = fx245_document(18.0, "End of report.");
            let chart = fx245_chart_name(&edited);
            match invalid {
                "missing" => {
                    edited
                        .package
                        .get_or_create_part_rels(&chart)
                        .add("urn:sidecar", "/missing.xml");
                }
                "duplicate" => {
                    let edge = edited.package.get_part_rels(&chart).unwrap().items[0].clone();
                    edited
                        .package
                        .get_or_create_part_rels(&chart)
                        .items
                        .push(edge);
                }
                "depth" | "shortcut-depth" => {
                    let root = chart.clone();
                    let mut owner = chart;
                    for index in 0..130 {
                        let name = format!("/word/chain{index}.xml");
                        edited.package.set_part(&name, b"<opaque/>".to_vec());
                        edited
                            .package
                            .content_types
                            .add_override(&name, "application/xml");
                        edited.package.get_or_create_part_rels(&owner).add_with_id(
                            "0-chain",
                            "urn:chain",
                            &name,
                        );
                        if invalid == "shortcut-depth" {
                            edited.package.get_or_create_part_rels(&root).add_with_id(
                                &format!("shortcut{index}"),
                                "urn:shortcut",
                                &name,
                            );
                        }
                        owner = name;
                    }
                }
                "external" => {
                    let owner = edited.doc_part_name.clone();
                    let edge = edited
                        .package
                        .get_part_rels_mut(&owner)
                        .unwrap()
                        .items
                        .iter_mut()
                        .find(|edge| edge.rel_type == rel_types::CHART)
                        .unwrap();
                    edge.target_mode = Some("External".to_owned());
                    edge.target = "https://example.invalid/chart.xml".to_owned();
                }
                _ => unreachable!(),
            }
            let original_before = original.to_bytes().unwrap();
            let edited_before = edited.package.clone();
            assert!(
                original
                    .compare(&edited, "Reviewer", "2026-01-01T00:00:00Z")
                    .is_err(),
                "{invalid}"
            );
            assert_eq!(original.to_bytes().unwrap(), original_before, "{invalid}");
            assert_eq!(edited.package.parts, edited_before.parts, "{invalid}");
            assert_eq!(
                edited.package.content_types.to_xml().unwrap(),
                edited_before.content_types.to_xml().unwrap(),
                "{invalid}"
            );
            for (owner, relationships) in &edited_before.part_rels {
                assert_eq!(
                    format!("{:?}", edited.package.get_part_rels(owner).unwrap()),
                    format!("{relationships:?}"),
                    "{invalid}"
                );
            }
        }
    }
    #[test]
    fn fx245_distinct_chart_roots_share_an_aggregate_resource_limit() {
        let mut original = fx245_document(14.0, "End of report.");
        let mut edited = fx245_document(18.0, "End of report.");
        edited
            .add_chart(
                oxml_chart::ChartKind::Bar,
                Length::inches(6.0),
                Length::inches(2.6),
                &oxml_chart::ChartData {
                    categories: vec!["Q1".into()],
                    series: vec![("Other".into(), vec![23.0])],
                    ..oxml_chart::ChartData::default()
                },
            )
            .unwrap();
        let roots = fx245_drawn_chart_names(&edited);
        assert_eq!(roots.len(), 2);
        for (root_index, root) in roots.iter().enumerate() {
            for index in 0..2048 {
                let name = format!("/word/leaf-{root_index}-{index}.xml");
                edited.package.set_part(&name, b"<opaque/>".to_vec());
                edited
                    .package
                    .content_types
                    .add_override(&name, "application/xml");
                edited.package.get_or_create_part_rels(root).add_with_id(
                    &format!("leaf{index}"),
                    "urn:leaf",
                    &name,
                );
            }
        }
        let before = original.to_bytes().unwrap();
        let edited_parts = edited.package.parts.clone();
        let error = original
            .compare(&edited, "Reviewer", "2026-01-01T00:00:00Z")
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("aggregate chart resources exceed part limit"),
            "{error}"
        );
        assert_eq!(original.to_bytes().unwrap(), before);
        assert_eq!(edited.package.parts, edited_parts);
    }
    #[test]
    fn fx245_existing_fragment_copy_keeps_distinct_case_alias_destinations() {
        let mut source = Document::new();
        let name = "/word/payload.xml";
        let alias = "/WORD/PAYLOAD.XML";
        source.package.set_part(name, b"<opaque/>".to_vec());
        source
            .package
            .content_types
            .add_override(name, "application/xml");
        source
            .package
            .get_or_create_part_rels(name)
            .add_with_id("cycle", "urn:cycle", alias);
        let mapping = std::collections::BTreeMap::from([
            (name.to_owned(), "/word/copy-one.xml".to_owned()),
            (alias.to_owned(), "/word/copy-two.xml".to_owned()),
        ]);
        let mut destination = Document::new();
        let mut copied = std::collections::HashSet::new();
        crate::field::copy_fragment_part(
            &source.package,
            &mut destination,
            name,
            &mapping,
            &mut copied,
        )
        .unwrap();
        for target in mapping.values() {
            assert_eq!(
                destination.package.get_part(target),
                Some(b"<opaque/>".as_slice())
            );
            let edge = destination
                .package
                .get_part_rels(target)
                .unwrap()
                .get_by_id("cycle")
                .unwrap();
            let actual = OpcPackage::resolve_rel_target(target, &edge.target);
            assert_eq!(actual, "/word/copy-two.xml");
            assert!(destination.package.contains_part(&actual));
        }
    }
}
