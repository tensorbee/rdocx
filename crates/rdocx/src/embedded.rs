use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::io::Cursor;
use std::ops::Range;

use oxml_opc::OpcPackage;
use oxml_opc::relationship::{Relationship, rel_types};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest, Sha256};

use crate::{Document, Error, Result};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W_NS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const O_NS: &str = "urn:schemas-microsoft-com:office:office";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_R_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const ACTIVEX_NS: &str = "http://schemas.microsoft.com/office/2006/activeX";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS: &str = "http://www.w3.org/2000/xmlns/";
const VML_NS: &str = "urn:schemas-microsoft-com:vml";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPG_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPC_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const DSIG_NS: &str = "http://www.w3.org/2000/09/xmldsig#";
const OPC_SIGNATURE_NS: &str = "http://schemas.openxmlformats.org/package/2006/digital-signature";
const ACTIVEX_PROPERTIES_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const VBA_SIGNATURE_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaProjectSignature";
const VBA_AGILE_SIGNATURE_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaProjectSignatureAgile";
const PACKAGE_SIGNATURE_ORIGIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-origin";
const PACKAGE_SIGNATURE_XML_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
const INVALIDATED_PACKAGE_SIGNATURE: &str = "urn:rdocx:relationships/invalidated-package-signature";
const INVALIDATED_VBA_SIGNATURE: &str = "urn:rdocx:relationships/invalidated-vba-project-signature";

/// Executable or host-activated content owned by a Word package.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum EmbeddedContentKind {
    OleObject,
    ActiveXControl,
    VbaProject,
}

/// Presence and known mutation state of signature evidence for one payload.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum EmbeddedSignatureState {
    Absent,
    Present,
    Invalidated,
}

/// Explicit handling for signature evidence invalidated by an embedded mutation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum EmbeddedMutationPolicy {
    PreserveInvalidatedSignatures,
    RemoveInvalidatedSignatures,
}

/// Stable audit facts for one relationship-owned executable payload.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EmbeddedContentInfo {
    pub kind: EmbeddedContentKind,
    pub source_part: String,
    pub relationship_id: String,
    pub target_part: String,
    pub content_type: String,
    pub byte_len: usize,
    pub sha256: [u8; 32],
    pub signature_state: EmbeddedSignatureState,
}

#[derive(Clone, Debug)]
struct OwnedEmbeddedContent {
    info: EmbeddedContentInfo,
    control_owners: Vec<(String, String)>,
}

#[derive(Clone, Copy, Debug)]
struct SignatureContext {
    package_present: bool,
    package_invalidated: bool,
    attached: bool,
    attached_invalidated: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum XmlReferenceKind {
    Ole,
    Control,
}

#[derive(Clone, Debug)]
struct XmlReference {
    relationship_id: String,
    range: Range<usize>,
}

#[derive(Debug)]
struct OpenNode {
    start: usize,
    is_document_element: bool,
    story_path_kind: StoryPathKind,
    is_paragraph: bool,
    is_run: bool,
    is_object: bool,
    is_control_owner: bool,
    mc_path_kind: McPathKind,
    mc_branch_valid: bool,
    mc_container_state: Option<McContainerState>,
    mc_rules_valid: bool,
    ignorable_namespaces: BTreeSet<Vec<u8>>,
    has_invalid_mc_descendant: bool,
    text_box_path_kind: TextBoxPathKind,
    run_owner_kind: RunOwnerKind,
    object_relationship_ids: Vec<String>,
    object_child_count: usize,
    control_relationship_id: Option<String>,
    control_child_count: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum RunOwnerKind {
    Other,
    Container,
    StructuredDocumentTag,
    StructuredDocumentTagContent,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum McPathKind {
    Other,
    AlternateContent,
    Choice,
    Fallback,
}

#[derive(Clone, Copy, Debug)]
struct McContainerState {
    choice_count: usize,
    fallback_seen: bool,
    grammar_valid: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TextBoxPathKind {
    Other,
    LegacyPicture,
    LegacyObject,
    VmlGroup,
    VmlShape,
    VmlTextBox,
    Drawing,
    WordprocessingDrawing,
    Graphic,
    GraphicDataShape,
    GraphicDataGroup,
    GraphicDataCanvas,
    WordprocessingGroup,
    WordprocessingCanvas,
    WordprocessingShape,
    WordprocessingTextBox,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StoryRootKind {
    Document,
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
    Glossary,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TextBoxNamespaceKind {
    Other,
    Word,
    Vml,
    WordprocessingDrawing,
    Drawing,
    WordprocessingGroup,
    WordprocessingShape,
    WordprocessingCanvas,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StoryPathKind {
    Other,
    DocumentRoot,
    HeaderFooterRoot,
    FootnotesRoot,
    EndnotesRoot,
    CommentsRoot,
    GlossaryRoot,
    TextBoxRoot,
    Body,
    Footnote,
    Endnote,
    Comment,
    DocParts,
    DocPart,
    DocPartBody,
    CustomXml,
    StructuredDocumentTag,
    StructuredDocumentTagContent,
    Table,
    TableRow,
    TableCell,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StoryContentKind {
    Block,
    TableRow,
    TableCell,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ExpectedStoryChild {
    Body,
    Footnote,
    Endnote,
    Comment,
    DocParts,
    DocPart,
    DocPartBody,
    Content(StoryContentKind),
    StructuredDocumentTagContent(StoryContentKind),
}

impl Document {
    /// Inventories relationship-owned OLE, ActiveX, and VBA payloads without decoding them.
    pub fn embedded_content(&self) -> Result<Vec<EmbeddedContentInfo>> {
        Ok(self
            .consolidated_embedded_candidate()?
            .owned_embedded_content_in_package()?
            .into_iter()
            .map(|owned| owned.info)
            .collect())
    }

    /// Extracts one relationship-owned payload byte for byte.
    pub fn extract_embedded_content(
        &self,
        source_part: &str,
        relationship_id: &str,
    ) -> Result<Vec<u8>> {
        validate_identity_source(source_part)?;
        let staged = self.consolidated_embedded_candidate()?;
        let owned = staged.find_embedded_content_in_package(source_part, relationship_id)?;
        Ok(required_part(&staged.package, &owned.info.target_part)?.to_vec())
    }

    /// Replaces one opaque payload while retaining its relationship identity and part metadata.
    pub fn replace_embedded_content(
        &mut self,
        source_part: &str,
        relationship_id: &str,
        bytes: &[u8],
        policy: EmbeddedMutationPolicy,
    ) -> Result<EmbeddedContentInfo> {
        validate_identity_source(source_part)?;
        let mut staged = self.consolidated_embedded_candidate()?;
        let selected = staged.find_embedded_content_in_package(source_part, relationship_id)?;
        staged
            .package
            .set_part(&selected.info.target_part, bytes.to_vec());
        staged.invalidate_embedded_signatures(&selected, policy)?;
        self.commit_embedded_candidate(staged)?;
        Ok(self
            .consolidated_embedded_candidate()?
            .find_embedded_content_in_package(source_part, relationship_id)?
            .info)
    }

    /// Removes one logical embedded object and only its newly unreachable owned candidates.
    pub fn remove_embedded_content(
        &mut self,
        source_part: &str,
        relationship_id: &str,
        policy: EmbeddedMutationPolicy,
    ) -> Result<()> {
        validate_identity_source(source_part)?;
        let mut staged = self.consolidated_embedded_candidate()?;
        let selected = staged.find_embedded_content_in_package(source_part, relationship_id)?;
        staged.remove_embedded_content_in_place(&selected, policy)?;
        self.commit_embedded_candidate(staged)
    }

    fn consolidated_embedded_candidate(&self) -> Result<Self> {
        let mut staged = self.clone_for_staging();
        staged.package_signatures_invalidated |=
            staged.retained_package_signature_would_be_invalidated()?;
        staged.flush_to_package()?;
        Ok(staged)
    }

    fn owned_embedded_content_in_package(&self) -> Result<Vec<OwnedEmbeddedContent>> {
        let package_signature_present = has_package_signature(&self.package)?;
        let package_signature_invalidated = self.package_signatures_invalidated
            || package_signature_invalidation_marked(&self.package)?;
        let mut found = BTreeMap::<(String, String), OwnedEmbeddedContent>::new();
        let mut controls = BTreeMap::<String, Vec<(String, String)>>::new();
        let mut source_parts = BTreeSet::<String>::new();

        for (source_part, relationships) in &self.package.part_rels {
            ensure_unique_relationship_ids(source_part, &relationships.items)?;
            let has_owner_candidate = relationships.items.iter().any(|relationship| {
                matches!(
                    relationship.rel_type.as_str(),
                    rel_types::OLE_OBJECT
                        | rel_types::STRICT_OLE_OBJECT
                        | rel_types::CONTROL
                        | rel_types::STRICT_CONTROL
                )
            });
            let has_xml_content_type = self
                .package
                .content_types
                .content_type_for(source_part)
                .is_some_and(|content_type| content_type.ends_with("xml"));
            if !has_owner_candidate && !has_xml_content_type && source_part != &self.doc_part_name {
                continue;
            }
            validate_identity_source(source_part)?;
            source_parts.insert(source_part.clone());
        }
        for source_part in self.package.parts.keys() {
            if self.word_story_root_kind(source_part).is_some() {
                validate_identity_source(source_part)?;
                source_parts.insert(source_part.clone());
            }
        }

        for source_part in source_parts {
            let xml = required_part(&self.package, &source_part)?;
            let Some(story_root_kind) = self.word_story_root_kind(&source_part) else {
                continue;
            };
            let ole_references = xml_references(xml, XmlReferenceKind::Ole, story_root_kind)?;
            let control_references =
                xml_references(xml, XmlReferenceKind::Control, story_root_kind)?;
            reject_owner_range_collisions(&source_part, &ole_references, &control_references)?;
            for reference in ole_references {
                let relationship =
                    required_relationship(&self.package, &source_part, &reference.relationship_id)?;
                require_relationship_kind(
                    &source_part,
                    relationship,
                    &[rel_types::OLE_OBJECT, rel_types::STRICT_OLE_OBJECT],
                    "OLE object",
                )?;
                let target = safe_internal_target(&source_part, relationship)?;
                let info = embedded_info(
                    self,
                    EmbeddedContentKind::OleObject,
                    &source_part,
                    relationship,
                    target,
                    SignatureContext {
                        package_present: package_signature_present,
                        package_invalidated: package_signature_invalidated,
                        attached: false,
                        attached_invalidated: false,
                    },
                )?;
                insert_unique(&mut found, info, Vec::new())?;
            }
            for reference in control_references {
                let relationship =
                    required_relationship(&self.package, &source_part, &reference.relationship_id)?;
                require_relationship_kind(
                    &source_part,
                    relationship,
                    &[rel_types::CONTROL, rel_types::STRICT_CONTROL],
                    "ActiveX control properties",
                )?;
                let control_part = safe_internal_target(&source_part, relationship)?;
                required_part(&self.package, &control_part)?;
                controls
                    .entry(control_part)
                    .or_default()
                    .push((source_part.clone(), reference.relationship_id));
            }
        }

        for (control_part, owners) in controls {
            require_exact_content_type(
                &self.package,
                &control_part,
                ACTIVEX_PROPERTIES_CONTENT_TYPE,
                "ActiveX properties",
            )?;
            let properties = required_part(&self.package, &control_part)?;
            let binary_relationship_id = active_x_binary_relationship_id(properties)?.ok_or_else(
                || {
                    invalid(
                        "inventory embedded content",
                        format!(
                            "{control_part}: ActiveX properties root has no relationship-owned binary"
                        ),
                    )
                },
            )?;
            let control_relationships =
                self.package.get_part_rels(&control_part).ok_or_else(|| {
                    invalid(
                        "inventory embedded content",
                        format!("{control_part}: ActiveX properties have no relationship set"),
                    )
                })?;
            ensure_unique_relationship_ids(&control_part, &control_relationships.items)?;
            let binary_relationships = control_relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::ACTIVEX_CONTROL_BINARY)
                .collect::<Vec<_>>();
            if binary_relationships.len() != 1 {
                return Err(invalid(
                    "inventory embedded content",
                    format!(
                        "{control_part}: found {} ActiveX binary relationships, expected exactly one",
                        binary_relationships.len()
                    ),
                ));
            }
            let relationship =
                required_relationship(&self.package, &control_part, &binary_relationship_id)?;
            require_relationship_kind(
                &control_part,
                relationship,
                &[rel_types::ACTIVEX_CONTROL_BINARY],
                "ActiveX binary",
            )?;
            let target = safe_internal_target(&control_part, relationship)?;
            let info = embedded_info(
                self,
                EmbeddedContentKind::ActiveXControl,
                &control_part,
                relationship,
                target,
                SignatureContext {
                    package_present: package_signature_present,
                    package_invalidated: package_signature_invalidated,
                    attached: false,
                    attached_invalidated: false,
                },
            )?;
            insert_unique(&mut found, info, owners)?;
        }

        if let Some(relationships) = self.package.get_part_rels(&self.doc_part_name) {
            ensure_unique_relationship_ids(&self.doc_part_name, &relationships.items)?;
            let vba_projects = relationships
                .items
                .iter()
                .filter(|relationship| relationship.rel_type == rel_types::VBA_PROJECT)
                .collect::<Vec<_>>();
            if vba_projects.len() > 1 {
                return Err(invalid(
                    "inventory embedded content",
                    format!(
                        "{}: found {} VBA project relationships, expected at most one",
                        self.doc_part_name,
                        vba_projects.len()
                    ),
                ));
            }
            for relationship in vba_projects {
                let target = safe_internal_target(&self.doc_part_name, relationship)?;
                let attached_signature = attached_vba_signature_state(&self.package, &target)?;
                let info = embedded_info(
                    self,
                    EmbeddedContentKind::VbaProject,
                    &self.doc_part_name,
                    relationship,
                    target,
                    SignatureContext {
                        package_present: package_signature_present,
                        package_invalidated: package_signature_invalidated,
                        attached: attached_signature.is_some(),
                        attached_invalidated: attached_signature == Some(true),
                    },
                )?;
                insert_unique(&mut found, info, Vec::new())?;
            }
        }

        Ok(found.into_values().collect())
    }

    fn word_story_root_kind(&self, part_name: &str) -> Option<StoryRootKind> {
        if part_name == self.doc_part_name {
            return Some(StoryRootKind::Document);
        }
        match self.package.content_types.content_type_for(part_name)? {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" => {
                Some(StoryRootKind::Header)
            }
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" => {
                Some(StoryRootKind::Footer)
            }
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml" => {
                Some(StoryRootKind::Footnotes)
            }
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" => {
                Some(StoryRootKind::Endnotes)
            }
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" => {
                Some(StoryRootKind::Comments)
            }
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml" => {
                Some(StoryRootKind::Glossary)
            }
            _ => None,
        }
    }

    fn find_embedded_content_in_package(
        &self,
        source_part: &str,
        relationship_id: &str,
    ) -> Result<OwnedEmbeddedContent> {
        self.owned_embedded_content_in_package()?
            .into_iter()
            .find(|owned| {
                owned.info.source_part == source_part
                    && owned.info.relationship_id == relationship_id
            })
            .ok_or_else(|| {
                invalid(
                    "resolve embedded content",
                    format!(
                        "{source_part}: relationship {relationship_id} is not an owned embedded payload"
                    ),
                )
            })
    }

    fn invalidate_embedded_signatures(
        &mut self,
        selected: &OwnedEmbeddedContent,
        policy: EmbeddedMutationPolicy,
    ) -> Result<()> {
        let identity = (
            selected.info.source_part.clone(),
            selected.info.relationship_id.clone(),
        );
        if selected.info.signature_state != EmbeddedSignatureState::Absent {
            self.embedded_invalidated_signatures.insert(identity);
        }
        if has_package_signature(&self.package)? {
            self.package_signatures_invalidated = true;
        }
        if policy == EmbeddedMutationPolicy::PreserveInvalidatedSignatures
            && selected.info.kind == EmbeddedContentKind::VbaProject
        {
            mark_vba_signature_invalidated(&mut self.package, &selected.info.target_part)?;
        }
        if policy == EmbeddedMutationPolicy::RemoveInvalidatedSignatures {
            remove_package_signatures(&mut self.package)?;
            if selected.info.kind == EmbeddedContentKind::VbaProject {
                remove_vba_signatures(&mut self.package, &selected.info.target_part)?;
            }
            self.embedded_invalidated_signatures.remove(&(
                selected.info.source_part.clone(),
                selected.info.relationship_id.clone(),
            ));
            self.package_signatures_invalidated = false;
        }
        Ok(())
    }

    fn remove_embedded_content_in_place(
        &mut self,
        selected: &OwnedEmbeddedContent,
        policy: EmbeddedMutationPolicy,
    ) -> Result<()> {
        if selected.info.kind == EmbeddedContentKind::VbaProject
            && relationship_target_is_reachable_except(
                &self.package,
                &selected.info.target_part,
                Some((&selected.info.source_part, &selected.info.relationship_id)),
            )?
        {
            return Err(invalid(
                "remove embedded content",
                format!(
                    "{}: VBA project target {} is shared by another relationship",
                    selected.info.source_part, selected.info.target_part
                ),
            ));
        }
        self.invalidate_embedded_signatures(selected, policy)?;
        match selected.info.kind {
            EmbeddedContentKind::OleObject => {
                let story_root_kind = self
                    .word_story_root_kind(&selected.info.source_part)
                    .ok_or_else(|| {
                        invalid(
                            "remove embedded content",
                            format!("{}: unsupported Word story part", selected.info.source_part),
                        )
                    })?;
                remove_xml_reference(
                    &mut self.package,
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                    XmlReferenceKind::Ole,
                    story_root_kind,
                )?;
                remove_relationship(
                    &mut self.package,
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                )?;
                delete_if_unreachable(&mut self.package, &selected.info.target_part)?;
            }
            EmbeddedContentKind::ActiveXControl => {
                for (owner_part, owner_relationship_id) in &selected.control_owners {
                    let story_root_kind =
                        self.word_story_root_kind(owner_part).ok_or_else(|| {
                            invalid(
                                "remove embedded content",
                                format!("{owner_part}: unsupported Word story part"),
                            )
                        })?;
                    remove_xml_reference(
                        &mut self.package,
                        owner_part,
                        owner_relationship_id,
                        XmlReferenceKind::Control,
                        story_root_kind,
                    )?;
                    remove_relationship(&mut self.package, owner_part, owner_relationship_id)?;
                }
                if relationship_target_is_reachable(&self.package, &selected.info.source_part)? {
                    return Ok(());
                }
                remove_relationship(
                    &mut self.package,
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                )?;
                delete_if_unreachable(&mut self.package, &selected.info.target_part)?;
                delete_if_unreachable(&mut self.package, &selected.info.source_part)?;
            }
            EmbeddedContentKind::VbaProject => {
                if policy == EmbeddedMutationPolicy::PreserveInvalidatedSignatures {
                    retain_vba_signature_parts_as_evidence(
                        &mut self.package,
                        &selected.info.target_part,
                    );
                }
                remove_relationship(
                    &mut self.package,
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                )?;
                delete_if_unreachable(&mut self.package, &selected.info.target_part)?;
            }
        }
        Ok(())
    }

    fn commit_embedded_candidate(&mut self, mut staged: Self) -> Result<()> {
        let embedded_invalidated_signatures = staged.embedded_invalidated_signatures.clone();
        let package_signatures_invalidated = staged.package_signatures_invalidated;
        persist_invalidated_package_signature(&mut staged.package, package_signatures_invalidated)?;
        staged.owned_embedded_content_in_package()?;
        let mut output = Cursor::new(Vec::new());
        staged.package.write_to(&mut output)?;
        let mut reopened = Self::from_bytes(output.get_ref())?;
        reopened.owned_embedded_content_in_package()?;
        reopened.embedded_invalidated_signatures = embedded_invalidated_signatures;
        reopened.package_signatures_invalidated = package_signatures_invalidated;
        self.commit_staged_mutation(reopened);
        Ok(())
    }

    pub(crate) fn retained_package_signature_would_be_invalidated(&self) -> Result<bool> {
        if !has_package_signature(&self.package)? {
            return Ok(false);
        }
        let mut staged = self.clone_for_staging();
        staged.flush_to_package()?;
        if self
            .package
            .get_part(&self.doc_part_name)
            .and_then(|xml| rdocx_oxml::document::CT_Document::from_xml(xml).ok())
            .as_ref()
            == Some(&self.document)
            && let Some(original) = self.package.get_part(&self.doc_part_name)
        {
            staged
                .package
                .set_part(&self.doc_part_name, original.to_vec());
        }
        Ok(!packages_are_semantically_equal(
            &staged.package,
            &self.package,
        ))
    }
}

#[cfg(feature = "digital-signatures")]
pub(crate) fn known_invalid_package_signature_on_open(package: &OpcPackage) -> bool {
    if package_signature_invalidation_marked(package).unwrap_or(false) {
        return true;
    }
    match package.verify_signatures() {
        Ok(reports) => {
            !reports.is_empty()
                && reports
                    .iter()
                    .any(|report| !report.cryptographically_valid || !report.coverage_complete)
        }
        Err(oxml_opc::OpcError::UnsupportedSignatureAlgorithm {
            kind: "part transform chain",
            ..
        }) => true,
        Err(_) => signature_manifest_has_missing_reference(package),
    }
}

#[cfg(not(feature = "digital-signatures"))]
pub(crate) fn known_invalid_package_signature_on_open(package: &OpcPackage) -> bool {
    let marked = package_signature_invalidation_marked(package).unwrap_or(false);
    let missing = signature_manifest_has_missing_reference(package);
    marked || missing
}

pub(crate) fn persist_invalidated_package_signature(
    package: &mut OpcPackage,
    invalidated: bool,
) -> Result<()> {
    if !invalidated || package_signature_invalidation_marked(package)? {
        return Ok(());
    }
    let Some(graph) = package_signature_graph(package)? else {
        return Ok(());
    };
    let origin_part = &graph.origins[0].1;
    package.package_rels.add(
        INVALIDATED_PACKAGE_SIGNATURE,
        origin_part.strip_prefix('/').unwrap_or(origin_part),
    );
    Ok(())
}

fn embedded_info(
    document: &Document,
    kind: EmbeddedContentKind,
    source_part: &str,
    relationship: &Relationship,
    target_part: String,
    signature: SignatureContext,
) -> Result<EmbeddedContentInfo> {
    let bytes = required_part(&document.package, &target_part)?;
    let content_type = document
        .package
        .content_types
        .content_type_for(&target_part)
        .ok_or_else(|| {
            invalid(
                "inventory embedded content",
                format!("{target_part}: embedded payload has no content type"),
            )
        })?
        .to_owned();
    let identity = (source_part.to_owned(), relationship.id.clone());
    let signature_state = if !signature.package_present && !signature.attached {
        EmbeddedSignatureState::Absent
    } else if signature.package_invalidated
        || signature.attached_invalidated
        || document.embedded_invalidated_signatures.contains(&identity)
    {
        EmbeddedSignatureState::Invalidated
    } else {
        EmbeddedSignatureState::Present
    };
    Ok(EmbeddedContentInfo {
        kind,
        source_part: source_part.to_owned(),
        relationship_id: relationship.id.clone(),
        target_part,
        content_type,
        byte_len: bytes.len(),
        sha256: Sha256::digest(bytes).into(),
        signature_state,
    })
}

fn insert_unique(
    found: &mut BTreeMap<(String, String), OwnedEmbeddedContent>,
    info: EmbeddedContentInfo,
    control_owners: Vec<(String, String)>,
) -> Result<()> {
    let key = (info.source_part.clone(), info.relationship_id.clone());
    if let Some(existing) = found.get_mut(&key) {
        if existing.info.kind != info.kind || existing.info.target_part != info.target_part {
            return Err(invalid(
                "inventory embedded content",
                format!(
                    "{}: relationship {} has ambiguous executable ownership",
                    key.0, key.1
                ),
            ));
        }
        existing.control_owners.extend(control_owners);
        existing.control_owners.sort();
        existing.control_owners.dedup();
    } else {
        found.insert(
            key,
            OwnedEmbeddedContent {
                info,
                control_owners,
            },
        );
    }
    Ok(())
}

fn reject_owner_range_collisions(
    source_part: &str,
    ole_references: &[XmlReference],
    control_references: &[XmlReference],
) -> Result<()> {
    for (kind, references) in [("OLE", ole_references), ("ActiveX", control_references)] {
        if let Some((left, right)) = references.iter().enumerate().find_map(|(index, left)| {
            references[index + 1..]
                .iter()
                .find(|right| owner_ranges_overlap(&left.range, &right.range))
                .map(|right| (left, right))
        }) {
            return Err(invalid(
                "inventory embedded content",
                format!(
                    "{source_part}: {kind} relationships {} and {} claim overlapping owner XML",
                    left.relationship_id, right.relationship_id
                ),
            ));
        }
    }
    if let Some((ole, control)) = ole_references.iter().find_map(|ole| {
        control_references
            .iter()
            .find(|control| owner_ranges_overlap(&ole.range, &control.range))
            .map(|control| (ole, control))
    }) {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "{source_part}: OLE relationship {} and ActiveX relationship {} claim overlapping owner XML",
                ole.relationship_id, control.relationship_id
            ),
        ));
    }
    Ok(())
}

fn owner_ranges_overlap(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
}

fn required_relationship<'a>(
    package: &'a OpcPackage,
    source_part: &str,
    relationship_id: &str,
) -> Result<&'a Relationship> {
    let matches = package
        .get_part_rels(source_part)
        .map(|relationships| {
            relationships
                .items
                .iter()
                .filter(|relationship| relationship.id == relationship_id)
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    match matches.as_slice() {
        [relationship] => Ok(*relationship),
        [] => Err(invalid(
            "resolve embedded content",
            format!("{source_part}: relationship {relationship_id} is missing"),
        )),
        _ => Err(invalid(
            "resolve embedded content",
            format!(
                "{source_part}: relationship id {relationship_id} is ambiguous across {} entries",
                matches.len()
            ),
        )),
    }
}

fn require_relationship_kind(
    source_part: &str,
    relationship: &Relationship,
    expected: &[&str],
    label: &str,
) -> Result<()> {
    if expected.contains(&relationship.rel_type.as_str()) {
        Ok(())
    } else {
        Err(invalid(
            "inventory embedded content",
            format!(
                "{source_part}: relationship {} has type {}, expected {label}",
                relationship.id, relationship.rel_type
            ),
        ))
    }
}

fn safe_internal_target(source_part: &str, relationship: &Relationship) -> Result<String> {
    if !relationship_is_internal(source_part, relationship)? {
        return Err(invalid(
            "resolve embedded content",
            format!(
                "{source_part}: relationship {} is external",
                relationship.id
            ),
        ));
    }
    if !relationship_target_is_normalized_pack_uri(&relationship.target)
        || target_escapes_package_root(source_part, &relationship.target)
    {
        return Err(invalid(
            "resolve embedded content",
            format!(
                "{source_part}: relationship {} has an unsafe internal target {}",
                relationship.id, relationship.target
            ),
        ));
    }
    Ok(OpcPackage::resolve_rel_target(
        source_part,
        &relationship.target,
    ))
}

fn relationship_target_is_normalized_pack_uri(target: &str) -> bool {
    if target.is_empty()
        || target.ends_with('/')
        || target.contains("//")
        || target.contains(['\\', '?', '#'])
        || !target.is_ascii()
    {
        return false;
    }
    let relative_first_segment =
        (!target.starts_with('/')).then(|| target.split('/').next().unwrap_or_default());
    if relative_first_segment.is_some_and(|segment| segment.contains(':')) {
        return false;
    }
    let bytes = target.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'%' {
            let (Some(&high), Some(&low)) = (bytes.get(index + 1), bytes.get(index + 2)) else {
                return false;
            };
            if !high.is_ascii_hexdigit()
                || !low.is_ascii_hexdigit()
                || high.is_ascii_lowercase()
                || low.is_ascii_lowercase()
            {
                return false;
            }
            let decoded = (hex_value(high) << 4) | hex_value(low);
            if decoded.is_ascii_alphanumeric()
                || matches!(decoded, b'-' | b'.' | b'_' | b'~' | b'/' | b'\\')
                || decoded.is_ascii_control()
            {
                return false;
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
                    | b'@'
                    | b'/'
            ))
        {
            return false;
        }
        index += 1;
    }
    target
        .split('/')
        .filter(|segment| !segment.is_empty())
        .all(|segment| matches!(segment, "." | "..") || !segment.ends_with('.'))
}

fn hex_value(byte: u8) -> u8 {
    match byte {
        b'0'..=b'9' => byte - b'0',
        b'A'..=b'F' => byte - b'A' + 10,
        _ => 0,
    }
}

fn target_escapes_package_root(source_part: &str, target: &str) -> bool {
    let mut depth = if target.starts_with('/') {
        0
    } else {
        source_part
            .trim_start_matches('/')
            .split('/')
            .count()
            .saturating_sub(1)
    };
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." if depth == 0 => return true,
            ".." => depth -= 1,
            _ => depth += 1,
        }
    }
    false
}

fn required_part<'a>(package: &'a OpcPackage, part_name: &str) -> Result<&'a [u8]> {
    package.get_part(part_name).ok_or_else(|| {
        invalid(
            "resolve embedded content",
            format!("{part_name}: required part is missing"),
        )
    })
}

fn validate_identity_source(source_part: &str) -> Result<()> {
    if !relationship_target_is_normalized_pack_uri(source_part)
        || !source_part.starts_with('/')
        || source_part == "/"
        || source_part.ends_with('/')
        || source_part.contains("//")
        || source_part.contains('\\')
        || source_part
            .split('/')
            .any(|segment| matches!(segment, "." | ".."))
        || source_part.bytes().any(|byte| byte.is_ascii_control())
    {
        return Err(invalid(
            "resolve embedded content",
            format!("unsafe source part identity {source_part}"),
        ));
    }
    Ok(())
}

fn validate_xml_declaration(declaration: &BytesDecl<'_>, operation: &'static str) -> Result<()> {
    let version = declaration
        .xml_version()
        .map_err(|error| invalid(operation, format!("invalid XML declaration: {error}")))?;
    if version != XmlVersion::Explicit1_0 {
        return Err(invalid(
            operation,
            "XML declaration version must be 1.0".to_owned(),
        ));
    }
    let declaration_text = std::str::from_utf8(declaration.as_ref())
        .map_err(|error| invalid(operation, format!("invalid XML declaration: {error}")))?;
    let start = BytesStart::from_content(declaration_text, 3);
    let mut position = 0usize;
    for attribute in start.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(operation, format!("invalid XML declaration: {error}")))?;
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| invalid(operation, format!("invalid XML declaration: {error}")))?;
        let valid = match (position, attribute.key.as_ref()) {
            (0, b"version") => value == "1.0",
            (1, b"encoding") => value.eq_ignore_ascii_case("UTF-8"),
            (1 | 2, b"standalone") => matches!(value, "yes" | "no"),
            _ => false,
        };
        if !valid {
            return Err(invalid(
                operation,
                "XML declaration attributes are invalid, duplicated, or out of order".to_owned(),
            ));
        }
        position = position.saturating_add(1);
    }
    if position == 0 {
        return Err(invalid(
            operation,
            "XML declaration is missing its version".to_owned(),
        ));
    }
    Ok(())
}

fn validate_literal_xml_characters(xml: &[u8], operation: &'static str) -> Result<()> {
    let xml = std::str::from_utf8(xml)
        .map_err(|error| invalid(operation, format!("XML is not UTF-8: {error}")))?;
    if xml.chars().all(xml_1_0_character_is_valid) {
        return Ok(());
    }
    Err(invalid(
        operation,
        "XML contains a forbidden literal XML 1.0 character".to_owned(),
    ))
}

fn xml_1_0_character_is_valid(character: char) -> bool {
    matches!(character, '\t' | '\n' | '\r')
        || ('\u{20}'..='\u{D7FF}').contains(&character)
        || ('\u{E000}'..='\u{FFFD}').contains(&character)
        || ('\u{10000}'..='\u{10FFFF}').contains(&character)
}

fn validate_scanned_xml_event(
    reader: &NsReader<&[u8]>,
    event: &Event<'_>,
    operation: &'static str,
) -> Result<()> {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            validate_scanned_element(reader, element, operation)
        }
        Event::End(element) => {
            let element_name = element.name();
            let prefix = validate_xml_qname(element_name.as_ref(), operation)?;
            let namespace = reader.resolver().resolve_element(element_name).0;
            validate_bound_prefix(&namespace, prefix, operation).map(|_| ())
        }
        Event::PI(instruction) => validate_xml_name(instruction.target(), operation),
        _ => Ok(()),
    }
}

fn validate_scanned_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    operation: &'static str,
) -> Result<()> {
    let element_name = element.name();
    let prefix = validate_xml_qname(element_name.as_ref(), operation)?;
    if prefix == Some(b"xmlns".as_slice()) {
        return Err(invalid(
            operation,
            "XML element uses the reserved xmlns prefix".to_owned(),
        ));
    }
    let namespace = reader.resolver().resolve_element(element_name).0;
    validate_bound_prefix(&namespace, prefix, operation)?;

    let mut expanded_names = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(operation, error.to_string()))?;
        let name = attribute.key.as_ref();
        let prefix = validate_xml_qname(name, operation)?;
        if attribute.value.contains(&b'<') {
            return Err(invalid(
                operation,
                "XML attribute contains a literal less-than sign".to_owned(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            .map_err(|error| invalid(operation, error.to_string()))?;
        if !value.chars().all(xml_1_0_character_is_valid) {
            return Err(invalid(
                operation,
                "XML attribute contains a forbidden XML 1.0 character".to_owned(),
            ));
        }

        if name == b"xmlns" {
            validate_namespace_declaration(None, value.as_bytes(), operation)?;
            continue;
        }
        if prefix == Some(b"xmlns".as_slice()) {
            validate_namespace_declaration(Some(local_name(name)), value.as_bytes(), operation)?;
            continue;
        }

        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let resolved = validate_bound_prefix(&namespace, prefix, operation)?;
        if !expanded_names.insert((resolved, local.as_ref().to_vec())) {
            return Err(invalid(
                operation,
                "XML element has duplicate expanded-name attributes".to_owned(),
            ));
        }
    }
    Ok(())
}

fn validate_namespace_declaration(
    prefix: Option<&[u8]>,
    namespace: &[u8],
    operation: &'static str,
) -> Result<()> {
    let valid = match prefix {
        None => namespace != XML_NS.as_bytes() && namespace != XMLNS_NS.as_bytes(),
        Some(b"xml") => namespace == XML_NS.as_bytes(),
        Some(b"xmlns") => false,
        Some(_) => {
            !namespace.is_empty()
                && namespace != XML_NS.as_bytes()
                && namespace != XMLNS_NS.as_bytes()
        }
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            operation,
            "XML contains an invalid namespace declaration".to_owned(),
        ))
    }
}

fn validate_bound_prefix(
    namespace: &ResolveResult<'_>,
    prefix: Option<&[u8]>,
    operation: &'static str,
) -> Result<Option<Vec<u8>>> {
    match namespace {
        ResolveResult::Bound(Namespace(namespace)) => Ok(Some(namespace.to_vec())),
        ResolveResult::Unbound if prefix.is_none() => Ok(None),
        ResolveResult::Unbound => Err(invalid(
            operation,
            format!(
                "XML uses unbound namespace prefix {}",
                String::from_utf8_lossy(prefix.unwrap_or_default())
            ),
        )),
        ResolveResult::Unknown(prefix) => Err(invalid(
            operation,
            format!(
                "XML uses unbound namespace prefix {}",
                String::from_utf8_lossy(prefix)
            ),
        )),
    }
}

fn validate_xml_qname<'a>(name: &'a [u8], operation: &'static str) -> Result<Option<&'a [u8]>> {
    let name_text = std::str::from_utf8(name)
        .map_err(|error| invalid(operation, format!("invalid XML qualified name: {error}")))?;
    let mut parts = name_text.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if !xml_ncname_is_valid(first)
        || second.is_some_and(|local| !xml_ncname_is_valid(local))
        || parts.next().is_some()
    {
        return Err(invalid(
            operation,
            format!("invalid XML qualified name {name_text}"),
        ));
    }
    Ok(second.map(|_| first.as_bytes()))
}

fn validate_xml_name(name: &[u8], operation: &'static str) -> Result<()> {
    let name = std::str::from_utf8(name)
        .map_err(|error| invalid(operation, format!("invalid XML name: {error}")))?;
    let mut characters = name.chars();
    let valid = characters
        .next()
        .is_some_and(|character| character == ':' || xml_ncname_start_character(character))
        && characters.all(|character| character == ':' || xml_ncname_character(character));
    if valid {
        Ok(())
    } else {
        Err(invalid(operation, format!("invalid XML name {name}")))
    }
}

fn require_predefined_or_character_reference(reference: &BytesRef<'_>) -> Result<char> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| invalid("scan embedded XML", error.to_string()))?
    {
        if matches!(
            character,
            '\u{9}'
                | '\u{A}'
                | '\u{D}'
                | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}'
                | '\u{10000}'..='\u{10FFFF}'
        ) {
            return Ok(character);
        }
        return Err(invalid(
            "scan embedded XML",
            "character reference is not legal in XML 1.0".to_owned(),
        ));
    }
    let name = reference
        .decode()
        .map_err(|error| invalid("scan embedded XML", error.to_string()))?;
    match name.as_ref() {
        "amp" => Ok('&'),
        "lt" => Ok('<'),
        "gt" => Ok('>'),
        "apos" => Ok('\''),
        "quot" => Ok('"'),
        _ => Err(invalid(
            "scan embedded XML",
            format!("undeclared general entity reference &{name};"),
        )),
    }
}

fn remove_xml_reference(
    package: &mut OpcPackage,
    source_part: &str,
    relationship_id: &str,
    kind: XmlReferenceKind,
    story_root_kind: StoryRootKind,
) -> Result<()> {
    let xml = required_part(package, source_part)?.to_vec();
    let ranges = xml_references(&xml, kind, story_root_kind)?
        .into_iter()
        .filter(|reference| reference.relationship_id == relationship_id)
        .map(|reference| reference.range)
        .collect::<Vec<_>>();
    if ranges.is_empty() {
        return Err(invalid(
            "remove embedded content",
            format!("{source_part}: owning XML reference {relationship_id} disappeared"),
        ));
    }
    package.set_part(source_part, remove_ranges(&xml, ranges)?);
    Ok(())
}

fn remove_relationship(
    package: &mut OpcPackage,
    source_part: &str,
    relationship_id: &str,
) -> Result<()> {
    required_relationship(package, source_part, relationship_id)?;
    let relationships = package.part_rels.get_mut(source_part).ok_or_else(|| {
        invalid(
            "remove embedded content",
            format!("{source_part}: relationship set disappeared"),
        )
    })?;
    relationships
        .items
        .retain(|relationship| relationship.id != relationship_id);
    Ok(())
}

fn delete_if_unreachable(package: &mut OpcPackage, candidate: &str) -> Result<()> {
    if relationship_target_is_reachable(package, candidate)? {
        return Ok(());
    }
    package.parts.remove(candidate);
    package.part_rels.remove(candidate);
    package.content_types.overrides.remove(candidate);
    Ok(())
}

fn relationship_target_is_reachable(package: &OpcPackage, candidate: &str) -> Result<bool> {
    relationship_target_is_reachable_except(package, candidate, None)
}

fn relationship_target_is_reachable_except(
    package: &OpcPackage,
    candidate: &str,
    excluded: Option<(&str, &str)>,
) -> Result<bool> {
    for relationship in &package.package_rels.items {
        if excluded == Some(("/", relationship.id.as_str())) {
            continue;
        }
        if relationship_is_internal("/", relationship)?
            && safe_internal_target("/", relationship)? == candidate
        {
            return Ok(true);
        }
    }
    for (source_part, relationships) in &package.part_rels {
        for relationship in &relationships.items {
            if excluded == Some((source_part.as_str(), relationship.id.as_str())) {
                continue;
            }
            if relationship_is_internal(source_part, relationship)?
                && safe_internal_target(source_part, relationship)? == candidate
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn relationship_is_internal(source_part: &str, relationship: &Relationship) -> Result<bool> {
    match relationship.target_mode.as_deref() {
        None | Some("Internal") => Ok(true),
        Some("External") => Ok(false),
        Some(mode) => Err(invalid(
            "resolve embedded content",
            format!(
                "{source_part}: relationship {} has invalid target mode {mode}",
                relationship.id
            ),
        )),
    }
}

fn active_x_binary_relationship_id(xml: &[u8]) -> Result<Option<String>> {
    validate_literal_xml_characters(xml, "inventory embedded content")?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut declaration_seen = false;
    let mut prolog_content_seen = false;
    let mut relationship_id = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid("inventory embedded content", error.to_string()))?;
        let root_namespace = namespace_is(&namespace, &[ACTIVEX_NS]);
        let event = event.into_owned();
        drop(namespace);
        validate_scanned_xml_event(&reader, &event, "inventory embedded content")?;
        match event {
            Event::Start(element) if depth == 0 => {
                if root_seen || !root_namespace || local_name(element.name().as_ref()) != b"ocx" {
                    return Err(invalid(
                        "inventory embedded content",
                        "ActiveX properties must contain one ax:ocx root".to_owned(),
                    ));
                }
                relationship_id = relationship_attribute(&reader, &element)?;
                root_seen = true;
                depth = 1;
            }
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    invalid(
                        "inventory embedded content",
                        "ActiveX properties nesting is too deep".to_owned(),
                    )
                })?
            }
            Event::Empty(element) if depth == 0 => {
                if root_seen || !root_namespace || local_name(element.name().as_ref()) != b"ocx" {
                    return Err(invalid(
                        "inventory embedded content",
                        "ActiveX properties must contain one ax:ocx root".to_owned(),
                    ));
                }
                relationship_id = relationship_attribute(&reader, &element)?;
                root_seen = true;
            }
            Event::Empty(_) => {}
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    invalid(
                        "inventory embedded content",
                        "ActiveX properties have an unmatched closing element".to_owned(),
                    )
                })?;
            }
            Event::Decl(declaration) => {
                if depth != 0 || root_seen || declaration_seen || prolog_content_seen {
                    return Err(invalid(
                        "inventory embedded content",
                        "ActiveX properties contain a misplaced XML declaration".to_owned(),
                    ));
                }
                validate_xml_declaration(&declaration, "inventory embedded content")?;
                declaration_seen = true;
            }
            Event::DocType(_) => {
                return Err(invalid(
                    "inventory embedded content",
                    "ActiveX properties cannot contain a document type".to_owned(),
                ));
            }
            Event::GeneralRef(reference) if depth > 0 => {
                require_predefined_or_character_reference(&reference)?;
            }
            Event::PI(instruction) if instruction.target().eq_ignore_ascii_case(b"xml") => {
                return Err(invalid(
                    "inventory embedded content",
                    "ActiveX properties contain a reserved XML processing instruction".to_owned(),
                ));
            }
            Event::Comment(_) | Event::PI(_) if depth == 0 => {
                if !root_seen {
                    prolog_content_seen = true;
                }
            }
            Event::Text(text)
                if depth == 0 && {
                    let bytes: &[u8] = text.as_ref();
                    bytes.iter().all(u8::is_ascii_whitespace)
                } =>
            {
                if !root_seen {
                    prolog_content_seen = true;
                }
            }
            Event::Eof if root_seen && depth == 0 => return Ok(relationship_id),
            Event::Eof => {
                return Err(invalid(
                    "inventory embedded content",
                    "ActiveX properties XML ended before its root closed".to_owned(),
                ));
            }
            _ if depth > 0 => {}
            _ => {
                return Err(invalid(
                    "inventory embedded content",
                    "ActiveX properties contain content outside ax:ocx".to_owned(),
                ));
            }
        }
        buffer.clear();
    }
}

fn xml_references(
    xml: &[u8],
    kind: XmlReferenceKind,
    expected_root: StoryRootKind,
) -> Result<Vec<XmlReference>> {
    validate_literal_xml_characters(xml, "scan embedded XML")?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut nodes = Vec::<OpenNode>::new();
    let mut references = Vec::new();
    let mut invalid_mc_ranges = Vec::<Range<usize>>::new();
    let mut document_element_count = 0usize;
    let mut document_element_is_expected = false;
    let mut outside_document_content_seen = false;
    let mut declaration_seen = false;
    let mut prolog_content_seen = false;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let is_word = namespace_is(&namespace, &[W_NS, STRICT_W_NS]);
        let is_office = namespace_is(&namespace, &[O_NS]);
        let is_mc = namespace_is(&namespace, &[MC_NS]);
        let is_vml = namespace_is(&namespace, &[VML_NS]);
        let is_wordprocessing_drawing = namespace_is(&namespace, &[WP_NS]);
        let is_drawing = namespace_is(&namespace, &[A_NS]);
        let is_wordprocessing_group = namespace_is(&namespace, &[WPG_NS]);
        let is_wordprocessing_shape = namespace_is(&namespace, &[WPS_NS]);
        let is_wordprocessing_canvas = namespace_is(&namespace, &[WPC_NS]);
        let text_box_namespace_kind = if is_word {
            TextBoxNamespaceKind::Word
        } else if is_vml {
            TextBoxNamespaceKind::Vml
        } else if is_wordprocessing_drawing {
            TextBoxNamespaceKind::WordprocessingDrawing
        } else if is_drawing {
            TextBoxNamespaceKind::Drawing
        } else if is_wordprocessing_group {
            TextBoxNamespaceKind::WordprocessingGroup
        } else if is_wordprocessing_shape {
            TextBoxNamespaceKind::WordprocessingShape
        } else if is_wordprocessing_canvas {
            TextBoxNamespaceKind::WordprocessingCanvas
        } else {
            TextBoxNamespaceKind::Other
        };
        let resolved_namespace = match &namespace {
            ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
            _ => None,
        };
        let event = event.into_owned();
        drop(namespace);
        let event_end = reader.buffer_position() as usize;
        validate_scanned_xml_event(&reader, &event, "scan embedded XML")?;
        match event {
            Event::Start(element) => {
                let is_document_element = nodes.is_empty()
                    && document_element_count == 0
                    && !outside_document_content_seen;
                if nodes.is_empty() {
                    document_element_count = document_element_count.saturating_add(1);
                }
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                if is_document_element {
                    document_element_is_expected =
                        story_root_matches(expected_root, is_word, local);
                }
                let is_text_box_story_root = is_word
                    && local == b"txbxContent"
                    && effective_text_box_parent(&nodes).is_some_and(|parent| {
                        matches!(
                            parent,
                            TextBoxPathKind::VmlTextBox | TextBoxPathKind::WordprocessingTextBox
                        )
                    });
                let story_path_kind = story_path_kind(is_word, local, is_text_box_story_root);
                let is_paragraph = is_word && local == b"p";
                let is_run = is_word && local == b"r";
                let run_owner_kind = run_owner_kind(is_word, local);
                let graphic_data_uri = if is_drawing && local == b"graphicData" {
                    drawing_graphic_data_uri(&reader, &element)?
                } else {
                    None
                };
                let text_box_path_kind = text_box_path_kind(
                    text_box_namespace_kind,
                    local,
                    graphic_data_uri.as_deref(),
                    &nodes,
                );
                let mc_path_kind = if is_mc {
                    match local {
                        b"AlternateContent" => McPathKind::AlternateContent,
                        b"Choice" => McPathKind::Choice,
                        b"Fallback" => McPathKind::Fallback,
                        _ => McPathKind::Other,
                    }
                } else {
                    McPathKind::Other
                };
                let inherited_ignorable_namespaces = nodes
                    .last()
                    .map(|node| &node.ignorable_namespaces)
                    .cloned()
                    .unwrap_or_default();
                let (mc_rules_valid, ignorable_namespaces) =
                    mc_rule_state(&reader, &element, &inherited_ignorable_namespaces)?;
                let mc_element_valid = mc_path_kind == McPathKind::Other
                    || mc_element_attributes_are_valid(
                        &reader,
                        &element,
                        mc_path_kind,
                        &ignorable_namespaces,
                    )?;
                let ignorable_extension_child = mc_path_kind == McPathKind::Other
                    && resolved_namespace
                        .as_deref()
                        .is_some_and(|uri| inherited_ignorable_namespaces.contains(uri));
                let mc_branch_valid = record_mc_child(
                    &reader,
                    &element,
                    mc_path_kind,
                    mc_rules_valid && mc_element_valid,
                    ignorable_extension_child,
                    &mut nodes,
                )?;
                let is_object = is_word
                    && local == b"object"
                    && mc_rules_valid
                    && run_child_is_schema_positioned(&nodes);
                let is_control_owner = is_word
                    && matches!(local, b"object" | b"pict")
                    && mc_rules_valid
                    && run_child_is_schema_positioned(&nodes);
                if is_word && local == b"control" && control_child_is_schema_positioned(&nodes) {
                    let relationship_id = if mc_rules_valid {
                        relationship_attribute(&reader, &element)?
                    } else {
                        None
                    };
                    record_control_child(&mut nodes, relationship_id)?;
                }
                if is_office && local == b"OLEObject" && object_child_is_schema_positioned(&nodes) {
                    let relationship_id = if mc_rules_valid {
                        relationship_attribute(&reader, &element)?
                    } else {
                        None
                    };
                    record_object_child(&mut nodes, relationship_id)?;
                }
                nodes.push(OpenNode {
                    start: event_start,
                    is_document_element,
                    story_path_kind,
                    is_paragraph,
                    is_run,
                    is_object,
                    is_control_owner,
                    mc_path_kind,
                    mc_branch_valid,
                    mc_container_state: (mc_path_kind == McPathKind::AlternateContent).then_some(
                        McContainerState {
                            choice_count: 0,
                            fallback_seen: false,
                            grammar_valid: mc_rules_valid && mc_element_valid,
                        },
                    ),
                    mc_rules_valid,
                    ignorable_namespaces,
                    has_invalid_mc_descendant: false,
                    text_box_path_kind,
                    run_owner_kind,
                    object_relationship_ids: Vec::new(),
                    object_child_count: 0,
                    control_relationship_id: None,
                    control_child_count: 0,
                });
            }
            Event::Empty(element) => {
                let is_document_element = nodes.is_empty()
                    && document_element_count == 0
                    && !outside_document_content_seen;
                if nodes.is_empty() {
                    document_element_count = document_element_count.saturating_add(1);
                }
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                if is_document_element {
                    document_element_is_expected =
                        story_root_matches(expected_root, is_word, local);
                }
                let mc_path_kind = if is_mc {
                    match local {
                        b"AlternateContent" => McPathKind::AlternateContent,
                        b"Choice" => McPathKind::Choice,
                        b"Fallback" => McPathKind::Fallback,
                        _ => McPathKind::Other,
                    }
                } else {
                    McPathKind::Other
                };
                let inherited_ignorable_namespaces = nodes
                    .last()
                    .map(|node| &node.ignorable_namespaces)
                    .cloned()
                    .unwrap_or_default();
                let (mc_rules_valid, ignorable_namespaces) =
                    mc_rule_state(&reader, &element, &inherited_ignorable_namespaces)?;
                let mc_element_valid = mc_path_kind == McPathKind::Other
                    || mc_element_attributes_are_valid(
                        &reader,
                        &element,
                        mc_path_kind,
                        &ignorable_namespaces,
                    )?;
                let ignorable_extension_child = mc_path_kind == McPathKind::Other
                    && resolved_namespace
                        .as_deref()
                        .is_some_and(|uri| inherited_ignorable_namespaces.contains(uri));
                record_mc_child(
                    &reader,
                    &element,
                    mc_path_kind,
                    mc_rules_valid && mc_element_valid,
                    ignorable_extension_child,
                    &mut nodes,
                )?;
                if is_word && local == b"control" && control_child_is_schema_positioned(&nodes) {
                    let relationship_id = if mc_rules_valid {
                        relationship_attribute(&reader, &element)?
                    } else {
                        None
                    };
                    record_control_child(&mut nodes, relationship_id)?;
                }
                if is_office && local == b"OLEObject" && object_child_is_schema_positioned(&nodes) {
                    let relationship_id = if mc_rules_valid {
                        relationship_attribute(&reader, &element)?
                    } else {
                        None
                    };
                    record_object_child(&mut nodes, relationship_id)?;
                }
            }
            Event::End(_) => {
                let node = nodes.pop().ok_or_else(|| {
                    invalid("scan embedded XML", "unmatched closing element".to_owned())
                })?;
                if node
                    .mc_container_state
                    .is_some_and(|state| !state.grammar_valid || state.choice_count == 0)
                {
                    invalid_mc_ranges.push(node.start..event_end);
                    for ancestor in &mut nodes {
                        if ancestor.is_object || ancestor.is_control_owner {
                            ancestor.has_invalid_mc_descendant = true;
                        }
                    }
                }
                if node.is_object
                    && !node.has_invalid_mc_descendant
                    && node.object_child_count > 0
                    && node.control_child_count > 0
                {
                    return Err(invalid(
                        "scan embedded XML",
                        format!(
                            "Word object at byte {} mixes OLE and ActiveX owner children",
                            node.start
                        ),
                    ));
                }
                if kind == XmlReferenceKind::Ole
                    && node.is_object
                    && !node.has_invalid_mc_descendant
                {
                    let relationship_ids = node.object_relationship_ids;
                    if node.object_child_count > 1 {
                        return Err(invalid(
                            "scan embedded XML",
                            format!(
                                "Word object at byte {} has ambiguous OLE relationship ids {}",
                                node.start,
                                relationship_ids.join(", ")
                            ),
                        ));
                    }
                    if node.object_child_count == 1 && relationship_ids.is_empty() {
                        return Err(invalid(
                            "scan embedded XML",
                            format!(
                                "Word object at byte {} has an OLE child without a relationship id",
                                node.start
                            ),
                        ));
                    }
                    for relationship_id in relationship_ids {
                        references.push(XmlReference {
                            relationship_id,
                            range: node.start..event_end,
                        });
                    }
                }
                if kind == XmlReferenceKind::Control
                    && node.is_control_owner
                    && !node.has_invalid_mc_descendant
                    && node.control_child_count > 1
                {
                    return Err(invalid(
                        "scan embedded XML",
                        format!(
                            "Word control owner at byte {} has {} control children",
                            node.start, node.control_child_count
                        ),
                    ));
                }
                if kind == XmlReferenceKind::Control
                    && node.is_control_owner
                    && !node.has_invalid_mc_descendant
                    && node.control_child_count == 1
                    && node.control_relationship_id.is_none()
                {
                    return Err(invalid(
                        "scan embedded XML",
                        format!(
                            "Word control owner at byte {} has a child without a relationship id",
                            node.start
                        ),
                    ));
                }
                if kind == XmlReferenceKind::Control
                    && node.is_control_owner
                    && !node.has_invalid_mc_descendant
                    && let Some(relationship_id) = node.control_relationship_id
                {
                    references.push(XmlReference {
                        relationship_id,
                        range: node.start..event_end,
                    });
                }
            }
            Event::Eof => {
                if !nodes.is_empty() {
                    return Err(invalid(
                        "scan embedded XML",
                        "unclosed XML element".to_owned(),
                    ));
                }
                if outside_document_content_seen
                    || document_element_count != 1
                    || !document_element_is_expected
                {
                    return Ok(Vec::new());
                }
                references.retain(|reference| {
                    !invalid_mc_ranges.iter().any(|range| {
                        range.start <= reference.range.start && reference.range.end <= range.end
                    })
                });
                references.sort_by_key(|reference| reference.range.start);
                return Ok(references);
            }
            Event::Text(text) if nodes.is_empty() => {
                let bytes: &[u8] = text.as_ref();
                outside_document_content_seen |= !bytes.iter().all(u8::is_ascii_whitespace);
                if document_element_count == 0 {
                    prolog_content_seen = true;
                }
            }
            Event::CData(_) | Event::GeneralRef(_) if nodes.is_empty() => {
                outside_document_content_seen = true;
            }
            Event::GeneralRef(reference) => {
                let character = require_predefined_or_character_reference(&reference)?;
                if nodes.last().is_some_and(|node| {
                    node.mc_path_kind == McPathKind::AlternateContent
                        && !character.is_ascii_whitespace()
                }) {
                    nodes
                        .last_mut()
                        .unwrap()
                        .mc_container_state
                        .as_mut()
                        .unwrap()
                        .grammar_valid = false;
                }
            }
            Event::Decl(declaration) => {
                let misplaced = !nodes.is_empty()
                    || document_element_count != 0
                    || declaration_seen
                    || prolog_content_seen;
                if misplaced {
                    outside_document_content_seen = true;
                } else {
                    validate_xml_declaration(&declaration, "scan embedded XML")?;
                }
                declaration_seen = true;
            }
            Event::DocType(_) => {
                return Err(invalid(
                    "scan embedded XML",
                    "Word story XML cannot contain a document type".to_owned(),
                ));
            }
            Event::PI(instruction) if instruction.target().eq_ignore_ascii_case(b"xml") => {
                return Err(invalid(
                    "scan embedded XML",
                    "reserved XML processing instruction target".to_owned(),
                ));
            }
            Event::Comment(_) | Event::PI(_) if nodes.is_empty() && document_element_count == 0 => {
                prolog_content_seen = true;
            }
            Event::Text(text)
                if nodes.last().is_some_and(|node| {
                    node.mc_path_kind == McPathKind::AlternateContent && {
                        let bytes: &[u8] = text.as_ref();
                        !bytes.iter().all(u8::is_ascii_whitespace)
                    }
                }) =>
            {
                nodes
                    .last_mut()
                    .unwrap()
                    .mc_container_state
                    .as_mut()
                    .unwrap()
                    .grammar_valid = false;
            }
            Event::CData(_)
                if nodes
                    .last()
                    .is_some_and(|node| node.mc_path_kind == McPathKind::AlternateContent) =>
            {
                nodes
                    .last_mut()
                    .unwrap()
                    .mc_container_state
                    .as_mut()
                    .unwrap()
                    .grammar_valid = false;
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn record_mc_child(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    child_kind: McPathKind,
    element_valid: bool,
    ignorable_extension_child: bool,
    nodes: &mut [OpenNode],
) -> Result<bool> {
    let Some(state) = nodes
        .last_mut()
        .and_then(|node| node.mc_container_state.as_mut())
    else {
        return Ok(false);
    };
    match child_kind {
        McPathKind::Choice => {
            let branch_valid = state.grammar_valid
                && !state.fallback_seen
                && element_valid
                && choice_has_valid_requires(reader, element)?;
            state.choice_count = state.choice_count.saturating_add(1);
            state.grammar_valid &= branch_valid;
            Ok(branch_valid)
        }
        McPathKind::Fallback => {
            let branch_valid = state.grammar_valid
                && state.choice_count > 0
                && !state.fallback_seen
                && element_valid
                && fallback_has_no_requires(reader, element)?;
            state.fallback_seen = true;
            state.grammar_valid &= branch_valid;
            Ok(branch_valid)
        }
        McPathKind::Other if ignorable_extension_child => Ok(false),
        McPathKind::Other | McPathKind::AlternateContent => {
            state.grammar_valid = false;
            Ok(false)
        }
    }
}

fn choice_has_valid_requires(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<bool> {
    let mut requires = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"Requires" {
            if requires.is_some() {
                return Ok(false);
            }
            requires = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .map_err(|error| invalid("scan embedded XML", error.to_string()))?
                    .into_owned(),
            );
        }
    }
    Ok(requires.is_some_and(|value| {
        let mut prefixes = value.split_ascii_whitespace();
        prefixes.clone().next().is_some()
            && prefixes.all(|prefix| namespace_prefix_is_valid(reader, prefix))
    }))
}

fn fallback_has_no_requires(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"Requires" {
            return Ok(false);
        }
    }
    Ok(true)
}

fn mc_rule_state(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    inherited_ignorable: &BTreeSet<Vec<u8>>,
) -> Result<(bool, BTreeSet<Vec<u8>>)> {
    let mut ignorable = inherited_ignorable.clone();
    let mut valid = true;
    let mut seen = HashSet::<Vec<u8>>::new();

    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !namespace_is(&namespace, &[MC_NS]) {
            continue;
        }
        if !seen.insert(local.as_ref().to_vec()) {
            valid = false;
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            .map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        if local.as_ref() == b"Ignorable" {
            for prefix in value.split_ascii_whitespace() {
                let Some(namespace) = namespace_for_prefix(reader, prefix) else {
                    valid = false;
                    continue;
                };
                ignorable.insert(namespace);
            }
        }
    }

    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !namespace_is(&namespace, &[MC_NS]) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
            .map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let attribute_valid = match local.as_ref() {
            b"Ignorable" => true,
            b"MustUnderstand" => value
                .split_ascii_whitespace()
                .all(|prefix| must_understand_prefix_is_valid(reader, prefix)),
            b"ProcessContent" => mc_qname_list_is_valid(reader, &value, true, Some(&ignorable)),
            b"PreserveElements" | b"PreserveAttributes" => {
                mc_qname_list_is_valid(reader, &value, true, Some(&ignorable))
            }
            _ => false,
        };
        valid &= attribute_valid;
    }
    Ok((valid, ignorable))
}

fn mc_element_attributes_are_valid(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    element_kind: McPathKind,
    ignorable_namespaces: &BTreeSet<Vec<u8>>,
) -> Result<bool> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_is(&namespace, &[XML_NS]) {
            return Ok(false);
        }
        if matches!(namespace, ResolveResult::Unbound) {
            if element_kind == McPathKind::Choice
                && !raw_name.contains(&b':')
                && local.as_ref() == b"Requires"
            {
                continue;
            }
            return Ok(false);
        }
        if namespace_is(&namespace, &[MC_NS])
            && !matches!(
                local.as_ref(),
                b"Ignorable"
                    | b"MustUnderstand"
                    | b"ProcessContent"
                    | b"PreserveElements"
                    | b"PreserveAttributes"
            )
        {
            return Ok(false);
        }
        if !namespace_is(&namespace, &[MC_NS])
            && !resolved_namespace_is_ignorable(&namespace, ignorable_namespaces)
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn namespace_prefix_is_valid(reader: &NsReader<&[u8]>, prefix: &str) -> bool {
    namespace_for_prefix(reader, prefix).is_some()
}

fn must_understand_prefix_is_valid(reader: &NsReader<&[u8]>, prefix: &str) -> bool {
    namespace_for_prefix(reader, prefix).is_some_and(|namespace| {
        [
            W_NS,
            STRICT_W_NS,
            O_NS,
            R_NS,
            STRICT_R_NS,
            VML_NS,
            WP_NS,
            A_NS,
            WPG_NS,
            WPS_NS,
            WPC_NS,
        ]
        .iter()
        .any(|understood| namespace == understood.as_bytes())
    })
}

fn namespace_for_prefix(reader: &NsReader<&[u8]>, prefix: &str) -> Option<Vec<u8>> {
    if !xml_ncname_is_valid(prefix) {
        return None;
    }
    let qualified_name = format!("{prefix}:mcProbe");
    match reader
        .resolver()
        .resolve_element(QName(qualified_name.as_bytes()))
        .0
    {
        ResolveResult::Bound(Namespace(namespace))
            if namespace != MC_NS.as_bytes() && namespace != XML_NS.as_bytes() =>
        {
            Some(namespace.to_vec())
        }
        _ => None,
    }
}

fn mc_qname_list_is_valid(
    reader: &NsReader<&[u8]>,
    value: &str,
    wildcard_allowed: bool,
    required_ignorable: Option<&BTreeSet<Vec<u8>>>,
) -> bool {
    value.split_ascii_whitespace().all(|token| {
        let mut parts = token.split(':');
        let (Some(prefix), Some(local), None) = (parts.next(), parts.next(), parts.next()) else {
            return false;
        };
        let Some(namespace) = namespace_for_prefix(reader, prefix) else {
            return false;
        };
        required_ignorable.is_none_or(|ignorable| ignorable.contains(&namespace))
            && ((wildcard_allowed && local == "*") || xml_ncname_is_valid(local))
    })
}

fn xml_ncname_is_valid(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(xml_ncname_start_character)
        && characters.all(xml_ncname_character)
}

fn xml_ncname_start_character(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00C0}'..='\u{00D6}'
            | '\u{00D8}'..='\u{00F6}'
            | '\u{00F8}'..='\u{02FF}'
            | '\u{0370}'..='\u{037D}'
            | '\u{037F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}'
            | '\u{2070}'..='\u{218F}'
            | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}'
            | '\u{F900}'..='\u{FDCF}'
            | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

fn xml_ncname_character(character: char) -> bool {
    xml_ncname_start_character(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00B7}' | '\u{0300}'..='\u{036F}' | '\u{203F}'..='\u{2040}'
        )
}

fn run_child_is_schema_positioned(nodes: &[OpenNode]) -> bool {
    if !compatibility_ancestry_is_valid(nodes) {
        return false;
    }
    let Some(mut parent) = nodes.len().checked_sub(1) else {
        return false;
    };
    while nodes[parent].mc_path_kind != McPathKind::Other {
        let Some(next) = parent.checked_sub(1) else {
            return false;
        };
        parent = next;
    }
    if !nodes[parent].is_run {
        return false;
    }
    let Some(paragraph) = nodes[..parent].iter().rposition(|node| node.is_paragraph) else {
        return false;
    };
    valid_story_owner_path(&nodes[..paragraph])
        && valid_run_owner_path(&nodes[paragraph + 1..parent])
}

fn story_path_kind(is_word: bool, local: &[u8], is_text_box_story_root: bool) -> StoryPathKind {
    if !is_word {
        return StoryPathKind::Other;
    }
    if is_text_box_story_root {
        return StoryPathKind::TextBoxRoot;
    }
    match local {
        b"document" => StoryPathKind::DocumentRoot,
        b"hdr" | b"ftr" => StoryPathKind::HeaderFooterRoot,
        b"footnotes" => StoryPathKind::FootnotesRoot,
        b"endnotes" => StoryPathKind::EndnotesRoot,
        b"comments" => StoryPathKind::CommentsRoot,
        b"glossaryDocument" => StoryPathKind::GlossaryRoot,
        b"body" => StoryPathKind::Body,
        b"footnote" => StoryPathKind::Footnote,
        b"endnote" => StoryPathKind::Endnote,
        b"comment" => StoryPathKind::Comment,
        b"docParts" => StoryPathKind::DocParts,
        b"docPart" => StoryPathKind::DocPart,
        b"docPartBody" => StoryPathKind::DocPartBody,
        b"customXml" => StoryPathKind::CustomXml,
        b"sdt" => StoryPathKind::StructuredDocumentTag,
        b"sdtContent" => StoryPathKind::StructuredDocumentTagContent,
        b"tbl" => StoryPathKind::Table,
        b"tr" => StoryPathKind::TableRow,
        b"tc" => StoryPathKind::TableCell,
        _ => StoryPathKind::Other,
    }
}

fn story_root_matches(expected: StoryRootKind, is_word: bool, local: &[u8]) -> bool {
    is_word
        && matches!(
            (expected, local),
            (StoryRootKind::Document, b"document")
                | (StoryRootKind::Header, b"hdr")
                | (StoryRootKind::Footer, b"ftr")
                | (StoryRootKind::Footnotes, b"footnotes")
                | (StoryRootKind::Endnotes, b"endnotes")
                | (StoryRootKind::Comments, b"comments")
                | (StoryRootKind::Glossary, b"glossaryDocument")
        )
}

fn valid_story_owner_path(nodes: &[OpenNode]) -> bool {
    let Some(root_index) = nodes
        .iter()
        .rposition(|node| node.story_path_kind == StoryPathKind::TextBoxRoot)
        .or_else(|| {
            nodes.first().and_then(|node| {
                (node.is_document_element
                    && matches!(
                        node.story_path_kind,
                        StoryPathKind::DocumentRoot
                            | StoryPathKind::HeaderFooterRoot
                            | StoryPathKind::FootnotesRoot
                            | StoryPathKind::EndnotesRoot
                            | StoryPathKind::CommentsRoot
                            | StoryPathKind::GlossaryRoot
                    ))
                .then_some(0)
            })
        })
    else {
        return false;
    };
    let mut nodes = nodes[root_index..]
        .iter()
        .filter(|node| node.mc_path_kind == McPathKind::Other);
    let Some(root) = nodes.next() else {
        return false;
    };
    let mut expected = match root.story_path_kind {
        StoryPathKind::DocumentRoot => ExpectedStoryChild::Body,
        StoryPathKind::HeaderFooterRoot => ExpectedStoryChild::Content(StoryContentKind::Block),
        StoryPathKind::FootnotesRoot => ExpectedStoryChild::Footnote,
        StoryPathKind::EndnotesRoot => ExpectedStoryChild::Endnote,
        StoryPathKind::CommentsRoot => ExpectedStoryChild::Comment,
        StoryPathKind::GlossaryRoot => ExpectedStoryChild::DocParts,
        StoryPathKind::TextBoxRoot => ExpectedStoryChild::Content(StoryContentKind::Block),
        _ => return false,
    };
    for node in nodes {
        let Some(next) = next_story_expectation(expected, node.story_path_kind) else {
            return false;
        };
        expected = next;
    }
    expected == ExpectedStoryChild::Content(StoryContentKind::Block)
}

fn text_box_path_kind(
    namespace: TextBoxNamespaceKind,
    local: &[u8],
    graphic_data_uri: Option<&str>,
    nodes: &[OpenNode],
) -> TextBoxPathKind {
    let parent = effective_text_box_parent(nodes).unwrap_or(TextBoxPathKind::Other);
    if namespace == TextBoxNamespaceKind::Word
        && local == b"pict"
        && run_child_is_schema_positioned(nodes)
    {
        return TextBoxPathKind::LegacyPicture;
    }
    if namespace == TextBoxNamespaceKind::Word
        && local == b"object"
        && run_child_is_schema_positioned(nodes)
    {
        return TextBoxPathKind::LegacyObject;
    }
    if namespace == TextBoxNamespaceKind::Word
        && local == b"drawing"
        && (run_child_is_schema_positioned(nodes) || parent == TextBoxPathKind::LegacyObject)
    {
        return TextBoxPathKind::Drawing;
    }
    if namespace == TextBoxNamespaceKind::Vml && local == b"group" {
        return match parent {
            TextBoxPathKind::LegacyPicture
            | TextBoxPathKind::LegacyObject
            | TextBoxPathKind::VmlGroup => TextBoxPathKind::VmlGroup,
            _ => TextBoxPathKind::Other,
        };
    }
    if namespace == TextBoxNamespaceKind::Vml
        && matches!(
            local,
            b"shape"
                | b"arc"
                | b"curve"
                | b"image"
                | b"line"
                | b"oval"
                | b"polyline"
                | b"rect"
                | b"roundrect"
        )
    {
        return match parent {
            TextBoxPathKind::LegacyPicture
            | TextBoxPathKind::LegacyObject
            | TextBoxPathKind::VmlGroup => TextBoxPathKind::VmlShape,
            _ => TextBoxPathKind::Other,
        };
    }
    if namespace == TextBoxNamespaceKind::Vml
        && local == b"textbox"
        && parent == TextBoxPathKind::VmlShape
    {
        return TextBoxPathKind::VmlTextBox;
    }
    if namespace == TextBoxNamespaceKind::WordprocessingDrawing
        && matches!(local, b"inline" | b"anchor")
        && parent == TextBoxPathKind::Drawing
    {
        return TextBoxPathKind::WordprocessingDrawing;
    }
    if namespace == TextBoxNamespaceKind::Drawing
        && local == b"graphic"
        && parent == TextBoxPathKind::WordprocessingDrawing
    {
        return TextBoxPathKind::Graphic;
    }
    if namespace == TextBoxNamespaceKind::Drawing
        && local == b"graphicData"
        && parent == TextBoxPathKind::Graphic
    {
        return match graphic_data_uri {
            Some(WPS_NS) => TextBoxPathKind::GraphicDataShape,
            Some(WPG_NS) => TextBoxPathKind::GraphicDataGroup,
            Some(WPC_NS) => TextBoxPathKind::GraphicDataCanvas,
            _ => TextBoxPathKind::Other,
        };
    }
    if namespace == TextBoxNamespaceKind::WordprocessingCanvas
        && local == b"wpc"
        && parent == TextBoxPathKind::GraphicDataCanvas
    {
        return TextBoxPathKind::WordprocessingCanvas;
    }
    if namespace == TextBoxNamespaceKind::WordprocessingGroup
        && local == b"wgp"
        && matches!(
            parent,
            TextBoxPathKind::GraphicDataGroup | TextBoxPathKind::WordprocessingCanvas
        )
    {
        return TextBoxPathKind::WordprocessingGroup;
    }
    if namespace == TextBoxNamespaceKind::WordprocessingGroup
        && local == b"grpSp"
        && parent == TextBoxPathKind::WordprocessingGroup
    {
        return TextBoxPathKind::WordprocessingGroup;
    }
    if namespace == TextBoxNamespaceKind::WordprocessingShape
        && local == b"wsp"
        && matches!(
            parent,
            TextBoxPathKind::GraphicDataShape
                | TextBoxPathKind::WordprocessingGroup
                | TextBoxPathKind::WordprocessingCanvas
        )
    {
        return TextBoxPathKind::WordprocessingShape;
    }
    if namespace == TextBoxNamespaceKind::WordprocessingShape
        && local == b"txbx"
        && parent == TextBoxPathKind::WordprocessingShape
    {
        return TextBoxPathKind::WordprocessingTextBox;
    }
    TextBoxPathKind::Other
}

fn effective_text_box_parent(nodes: &[OpenNode]) -> Option<TextBoxPathKind> {
    compatibility_ancestry_is_valid(nodes).then(|| {
        nodes
            .iter()
            .rev()
            .find(|node| node.mc_path_kind == McPathKind::Other)
            .map_or(TextBoxPathKind::Other, |node| node.text_box_path_kind)
    })
}

fn drawing_graphic_data_uri(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<String>> {
    let mut uri = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound)
            && !attribute.key.as_ref().contains(&b':')
            && local.as_ref() == b"uri"
        {
            if uri.is_some() {
                return Ok(None);
            }
            uri = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                    .map_err(|error| invalid("scan embedded XML", error.to_string()))?
                    .into_owned(),
            );
        }
    }
    Ok(uri)
}

fn next_story_expectation(
    expected: ExpectedStoryChild,
    actual: StoryPathKind,
) -> Option<ExpectedStoryChild> {
    match (expected, actual) {
        (ExpectedStoryChild::Body, StoryPathKind::Body)
        | (ExpectedStoryChild::Footnote, StoryPathKind::Footnote)
        | (ExpectedStoryChild::Endnote, StoryPathKind::Endnote)
        | (ExpectedStoryChild::Comment, StoryPathKind::Comment)
        | (ExpectedStoryChild::DocPartBody, StoryPathKind::DocPartBody) => {
            Some(ExpectedStoryChild::Content(StoryContentKind::Block))
        }
        (ExpectedStoryChild::DocParts, StoryPathKind::DocParts) => {
            Some(ExpectedStoryChild::DocPart)
        }
        (ExpectedStoryChild::DocPart, StoryPathKind::DocPart) => {
            Some(ExpectedStoryChild::DocPartBody)
        }
        (ExpectedStoryChild::Content(content), StoryPathKind::CustomXml) => {
            Some(ExpectedStoryChild::Content(content))
        }
        (ExpectedStoryChild::Content(content), StoryPathKind::StructuredDocumentTag) => {
            Some(ExpectedStoryChild::StructuredDocumentTagContent(content))
        }
        (
            ExpectedStoryChild::StructuredDocumentTagContent(content),
            StoryPathKind::StructuredDocumentTagContent,
        ) => Some(ExpectedStoryChild::Content(content)),
        (ExpectedStoryChild::Content(StoryContentKind::Block), StoryPathKind::Table) => {
            Some(ExpectedStoryChild::Content(StoryContentKind::TableRow))
        }
        (ExpectedStoryChild::Content(StoryContentKind::TableRow), StoryPathKind::TableRow) => {
            Some(ExpectedStoryChild::Content(StoryContentKind::TableCell))
        }
        (ExpectedStoryChild::Content(StoryContentKind::TableCell), StoryPathKind::TableCell) => {
            Some(ExpectedStoryChild::Content(StoryContentKind::Block))
        }
        _ => None,
    }
}

fn run_owner_kind(is_word: bool, local: &[u8]) -> RunOwnerKind {
    if !is_word {
        return RunOwnerKind::Other;
    }
    match local {
        b"customXml" | b"smartTag" | b"hyperlink" | b"fldSimple" | b"ins" | b"del"
        | b"moveFrom" | b"moveTo" | b"dir" | b"bdo" => RunOwnerKind::Container,
        b"sdt" => RunOwnerKind::StructuredDocumentTag,
        b"sdtContent" => RunOwnerKind::StructuredDocumentTagContent,
        _ => RunOwnerKind::Other,
    }
}

fn valid_run_owner_path(nodes: &[OpenNode]) -> bool {
    let mut needs_sdt_content = false;
    for node in nodes
        .iter()
        .filter(|node| node.mc_path_kind == McPathKind::Other)
    {
        if needs_sdt_content {
            if node.run_owner_kind != RunOwnerKind::StructuredDocumentTagContent {
                return false;
            }
            needs_sdt_content = false;
            continue;
        }
        match node.run_owner_kind {
            RunOwnerKind::Container => {}
            RunOwnerKind::StructuredDocumentTag => needs_sdt_content = true,
            RunOwnerKind::Other | RunOwnerKind::StructuredDocumentTagContent => return false,
        }
    }
    !needs_sdt_content
}

fn object_child_is_schema_positioned(nodes: &[OpenNode]) -> bool {
    if !compatibility_ancestry_is_valid(nodes) {
        return false;
    }
    let Some(mut parent) = nodes.len().checked_sub(1) else {
        return false;
    };
    while nodes[parent].mc_path_kind != McPathKind::Other {
        let Some(next) = parent.checked_sub(1) else {
            return false;
        };
        parent = next;
    }
    nodes[parent].is_object
}

fn control_child_is_schema_positioned(nodes: &[OpenNode]) -> bool {
    if !compatibility_ancestry_is_valid(nodes) {
        return false;
    }
    let Some(mut parent) = nodes.len().checked_sub(1) else {
        return false;
    };
    while nodes[parent].mc_path_kind != McPathKind::Other {
        let Some(next) = parent.checked_sub(1) else {
            return false;
        };
        parent = next;
    }
    nodes[parent].is_control_owner
}

fn record_object_child(nodes: &mut [OpenNode], relationship_id: Option<String>) -> Result<()> {
    let owner = nodes
        .iter_mut()
        .rev()
        .find(|node| node.is_object)
        .ok_or_else(|| {
            invalid(
                "scan embedded XML",
                "schema-positioned OLE object has no owning Word object".to_owned(),
            )
        })?;
    owner.object_child_count = owner.object_child_count.saturating_add(1);
    if let Some(relationship_id) = relationship_id {
        owner.object_relationship_ids.push(relationship_id);
    }
    Ok(())
}

fn record_control_child(nodes: &mut [OpenNode], relationship_id: Option<String>) -> Result<()> {
    let owner = nodes
        .iter_mut()
        .rev()
        .find(|node| node.is_control_owner)
        .ok_or_else(|| {
            invalid(
                "scan embedded XML",
                "schema-positioned Word control has no owning object or picture".to_owned(),
            )
        })?;
    owner.control_child_count = owner.control_child_count.saturating_add(1);
    if let Some(relationship_id) = relationship_id {
        if owner.control_relationship_id.is_some() {
            return Err(invalid(
                "scan embedded XML",
                format!(
                    "Word control owner at byte {} has multiple relationship ids",
                    owner.start
                ),
            ));
        }
        owner.control_relationship_id = Some(relationship_id);
    }
    Ok(())
}

fn compatibility_ancestry_is_valid(nodes: &[OpenNode]) -> bool {
    if nodes.iter().any(|node| !node.mc_rules_valid) {
        return false;
    }
    let mut index = 0usize;
    while index < nodes.len() {
        match nodes[index].mc_path_kind {
            McPathKind::Other => index += 1,
            McPathKind::AlternateContent => {
                if nodes.get(index + 1).is_none_or(|node| {
                    !matches!(node.mc_path_kind, McPathKind::Choice | McPathKind::Fallback)
                        || !node.mc_branch_valid
                }) {
                    return false;
                }
                index += 2;
            }
            McPathKind::Choice | McPathKind::Fallback => return false,
        }
    }
    true
}

fn relationship_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<String>> {
    let mut relationship_id = None;
    let mut semantic_attribute_seen = false;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == b"id" && namespace_is(&namespace, &[R_NS, STRICT_R_NS]) {
            if semantic_attribute_seen {
                return Err(invalid(
                    "scan embedded XML",
                    format!(
                        "{} has duplicate relationship id attributes",
                        String::from_utf8_lossy(element.name().as_ref())
                    ),
                ));
            }
            semantic_attribute_seen = true;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .map_err(|error| invalid("scan embedded XML", error.to_string()))?
                .into_owned();
            relationship_id = (!value.is_empty()).then_some(value);
        }
    }
    Ok(relationship_id)
}

fn namespace_is(namespace: &ResolveResult<'_>, expected: &[&str]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if expected.iter().any(|value| *uri == value.as_bytes()))
}

fn resolved_namespace_is_ignorable(
    namespace: &ResolveResult<'_>,
    ignorable_namespaces: &BTreeSet<Vec<u8>>,
) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if ignorable_namespaces.contains(*uri))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn remove_ranges(xml: &[u8], mut ranges: Vec<Range<usize>>) -> Result<Vec<u8>> {
    ranges.sort_by_key(|range| range.start);
    ranges.dedup();
    let mut output = Vec::with_capacity(xml.len());
    let mut copied = 0usize;
    for range in ranges {
        if range.start < copied || range.end > xml.len() || range.start > range.end {
            return Err(invalid(
                "remove embedded content",
                "overlapping or invalid owning XML ranges".to_owned(),
            ));
        }
        output.extend_from_slice(&xml[copied..range.start]);
        copied = range.end;
    }
    output.extend_from_slice(&xml[copied..]);
    Ok(output)
}

fn is_vba_signature(relationship: &Relationship) -> bool {
    matches!(
        relationship.rel_type.as_str(),
        rel_types::VBA_PROJECT_SIGNATURE | rel_types::VBA_PROJECT_SIGNATURE_AGILE
    )
}

fn attached_vba_signature_state(package: &OpcPackage, project_part: &str) -> Result<Option<bool>> {
    let Some(relationships) = package.get_part_rels(project_part) else {
        return Ok(None);
    };
    ensure_unique_relationship_ids(project_part, &relationships.items)?;
    let signatures = relationships
        .items
        .iter()
        .filter(|relationship| is_vba_signature(relationship))
        .collect::<Vec<_>>();
    if signatures.len() > 1 {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "{project_part}: found {} VBA project signature relationships, expected at most one",
                signatures.len()
            ),
        ));
    }
    let markers = relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == INVALIDATED_VBA_SIGNATURE)
        .collect::<Vec<_>>();
    if markers.len() > 1 {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "{project_part}: found {} VBA signature invalidation markers, expected at most one",
                markers.len()
            ),
        ));
    }
    let Some(signature) = signatures.first() else {
        if markers.is_empty() {
            return Ok(None);
        }
        return Err(invalid(
            "inventory embedded content",
            format!("{project_part}: VBA invalidation marker has no signature relationship"),
        ));
    };
    let target = safe_internal_target(project_part, signature)?;
    required_part(package, &target)?;
    let expected_content_type = if signature.rel_type == rel_types::VBA_PROJECT_SIGNATURE {
        VBA_SIGNATURE_CONTENT_TYPE
    } else {
        VBA_AGILE_SIGNATURE_CONTENT_TYPE
    };
    require_exact_content_type(package, &target, expected_content_type, "VBA signature")?;
    if let Some(marker) = markers.first() {
        let marker_target = safe_internal_target(project_part, marker)?;
        if marker_target != target {
            return Err(invalid(
                "inventory embedded content",
                format!(
                    "{project_part}: VBA invalidation marker targets {marker_target}, expected {target}"
                ),
            ));
        }
        Ok(Some(true))
    } else {
        Ok(Some(false))
    }
}

fn mark_vba_signature_invalidated(package: &mut OpcPackage, project_part: &str) -> Result<()> {
    let state = attached_vba_signature_state(package, project_part)?;
    if state.is_none() || state == Some(true) {
        return Ok(());
    }
    let target = package
        .get_part_rels(project_part)
        .and_then(|relationships| {
            relationships
                .items
                .iter()
                .find(|relationship| is_vba_signature(relationship))
        })
        .map(|relationship| relationship.target.clone())
        .ok_or_else(|| {
            invalid(
                "invalidate VBA signature",
                format!("{project_part}: signature relationship disappeared"),
            )
        })?;
    package
        .get_or_create_part_rels(project_part)
        .add(INVALIDATED_VBA_SIGNATURE, &target);
    Ok(())
}

#[derive(Debug)]
struct PackageSignatureGraph {
    origins: Vec<(String, String)>,
    signatures: Vec<(String, String)>,
}

fn package_signature_graph(package: &OpcPackage) -> Result<Option<PackageSignatureGraph>> {
    ensure_unique_relationship_ids("/", &package.package_rels.items)?;
    let origins = package
        .package_rels
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN)
        .collect::<Vec<_>>();
    if package
        .package_rels
        .items
        .iter()
        .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE)
    {
        return Err(invalid(
            "inventory embedded content",
            "/: digital-signature relationship is outside an origin part".to_owned(),
        ));
    }
    if origins.len() > 1 {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "/: found {} digital-signature origins, expected at most one",
                origins.len()
            ),
        ));
    }
    let markers = package
        .package_rels
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == INVALIDATED_PACKAGE_SIGNATURE)
        .collect::<Vec<_>>();
    if markers.len() > 1 {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "/: found {} package-signature invalidation markers, expected at most one",
                markers.len()
            ),
        ));
    }
    let Some(origin) = origins.first() else {
        if !markers.is_empty() {
            return Err(invalid(
                "inventory embedded content",
                "/: package-signature invalidation marker has no origin".to_owned(),
            ));
        }
        for (source_part, relationships) in &package.part_rels {
            if relationships.items.iter().any(|relationship| {
                matches!(
                    relationship.rel_type.as_str(),
                    rel_types::DIGITAL_SIGNATURE_ORIGIN | rel_types::DIGITAL_SIGNATURE
                )
            }) {
                return Err(invalid(
                    "inventory embedded content",
                    format!("{source_part}: misplaced digital-signature relationship"),
                ));
            }
        }
        return Ok(None);
    };
    let origin_part = safe_internal_target("/", origin)?;
    if let Some(marker) = markers.first()
        && safe_internal_target("/", marker)? != origin_part
    {
        return Err(invalid(
            "inventory embedded content",
            format!("/: package-signature invalidation marker does not target {origin_part}"),
        ));
    }
    for (source_part, relationships) in &package.part_rels {
        if relationships.items.iter().any(|relationship| {
            matches!(
                relationship.rel_type.as_str(),
                rel_types::DIGITAL_SIGNATURE_ORIGIN | rel_types::DIGITAL_SIGNATURE
            )
        }) {
            ensure_unique_relationship_ids(source_part, &relationships.items)?;
        }
        if relationships
            .items
            .iter()
            .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN)
        {
            return Err(invalid(
                "inventory embedded content",
                format!("{source_part}: misplaced digital-signature origin relationship"),
            ));
        }
        if source_part != &origin_part
            && relationships
                .items
                .iter()
                .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE)
        {
            return Err(invalid(
                "inventory embedded content",
                format!("{source_part}: misplaced digital-signature relationship"),
            ));
        }
    }
    required_part(package, &origin_part)?;
    require_exact_content_type(
        package,
        &origin_part,
        PACKAGE_SIGNATURE_ORIGIN_CONTENT_TYPE,
        "package signature origin",
    )?;
    let origin_relationships = package.get_part_rels(&origin_part).ok_or_else(|| {
        invalid(
            "inventory embedded content",
            format!("{origin_part}: digital-signature origin has no relationship set"),
        )
    })?;
    ensure_unique_relationship_ids(&origin_part, &origin_relationships.items)?;
    if let Some(unrelated) = origin_relationships
        .items
        .iter()
        .find(|relationship| relationship.rel_type != rel_types::DIGITAL_SIGNATURE)
    {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "{origin_part}: digital-signature origin has unrelated relationship {}",
                unrelated.id
            ),
        ));
    }
    let signatures = origin_relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE)
        .collect::<Vec<_>>();
    if signatures.is_empty() {
        return Err(invalid(
            "inventory embedded content",
            format!("{origin_part}: digital-signature origin has no signature relationship"),
        ));
    }
    let mut signature_parts = HashSet::new();
    let mut graph = PackageSignatureGraph {
        origins: vec![(origin.id.clone(), origin_part.clone())],
        signatures: Vec::new(),
    };
    for signature in signatures {
        let signature_part = safe_internal_target(&origin_part, signature)?;
        if !signature_parts.insert(signature_part.clone()) {
            return Err(invalid(
                "inventory embedded content",
                format!("duplicate digital-signature target {signature_part}"),
            ));
        }
        required_part(package, &signature_part)?;
        require_exact_content_type(
            package,
            &signature_part,
            PACKAGE_SIGNATURE_XML_CONTENT_TYPE,
            "package XML signature",
        )?;
        graph.signatures.push((origin_part.clone(), signature_part));
    }
    reject_unrelated_signature_incoming(package, &graph)?;
    Ok(Some(graph))
}

fn require_exact_content_type(
    package: &OpcPackage,
    part: &str,
    expected: &str,
    role: &str,
) -> Result<()> {
    match package.content_types.content_type_for(part) {
        Some(actual) if actual == expected => Ok(()),
        Some(actual) => Err(invalid(
            "inventory embedded content",
            format!("{part}: {role} has content type {actual}, expected {expected}"),
        )),
        None => Err(invalid(
            "inventory embedded content",
            format!("{part}: {role} has no content type, expected {expected}"),
        )),
    }
}

fn reject_unrelated_signature_incoming(
    package: &OpcPackage,
    graph: &PackageSignatureGraph,
) -> Result<()> {
    let origin_part = &graph.origins[0].1;
    let signature_parts = graph
        .signatures
        .iter()
        .map(|(_, part)| part.as_str())
        .collect::<HashSet<_>>();
    for relationship in &package.package_rels.items {
        if !relationship_is_internal("/", relationship)? {
            continue;
        }
        let target = safe_internal_target("/", relationship)?;
        let allowed_origin = target == *origin_part
            && matches!(
                relationship.rel_type.as_str(),
                rel_types::DIGITAL_SIGNATURE_ORIGIN | INVALIDATED_PACKAGE_SIGNATURE
            );
        if (target == *origin_part && !allowed_origin) || signature_parts.contains(target.as_str())
        {
            return Err(invalid(
                "inventory embedded content",
                format!(
                    "/: unrelated relationship {} targets package signature part {target}",
                    relationship.id
                ),
            ));
        }
    }
    for (source_part, relationships) in &package.part_rels {
        for relationship in &relationships.items {
            if !relationship_is_internal(source_part, relationship)? {
                continue;
            }
            let target = safe_internal_target(source_part, relationship)?;
            let allowed_signature = source_part == origin_part
                && relationship.rel_type == rel_types::DIGITAL_SIGNATURE
                && signature_parts.contains(target.as_str());
            if target == *origin_part
                || (signature_parts.contains(target.as_str()) && !allowed_signature)
            {
                return Err(invalid(
                    "inventory embedded content",
                    format!(
                        "{source_part}: unrelated relationship {} targets package signature part {target}",
                        relationship.id
                    ),
                ));
            }
        }
    }
    Ok(())
}

fn has_package_signature(package: &OpcPackage) -> Result<bool> {
    package_signature_graph(package).map(|graph| graph.is_some())
}

fn package_signature_invalidation_marked(package: &OpcPackage) -> Result<bool> {
    package_signature_graph(package)?;
    Ok(package
        .package_rels
        .items
        .iter()
        .any(|relationship| relationship.rel_type == INVALIDATED_PACKAGE_SIGNATURE))
}

fn ensure_unique_relationship_ids(source_part: &str, relationships: &[Relationship]) -> Result<()> {
    let mut seen = HashSet::new();
    if let Some(duplicate) = relationships
        .iter()
        .map(|relationship| relationship.id.as_str())
        .find(|id| !seen.insert(*id))
    {
        return Err(invalid(
            "resolve embedded content",
            format!("{source_part}: duplicate relationship id {duplicate}"),
        ));
    }
    Ok(())
}

fn remove_package_signatures(package: &mut OpcPackage) -> Result<()> {
    let Some(graph) = package_signature_graph(package)? else {
        return Ok(());
    };
    package.package_rels.items.retain(|relationship| {
        relationship.rel_type != rel_types::DIGITAL_SIGNATURE_ORIGIN
            && relationship.rel_type != INVALIDATED_PACKAGE_SIGNATURE
    });
    for (_, signature_part) in graph.signatures {
        package.parts.remove(&signature_part);
        package.part_rels.remove(&signature_part);
        package.content_types.overrides.remove(&signature_part);
    }
    for (_, origin_part) in graph.origins {
        package.parts.remove(&origin_part);
        package.part_rels.remove(&origin_part);
        package.content_types.overrides.remove(&origin_part);
    }
    Ok(())
}

fn remove_vba_signatures(package: &mut OpcPackage, project_part: &str) -> Result<()> {
    attached_vba_signature_state(package, project_part)?;
    let signatures = package
        .get_part_rels(project_part)
        .map(|relationships| {
            relationships
                .items
                .iter()
                .filter(|relationship| is_vba_signature(relationship))
                .cloned()
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    if let Some(relationships) = package.part_rels.get_mut(project_part) {
        relationships.items.retain(|relationship| {
            !is_vba_signature(relationship) && relationship.rel_type != INVALIDATED_VBA_SIGNATURE
        });
    }
    for signature in signatures {
        if relationship_is_internal(project_part, &signature)? {
            let signature_part = safe_internal_target(project_part, &signature)?;
            delete_if_unreachable(package, &signature_part)?;
        }
    }
    Ok(())
}

fn retain_vba_signature_parts_as_evidence(package: &mut OpcPackage, project_part: &str) {
    if let Some(relationships) = package.part_rels.get_mut(project_part) {
        relationships.items.retain(|relationship| {
            !is_vba_signature(relationship) && relationship.rel_type != INVALIDATED_VBA_SIGNATURE
        });
    }
}

#[derive(Debug)]
struct SignatureReference {
    path: String,
    relationship_ids: Vec<String>,
}

fn signature_manifest_has_missing_reference(package: &OpcPackage) -> bool {
    let Some(graph) = package_signature_graph(package).ok().flatten() else {
        return false;
    };
    graph.signatures.iter().any(|(_, signature_part)| {
        required_part(package, signature_part)
            .ok()
            .and_then(|xml| signature_references(xml).ok())
            .is_some_and(|references| {
                references.into_iter().any(|reference| {
                    if matches!(
                        reference.path.as_str(),
                        "/[Content_Types].xml" | "/_rels/.rels"
                    ) {
                        return false;
                    }
                    if let Some(source) = relationship_source_from_path(&reference.path) {
                        let Some(relationships) = package.get_part_rels(&source) else {
                            return true;
                        };
                        return reference.relationship_ids.iter().any(|id| {
                            relationships
                                .items
                                .iter()
                                .filter(|relationship| relationship.id == *id)
                                .count()
                                != 1
                        });
                    }
                    !package.parts.contains_key(&reference.path)
                })
            })
    })
}

fn signature_references(xml: &[u8]) -> Result<Vec<SignatureReference>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut references = Vec::new();
    let mut active_reference = None::<(usize, SignatureReference)>;
    let mut depth = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid("inspect package signature", error.to_string()))?;
        let is_dsig = namespace_is(&namespace, &[DSIG_NS]);
        let is_opc_signature = namespace_is(&namespace, &[OPC_SIGNATURE_NS]);
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                if is_dsig
                    && local_name(element.name().as_ref()) == b"Reference"
                    && let Some(uri) = unqualified_attribute(&element, b"URI")?
                    && uri.starts_with('/')
                {
                    let path = uri.split('?').next().unwrap_or(&uri);
                    active_reference = Some((
                        depth,
                        SignatureReference {
                            path: percent_decode_path(path).ok_or_else(|| {
                                invalid(
                                    "inspect package signature",
                                    format!("invalid package reference URI {uri}"),
                                )
                            })?,
                            relationship_ids: Vec::new(),
                        },
                    ));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    invalid(
                        "inspect package signature",
                        "signature XML nesting is too deep".to_owned(),
                    )
                })?;
            }
            Event::Empty(element) => {
                if is_dsig
                    && local_name(element.name().as_ref()) == b"Reference"
                    && let Some(uri) = unqualified_attribute(&element, b"URI")?
                    && uri.starts_with('/')
                {
                    let path = uri.split('?').next().unwrap_or(&uri);
                    references.push(SignatureReference {
                        path: percent_decode_path(path).ok_or_else(|| {
                            invalid(
                                "inspect package signature",
                                format!("invalid package reference URI {uri}"),
                            )
                        })?,
                        relationship_ids: Vec::new(),
                    });
                } else if is_opc_signature
                    && local_name(element.name().as_ref()) == b"RelationshipReference"
                    && let Some((_, reference)) = active_reference.as_mut()
                    && let Some(id) = unqualified_attribute(&element, b"SourceId")?
                {
                    reference.relationship_ids.push(id);
                }
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    invalid(
                        "inspect package signature",
                        "signature XML has an unmatched closing element".to_owned(),
                    )
                })?;
                if active_reference
                    .as_ref()
                    .is_some_and(|(reference_depth, _)| *reference_depth == depth)
                    && let Some((_, reference)) = active_reference.take()
                {
                    references.push(reference);
                }
            }
            Event::Eof if depth == 0 => return Ok(references),
            Event::Eof => {
                return Err(invalid(
                    "inspect package signature",
                    "signature XML ended before its root closed".to_owned(),
                ));
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn unqualified_attribute(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        if attribute.key.as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| invalid("scan embedded XML", error.to_string()));
        }
    }
    Ok(None)
}

fn relationship_source_from_path(path: &str) -> Option<String> {
    let marker = "/_rels/";
    let marker_index = path.rfind(marker)?;
    let filename = path
        .get(marker_index + marker.len()..)?
        .strip_suffix(".rels")?;
    Some(format!("{}{filename}", &path[..marker_index + 1]))
}

fn percent_decode_path(value: &str) -> Option<String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            let high = hex_digit(*bytes.get(index + 1)?)?;
            let low = hex_digit(*bytes.get(index + 2)?)?;
            decoded.push((high << 4) | low);
            index += 3;
        } else {
            decoded.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(decoded).ok()
}

fn hex_digit(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn packages_are_semantically_equal(left: &OpcPackage, right: &OpcPackage) -> bool {
    left.parts == right.parts
        && left.content_types.defaults == right.content_types.defaults
        && left.content_types.overrides == right.content_types.overrides
        && left.package_rels.items == right.package_rels.items
        && left.part_rels.len() == right.part_rels.len()
        && left.part_rels.iter().all(|(source, relationships)| {
            right
                .part_rels
                .get(source)
                .is_some_and(|other| relationships.items == other.items)
        })
}

pub(crate) fn synchronized_package_mutation_invalidates_signature(
    original: &OpcPackage,
    candidate: &OpcPackage,
) -> bool {
    let original_signed = original
        .package_rels
        .items
        .iter()
        .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN);
    let candidate_signed = candidate
        .package_rels
        .items
        .iter()
        .any(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN);
    original_signed && candidate_signed && !packages_are_semantically_equal(original, candidate)
}

fn invalid(operation: &'static str, message: String) -> Error {
    Error::InvalidEmbeddedMutation { operation, message }
}
