use std::collections::{BTreeMap, HashSet};
use std::io::Cursor;
use std::ops::Range;

use oxml_opc::OpcPackage;
use oxml_opc::relationship::{Relationship, rel_types};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rpptx_oxml::presentation::CT_Presentation;
use rpptx_oxml::slide_parts::{CT_Slide, CT_SlideLayout, CT_SlideMaster};
use sha2::{Digest, Sha256};

use crate::{Error, Presentation, Result};

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_P_NS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_A_NS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_R_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const ACTIVEX_NS: &str = "http://schemas.microsoft.com/office/2006/activeX";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const DSIG_NS: &str = "http://www.w3.org/2000/09/xmldsig#";
const OPC_SIGNATURE_NS: &str = "http://schemas.openxmlformats.org/package/2006/digital-signature";
const OLE_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";
const INVALIDATED_PACKAGE_SIGNATURE: &str = "urn:rdocx:relationships/invalidated-package-signature";
const INVALIDATED_VBA_SIGNATURE: &str = "urn:rdocx:relationships/invalidated-vba-project-signature";

/// Executable or host-activated content owned by a presentation package.
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
    is_slide_root: bool,
    is_presentation_root: bool,
    is_common_slide_data: bool,
    is_shape_tree: bool,
    is_group_shape: bool,
    is_graphic_frame: bool,
    is_graphic: bool,
    graphic_data_is_ole: bool,
    is_ole_path_container: bool,
    is_controls: bool,
    ole_relationship_ids: Vec<String>,
    control_relationship_id: Option<String>,
}

impl Presentation {
    /// Inventories relationship-owned OLE, ActiveX, and VBA payloads without decoding them.
    pub fn embedded_content(&self) -> Result<Vec<EmbeddedContentInfo>> {
        Ok(self
            .owned_embedded_content()?
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
        let owned = self.find_embedded_content(source_part, relationship_id)?;
        Ok(required_part(&self.staged_package(true)?, &owned.info.target_part)?.to_vec())
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
        let selected = self.find_embedded_content(source_part, relationship_id)?;
        let mut staged = self.consolidated_embedded_candidate()?;
        staged
            .package
            .set_part(&selected.info.target_part, bytes.to_vec());
        staged.invalidate_embedded_signatures(&selected, policy)?;
        self.commit_embedded_candidate(staged)?;
        Ok(self
            .find_embedded_content(source_part, relationship_id)?
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
        let selected = self.find_embedded_content(source_part, relationship_id)?;
        let mut staged = self.consolidated_embedded_candidate()?;
        staged.remove_embedded_content_in_place(&selected, policy)?;
        self.commit_embedded_candidate(staged)
    }

    fn owned_embedded_content(&self) -> Result<Vec<OwnedEmbeddedContent>> {
        let package = self.staged_package(true)?;
        let package_signature_present = has_package_signature(&package)?;
        let package_signature_invalidated = package_signature_present
            && (self.package_signatures_invalidated || self.has_uncommitted_package_changes()?);
        let mut found = BTreeMap::<(String, String), OwnedEmbeddedContent>::new();

        let mut ole_sources = BTreeMap::<String, Vec<u8>>::new();
        for slide in &self.slides {
            ole_sources.insert(
                slide.part_name.clone(),
                slide
                    .slide
                    .to_xml()
                    .map_err(|error| malformed(&slide.part_name, error))?,
            );
        }
        for layout in &self.layouts {
            ole_sources.insert(
                layout.part_name.clone(),
                required_part(&package, &layout.part_name)?.to_vec(),
            );
        }
        for master_id in &self.presentation.slide_master_ids {
            let relationship = required_relationship(
                &package,
                &self.presentation_part,
                &master_id.relationship_id,
            )?;
            require_relationship_kind(
                &self.presentation_part,
                relationship,
                &[rel_types::SLIDE_MASTER],
                "slide master",
            )?;
            let master_part = safe_internal_target(&self.presentation_part, relationship)?;
            let xml = required_part(&package, &master_part)?;
            CT_SlideMaster::from_xml(xml).map_err(|error| malformed(&master_part, error))?;
            ole_sources.insert(master_part, xml.to_vec());
        }
        for (source_part, xml) in ole_sources {
            collect_ole_content(
                self,
                &package,
                &source_part,
                &xml,
                package_signature_present,
                package_signature_invalidated,
                &mut found,
            )?;
        }

        let mut controls = BTreeMap::<String, Vec<(String, String)>>::new();
        let mut control_sources = self
            .slides
            .iter()
            .map(|slide| {
                slide
                    .slide
                    .to_xml()
                    .map(|xml| (slide.part_name.clone(), xml))
                    .map_err(|error| malformed(&slide.part_name, error))
            })
            .collect::<Result<Vec<_>>>()?;
        control_sources.push((
            self.presentation_part.clone(),
            self.presentation
                .to_xml()
                .map_err(|error| malformed(&self.presentation_part, error))?,
        ));
        for (source_part, xml) in control_sources {
            for reference in xml_references(&xml, XmlReferenceKind::Control)? {
                let relationship =
                    required_relationship(&package, &source_part, &reference.relationship_id)?;
                require_relationship_kind(
                    &source_part,
                    relationship,
                    &[rel_types::CONTROL, rel_types::STRICT_CONTROL],
                    "ActiveX control properties",
                )?;
                let control_part = safe_internal_target(&source_part, relationship)?;
                required_part(&package, &control_part)?;
                controls
                    .entry(control_part)
                    .or_default()
                    .push((source_part.clone(), reference.relationship_id));
            }
        }
        for (control_part, owners) in controls {
            let properties = required_part(&package, &control_part)?;
            let binary_relationship_id = active_x_binary_relationship_id(properties)?
                .ok_or_else(|| invalid("inventory embedded content", format!(
                    "{control_part}: ActiveX properties root has no relationship-owned binary"
                )))?;
            let relationship =
                required_relationship(&package, &control_part, &binary_relationship_id)?;
            require_relationship_kind(
                &control_part,
                relationship,
                &[rel_types::ACTIVEX_CONTROL_BINARY],
                "ActiveX binary",
            )?;
            let target = safe_internal_target(&control_part, relationship)?;
            let info = embedded_info(
                self,
                &package,
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

        if let Some(relationships) = package.get_part_rels(&self.presentation_part) {
            ensure_unique_relationship_ids(&self.presentation_part, &relationships.items)?;
            for relationship in &relationships.items {
                if relationship.rel_type != rel_types::VBA_PROJECT {
                    continue;
                }
                let target = safe_internal_target(&self.presentation_part, relationship)?;
                let attached_signature = attached_vba_signature_state(&package, &target)?;
                let info = embedded_info(
                    self,
                    &package,
                    EmbeddedContentKind::VbaProject,
                    &self.presentation_part,
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

    fn find_embedded_content(
        &self,
        source_part: &str,
        relationship_id: &str,
    ) -> Result<OwnedEmbeddedContent> {
        self.owned_embedded_content()?
            .into_iter()
            .find(|owned| {
                owned.info.source_part == source_part
                    && owned.info.relationship_id == relationship_id
            })
            .ok_or_else(|| invalid(
                "resolve embedded content",
                format!("{source_part}: relationship {relationship_id} is not an owned embedded payload"),
            ))
    }

    fn consolidated_embedded_candidate(&self) -> Result<Self> {
        let package = self.staged_package(true)?;
        let mut staged = Self::from_package(package)?;
        staged.embedded_invalidated_signatures = self.embedded_invalidated_signatures.clone();
        staged.package_signatures_invalidated = self.package_signatures_invalidated;
        Ok(staged)
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
        self.invalidate_embedded_signatures(selected, policy)?;
        match selected.info.kind {
            EmbeddedContentKind::OleObject => {
                self.remove_xml_reference(
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                    XmlReferenceKind::Ole,
                )?;
                remove_relationship(
                    &mut self.package,
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                )?;
                delete_if_unreachable(&mut self.package, &selected.info.target_part);
            }
            EmbeddedContentKind::ActiveXControl => {
                let owned_candidates = self
                    .package
                    .get_part_rels(&selected.info.source_part)
                    .map(|relationships| {
                        relationships
                            .items
                            .iter()
                            .filter(|relationship| !is_external(relationship))
                            .map(|relationship| {
                                OpcPackage::resolve_rel_target(
                                    &selected.info.source_part,
                                    &relationship.target,
                                )
                            })
                            .collect::<Vec<_>>()
                    })
                    .unwrap_or_default();
                for (owner_part, owner_relationship_id) in &selected.control_owners {
                    self.remove_xml_reference(
                        owner_part,
                        owner_relationship_id,
                        XmlReferenceKind::Control,
                    )?;
                    remove_relationship(&mut self.package, owner_part, owner_relationship_id)?;
                }
                remove_relationship(
                    &mut self.package,
                    &selected.info.source_part,
                    &selected.info.relationship_id,
                )?;
                delete_if_unreachable(&mut self.package, &selected.info.target_part);
                delete_if_unreachable(&mut self.package, &selected.info.source_part);
                for candidate in owned_candidates {
                    delete_if_unreachable(&mut self.package, &candidate);
                }
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
                delete_if_unreachable(&mut self.package, &selected.info.target_part);
            }
        }
        Ok(())
    }

    fn remove_xml_reference(
        &mut self,
        source_part: &str,
        relationship_id: &str,
        kind: XmlReferenceKind,
    ) -> Result<()> {
        let xml = if source_part == self.presentation_part {
            self.presentation
                .to_xml()
                .map_err(|error| malformed(source_part, error))?
        } else if let Some(slide) = self
            .slides
            .iter()
            .find(|slide| slide.part_name == source_part)
        {
            slide
                .slide
                .to_xml()
                .map_err(|error| malformed(source_part, error))?
        } else if self
            .layouts
            .iter()
            .any(|layout| layout.part_name == source_part)
        {
            required_part(&self.package, source_part)?.to_vec()
        } else if let Some(xml) = self.package.get_part(source_part) {
            CT_SlideMaster::from_xml(xml).map_err(|error| malformed(source_part, error))?;
            xml.to_vec()
        } else {
            return Err(invalid(
                "remove embedded content",
                format!("{source_part}: owning XML part is outside the typed presentation model"),
            ));
        };
        let references = xml_references(&xml, kind)?;
        let ranges = references
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
        let rewritten = remove_ranges(&xml, ranges)?;
        if source_part == self.presentation_part {
            self.presentation = CT_Presentation::from_xml(&rewritten)
                .map_err(|error| malformed(source_part, error))?;
        } else if let Some(slide) = self
            .slides
            .iter_mut()
            .find(|slide| slide.part_name == source_part)
        {
            slide.slide =
                CT_Slide::from_xml(&rewritten).map_err(|error| malformed(source_part, error))?;
        } else if let Some(layout) = self
            .layouts
            .iter_mut()
            .find(|layout| layout.part_name == source_part)
        {
            layout.layout = CT_SlideLayout::from_xml(&rewritten)
                .map_err(|error| malformed(source_part, error))?;
            self.package.set_part(source_part, rewritten);
        } else {
            CT_SlideMaster::from_xml(&rewritten).map_err(|error| malformed(source_part, error))?;
            self.package.set_part(source_part, rewritten);
        }
        Ok(())
    }

    fn commit_embedded_candidate(&mut self, staged: Self) -> Result<()> {
        let embedded_invalidated_signatures = staged.embedded_invalidated_signatures.clone();
        let package_signatures_invalidated = staged.package_signatures_invalidated;
        let mut package = staged.staged_package(true)?;
        persist_invalidated_package_signature(&mut package, package_signatures_invalidated)?;
        let mut output = Cursor::new(Vec::new());
        package.write_to(&mut output)?;
        let mut reopened = Self::from_bytes(output.get_ref())?;
        let issues = reopened.validate();
        if !issues.is_empty() {
            return Err(invalid(
                "mutate embedded content",
                format!("staged presentation validation failed: {issues:?}"),
            ));
        }
        reopened.embedded_invalidated_signatures = embedded_invalidated_signatures;
        reopened.package_signatures_invalidated = package_signatures_invalidated;
        *self = reopened;
        Ok(())
    }

    pub(crate) fn retained_package_signature_would_be_invalidated(&self) -> Result<bool> {
        Ok(has_package_signature(&self.package)? && self.has_uncommitted_package_changes()?)
    }

    fn has_uncommitted_package_changes(&self) -> Result<bool> {
        let staged = self.staged_package(true)?;
        Ok(package_bytes(&staged)? != package_bytes(&self.package)?)
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
    package_signature_invalidation_marked(package).unwrap_or(false)
        || signature_manifest_has_missing_reference(package)
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

fn package_signature_invalidation_marked(package: &OpcPackage) -> Result<bool> {
    package_signature_graph(package)?;
    Ok(package
        .package_rels
        .items
        .iter()
        .any(|relationship| relationship.rel_type == INVALIDATED_PACKAGE_SIGNATURE))
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
                    let path = reference.path;
                    if signature_reference_is_package_manifest(&path) {
                        return false;
                    }
                    if let Some(source) = relationship_source_from_path(&path) {
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
                    !package.contains_part(&path)
                })
            })
    })
}

fn signature_reference_is_package_manifest(path: &str) -> bool {
    path.eq_ignore_ascii_case("/[Content_Types].xml") || path.eq_ignore_ascii_case("/_rels/.rels")
}

#[derive(Debug)]
struct SignatureReference {
    path: String,
    relationship_ids: Vec<String>,
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

fn relationship_source_from_path(path: &str) -> Option<String> {
    let marker = "/_rels/";
    let identity = path.to_ascii_lowercase();
    let marker_index = identity.rfind(marker)?;
    let filename = path.get(marker_index + marker.len()..)?;
    if !filename
        .get(filename.len().saturating_sub(".rels".len())..)
        .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".rels"))
    {
        return None;
    }
    let filename = &filename[..filename.len() - ".rels".len()];
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

fn collect_ole_content(
    presentation: &Presentation,
    package: &OpcPackage,
    source_part: &str,
    xml: &[u8],
    package_signature_present: bool,
    package_signature_invalidated: bool,
    found: &mut BTreeMap<(String, String), OwnedEmbeddedContent>,
) -> Result<()> {
    for reference in xml_references(xml, XmlReferenceKind::Ole)? {
        let relationship = required_relationship(package, source_part, &reference.relationship_id)?;
        require_relationship_kind(
            source_part,
            relationship,
            &[rel_types::OLE_OBJECT, rel_types::STRICT_OLE_OBJECT],
            "OLE object",
        )?;
        let target = safe_internal_target(source_part, relationship)?;
        let info = embedded_info(
            presentation,
            package,
            EmbeddedContentKind::OleObject,
            source_part,
            relationship,
            target,
            SignatureContext {
                package_present: package_signature_present,
                package_invalidated: package_signature_invalidated,
                attached: false,
                attached_invalidated: false,
            },
        )?;
        insert_unique(found, info, Vec::new())?;
    }
    Ok(())
}

fn embedded_info(
    presentation: &Presentation,
    package: &OpcPackage,
    kind: EmbeddedContentKind,
    source_part: &str,
    relationship: &Relationship,
    target_part: String,
    signature: SignatureContext,
) -> Result<EmbeddedContentInfo> {
    let bytes = required_part(package, &target_part)?;
    let content_type = package
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
        || presentation
            .embedded_invalidated_signatures
            .contains(&identity)
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
        [] => Err(Error::MissingRelationship {
            source_part: source_part.to_owned(),
            relationship_id: relationship_id.to_owned(),
        }),
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
    if relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
    {
        return Err(Error::ExternalRelationship {
            source_part: source_part.to_owned(),
            relationship_id: relationship.id.clone(),
        });
    }
    if relationship.target.is_empty()
        || relationship
            .target
            .bytes()
            .any(|byte| byte.is_ascii_control())
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
    package
        .get_part(part_name)
        .ok_or_else(|| Error::MissingPart {
            part_name: part_name.to_owned(),
        })
}

fn validate_identity_source(source_part: &str) -> Result<()> {
    if !source_part.starts_with('/')
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
    let invalidation_markers = relationships
        .items
        .iter()
        .filter(|relationship| relationship.rel_type == INVALIDATED_VBA_SIGNATURE)
        .collect::<Vec<_>>();
    if invalidation_markers.len() > 1 {
        return Err(invalid(
            "inventory embedded content",
            format!(
                "{project_part}: found {} VBA signature invalidation markers, expected at most one",
                invalidation_markers.len()
            ),
        ));
    }
    let Some(signature) = signatures.first() else {
        if invalidation_markers.is_empty() {
            return Ok(None);
        }
        return Err(invalid(
            "inventory embedded content",
            format!("{project_part}: VBA invalidation marker has no signature relationship"),
        ));
    };
    let target = safe_internal_target(project_part, signature)?;
    required_part(package, &target)?;
    if let Some(marker) = invalidation_markers.first() {
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
    let signature_target = package
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
        .add(INVALIDATED_VBA_SIGNATURE, &signature_target);
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
        if !source_part.eq_ignore_ascii_case(&origin_part)
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
        origins: Vec::with_capacity(1),
        signatures: Vec::new(),
    };
    graph.origins.push((origin.id.clone(), origin_part.clone()));
    for signature in signatures {
        let signature_part = safe_internal_target(&origin_part, signature)?;
        if !signature_parts.insert(signature_part.clone()) {
            return Err(invalid(
                "inventory embedded content",
                format!("duplicate digital-signature target {signature_part}"),
            ));
        }
        required_part(package, &signature_part)?;
        graph.signatures.push((origin_part.clone(), signature_part));
    }
    Ok(Some(graph))
}

fn has_package_signature(package: &OpcPackage) -> Result<bool> {
    package_signature_graph(package).map(|graph| graph.is_some())
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

fn remove_relationship(
    package: &mut OpcPackage,
    source_part: &str,
    relationship_id: &str,
) -> Result<()> {
    required_relationship(package, source_part, relationship_id)?;
    let relationships =
        package
            .get_part_rels_mut(source_part)
            .ok_or_else(|| Error::MissingRelationship {
                source_part: source_part.to_owned(),
                relationship_id: relationship_id.to_owned(),
            })?;
    let before = relationships.items.len();
    relationships
        .items
        .retain(|relationship| relationship.id != relationship_id);
    if relationships.items.len() == before {
        return Err(Error::MissingRelationship {
            source_part: source_part.to_owned(),
            relationship_id: relationship_id.to_owned(),
        });
    }
    Ok(())
}

fn delete_if_unreachable(package: &mut OpcPackage, candidate: &str) {
    if relationship_target_is_reachable(package, candidate) {
        return;
    }
    package.remove_part(candidate);
    package.remove_part_rels(candidate);
    package.content_types.remove_override(candidate);
}

fn relationship_target_is_reachable(package: &OpcPackage, candidate: &str) -> bool {
    package.package_rels.items.iter().any(|relationship| {
        !is_external(relationship)
            && OpcPackage::resolve_rel_target("/", &relationship.target)
                .eq_ignore_ascii_case(candidate)
    }) || package
        .part_rels
        .iter()
        .any(|(source_part, relationships)| {
            relationships.items.iter().any(|relationship| {
                !is_external(relationship)
                    && OpcPackage::resolve_rel_target(source_part, &relationship.target)
                        .eq_ignore_ascii_case(candidate)
            })
        })
}

fn is_external(relationship: &Relationship) -> bool {
    relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
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
        package.remove_part(&signature_part);
        package.remove_part_rels(&signature_part);
        package.content_types.remove_override(&signature_part);
    }
    for (_, origin_part) in graph.origins {
        package.remove_part(&origin_part);
        package.remove_part_rels(&origin_part);
        package.content_types.remove_override(&origin_part);
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
    if let Some(relationships) = package.get_part_rels_mut(project_part) {
        relationships.items.retain(|relationship| {
            !is_vba_signature(relationship) && relationship.rel_type != INVALIDATED_VBA_SIGNATURE
        });
    }
    for signature in signatures {
        if !is_external(&signature) {
            let signature_part = safe_internal_target(project_part, &signature)?;
            delete_if_unreachable(package, &signature_part);
        }
    }
    Ok(())
}

fn retain_vba_signature_parts_as_evidence(package: &mut OpcPackage, project_part: &str) {
    let retained = package
        .get_part_rels(project_part)
        .map(|relationships| {
            relationships
                .items
                .iter()
                .filter(|relationship| is_vba_signature(relationship) && !is_external(relationship))
                .map(|relationship| {
                    OpcPackage::resolve_rel_target(project_part, &relationship.target)
                })
                .collect::<HashSet<_>>()
        })
        .unwrap_or_default();
    if let Some(relationships) = package.get_part_rels_mut(project_part) {
        relationships.items.retain(|relationship| {
            !is_vba_signature(relationship) && relationship.rel_type != INVALIDATED_VBA_SIGNATURE
        });
    }
    for part in retained {
        let _ = package.get_part(&part);
    }
}

fn package_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn active_x_binary_relationship_id(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut relationship_id = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                invalid(
                    "inventory embedded content",
                    format!("invalid ActiveX properties XML: {error}"),
                )
            })?;
        let root_namespace = namespace_is(&namespace, &[ACTIVEX_NS]);
        let event = event.into_owned();
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
                })?;
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
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) if depth == 0 => {}
            Event::Text(text)
                if depth == 0 && {
                    let bytes: &[u8] = text.as_ref();
                    bytes.iter().all(u8::is_ascii_whitespace)
                } => {}
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

fn xml_references(xml: &[u8], kind: XmlReferenceKind) -> Result<Vec<XmlReference>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut nodes = Vec::<OpenNode>::new();
    let mut references = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid("scan embedded XML", error.to_string()))?;
        let is_p_namespace = namespace_is(&namespace, &[P_NS, STRICT_P_NS]);
        let is_a_namespace = namespace_is(&namespace, &[A_NS, STRICT_A_NS]);
        let is_mc_namespace = namespace_is(&namespace, &[MC_NS]);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                let is_slide_root =
                    is_p_namespace && matches!(local, b"sld" | b"sldLayout" | b"sldMaster");
                let is_presentation_root = is_p_namespace && local == b"presentation";
                let is_common_slide_data = is_p_namespace && local == b"cSld";
                let is_shape_tree = is_p_namespace && local == b"spTree";
                let is_group_shape = is_p_namespace && local == b"grpSp";
                let is_graphic_frame = is_p_namespace && local == b"graphicFrame";
                let is_graphic = is_a_namespace && local == b"graphic";
                let is_graphic_data = is_a_namespace && local == b"graphicData";
                let graphic_data_is_ole = is_graphic_data
                    && unqualified_attribute(&element, b"uri")?.as_deref()
                        == Some(OLE_GRAPHIC_DATA_URI);
                let is_controls = is_p_namespace && local == b"controls";
                let is_ole_path_container = is_mc_namespace
                    && matches!(local, b"AlternateContent" | b"Choice" | b"Fallback");
                let control_relationship_id = if kind == XmlReferenceKind::Control
                    && is_p_namespace
                    && local == b"control"
                    && control_reference_is_schema_positioned(&nodes)
                {
                    relationship_attribute(&reader, &element)?
                } else {
                    None
                };
                if kind == XmlReferenceKind::Ole
                    && is_p_namespace
                    && local == b"oleObj"
                    && ole_reference_is_schema_positioned(&nodes)
                    && let Some(relationship_id) = relationship_attribute(&reader, &element)?
                    && let Some(frame) = nodes.iter_mut().rev().find(|node| node.is_graphic_frame)
                {
                    frame.ole_relationship_ids.push(relationship_id);
                }
                nodes.push(OpenNode {
                    start: event_start,
                    is_slide_root,
                    is_presentation_root,
                    is_common_slide_data,
                    is_shape_tree,
                    is_group_shape,
                    is_graphic_frame,
                    is_graphic,
                    graphic_data_is_ole,
                    is_ole_path_container,
                    is_controls,
                    ole_relationship_ids: Vec::new(),
                    control_relationship_id,
                });
            }
            Event::Empty(element) => {
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                if kind == XmlReferenceKind::Control
                    && is_p_namespace
                    && local == b"control"
                    && control_reference_is_schema_positioned(&nodes)
                    && let Some(relationship_id) = relationship_attribute(&reader, &element)?
                {
                    references.push(XmlReference {
                        relationship_id,
                        range: event_start..event_end,
                    });
                }
                if kind == XmlReferenceKind::Ole
                    && is_p_namespace
                    && local == b"oleObj"
                    && ole_reference_is_schema_positioned(&nodes)
                    && let Some(relationship_id) = relationship_attribute(&reader, &element)?
                    && let Some(frame) = nodes.iter_mut().rev().find(|node| node.is_graphic_frame)
                {
                    frame.ole_relationship_ids.push(relationship_id);
                }
            }
            Event::End(_) => {
                let node = nodes.pop().ok_or_else(|| {
                    invalid("scan embedded XML", "unmatched closing element".to_owned())
                })?;
                if kind == XmlReferenceKind::Ole && node.is_graphic_frame {
                    let mut relationship_ids = node.ole_relationship_ids;
                    relationship_ids.sort();
                    relationship_ids.dedup();
                    if relationship_ids.len() > 1 {
                        return Err(invalid(
                            "scan embedded XML",
                            format!(
                                "graphic frame at byte {} has ambiguous OLE relationship ids {}",
                                node.start,
                                relationship_ids.join(", ")
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
                references.sort_by_key(|reference| reference.range.start);
                return Ok(references);
            }
            _ => {}
        }
        buffer.clear();
    }
}

fn ole_reference_is_schema_positioned(nodes: &[OpenNode]) -> bool {
    let Some(graphic_data) = nodes.iter().rposition(|node| node.graphic_data_is_ole) else {
        return false;
    };
    if graphic_data < 2
        || !nodes[graphic_data - 1].is_graphic
        || !nodes[graphic_data - 2].is_graphic_frame
        || !schema_owned_graphic_frame(nodes, graphic_data - 2)
    {
        return false;
    }
    nodes[graphic_data + 1..]
        .iter()
        .all(|node| node.is_ole_path_container)
}

fn schema_owned_graphic_frame(nodes: &[OpenNode], frame: usize) -> bool {
    let Some(mut parent) = frame.checked_sub(1) else {
        return false;
    };
    while nodes[parent].is_ole_path_container {
        let Some(next) = parent.checked_sub(1) else {
            return false;
        };
        parent = next;
    }
    while nodes[parent].is_group_shape {
        let Some(next) = parent.checked_sub(1) else {
            return false;
        };
        parent = next;
        while nodes[parent].is_ole_path_container {
            let Some(next) = parent.checked_sub(1) else {
                return false;
            };
            parent = next;
        }
    }
    if !nodes[parent].is_shape_tree || parent < 2 {
        return false;
    }
    nodes[parent - 1].is_common_slide_data && nodes[parent - 2].is_slide_root
}

fn control_reference_is_schema_positioned(nodes: &[OpenNode]) -> bool {
    let Some(controls) = nodes.last() else {
        return false;
    };
    if !controls.is_controls || nodes.len() < 2 {
        return false;
    }
    let mut parent = nodes.len() - 2;
    while nodes[parent].is_ole_path_container {
        let Some(next) = parent.checked_sub(1) else {
            return false;
        };
        parent = next;
    }
    (nodes[parent].is_common_slide_data && parent == 1 && nodes[0].is_slide_root)
        || (nodes[parent].is_presentation_root && parent == 0)
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

fn namespace_is(namespace: &ResolveResult<'_>, expected: &[&str]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if expected.iter().any(|value| *uri == value.as_bytes()))
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

fn malformed(part_name: &str, error: impl std::fmt::Display) -> Error {
    Error::MalformedPart {
        part_name: part_name.to_owned(),
        message: error.to_string(),
    }
}

fn invalid(operation: &'static str, message: String) -> Error {
    Error::InvalidEmbeddedMutation { operation, message }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn signature_manifest_and_origin_owner_paths_use_case_equivalent_identity() {
        assert!(signature_reference_is_package_manifest(
            "/[CONTENT_TYPES].XML"
        ));
        assert!(signature_reference_is_package_manifest("/_RELS/.RELS"));
        let owner =
            relationship_source_from_path("/_XMLSIGNATURES/_RELS/ORIGIN.SIGS.RELS").unwrap();
        assert!(owner.eq_ignore_ascii_case("/_xmlsignatures/origin.sigs"));

        let mut package = OpcPackage::new();
        package.set_part("/_XMLSIGNATURES/ORIGIN.SIGS", Vec::new());
        package.set_part(
            "/_xmlsignatures/sig1.xml",
            br#"<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><SignedInfo><Reference URI="/[CONTENT_TYPES].XML"/></SignedInfo></Signature>"#.to_vec(),
        );
        package.package_rels.add_with_id(
            "origin",
            rel_types::DIGITAL_SIGNATURE_ORIGIN,
            "_xmlsignatures/origin.sigs",
        );
        package
            .get_or_create_part_rels("/_XMLSIGNATURES/ORIGIN.SIGS")
            .add_with_id("signature", rel_types::DIGITAL_SIGNATURE, "sig1.xml");
        assert!(package_signature_graph(&package).unwrap().is_some());
        assert!(!signature_manifest_has_missing_reference(&package));
    }
}
