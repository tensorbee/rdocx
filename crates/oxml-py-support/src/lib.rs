//! Shared path, revision, and unit support for Python OOXML bindings.

use smallvec::SmallVec;
use thiserror::Error;

use oxml_core::Length;
use pyo3::exceptions::{PyOverflowError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyString;

/// Validate an RGBColor, a legacy channel triple, or optional-hash ASCII hex.
/// Callers handle their own None and auto policies before extraction.
pub fn rgb_channels(value: &Bound<'_, PyAny>, parameter: &str) -> PyResult<(u8, u8, u8)> {
    if value.is_instance_of::<PyString>() {
        let malformed = || {
            PyValueError::new_err(format!(
                "{parameter} must contain exactly six ASCII hexadecimal digits, optionally preceded by one #"
            ))
        };
        let source = value.extract::<String>().map_err(|_| malformed())?;
        let hex = source.strip_prefix('#').unwrap_or(&source);
        if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(malformed());
        }
        let channel = |start| {
            u8::from_str_radix(&hex[start..start + 2], 16)
                .expect("six ASCII hexadecimal digits were validated")
        };
        return Ok((channel(0), channel(2), channel(4)));
    }
    let (red, green, blue) = value.extract::<(i64, i64, i64)>().map_err(|error| {
        if error.is_instance_of::<PyOverflowError>(value.py()) {
            return PyValueError::new_err(format!("{parameter} channels must be integers from 0 to 255"));
        }
        PyTypeError::new_err(format!(
            "{parameter} accepts RGBColor, a legacy triple of integer channels, or six ASCII hexadecimal digits with optional #, got {}",
            value.get_type().name().map(|name| name.to_string()).unwrap_or_else(|_| "unknown type".to_owned())
        ))
    })?;
    let channel = |channel| {
        u8::try_from(channel).map_err(|_| {
            PyValueError::new_err(format!(
                "{parameter} channels must be integers from 0 to 255"
            ))
        })
    };
    Ok((channel(red)?, channel(green)?, channel(blue)?))
}

/// One index step in a Word or PowerPoint content handle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PathSeg {
    /// A slide in a presentation.
    Slide(usize),
    /// A layout in a presentation.
    Layout(usize),
    /// A master in a presentation.
    Master(usize),
    /// A title, body or other master text style selected by its fixed slot.
    TextStyle(usize),
    /// A shape in a slide or group shape.
    Shape(usize),
    /// An item in the document body's block content.
    Body(usize),
    /// A row in a table.
    Row(usize),
    /// A cell in a table row.
    Cell(usize),
    /// A paragraph in its parent container.
    Para(usize),
    /// A run in a paragraph.
    Run(usize),
}

/// An index path paired with the document revision at which it was captured.
#[derive(Debug, Clone)]
pub struct ContentPath {
    /// Ordered index steps from the document root to the content.
    pub segs: SmallVec<[PathSeg; 5]>,
    /// Document revision at handle creation.
    pub revision: u64,
    /// Optional cell-text generation inherited by nested presentation handles.
    pub text_revision: Option<u64>,
    binding: Option<PathBinding>,
}

impl ContentPath {
    /// Capture a path at `revision`.
    pub const fn new(segs: SmallVec<[PathSeg; 5]>, revision: u64) -> Self {
        Self {
            segs,
            revision,
            text_revision: None,
            binding: None,
        }
    }

    /// Reject a handle captured before or after `current_revision`.
    pub fn validate_revision(
        &self,
        current_revision: u64,
        element_kind: &str,
        recovery_hint: &str,
    ) -> Result<(), StaleElementError> {
        if self.revision == current_revision {
            return Ok(());
        }

        Err(StaleElementError {
            element_kind: element_kind.to_owned(),
            captured_revision: self.revision,
            current_revision,
            recovery_hint: recovery_hint.to_owned(),
        })
    }
}

/// A content handle whose captured revision no longer matches its document.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[error(
    "stale {element_kind} handle was created at document revision {captured_revision}, but the document is now at revision {current_revision} (a structural change invalidated it). {recovery_hint}"
)]
pub struct StaleElementError {
    /// Human-readable handle kind, such as `paragraph` or `run`.
    pub element_kind: String,
    /// Revision captured when the handle was created.
    pub captured_revision: u64,
    /// Revision owned by the document when the handle was used.
    pub current_revision: u64,
    /// Caller-supplied guidance for obtaining a fresh handle.
    pub recovery_hint: String,
}

/// Relevant invalidation partitions shared by the two binding families.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum RevisionScope {
    Document,
    Owner,
    Grid,
    Text,
    SlideInventory,
    LayoutInventory,
    MasterInventory,
}

#[derive(Debug)]
struct ScopeGeneration {
    origin: std::sync::Arc<()>,
    generation: u64,
    scope: RevisionScope,
    retired: std::sync::OnceLock<InvalidationCause>,
}

/// Hashes an allocation address, unique while the allocation is held, for
/// the scope set of one preparation.
#[derive(Default)]
struct AddressHasher(u64);

impl std::hash::Hasher for AddressHasher {
    fn finish(&self) -> u64 {
        self.0
    }

    fn write(&mut self, bytes: &[u8]) {
        for byte in bytes {
            self.0 = (self.0.rotate_left(8) ^ u64::from(*byte)).wrapping_mul(0x9e37_79b9_7f4a_7c15);
        }
    }

    fn write_usize(&mut self, address: usize) {
        self.0 = (address as u64).wrapping_mul(0x9e37_79b9_7f4a_7c15);
    }
}

/// One immutable generation retained by the current inventory and live handles.
#[derive(Debug, Clone)]
pub struct RevisionState(std::sync::Arc<ScopeGeneration>);

/// The first successful public mutation that retired a generation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct InvalidationCause {
    pub serial: u64,
    pub operation: Option<&'static str>,
    pub scope: RevisionScope,
}

impl RevisionState {
    pub fn generation(&self) -> u64 {
        self.0.generation
    }
    pub fn scope(&self) -> RevisionScope {
        self.0.scope
    }
    pub fn first_cause(&self) -> Option<&InvalidationCause> {
        self.0.retired.get()
    }
    pub fn same_state(&self, other: &Self) -> bool {
        std::sync::Arc::ptr_eq(&self.0, &other.0)
    }
}

#[derive(Debug, Clone)]
struct PathBinding {
    origin: std::sync::Arc<()>,
    states: SmallVec<[RevisionState; 4]>,
}

/// Prepublication validation failure, with no counter or token mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub enum RevisionError {
    #[error("document transaction serial or owner generation overflowed")]
    Overflow,
    #[error("scope belongs to a different document")]
    ForeignOwner,
    #[error("scope generation is already retired")]
    Retired,
    #[error("scope occurs more than once in one transaction")]
    DuplicateScope,
    #[error("prepared transaction no longer matches the owning counter")]
    StalePreparation,
    #[error("initial owner inventory was already initialized")]
    AlreadyInitialized,
    #[error("unable to allocate the prepared owner inventory")]
    Allocation,
}

/// Preserve checked overflow and allocation failure before native publication.
pub fn revision_to_pyerr(error: RevisionError) -> PyErr {
    match error {
        RevisionError::Overflow => PyOverflowError::new_err(error.to_string()),
        RevisionError::Allocation => pyo3::exceptions::PyMemoryError::new_err(error.to_string()),
        _ => PyValueError::new_err(error.to_string()),
    }
}

/// Fully allocated successors. Dropping this value has no observable effect.
#[derive(Debug)]
pub struct PreparedRevision {
    origin: std::sync::Arc<()>,
    expected_serial: u64,
    expected_generation: u64,
    next_generation: u64,
    operation: Option<&'static str>,
    retire_document: bool,
    retired: Vec<RevisionState>,
    successors: Vec<RevisionState>,
    born: Vec<RevisionState>,
}

impl PreparedRevision {
    /// Successors have the same ordering as the caller's affected scopes.
    pub fn successor(&self, index: usize) -> Option<RevisionState> {
        self.successors.get(index.checked_add(1)?).cloned()
    }
    pub fn committed_serial(&self) -> u64 {
        self.expected_serial + 1
    }

    /// The actual document state that this preparation will publish.
    pub fn document_scope(&self) -> RevisionState {
        self.successors[0].clone()
    }
    pub fn born(&self, index: usize) -> Option<RevisionState> {
        self.born.get(index).cloned()
    }
}

/// One document's common committed serial and checked generation allocator.
#[derive(Debug)]
pub struct RevisionCounter {
    current: u64,
    next_generation: u64,
    origin: std::sync::Arc<()>,
    document: RevisionState,
    initialized: bool,
}

impl Default for RevisionCounter {
    fn default() -> Self {
        Self::new()
    }
}
impl PartialEq for RevisionCounter {
    fn eq(&self, other: &Self) -> bool {
        self.current == other.current
    }
}
impl Eq for RevisionCounter {}

impl RevisionCounter {
    pub fn new() -> Self {
        let origin = std::sync::Arc::new(());
        let document = RevisionState(std::sync::Arc::new(ScopeGeneration {
            origin: origin.clone(),
            generation: 0,
            scope: RevisionScope::Document,
            retired: std::sync::OnceLock::new(),
        }));
        Self {
            current: 0,
            next_generation: 0,
            origin,
            document,
            initialized: false,
        }
    }
    pub const fn current(&self) -> u64 {
        self.current
    }

    /// Capture the real document epoch after eager owner initialization.
    /// Global commits retire this state, while local owner commits retain it.
    pub fn current_document_scope(&self) -> RevisionState {
        self.document.clone()
    }

    fn owns(&self, state: &RevisionState) -> bool {
        std::sync::Arc::ptr_eq(&self.origin, &state.0.origin)
    }

    fn state(&self, generation: u64, scope: RevisionScope) -> RevisionState {
        RevisionState(std::sync::Arc::new(ScopeGeneration {
            origin: self.origin.clone(),
            generation,
            scope,
            retired: std::sync::OnceLock::new(),
        }))
    }

    /// Only construction uses this operation, before returning captured handles.
    pub fn initialize_scopes(
        &mut self,
        scopes: &[RevisionScope],
    ) -> Result<Vec<RevisionState>, RevisionError> {
        if self.initialized
            || self.current != 0
            || std::sync::Arc::strong_count(&self.document.0) != 1
        {
            return Err(RevisionError::AlreadyInitialized);
        }
        let count = u64::try_from(scopes.len()).map_err(|_| RevisionError::Overflow)?;
        let last = self
            .next_generation
            .checked_add(count)
            .ok_or(RevisionError::Overflow)?;
        let mut result = Vec::new();
        result
            .try_reserve_exact(scopes.len())
            .map_err(|_| RevisionError::Allocation)?;
        for (index, scope) in scopes.iter().enumerate() {
            result.push(self.state(self.next_generation + index as u64 + 1, *scope));
        }
        self.next_generation = last;
        self.initialized = true;
        Ok(result)
    }

    pub fn capture(&self, segs: SmallVec<[PathSeg; 5]>) -> ContentPath {
        let mut path = ContentPath::new(segs, self.current);
        path.binding = Some(PathBinding {
            origin: self.origin.clone(),
            states: smallvec::smallvec![self.document.clone()],
        });
        path
    }

    pub fn capture_scopes(
        &self,
        segs: SmallVec<[PathSeg; 5]>,
        scopes: &[RevisionState],
    ) -> Result<ContentPath, RevisionError> {
        if scopes.iter().any(|state| !self.owns(state)) {
            return Err(RevisionError::ForeignOwner);
        }
        if scopes.iter().any(|state| state.first_cause().is_some()) {
            return Err(RevisionError::Retired);
        }
        let mut states = SmallVec::new();
        states
            .try_reserve(scopes.len())
            .map_err(|_| RevisionError::Allocation)?;
        states.extend(scopes.iter().cloned());
        let mut path = ContentPath::new(segs, self.current);
        path.binding = Some(PathBinding {
            origin: self.origin.clone(),
            states,
        });
        Ok(path)
    }

    /// Capture an already allocated future document state before allocating a
    /// returned Python handle. Failure leaves the current counter unchanged.
    pub fn capture_prepared(
        &self,
        prepared: &PreparedRevision,
        segs: SmallVec<[PathSeg; 5]>,
    ) -> Result<ContentPath, RevisionError> {
        if !std::sync::Arc::ptr_eq(&self.origin, &prepared.origin) {
            return Err(RevisionError::ForeignOwner);
        }
        if self.current != prepared.expected_serial
            || self.next_generation != prepared.expected_generation
        {
            return Err(RevisionError::StalePreparation);
        }
        let mut path = ContentPath::new(segs, prepared.committed_serial());
        path.binding = Some(PathBinding {
            origin: self.origin.clone(),
            states: smallvec::smallvec![prepared.successors[0].clone()],
        });
        Ok(path)
    }

    /// Allocate all affected successors and new owner scopes before native publication.
    pub fn prepare_scopes(
        &self,
        operation: &'static str,
        scopes: &[RevisionState],
        born: &[RevisionScope],
    ) -> Result<PreparedRevision, RevisionError> {
        self.prepare(Some(operation), scopes, born, true)
    }

    /// Advance the commit serial while retiring only the supplied owner scopes.
    /// Captured document authority stays live for checked local replacements.
    pub fn prepare_local_scopes(
        &self,
        operation: &'static str,
        scopes: &[RevisionState],
        born: &[RevisionScope],
    ) -> Result<PreparedRevision, RevisionError> {
        self.prepare(Some(operation), scopes, born, false)
    }

    fn prepare(
        &self,
        operation: Option<&'static str>,
        scopes: &[RevisionState],
        born: &[RevisionScope],
        retire_document: bool,
    ) -> Result<PreparedRevision, RevisionError> {
        self.current.checked_add(1).ok_or(RevisionError::Overflow)?;
        let total = scopes
            .len()
            .checked_add(born.len())
            .and_then(|n| n.checked_add(usize::from(retire_document)))
            .ok_or(RevisionError::Overflow)?;
        let next_generation = self
            .next_generation
            .checked_add(u64::try_from(total).map_err(|_| RevisionError::Overflow)?)
            .ok_or(RevisionError::Overflow)?;
        let mut retired = Vec::new();
        retired
            .try_reserve_exact(scopes.len() + 1)
            .map_err(|_| RevisionError::Allocation)?;
        retired.push(self.document.clone());
        // Identity is the shared allocation, as in `same_state`. A set keeps a
        // whole-inventory retirement linear in the number of scopes.
        let mut seen = std::collections::HashSet::<
            *const ScopeGeneration,
            std::hash::BuildHasherDefault<AddressHasher>,
        >::default();
        seen.try_reserve(scopes.len() + 1)
            .map_err(|_| RevisionError::Allocation)?;
        seen.insert(std::sync::Arc::as_ptr(&self.document.0));
        for scope in scopes {
            if !self.owns(scope) {
                return Err(RevisionError::ForeignOwner);
            }
            if scope.first_cause().is_some() {
                return Err(RevisionError::Retired);
            }
            if !seen.insert(std::sync::Arc::as_ptr(&scope.0)) {
                return Err(RevisionError::DuplicateScope);
            }
            retired.push(scope.clone());
        }
        let mut successors = Vec::new();
        successors
            .try_reserve_exact(retired.len())
            .map_err(|_| RevisionError::Allocation)?;
        let mut generation = self.next_generation;
        for (index, old) in retired.iter().enumerate() {
            if index == 0 && !retire_document {
                successors.push(old.clone());
                continue;
            }
            generation += 1;
            successors.push(self.state(generation, old.scope()));
        }
        let mut births = Vec::new();
        births
            .try_reserve_exact(born.len())
            .map_err(|_| RevisionError::Allocation)?;
        for scope in born {
            generation += 1;
            births.push(self.state(generation, *scope));
        }
        Ok(PreparedRevision {
            origin: self.origin.clone(),
            expected_serial: self.current,
            expected_generation: self.next_generation,
            next_generation,
            operation,
            retire_document,
            retired,
            successors,
            born: births,
        })
    }

    /// Under the same exclusive wrapper borrow, commit before infallible native swaps.
    /// All rejection is before the first retirement. No fallible action follows success.
    pub fn commit_scopes(&mut self, prepared: PreparedRevision) -> Result<u64, RevisionError> {
        if !std::sync::Arc::ptr_eq(&self.origin, &prepared.origin) {
            return Err(RevisionError::ForeignOwner);
        }
        if self.current != prepared.expected_serial
            || self.next_generation != prepared.expected_generation
            || !self.document.same_state(&prepared.retired[0])
        {
            return Err(RevisionError::StalePreparation);
        }
        if prepared
            .retired
            .iter()
            .any(|state| state.first_cause().is_some())
        {
            return Err(RevisionError::Retired);
        }
        let serial = self.current + 1;
        for state in prepared
            .retired
            .iter()
            .skip(usize::from(!prepared.retire_document))
        {
            state
                .0
                .retired
                .set(InvalidationCause {
                    serial,
                    operation: prepared.operation,
                    scope: state.scope(),
                })
                .expect("private retirement under the owning exclusive borrow was prevalidated");
        }
        self.current = serial;
        self.next_generation = prepared.next_generation;
        self.document = prepared.successors[0].clone();
        Ok(serial)
    }

    /// Numeric compatibility. Migrated consumers prepare with the real public label.
    pub fn bump(&mut self) -> u64 {
        let prepared = self
            .prepare(None, &[], &[], true)
            .expect("document revision counter overflowed");
        self.commit_scopes(prepared)
            .expect("legacy numeric transaction was just prepared")
    }
}

impl PartialEq for ContentPath {
    fn eq(&self, other: &Self) -> bool {
        self.segs == other.segs
            && self.revision == other.revision
            && self.text_revision == other.text_revision
    }
}
impl Eq for ContentPath {}

impl ContentPath {
    /// Unbound numeric selectors remain current-document explicit addressing.
    pub fn belongs_to(&self, counter: &RevisionCounter) -> bool {
        self.binding
            .as_ref()
            .is_none_or(|binding| std::sync::Arc::ptr_eq(&binding.origin, &counter.origin))
    }

    /// The physical owner is resolved by its checked generation, never visible equality.
    pub fn owner_generation(&self) -> Option<u64> {
        self.binding
            .as_ref()?
            .states
            .iter()
            .rev()
            .find(|state| state.scope() == RevisionScope::Owner)
            .map(RevisionState::generation)
    }

    /// Select the earliest committed retirement, with fixed scope ordering for ties.
    pub fn validate_scopes(
        &self,
        counter: &RevisionCounter,
        kind: &str,
        hint: &str,
    ) -> Result<(), StaleElementError> {
        if !self.belongs_to(counter) {
            return Err(StaleElementError {
                element_kind: kind.to_owned(),
                captured_revision: self.revision,
                current_revision: counter.current(),
                recovery_hint: "belongs to a different document".to_owned(),
            });
        }
        let Some(binding) = &self.binding else {
            return self.validate_revision(
                counter.current(),
                kind,
                &format!(
                    "Causal provenance is unavailable for this explicit numeric selector. {hint}"
                ),
            );
        };
        let first = binding
            .states
            .iter()
            .filter_map(RevisionState::first_cause)
            .min_by_key(|cause| (cause.serial, cause.scope));
        let Some(cause) = first else {
            return Ok(());
        };
        let reason = cause.operation.map_or_else(
            || "Causal provenance is unavailable for this legacy mutation.".to_owned(),
            |operation| {
                format!(
                    "First invalidated by {operation} at committed revision {}.",
                    cause.serial
                )
            },
        );
        Err(StaleElementError {
            element_kind: kind.to_owned(),
            captured_revision: self.revision,
            current_revision: counter.current(),
            recovery_hint: format!("{reason} {hint}"),
        })
    }
}

/// Convert inches to EMU with the canonical truncation behavior.
pub fn emu_from_inches(value: f64) -> i64 {
    Length::inches(value).to_emu()
}

/// Convert centimetres to EMU with the canonical truncation behavior.
pub fn emu_from_centimetres(value: f64) -> i64 {
    Length::cm(value).to_emu()
}

/// Convert millimetres to EMU with the canonical truncation behavior.
pub fn emu_from_millimetres(value: f64) -> i64 {
    Length::mm(value).to_emu()
}

/// Convert points to EMU with the canonical truncation behavior.
pub fn emu_from_points(value: f64) -> i64 {
    Length::pt(value).to_emu()
}

/// Convert twips to EMU.
pub fn emu_from_twips(value: i32) -> i64 {
    Length::twips(value).to_emu()
}

/// Convert EMU to inches.
pub fn inches_from_emu(value: i64) -> f64 {
    Length::emu(value).to_inches()
}

/// Convert EMU to centimetres.
pub fn centimetres_from_emu(value: i64) -> f64 {
    Length::emu(value).to_cm()
}

/// Convert EMU to millimetres.
pub fn millimetres_from_emu(value: i64) -> f64 {
    Length::emu(value).to_cm() * 10.0
}

/// Convert EMU to points.
pub fn points_from_emu(value: i64) -> f64 {
    Length::emu(value).to_pt()
}

/// Convert EMU to twips.
pub fn twips_from_emu(value: i64) -> i32 {
    Length::emu(value).to_twips()
}

/// Convert a native signature report to independent Python dictionaries and lists.
pub fn signature_report_to_py(
    py: Python<'_>,
    report: &oxml_opc::SignatureReport,
) -> PyResult<Py<pyo3::types::PyDict>> {
    use oxml_opc::SignatureIssue;
    use pyo3::types::{PyDict, PyList};
    let result = PyDict::new(py);
    result.set_item("signature_part", &report.signature_part)?;
    if let Some(signer) = &report.signer {
        let identity = PyDict::new(py);
        identity.set_item("subject", &signer.subject)?;
        identity.set_item("issuer", &signer.issuer)?;
        identity.set_item("serial_number", &signer.serial_number)?;
        result.set_item("signer", identity)?;
    } else {
        result.set_item("signer", py.None())?;
    }
    result.set_item("cryptographically_valid", report.cryptographically_valid)?;
    result.set_item("coverage_complete", report.coverage_complete)?;
    result.set_item("covered_parts", &report.covered_parts)?;
    let relationships = PyList::empty(py);
    for relationship in &report.covered_relationships {
        let item = PyDict::new(py);
        item.set_item("source_part", &relationship.source_part)?;
        item.set_item("relationship_id", &relationship.relationship_id)?;
        relationships.append(item)?;
    }
    result.set_item("covered_relationships", relationships)?;
    let issues = PyList::empty(py);
    for issue in &report.issues {
        let item = PyDict::new(py);
        let details = PyDict::new(py);
        let kind = match issue {
            SignatureIssue::MissingPart { uri } => {
                details.set_item("uri", uri)?;
                "missing_part"
            }
            SignatureIssue::ChangedReference {
                uri,
                expected_digest,
                actual_digest,
            } => {
                details.set_item("uri", uri)?;
                details.set_item("expected_digest", expected_digest)?;
                details.set_item("actual_digest", actual_digest)?;
                "changed_reference"
            }
            SignatureIssue::DuplicateReference { uri } => {
                details.set_item("uri", uri)?;
                "duplicate_reference"
            }
            SignatureIssue::MissingRelationship {
                source_part,
                relationship_id,
            } => {
                details.set_item("source_part", source_part)?;
                details.set_item("relationship_id", relationship_id)?;
                "missing_relationship"
            }
            SignatureIssue::DuplicateRelationship {
                source_part,
                relationship_id,
            } => {
                details.set_item("source_part", source_part)?;
                details.set_item("relationship_id", relationship_id)?;
                "duplicate_relationship"
            }
            SignatureIssue::ExternalRelationship {
                source_part,
                relationship_id,
                target,
            } => {
                details.set_item("source_part", source_part)?;
                details.set_item("relationship_id", relationship_id)?;
                details.set_item("target", target)?;
                "external_relationship"
            }
            SignatureIssue::MissingRelationshipTarget {
                source_part,
                relationship_id,
                target,
            } => {
                details.set_item("source_part", source_part)?;
                details.set_item("relationship_id", relationship_id)?;
                details.set_item("target", target)?;
                "missing_relationship_target"
            }
            SignatureIssue::UncoveredPart { part_name } => {
                details.set_item("part_name", part_name)?;
                "uncovered_part"
            }
            SignatureIssue::UncoveredRelationship {
                source_part,
                relationship_id,
            } => {
                details.set_item("source_part", source_part)?;
                details.set_item("relationship_id", relationship_id)?;
                "uncovered_relationship"
            }
            SignatureIssue::SignatureValueMismatch => "signature_value_mismatch",
            SignatureIssue::OrphanSignaturePart { part_name } => {
                details.set_item("part_name", part_name)?;
                "orphan_signature_part"
            }
        };
        item.set_item("kind", kind)?;
        item.set_item("details", details)?;
        issues.append(item)?;
    }
    result.set_item("issues", issues)?;
    Ok(result.unbind())
}

#[cfg(test)]
mod tests {
    use super::*;
    use smallvec::smallvec;

    #[test]
    fn rgb_inputs_are_strict_and_keep_parameter_errors() {
        Python::initialize();
        Python::attach(|py| {
            for source in ["123abc", "#123ABC"] {
                let value = PyString::new(py, source);
                assert_eq!(
                    rgb_channels(value.as_any(), "Font.color").unwrap(),
                    (18, 58, 188)
                );
            }
            for source in ["##123456", "#12345", "１２３４５６", "12 456", "1234567"] {
                let value = PyString::new(py, source);
                let error = rgb_channels(value.as_any(), "Font.color").unwrap_err();
                assert!(error.is_instance_of::<PyValueError>(py));
                assert!(error.to_string().contains("Font.color"));
            }
            let surrogate = py
                .import("builtins")
                .unwrap()
                .getattr("chr")
                .unwrap()
                .call1((0xD800_u32,))
                .unwrap();
            let error = rgb_channels(&surrogate, "Font.color.rgb").unwrap_err();
            assert!(error.is_instance_of::<PyValueError>(py));
            assert!(error.to_string().contains("Font.color.rgb"));
            assert!(error.to_string().contains("ASCII hexadecimal"));
            let value = (1_i64, 2_i64, 3_i64).into_pyobject(py).unwrap();
            assert_eq!(
                rgb_channels(value.as_any(), "Cell.shading").unwrap(),
                (1, 2, 3)
            );
            let invalid = 42_i64.into_pyobject(py).unwrap();
            let error = rgb_channels(invalid.as_any(), "Cell.shading").unwrap_err();
            assert!(error.is_instance_of::<PyTypeError>(py));
            assert!(error.to_string().contains("Cell.shading"));
            assert!(error.to_string().contains("int"));
            let overflow = (i128::MAX, 2_i64, 3_i64).into_pyobject(py).unwrap();
            assert!(
                rgb_channels(overflow.as_any(), "Font.color")
                    .unwrap_err()
                    .is_instance_of::<PyValueError>(py)
            );
            let invalid = (-1_i64, 2_i64, 3_i64).into_pyobject(py).unwrap();
            assert!(
                rgb_channels(invalid.as_any(), "Font.color")
                    .unwrap_err()
                    .is_instance_of::<PyValueError>(py)
            );
        });
    }

    #[test]
    fn stale_path_reports_both_revisions() {
        let path = ContentPath::new(smallvec![PathSeg::Body(0), PathSeg::Para(3)], 4);

        let error: StaleElementError = path
            .validate_revision(5, "paragraph", "Re-fetch it with doc.paragraphs[i].")
            .unwrap_err();

        assert_eq!(error.captured_revision, 4);
        assert_eq!(error.current_revision, 5);
        assert_eq!(
            error.to_string(),
            "stale paragraph handle was created at document revision 4, but the document is now at revision 5 (a structural change invalidated it). Re-fetch it with doc.paragraphs[i]."
        );
    }

    #[test]
    fn matching_revision_accepts_content_path() {
        let path = ContentPath::new(smallvec![PathSeg::Body(0), PathSeg::Para(3)], 4);

        assert!(
            path.validate_revision(4, "paragraph", "Re-fetch it with doc.paragraphs[i].")
                .is_ok()
        );
    }

    #[test]
    fn revision_counter_bumps_after_successful_structure_change() {
        let mut revisions = RevisionCounter::default();
        let before = revisions.capture(smallvec![PathSeg::Body(0)]);

        revisions.bump();
        let after = revisions.capture(smallvec![PathSeg::Body(0)]);

        assert_eq!(before.revision, 0);
        assert_eq!(after.revision, 1);
    }

    #[test]
    fn fx211_prepared_document_scope_captures_returned_owner_before_publication() {
        let mut clock = RevisionCounter::new();
        let owners = clock.initialize_scopes(&[RevisionScope::Owner]).unwrap();
        let old = clock
            .capture_scopes(
                smallvec![PathSeg::Run(0)],
                &[clock.current_document_scope(), owners[0].clone()],
            )
            .unwrap();
        let prepared = clock
            .prepare_scopes(
                "Document.add_paragraph",
                &[],
                &[RevisionScope::Owner, RevisionScope::Text],
            )
            .unwrap();
        // Validate the same origin and preparation serial before combining
        // its actual document state with already allocated owner births.
        clock.capture_prepared(&prepared, smallvec![]).unwrap();
        let mut future = clock
            .capture_scopes(
                smallvec![PathSeg::Run(1)],
                &[
                    prepared.document_scope(),
                    prepared.born(0).unwrap(),
                    prepared.born(1).unwrap(),
                ],
            )
            .unwrap();
        future.revision = prepared.committed_serial();
        let born = prepared.born(1).unwrap();
        assert!(old.validate_scopes(&clock, "run", "refetch").is_ok());
        clock.commit_scopes(prepared).unwrap();
        assert_eq!(future.revision, clock.current());
        assert!(future.validate_scopes(&clock, "run", "refetch").is_ok());
        let local = clock
            .prepare_local_scopes("Run.element.replace_xml", &[born], &[])
            .unwrap();
        clock.commit_scopes(local).unwrap();
        assert!(
            old.validate_scopes(&clock, "run", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("Document.add_paragraph")
        );
        assert!(
            future
                .validate_scopes(&clock, "run", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("Run.element.replace_xml")
        );
    }

    #[test]
    fn fx211_document_scope_accessor_keeps_global_and_owner_authority_distinct() {
        let mut clock = RevisionCounter::new();
        let states = clock
            .initialize_scopes(&[RevisionScope::Owner, RevisionScope::Text])
            .unwrap();
        let document = clock.current_document_scope();
        let owner = clock
            .capture_scopes(smallvec![PathSeg::Para(0)], &states[..1])
            .unwrap();
        let run = clock
            .capture_scopes(
                smallvec![PathSeg::Para(0), PathSeg::Run(0)],
                &[document.clone(), states[0].clone(), states[1].clone()],
            )
            .unwrap();
        let append = clock
            .prepare_scopes("Document.add_paragraph", &[], &[])
            .unwrap();
        clock.commit_scopes(append).unwrap();
        assert!(
            owner
                .validate_scopes(&clock, "paragraph", "refetch")
                .is_ok()
        );
        let later = clock
            .prepare_local_scopes("Paragraph.element.replace_xml", &states[1..], &[])
            .unwrap();
        clock.commit_scopes(later).unwrap();
        assert!(
            owner
                .validate_scopes(&clock, "paragraph", "refetch")
                .is_ok()
        );
        let error = run.validate_scopes(&clock, "run", "refetch").unwrap_err();
        assert!(error.recovery_hint.contains("Document.add_paragraph"));
        assert!(
            !error
                .recovery_hint
                .contains("Paragraph.element.replace_xml")
        );
        assert_eq!(document.first_cause().unwrap().serial, 1);
        assert_eq!(states[1].first_cause().unwrap().serial, 2);
        assert!(!clock.current_document_scope().same_state(&document));
        let mut foreign = RevisionCounter::new();
        foreign.initialize_scopes(&[]).unwrap();
        assert!(matches!(
            foreign.capture_scopes(smallvec![], &states[..1]),
            Err(RevisionError::ForeignOwner)
        ));
        let mut early = RevisionCounter::new();
        let early_document = early.current_document_scope();
        assert!(matches!(
            early.initialize_scopes(&[]),
            Err(RevisionError::AlreadyInitialized)
        ));
        assert_eq!(early.current(), 0);
        assert!(early_document.first_cause().is_none());
    }

    #[test]
    fn fx211_local_owner_publication_keeps_document_and_ancestors_live() {
        let mut clock = RevisionCounter::new();
        let scopes = clock
            .initialize_scopes(&[RevisionScope::Owner, RevisionScope::Text])
            .unwrap();
        let ancestor = clock.capture(smallvec![PathSeg::Para(0)]);
        let child = clock
            .capture_scopes(
                smallvec![PathSeg::Para(0), PathSeg::Run(0)],
                &[clock.document.clone(), scopes[0].clone(), scopes[1].clone()],
            )
            .unwrap();
        let prepared = clock
            .prepare_local_scopes("Paragraph.element.replace_xml", &scopes[1..], &[])
            .unwrap();
        let successor = prepared.successor(0).unwrap();
        let future = clock.capture_prepared(&prepared, smallvec![]).unwrap();
        clock.commit_scopes(prepared).unwrap();
        assert_eq!(clock.current(), 1);
        assert!(
            ancestor
                .validate_scopes(&clock, "paragraph", "refetch")
                .is_ok()
        );
        assert!(scopes[0].first_cause().is_none());
        assert!(
            child
                .validate_scopes(&clock, "run", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("Paragraph.element.replace_xml")
        );
        assert!(
            future
                .validate_scopes(&clock, "paragraph", "refetch")
                .is_ok()
        );
        assert!(successor.first_cause().is_none());
        let later = clock
            .prepare_scopes("Document.add_paragraph", &[], &[])
            .unwrap();
        clock.commit_scopes(later).unwrap();
        assert!(
            ancestor
                .validate_scopes(&clock, "paragraph", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("Document.add_paragraph")
        );
        assert!(
            child
                .validate_scopes(&clock, "run", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("Paragraph.element.replace_xml")
        );
    }

    #[test]
    fn fx211_local_preparation_rejections_and_drops_publish_nothing() {
        let mut clock = RevisionCounter::new();
        let states = clock
            .initialize_scopes(&[RevisionScope::Owner, RevisionScope::Text])
            .unwrap();
        let document = clock.document.clone();
        let serial = clock.current;
        let generation = clock.next_generation;
        let count = std::sync::Arc::strong_count(&document.0);
        let prepared = clock
            .prepare_local_scopes(
                "Run.element.replace_xml",
                &states[1..],
                &[RevisionScope::Owner],
            )
            .unwrap();
        let successor = prepared.successor(0).unwrap();
        let successor_weak = std::sync::Arc::downgrade(&successor.0);
        let born = prepared.born(0).unwrap();
        let born_weak = std::sync::Arc::downgrade(&born.0);
        assert!(prepared.successors[0].same_state(&document));
        drop(successor);
        drop(born);
        drop(prepared);
        assert!(successor_weak.upgrade().is_none());
        assert!(born_weak.upgrade().is_none());
        assert_eq!(std::sync::Arc::strong_count(&document.0), count);
        let mut other = RevisionCounter::new();
        let foreign = other
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        assert!(matches!(
            clock.prepare_local_scopes("replace", &[foreign], &[]),
            Err(RevisionError::ForeignOwner)
        ));
        assert!(matches!(
            clock.prepare_local_scopes("replace", &[states[0].clone(), states[0].clone()], &[]),
            Err(RevisionError::DuplicateScope)
        ));
        let prepared = clock
            .prepare_local_scopes("replace", &states[1..], &[])
            .unwrap();
        assert_eq!(
            other.commit_scopes(prepared),
            Err(RevisionError::ForeignOwner)
        );
        assert_eq!((clock.current, clock.next_generation), (serial, generation));
        assert!(clock.document.same_state(&document));
        assert!(states.iter().all(|state| state.first_cause().is_none()));
        let first = clock
            .prepare_local_scopes("first", &states[1..], &[])
            .unwrap();
        let stale = clock
            .prepare_local_scopes("later", &states[..1], &[])
            .unwrap();
        clock.commit_scopes(first).unwrap();
        assert_eq!(
            clock.commit_scopes(stale),
            Err(RevisionError::StalePreparation)
        );
        assert!(states[0].first_cause().is_none());
        assert!(matches!(
            clock.prepare_local_scopes("replace", &states[1..], &[]),
            Err(RevisionError::Retired)
        ));
        assert_eq!(states[1].first_cause().unwrap().operation, Some("first"));
    }

    #[test]
    fn fx211_local_overflow_counts_only_allocated_successors_and_births() {
        let mut clock = RevisionCounter::new();
        let owner = clock
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        let document = clock.document.clone();
        clock.next_generation = u64::MAX;
        assert!(matches!(
            clock.prepare_local_scopes("replace", std::slice::from_ref(&owner), &[]),
            Err(RevisionError::Overflow)
        ));
        assert!(matches!(
            clock.prepare_local_scopes("add", &[], &[RevisionScope::Owner]),
            Err(RevisionError::Overflow)
        ));
        let empty = clock.prepare_local_scopes("local", &[], &[]).unwrap();
        clock.commit_scopes(empty).unwrap();
        assert_eq!(clock.current, 1);
        assert_eq!(clock.next_generation, u64::MAX);
        assert!(clock.document.same_state(&document));
        assert!(owner.first_cause().is_none());
        clock.current = u64::MAX;
        assert!(matches!(
            clock.prepare_local_scopes("local", &[], &[]),
            Err(RevisionError::Overflow)
        ));
        assert_eq!(clock.current, u64::MAX);
        assert!(document.first_cause().is_none());
    }

    #[test]
    fn fx211_local_future_capture_and_first_cause_use_existing_scope_authority() {
        let mut clock = RevisionCounter::new();
        let states = clock
            .initialize_scopes(&[
                RevisionScope::Owner,
                RevisionScope::Text,
                RevisionScope::Grid,
            ])
            .unwrap();
        let document = clock.document.clone();
        let child = clock
            .capture_scopes(
                smallvec![],
                &[
                    states[2].clone(),
                    document.clone(),
                    states[1].clone(),
                    states[0].clone(),
                ],
            )
            .unwrap();
        let first = clock
            .prepare_local_scopes("Text.replace_xml", &states[1..2], &[RevisionScope::Owner])
            .unwrap();
        assert_eq!(first.successor(0).unwrap().scope(), RevisionScope::Text);
        assert!(first.successor(1).is_none());
        assert!(first.successor(usize::MAX).is_none());
        assert!(first.born(usize::MAX).is_none());
        let born = first.born(0).unwrap();
        let future = clock.capture_prepared(&first, smallvec![]).unwrap();
        let born_path = clock
            .capture_scopes(smallvec![], std::slice::from_ref(&born))
            .unwrap();
        let stale_capture = clock.prepare_local_scopes("pending", &[], &[]).unwrap();
        let foreign = RevisionCounter::new();
        assert!(matches!(
            foreign.capture_prepared(&first, smallvec![]),
            Err(RevisionError::ForeignOwner)
        ));
        clock.commit_scopes(first).unwrap();
        assert!(matches!(
            clock.capture_prepared(&stale_capture, smallvec![]),
            Err(RevisionError::StalePreparation)
        ));
        assert_eq!(future.revision, 1);
        assert!(future.validate_scopes(&clock, "parent", "refetch").is_ok());
        assert!(
            born_path
                .validate_scopes(&clock, "owner", "refetch")
                .is_ok()
        );
        let second = clock
            .prepare_local_scopes(
                "Grid.replace_xml",
                &[states[0].clone(), states[2].clone()],
                &[],
            )
            .unwrap();
        clock.commit_scopes(second).unwrap();
        assert!(
            child
                .validate_scopes(&clock, "run", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("Text.replace_xml at committed revision 1")
        );
        assert_eq!(states[0].first_cause().unwrap().serial, 2);
        assert_eq!(states[2].first_cause().unwrap().serial, 2);
        assert!(clock.document.same_state(&document));
        assert!(
            ContentPath::new(smallvec![], 0)
                .validate_scopes(&clock, "manual", "refetch")
                .is_err()
        );
    }

    #[test]
    fn fx206_scope_tokens_preserve_ancestors_and_first_causal_order() {
        let mut clock = RevisionCounter::new();
        let scopes = clock
            .initialize_scopes(&[
                RevisionScope::Owner,
                RevisionScope::Text,
                RevisionScope::Grid,
            ])
            .unwrap();
        let owner = clock
            .capture_scopes(
                smallvec![PathSeg::Slide(0), PathSeg::Shape(0)],
                &scopes[..1],
            )
            .unwrap();
        let child = clock.capture_scopes(owner.segs.clone(), &scopes).unwrap();
        let first = clock
            .prepare_scopes("TextFrame.text", &scopes[1..2], &[])
            .unwrap();
        let next_text = first.successor(0).unwrap();
        clock.commit_scopes(first).unwrap();
        assert!(owner.validate_scopes(&clock, "shape", "re-fetch").is_ok());
        let second = clock
            .prepare_scopes("Table.rows.delete", &scopes[2..], &[])
            .unwrap();
        clock.commit_scopes(second).unwrap();
        let error = child
            .validate_scopes(&clock, "run", "re-fetch")
            .unwrap_err();
        assert!(
            error
                .recovery_hint
                .contains("TextFrame.text at committed revision 1")
        );
        assert!(!error.recovery_hint.contains("Table.rows.delete"));
        let later = clock
            .prepare_scopes("TextFrame.clear", &[next_text], &[])
            .unwrap();
        clock.commit_scopes(later).unwrap();
        assert_eq!(
            child
                .validate_scopes(&clock, "run", "re-fetch")
                .unwrap_err()
                .recovery_hint,
            error.recovery_hint
        );
    }

    #[test]
    fn fx206_scope_prepare_failure_noop_and_stale_commit_are_atomic() {
        let mut clock = RevisionCounter::new();
        let owner = clock
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        let captured = clock
            .capture_scopes(smallvec![PathSeg::Shape(0)], std::slice::from_ref(&owner))
            .unwrap();
        let prepared = clock
            .prepare_scopes("Shapes.remove", std::slice::from_ref(&owner), &[])
            .unwrap();
        drop(prepared);
        assert_eq!(clock.current(), 0);
        assert_eq!(owner.first_cause(), None);
        let first = clock
            .prepare_scopes("Shapes.remove", std::slice::from_ref(&owner), &[])
            .unwrap();
        let stale = clock
            .prepare_scopes("Shapes.remove", std::slice::from_ref(&owner), &[])
            .unwrap();
        clock.commit_scopes(first).unwrap();
        let cause = *owner.first_cause().unwrap();
        assert_eq!(
            clock.commit_scopes(stale),
            Err(RevisionError::StalePreparation)
        );
        assert_eq!(clock.current(), 1);
        assert_eq!(owner.first_cause(), Some(&cause));
        assert!(
            captured
                .validate_scopes(&clock, "shape", "refetch")
                .is_err()
        );
    }

    #[test]
    fn fx206_bound_paths_keep_numeric_equality_without_foreign_or_reused_owner_authority() {
        let mut clock = RevisionCounter::new();
        let owner = clock
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        let captured = clock
            .capture_scopes(smallvec![PathSeg::Shape(0)], std::slice::from_ref(&owner))
            .unwrap();
        let manual = ContentPath::new(captured.segs.clone(), captured.revision);
        assert_eq!(captured, manual);
        let foreign = RevisionCounter::new();
        assert!(!captured.belongs_to(&foreign));
        assert!(manual.belongs_to(&foreign));
        let prepared = clock
            .prepare_scopes("Shapes.replace", &[owner], &[RevisionScope::Owner])
            .unwrap();
        let recreated = prepared.born(0).unwrap();
        clock.commit_scopes(prepared).unwrap();
        assert_ne!(captured.owner_generation(), Some(recreated.generation()));
        assert!(
            captured
                .validate_scopes(&clock, "shape", "refetch")
                .is_err()
        );
        let fresh = clock.capture_scopes(manual.segs, &[recreated]).unwrap();
        assert!(fresh.validate_scopes(&clock, "shape", "refetch").is_ok());
    }

    #[test]
    fn fx206_scope_lifetimes_have_no_history_and_overflow_has_no_side_effect() {
        let mut clock = RevisionCounter::new();
        let owner = clock
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        let weak = std::sync::Arc::downgrade(&owner.0);
        let prepared = clock
            .prepare_scopes("Shapes.remove", std::slice::from_ref(&owner), &[])
            .unwrap();
        clock.commit_scopes(prepared).unwrap();
        drop(owner);
        assert!(weak.upgrade().is_none());
        clock.current = u64::MAX;
        assert!(matches!(
            clock.prepare_scopes("Shapes.add", &[], &[]),
            Err(RevisionError::Overflow)
        ));
        assert_eq!(clock.current(), u64::MAX);
        assert!(clock.document.first_cause().is_none());
    }

    #[test]
    fn fx206_generation_overflow_and_foreign_duplicate_preflight_preserve_state() {
        let mut clock = RevisionCounter::new();
        let owner = clock
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        let mut other = RevisionCounter::new();
        let foreign = other
            .initialize_scopes(&[RevisionScope::Owner])
            .unwrap()
            .remove(0);
        assert!(matches!(
            clock.prepare_scopes("Shape.remove", &[foreign], &[]),
            Err(RevisionError::ForeignOwner)
        ));
        assert!(matches!(
            clock.prepare_scopes("Shape.remove", &[owner.clone(), owner.clone()], &[]),
            Err(RevisionError::DuplicateScope)
        ));
        assert_eq!(clock.current(), 0);
        assert!(owner.first_cause().is_none());
        let generation = owner.generation();
        clock.next_generation = u64::MAX;
        assert!(matches!(
            clock.prepare_scopes("Shapes.add", &[], &[RevisionScope::Owner]),
            Err(RevisionError::Overflow)
        ));
        assert_eq!(owner.generation(), generation);
        assert!(owner.first_cause().is_none());
        assert_eq!(clock.current(), 0);
        let mut construction = RevisionCounter::new();
        construction.next_generation = u64::MAX;
        assert!(matches!(
            construction.initialize_scopes(&[RevisionScope::Owner]),
            Err(RevisionError::Overflow)
        ));
        assert!(!construction.initialized);
        assert!(construction.document.first_cause().is_none());
    }

    #[test]
    fn fx206_same_transaction_scope_ties_have_one_label_and_future_capture_is_checked() {
        let mut clock = RevisionCounter::new();
        let scopes = clock
            .initialize_scopes(&[
                RevisionScope::Text,
                RevisionScope::Owner,
                RevisionScope::Grid,
            ])
            .unwrap();
        let captured = clock
            .capture_scopes(smallvec![PathSeg::Shape(0)], &scopes)
            .unwrap();
        let prepared = clock
            .prepare_scopes("ShapeCollection.remove", &scopes, &[])
            .unwrap();
        assert!(prepared.successor(usize::MAX).is_none());
        assert!(prepared.born(usize::MAX).is_none());
        let future = clock
            .capture_prepared(&prepared, smallvec![PathSeg::Shape(0)])
            .unwrap();
        let foreign = RevisionCounter::new();
        assert!(matches!(
            foreign.capture_prepared(&prepared, smallvec![]),
            Err(RevisionError::ForeignOwner)
        ));
        clock.commit_scopes(prepared).unwrap();
        assert_eq!(clock.current(), 1);
        assert!(future.validate_scopes(&clock, "shape", "refetch").is_ok());
        let causes = scopes
            .iter()
            .map(|state| *state.first_cause().unwrap())
            .collect::<Vec<_>>();
        assert!(causes.iter().all(|cause| cause.serial == 1 && cause.operation == Some("ShapeCollection.remove")));
        assert_eq!(
            causes
                .iter()
                .min_by_key(|cause| (cause.serial, cause.scope))
                .unwrap()
                .scope,
            RevisionScope::Owner
        );
        assert!(
            captured
                .validate_scopes(&clock, "shape", "refetch")
                .unwrap_err()
                .recovery_hint
                .contains("ShapeCollection.remove at committed revision 1")
        );
        assert!(matches!(
            clock.capture_scopes(smallvec![], &scopes),
            Err(RevisionError::Retired)
        ));
    }

    #[test]
    fn word_path_segments_preserve_nested_order() {
        let segments = smallvec![
            PathSeg::Body(0),
            PathSeg::Row(1),
            PathSeg::Cell(2),
            PathSeg::Para(3),
            PathSeg::Run(4),
        ];

        let path = ContentPath::new(segments.clone(), 7);

        assert_eq!(path.segs, segments);
    }

    #[test]
    fn presentation_path_segments_preserve_repeatable_shape_order() {
        let segments = smallvec![
            PathSeg::Slide(1),
            PathSeg::Shape(2),
            PathSeg::Shape(3),
            PathSeg::Para(4),
            PathSeg::Run(5),
        ];

        let path = ContentPath::new(segments.clone(), 7);

        assert_eq!(path.segs, segments);
    }

    #[test]
    fn python_length_helpers_preserve_rust_truncation() {
        assert_eq!(emu_from_inches(1.75 / 914_400.0), 1);
        assert_eq!(emu_from_inches(-1.75 / 914_400.0), -1);
        assert_eq!(emu_from_centimetres(1.75 / 360_000.0), 1);
        assert_eq!(emu_from_centimetres(-1.75 / 360_000.0), -1);
        assert_eq!(emu_from_millimetres(1.75 / 36_000.0), 1);
        assert_eq!(emu_from_millimetres(-1.75 / 36_000.0), -1);
        assert_eq!(emu_from_points(1.75 / 12_700.0), 1);
        assert_eq!(emu_from_points(-1.75 / 12_700.0), -1);
        assert_eq!(emu_from_twips(1_440), 914_400);
        assert_eq!(emu_from_twips(-1_440), -914_400);
        assert_eq!(inches_from_emu(914_400), 1.0);
        assert_eq!(centimetres_from_emu(914_400), 2.54);
        assert_eq!(millimetres_from_emu(914_400), 25.4);
        assert_eq!(points_from_emu(914_400), 72.0);
        assert_eq!(twips_from_emu(914_400), 1_440);
    }
}

#[cfg(test)]
#[test]
fn signature_reports_keep_all_issue_fields_and_owned_values() {
    use oxml_opc::{
        CoveredRelationship, SignatureIssue, SignatureReport, SignerCertificateIdentity,
    };
    let cases = vec![
        (
            SignatureIssue::MissingPart {
                uri: "uri-value".into(),
            },
            "missing_part",
            vec![("uri", "uri-value")],
        ),
        (
            SignatureIssue::ChangedReference {
                uri: "uri-value".into(),
                expected_digest: "expected_digest-value".into(),
                actual_digest: "actual_digest-value".into(),
            },
            "changed_reference",
            vec![
                ("uri", "uri-value"),
                ("expected_digest", "expected_digest-value"),
                ("actual_digest", "actual_digest-value"),
            ],
        ),
        (
            SignatureIssue::DuplicateReference {
                uri: "uri-value".into(),
            },
            "duplicate_reference",
            vec![("uri", "uri-value")],
        ),
        (
            SignatureIssue::MissingRelationship {
                source_part: "source_part-value".into(),
                relationship_id: "relationship_id-value".into(),
            },
            "missing_relationship",
            vec![
                ("source_part", "source_part-value"),
                ("relationship_id", "relationship_id-value"),
            ],
        ),
        (
            SignatureIssue::DuplicateRelationship {
                source_part: "source_part-value".into(),
                relationship_id: "relationship_id-value".into(),
            },
            "duplicate_relationship",
            vec![
                ("source_part", "source_part-value"),
                ("relationship_id", "relationship_id-value"),
            ],
        ),
        (
            SignatureIssue::ExternalRelationship {
                source_part: "source_part-value".into(),
                relationship_id: "relationship_id-value".into(),
                target: "target-value".into(),
            },
            "external_relationship",
            vec![
                ("source_part", "source_part-value"),
                ("relationship_id", "relationship_id-value"),
                ("target", "target-value"),
            ],
        ),
        (
            SignatureIssue::MissingRelationshipTarget {
                source_part: "source_part-value".into(),
                relationship_id: "relationship_id-value".into(),
                target: "target-value".into(),
            },
            "missing_relationship_target",
            vec![
                ("source_part", "source_part-value"),
                ("relationship_id", "relationship_id-value"),
                ("target", "target-value"),
            ],
        ),
        (
            SignatureIssue::UncoveredPart {
                part_name: "part_name-value".into(),
            },
            "uncovered_part",
            vec![("part_name", "part_name-value")],
        ),
        (
            SignatureIssue::UncoveredRelationship {
                source_part: "source_part-value".into(),
                relationship_id: "relationship_id-value".into(),
            },
            "uncovered_relationship",
            vec![
                ("source_part", "source_part-value"),
                ("relationship_id", "relationship_id-value"),
            ],
        ),
        (
            SignatureIssue::SignatureValueMismatch,
            "signature_value_mismatch",
            vec![],
        ),
        (
            SignatureIssue::OrphanSignaturePart {
                part_name: "part_name-value".into(),
            },
            "orphan_signature_part",
            vec![("part_name", "part_name-value")],
        ),
    ];
    let mut report = SignatureReport {
        signature_part: "/_xmlsignatures/sig1.xml".into(),
        signer: None,
        cryptographically_valid: true,
        coverage_complete: false,
        covered_parts: vec!["/word/document.xml".into()],
        covered_relationships: vec![CoveredRelationship {
            source_part: "/".into(),
            relationship_id: "rId1".into(),
        }],
        issues: vec![],
    };
    use pyo3::types::{PyDict, PyList};
    Python::initialize();
    Python::attach(|py| {
        for (issue, kind, fields) in cases {
            report.issues = vec![issue];
            let result = signature_report_to_py(py, &report).unwrap();
            let result = result.bind(py);
            assert!(result.get_item("signer").unwrap().unwrap().is_none());
            assert_eq!(
                result
                    .get_item("signature_part")
                    .unwrap()
                    .unwrap()
                    .extract::<String>()
                    .unwrap(),
                report.signature_part
            );
            assert!(
                result
                    .get_item("cryptographically_valid")
                    .unwrap()
                    .unwrap()
                    .extract::<bool>()
                    .unwrap()
            );
            assert!(
                !result
                    .get_item("coverage_complete")
                    .unwrap()
                    .unwrap()
                    .extract::<bool>()
                    .unwrap()
            );
            let issues = result.get_item("issues").unwrap().unwrap();
            let issue = issues.cast::<PyList>().unwrap().get_item(0).unwrap();
            let issue = issue.cast::<PyDict>().unwrap();
            assert_eq!(
                issue
                    .get_item("kind")
                    .unwrap()
                    .unwrap()
                    .extract::<String>()
                    .unwrap(),
                kind
            );
            let details = issue.get_item("details").unwrap().unwrap();
            let details = details.cast::<PyDict>().unwrap();
            assert_eq!(details.len(), fields.len());
            for (key, value) in fields {
                assert_eq!(
                    details
                        .get_item(key)
                        .unwrap()
                        .unwrap()
                        .extract::<String>()
                        .unwrap(),
                    value
                );
            }
            result.set_item("covered_parts", vec!["mutated"]).unwrap();
            assert_eq!(report.covered_parts[0], "/word/document.xml");
        }
        report.signer = Some(SignerCertificateIdentity {
            subject: "subject".into(),
            issuer: "issuer".into(),
            serial_number: "serial".into(),
        });
        let result = signature_report_to_py(py, &report).unwrap();
        let signer = result.bind(py).get_item("signer").unwrap().unwrap();
        let signer = signer.cast::<PyDict>().unwrap();
        for (key, expected) in [
            ("subject", "subject"),
            ("issuer", "issuer"),
            ("serial_number", "serial"),
        ] {
            assert_eq!(
                signer
                    .get_item(key)
                    .unwrap()
                    .unwrap()
                    .extract::<String>()
                    .unwrap(),
                expected
            );
        }
    });
}
