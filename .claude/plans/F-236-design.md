# F-236, Embedded object and macro inventory

**Status**: completed
**Sprint**: S68
**Size**: L
**Depends on**: F-171, F-172

## Problem

The Word facade already preserves legacy object XML as an unsupported raw
subtree, but exposes no relationship-owned inventory or mutation surface. The
preservation regression proves a `w:object` containing `o:OLEObject` remains
opaque and byte-preserved in `crates/rdocx/tests/regression_test.rs:12810`,
while `rdocx` exports no embedded-content types or methods in
`crates/rdocx/src/lib.rs:24` and `Document` only retains the underlying package
and typed main document at `crates/rdocx/src/document.rs:1352`. Callers cannot
identify, hash, extract, replace, or safely remove an OLE object, ActiveX
control, or VBA project.

These payloads can execute in Office. F-236 requires them to remain opaque and
non-executing while inventory identities and hashes remain stable, removal
leaves a valid package, and unrelated edits retain exact payload bytes. The
implementation must also build on the completed signature APIs, which
currently expose package verification and atomic signing at
`crates/rdocx/src/document.rs:1881`.

## Spec reference

- ECMA-376 Part 1, WordprocessingML run-level `w:object` and `w:control`
  ownership and VML Office `o:OLEObject` relationship references.
- ECMA-376 Part 2, normalized internal relationships, content types, and
  package signature invalidation.
- Microsoft Office relationship types for ActiveX binaries, VBA projects, and
  legacy and Agile VBA signatures.
- `docs/hld/02-scope-and-non-goals.md`, "Beyond v1" and "Still non-goals, and
  still permanent".
- `docs/hld/03-architecture.md`, "Why these seams", "Crate-level conventions",
  and "Facade conventions".
- `docs/hld/04-opc-and-packaging.md`, "Relationship types" and "Package
  integrity".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability" and "Native
  PowerPoint executable-content inventory".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy" and "The Word corpus".
- `docs/hld/14-development-backlog.md`, "F-236, Embedded object and macro
  inventory".

## Approach

Add a private `rdocx::embedded` implementation module and re-export the same
concrete vocabulary already proven by `rpptx`, without depending on `rpptx` or
introducing a shared trait or wrapper:

```rust
pub enum EmbeddedContentKind {
    OleObject,
    ActiveXControl,
    VbaProject,
}

pub enum EmbeddedSignatureState {
    Absent,
    Present,
    Invalidated,
}

pub enum EmbeddedMutationPolicy {
    PreserveInvalidatedSignatures,
    RemoveInvalidatedSignatures,
}

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

impl Document {
    pub fn embedded_content(&self) -> Result<Vec<EmbeddedContentInfo>>;
    pub fn extract_embedded_content(
        &self,
        source_part: &str,
        relationship_id: &str,
    ) -> Result<Vec<u8>>;
    pub fn replace_embedded_content(
        &mut self,
        source_part: &str,
        relationship_id: &str,
        bytes: &[u8],
        policy: EmbeddedMutationPolicy,
    ) -> Result<EmbeddedContentInfo>;
    pub fn remove_embedded_content(
        &mut self,
        source_part: &str,
        relationship_id: &str,
        policy: EmbeddedMutationPolicy,
    ) -> Result<()>;
}
```

Use normalized `(source_part, relationship_id)` as identity. Inventory starts
from a flushed staged clone so pending typed edits are represented without
mutating `self`. It scans only schema-positioned Word owner XML in package
relationship scopes: `o:OLEObject/@r:id` within a run-owned `w:object`,
`w:control/@r:id` for an ActiveX properties part, and the main document part's
VBA relationship. ActiveX follows the properties part to exactly one
`ACTIVEX_CONTROL_BINARY` relationship. VBA follows its optional legacy or
Agile signature relationship. Require one relationship identity, the expected
relationship type, an internal non-root-escaping target, an existing target
part, and a resolved content type. Sort results by source part and relationship
id and hash exact stored bytes with `sha2`.

Replacement retains target path, content type, relationship id, and owner XML.
Removal patches only the complete validated `w:object` or `w:control` owner
range, or removes the main-part VBA relationship, then deletes only newly
unreachable owned candidates and their stale relationship set and content-type
override. Shared targets and unrelated orphans survive.

Mutations clone, flush, inventory and validate the complete graph, apply
changes, serialize, reopen as `Document`, re-inventory, and commit with
`commit_staged_mutation` only after success. `PreserveInvalidatedSignatures`
retains exact package and VBA signature bytes and records deterministic
invalidation relationships. `RemoveInvalidatedSignatures` removes only the
validated package-signature graph and the selected VBA signature
infrastructure. Ordinary edits never decode or touch executable bytes.

Reuse the existing `oxml-opc` relationship constants. Add only
`sha2.workspace = true` to `rdocx`. Add a specific embedded-mutation error for
fail-closed diagnostics. No `oxml-opc`, `rdocx-oxml`, binding, WASM, CLI,
feature, trait, generic, or binary-fixture change is needed.

## Rejected alternatives

- Scanning filenames or extensions does not prove producing-scope ownership.
- Parsing OLE CFB, ActiveX binaries, VBA projects, or macro code expands the
  attack surface beyond the opaque inventory contract.
- Deleting every unreferenced embedding part would destroy producer orphans.
- Depending on `rpptx` or moving format-specific owner parsing into `oxml-opc`
  would put Word ownership in the wrong crate.
- Adding the graph and XML mutation implementation to the already large
  `document.rs` would make the security-sensitive path harder to inspect.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `word_embedded_inventory_reports_exact_hashes_relationship_paths_and_signature_state` | Source-built OLE, ActiveX, VBA, package-signature, legacy VBA-signature, and Agile VBA-signature graphs report stable identity, metadata, SHA-256, and signature state in order. |
| integration | `word_embedded_extract_replace_and_remove_are_atomic` | Extraction is exact, replacement retains identity, removal is ownership-aware, save and reopen succeeds, and every failure leaves bytes unchanged. |
| regression | `ordinary_document_edits_preserve_every_embedded_payload_byte` | Main-story and unrelated typed edits retain OLE, ActiveX, VBA, and signature payload bytes. |
| regression | `word_embedded_removal_deletes_only_newly_unreachable_owned_candidates` | Shared payloads and unrelated orphans remain while selected owner XML, relationships, parts, and overrides disappear. |
| regression | `word_embedded_mutation_policy_preserves_or_removes_invalidated_signature_evidence` | Both policies affect only package and selected VBA signature infrastructure and never report stale evidence as valid. |
| regression | `unsafe_or_malformed_word_embedded_graphs_fail_closed_without_mutation` | External and traversal targets, wrong types, missing metadata, duplicate identities, ambiguous owners, and malformed signature graphs fail before mutation. |
| round-trip | `word_embedded_owner_removal_preserves_every_unrelated_raw_xml_byte` | Prefix aliases and compatibility wrappers remain byte-exact outside the selected complete owner range after save and reopen. |

The exact backlog **test gate is regression**: "Inventory hashes and
relationship paths remain stable, safe removal leaves a valid document, and
unrelated edits preserve payload bytes."

Construct every fixture in the existing
`crates/rdocx/tests/regression_test.rs` binary. Do not add a test binary or
binary fixture.

## HLD impact

- `docs/hld/02-scope-and-non-goals.md`
- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- Any parser or serializer: re-read `docs/hld/04-opc-and-packaging.md` and
  `docs/hld/06-presentationml-model.md`. Add prefix-alias, schema-position,
  structural-reopen, and byte-exact unsupported-subtree checks. Source-range
  removal must not canonicalize retained XML or disturb schema child order.
- Public API of a published crate: this is additive pre-1.0 native Rust API.
  Run `cargo publish --dry-run -p rdocx` and assert the `.crate` remains below
  10 MiB. Confirm Python, WASM, and CLI surfaces do not change.
- New module or file: explicit approval is required for
  `crates/rdocx/src/embedded.rs`. It isolates security-sensitive graph walking
  and mutation from the large document facade. No trait, generic, crate, or
  feature is introduced.
- Crate dependency graph: `rdocx` adds its existing workspace `sha2`
  dependency, with no new cross-family edge. Run `cargo tree -p rdocx -e
  normal` and `no_shared_crate_depends_on_a_format_crate`.

## Hash harness

Expected unchanged. Inventory is read-only and replacement or removal is
opt-in. Any existing sample output delta is unexplained and blocks integration.
M22 is outside the mandatory M1 to M6 window, but `/verify` still runs the
repository harness.

## Implementation checklist

- [x] Add the approved private embedded module, `sha2` dependency, facade
  exports, and specific mutation error.
- [x] Discover Word OLE, ActiveX, and VBA ownership only through normalized,
  exact-type internal relationships and schema-position owner XML.
- [x] Produce deterministic inventory facts and exact SHA-256 hashes without
  decoding payloads.
- [x] Implement byte-exact extraction.
- [x] Implement staged identity-preserving replacement.
- [x] Implement complete owner-range removal and newly-unreachable candidate
  cleanup.
- [x] Implement both signature policies for package, legacy VBA, and Agile VBA
  signature evidence.
- [x] Validate malformed graphs and guarantee byte-for-byte failure atomicity.
- [x] Add all source-built cases to the existing regression binary.
- [x] Run focused `rdocx` checks, every risk rider, and `/verify`.

## Open questions

None. The focused private module, ActiveX scope, and explicit
preserve-or-remove signature policy are approved. No decoder, preview
renderer, binary `.doc` reader, package-class selector, Python, WASM, or CLI
API is included.
