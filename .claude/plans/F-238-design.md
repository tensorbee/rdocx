# F-238, Flat OPC and modern Word package variants

**Status**: completed
**Sprint**: S69
**Size**: M
**Depends on**: F-236, F-X077, F-X079

## Problem

`Document` opens only ZIP OPC through `OpcPackage` at
`crates/rdocx/src/document.rs:1868`, and new packages still hard-code the DOCX
main content type near `crates/rdocx/src/document.rs:1437`. The facade cannot
inspect or select DOCM, DOTX, or DOTM identity, and it has no Flat OPC input or
output boundary.

F-236 already owns exact relationship-backed VBA, OLE, ActiveX, and signature
payload preservation. F-238 must reuse that package ownership while adding
modern class identity and Flat OPC conversion. File extensions cannot be the
authority, and class conversion must not remove executable content.

## Spec reference

- `docs/hld/02-scope-and-non-goals.md`, modern OOXML formats and permanent
  legacy-format non-goals.
- `docs/hld/03-architecture.md`, facade ownership, shared package seams, and
  "Facade conventions".
- `docs/hld/04-opc-and-packaging.md`, "The package", "Generalising the
  constructors", "What transfers unmodified", and "Package integrity".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability" and "Native
  Word executable-content inventory".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", "The hash harness", and
  "The Word corpus".
- `docs/hld/14-development-backlog.md`, "F-238, Flat OPC and modern Word
  package variants".
- `docs/hld/15-build-and-toolchain.md`, native public API and archive gates.
- ECMA-376 Part 2, Flat OPC XML package representation.

## Approach

Add the four exact Word main-part content-type constants to the existing
`oxml-opc` content-type module and expose one native package-class enum:

```rust
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum WordPackageClass {
    Document,
    MacroEnabledDocument,
    Template,
    MacroEnabledTemplate,
}

impl Document {
    pub fn package_class(&self) -> Result<WordPackageClass>;
    pub fn to_bytes_as(&self, class: WordPackageClass) -> Result<Vec<u8>>;
    pub fn save_as_package_class<P: AsRef<Path>>(
        &self,
        path: P,
        class: WordPackageClass,
    ) -> Result<()>;
    pub fn from_flat_opc_bytes(bytes: &[u8]) -> Result<Self>;
    pub fn from_flat_opc_bytes_with_limits(
        bytes: &[u8],
        limits: PackageReadLimits,
    ) -> Result<Self>;
    pub fn open_flat_opc<P: AsRef<Path>>(path: P) -> Result<Self>;
    pub fn to_flat_opc_bytes(&self) -> Result<Vec<u8>>;
    pub fn save_flat_opc<P: AsRef<Path>>(&self, path: P) -> Result<()>;
}
```

Keep the bounded Flat OPC parser and deterministic writer in one private
`crates/rdocx/src/flat_opc.rs` module. Parse expanded package names after the
F-X077 lexical validator. Require one `pkg:package`, unique normalized absolute
part names, one content type per part, and exactly one `pkg:xmlData` or
`pkg:binaryData`. Reject schema-position lookalikes, duplicates, malformed
base64, unsafe names, malformed relationship owners, and limit excess before
publishing a document.

Project directly into the existing `OpcPackage`. Relationship parts populate
the existing relationship maps, and all other parts retain their exact bytes
and content types. The writer flushes a staged document, emits fixed `pkg:`
markup in sorted part-name order, writes XML parts as `xmlData`, writes opaque
binary parts as strict base64 `binaryData`, and reopens the result before
returning it. Path saves use the existing atomic publication helper.

The main-part content-type override is the sole class authority. Every ZIP and
Flat OPC open accepts exactly DOCX, DOCM, DOTX, or DOTM. Ordinary saves preserve
the opened class. Output-only conversion changes only that override on a staged
clone and uses the existing signature invalidation contract. VBA, OLE,
ActiveX, relationships, and unrelated parts remain byte-exact and the live
document remains unchanged.

Python, WASM, CLI, and rendering surfaces remain unchanged. The enum and
methods are additive native pre-1.0 API. No new dependency, crate, feature,
trait, generic parameter, public package model, or test binary is added.

## Rejected alternatives

- Infer class from the output extension. The serialized main content type is
  authoritative.
- Store a second class field. It could drift from `[Content_Types].xml`.
- Remove VBA when selecting an ordinary class. Preservation-first conversion
  must not destroy executable payloads.
- Put a Word-specific Flat OPC surface in `oxml-opc`. This story converts
  directly to the current Word document model and does not add a second shared
  public package model.
- Encode every part as binary. Word-compatible Flat OPC uses `xmlData` for XML
  parts and `binaryData` for opaque binary parts.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| round-trip | `flat_opc_and_modern_word_package_classes_reopen_without_repair_and_preserve_payloads` | DOCX, DOCM, DOTX, and DOTM survive ZIP, Flat OPC, document, and ZIP conversion with class, template meaning, relationships, raw XML, and executable bytes intact. |
| round-trip | `flat_opc_xml_and_binary_parts_round_trip_without_loss` | Prefix aliases and default namespaces read, fixed `pkg:` output writes, and XML and binary parts retain their required fidelity. |
| regression | `ordinary_save_preserves_opened_word_template_and_macro_classes` | Ordinary ZIP and Flat OPC saves do not collapse class identity to DOCX. |
| integration | `word_package_class_conversion_changes_only_the_main_content_type` | Every class conversion changes only the main override, retains payloads, invalidates signatures as declared, and leaves the live document unchanged. |
| regression | `unknown_word_main_content_type_fails_closed` | Unknown or ambiguous main types never acquire a supported class. |
| regression | `malformed_or_unsafe_flat_opc_fails_before_document_publication` | Namespace, schema order, duplicate, data-kind, base64, path, relationship, lexical, and limit failures are atomic. |

The **test gate is round-trip**. Each modern package class reopens without
repair and retains its executable payload and template semantics. The
source-built structural gate covers all four classes. A pinned Word 16.104
manual acceptance record confirms representative outputs open without repair.

## HLD impact

- `docs/hld/02-scope-and-non-goals.md`
- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Any parser or serialiser**. Parse expanded names, emit fixed `pkg:` order,
  reopen structurally, and prove unmodelled XML subtrees and opaque binary
  payloads retain their required bytes.
- **Public API of a published crate**. State the additive pre-1.0 impact, run
  rustdoc with warnings denied, run the patched publish dry run, and assert
  every archive remains below 10 MiB.
- **New module or file**. Obtain explicit approval for
  `crates/rdocx/src/flat_opc.rs`. It keeps one cohesive two-way package parser
  out of the already large `document.rs`. Add no trait, generic, crate, or
  feature.
- **External oracle comparison**. Pin Word 16.104, record exact build identity,
  use it only for the no-repair acceptance fact, and keep structural truth in
  source-built tests.
- **Crate dependency graph**. Reuse existing `rdocx` edges and add only content
  type constants to `oxml-opc`. Verify no shared crate gains a format edge.

## Hash harness

Expected unchanged. New opt-in inputs and outputs do not affect current DOCX
serialization or rendering. Any delta blocks integration.

## Implementation checklist

- [x] Complete F-X077 and publish its shared API through F-X079.
- [x] Add the approved private Flat OPC module.
- [x] Add exact modern Word content-type constants and package-class mapping.
- [x] Validate class during every package open and preserve it on ordinary
      saves.
- [x] Implement staged output-only class conversion with signature
      invalidation.
- [x] Implement bounded strict Flat OPC import into the existing `OpcPackage`.
- [x] Implement deterministic fixed-prefix Flat OPC export and atomic path
      save.
- [x] Add all source-built round-trip, preservation, malformed, and limit cases
      to existing test binaries.
- [x] Capture pinned Word no-repair evidence and run every routed gate.

## Open questions

None. The user approved the private `crates/rdocx/src/flat_opc.rs` module as
the cohesive owner for the bounded parser and writer.
