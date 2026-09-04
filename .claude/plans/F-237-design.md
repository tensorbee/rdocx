# F-237, Forms, glossary, and building blocks

**Status**: approved
**Sprint**: S68
**Size**: L
**Depends on**: none

## Problem

The Word field parser recognizes a complex field only from its
`w:fldCharType` and dirty marker at `crates/rdocx-oxml/src/text.rs:1367`, then
skips every child of `w:fldChar`. Legacy form metadata in `w:ffData` is
therefore retained only inside the field's raw source. `FieldRef` exposes the
instruction and cached result at `crates/rdocx/src/run.rs:223`, but has no typed
form kind, constraints, choices, or value mutation.

The package loader resolves the main document, styles, numbering, settings,
footnotes, and comments from `crates/rdocx/src/document.rs:1910`, but neither
`oxml-opc` nor `rdocx-oxml` defines the glossary-document relationship and
root model. Content controls already recognize document-part markers at
`crates/rdocx-oxml/src/content_control.rs:17`, but callers cannot inventory or
edit the glossary entries, AutoText, and building blocks those markers select.

## Spec reference

- ECMA-376 Part 1, WordprocessingML `w:fldChar/w:ffData`,
  `w:glossaryDocument`, `w:docParts`, `w:docPartPr`, and `w:docPartBody`.
- ECMA-376 Part 2, normalized internal relationships and content types.
- `docs/hld/02-scope-and-non-goals.md`, "Beyond v1" and the permanent binary
  `.doc` boundary.
- `docs/hld/03-architecture.md`, "Why these seams", "What stays put",
  "Crate-level conventions", and "Facade conventions".
- `docs/hld/04-opc-and-packaging.md`, "Relationship types", "Part naming",
  and "Package integrity".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy" and "The Word corpus".
- `docs/hld/14-development-backlog.md`, "F-237, Forms, glossary, and building
  blocks".

## Approach

Extend the existing field model with a bounded legacy-form projection while
retaining the complete complex-field XML as the serialization source:

```rust
pub enum LegacyFormFieldKind {
    TextInput,
    CheckBox,
    DropDownList,
}

pub enum LegacyFormFieldValue {
    Text(String),
    Checked(bool),
    SelectedIndex(usize),
}

pub struct LegacyFormFieldInfo {
    pub source_part: String,
    pub ordinal: usize,
    pub name: Option<String>,
    pub enabled: bool,
    pub calculate_on_exit: bool,
    pub kind: LegacyFormFieldKind,
    pub value: LegacyFormFieldValue,
    pub choices: Vec<String>,
}

impl Document {
    pub fn legacy_form_fields(&self) -> Result<Vec<LegacyFormFieldInfo>>;
    pub fn set_legacy_form_field_value(
        &mut self,
        source_part: &str,
        ordinal: usize,
        value: LegacyFormFieldValue,
    ) -> Result<LegacyFormFieldInfo>;
}
```

Inventory only `w:ffData` beneath a begin `w:fldChar` in supported modern Word
story parts. Use normalized source part plus source-order ordinal as identity,
because form names are optional and not unique. Parse common `name`, `enabled`,
`calcOnExit`, text-input, checkbox, and drop-down properties with namespace
aware matching. Retain unknown attributes and children in their exact slots.
Mutation changes only the typed value and related cached display where
applicable, validates kind-specific bounds, serializes a staged clone, reopens
it, and commits only after the selected identity and value round trip.

Add the shared glossary-document relationship and content-type constants to
`oxml-opc`. Add a focused `rdocx-oxml::glossary` root model that parses one
`w:glossaryDocument`, ordered `w:docPart` entries, their bounded properties,
and `w:docPartBody` content. The parser is prefix-tolerant and preserves every
unmodelled attribute, child, whitespace boundary, compatibility wrapper, and
unsupported body subtree. The writer emits fixed Word prefixes and schema
order only for changed typed slots, while an unchanged entry reuses its exact
source bytes.

Expose owned building-block values through the native facade:

```rust
pub enum BuildingBlockKind {
    AutoText,
    BuildingBlock,
}

pub struct BuildingBlock {
    pub name: String,
    pub kind: BuildingBlockKind,
    pub category: Option<String>,
    pub description: Option<String>,
    pub guid: Option<String>,
    pub gallery: Option<String>,
    pub behaviors: Vec<String>,
    pub body: CT_Body,
}

pub struct BuildingBlockInfo {
    pub glossary_part: String,
    pub ordinal: usize,
    pub block: BuildingBlock,
}

impl Document {
    pub fn building_blocks(&self) -> Result<Vec<BuildingBlockInfo>>;
    pub fn replace_building_block(
        &mut self,
        glossary_part: &str,
        ordinal: usize,
        block: BuildingBlock,
    ) -> Result<BuildingBlockInfo>;
}
```

Resolve exactly one internal glossary relationship from the main document,
require the expected content type and existing target, and reject duplicate,
external, traversal-shaped, or missing graphs. Use glossary part plus ordinal
as stable identity because entry names are not guaranteed unique. Classify an
entry as AutoText only from its typed `w:types/w:type` value. Replacement
changes only the selected entry, preserves every unrelated entry and raw
subtree, stages the package, serializes, reopens, validates, and commits
atomically.

Keep the public surface native Rust only. Do not add a binary `.doc` reader,
field execution, implicit building-block expansion, Python, WASM, CLI, trait,
generic, feature, new crate, integration binary, or binary fixture.

## Rejected alternatives

- Using form names or building-block names as identity would be ambiguous when
  a producer emits duplicates or omits a form name.
- Treating cached field text as the form definition would discard kind,
  constraints, and producer metadata.
- Parsing every glossary property would exceed the bounded native surface and
  increase the chance of rewriting safe unsupported content.
- Expanding AutoText or building blocks into the main story implicitly would
  make inventory mutate document content and would guess caller intent.
- Adding a binary `.doc` reader is outside the permanent legacy-format scope.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| round-trip | `legacy_form_fields_round_trip_typed_values_and_preserve_unmodelled_ffdata` | Text, checkbox, and drop-down forms expose typed values, edits persist after save and reopen, prefix aliases are accepted, and unsupported `w:ffData` children remain byte-exact. |
| regression | `legacy_form_field_identity_is_story_part_and_source_ordinal` | Missing and duplicate names remain independently addressable in deterministic source order. |
| regression | `invalid_legacy_form_mutations_are_atomic` | Wrong value kinds, out-of-range selections, malformed owners, and stale identities reject without changing document bytes. |
| round-trip | `glossary_entries_autotext_and_building_blocks_round_trip` | Relationship-resolved entries expose typed metadata and body content, selected replacement persists, and AutoText classification is stable. |
| round-trip | `unrelated_building_block_edits_preserve_every_unsupported_subtree_byte` | Unknown properties, compatibility wrappers, unsupported body children, whitespace, and unselected entries remain byte-identical. |
| regression | `unsafe_or_malformed_glossary_graphs_fail_closed` | Duplicate, external, traversal-shaped, wrong-content-type, missing-part, malformed-root, and stale-ordinal cases reject atomically. |

The exact backlog **test gate is round-trip**: "Supported entries remain
editable and every unsupported subtree survives unrelated document edits."

Construct fixtures in the existing `rdocx-oxml` unit target and
`crates/rdocx/tests/integration_test.rs` binary. Do not add another integration
binary or a binary fixture.

## HLD impact

- `docs/hld/02-scope-and-non-goals.md`
- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- Any parser or serializer: re-read `docs/hld/04-opc-and-packaging.md` and
  `docs/hld/06-presentationml-model.md`. Add prefix-alias, fixed-prefix,
  schema-order, structural-reopen, and byte-exact retained-subtree tests for
  form and glossary XML.
- Public API of a published crate: this is additive native Rust API for the
  pre-1.0 `rdocx` facade and stable-family `rdocx-oxml` and `oxml-opc` crates.
  Run publish dry runs for all three and assert every `.crate` remains below
  10 MiB. Confirm Python, WASM, and CLI surfaces do not change.
- Crate dependency graph: shared glossary constants remain in `oxml-opc`, the
  WordprocessingML model remains in `rdocx-oxml`, and package mutation remains
  in `rdocx`. Run `cargo tree -p rdocx -e normal` and
  `no_shared_crate_depends_on_a_format_crate`.
- New module or file: explicit approval is required for
  `crates/rdocx-oxml/src/glossary.rs` and
  `crates/rdocx/src/building_block.rs`. They keep the root glossary grammar and
  facade mutation path locally inspectable without creating a trait, generic,
  crate, or feature.

## Hash harness

Expected unchanged. The new readers are inert and mutations are opt-in. Any
existing sample output delta is unexplained and blocks integration. M22 is
outside the mandatory M1 to M6 window, but `/verify` still runs the repository
harness.

## Implementation checklist

- [ ] Add typed `w:ffData` parsing and exact-slot serialization to the existing
  field model.
- [ ] Inventory legacy forms in deterministic story-part order with explicit
  source-part and ordinal identity.
- [ ] Implement bounded staged form-value mutation and cached-display updates.
- [ ] Add the glossary relationship and content-type constants.
- [ ] Add the approved glossary OXML model and facade building-block module.
- [ ] Parse and serialize bounded building-block metadata and body content
  while preserving every unsupported subtree.
- [ ] Implement deterministic glossary inventory and atomic selected-entry
  replacement.
- [ ] Reject unsafe graphs, stale identities, malformed values, and partial
  mutations.
- [ ] Add source-built unit, round-trip, regression, and atomicity cases to
  existing test targets.
- [ ] Run focused checks, every risk rider, and `/verify`.

## Open questions

None. The two focused implementation files and part-scoped ordinal identities
are approved. Inventory covers supported internal Word stories. Form mutation
is limited to typed values and cached display, while building-block mutation
replaces one existing typed entry's supported metadata and body. The feature
does not execute fields, expand AutoText, author a new glossary part, or expose
a second binding surface.
