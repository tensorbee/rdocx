# F-007, Resolve core properties through the relationship

**Status**: completed
**Sprint**: S02
**Size**: S
**Depends on**: none

## Problem

`Document::from_package` loads metadata only from `/docProps/core.xml` at
`crates/rdocx/src/document.rs:132`, and `flush_to_package` writes it back to the
same hardcoded part at line 224. A valid package may place core properties at a
different target named by the package-level core-properties relationship, so
rdocx currently loses that metadata on load and can create a second orphaned
part on save.

## Spec reference

- `docs/hld/04-opc-and-packaging.md`, "Relationship types".
- `docs/hld/04-opc-and-packaging.md`, "Part naming".

## Approach

Add the package relationship constant
`rdocx_opc::relationship::rel_types::CORE_PROPERTIES`. Resolve the existing
package-level relationship on load, retain its normalized target on `Document`,
and write metadata plus its content-type override back to that target. A new
document, or an imported document with metadata but no relationship, uses
`/docProps/core.xml` and creates the missing relationship.

This is OPC routing around the existing `CoreProperties::from_xml` and
`to_xml`. It does not change either XML parser or serializer.

## Rejected alternatives

- Continue probing the conventional path as a fallback. That preserves the
  defect because the package relationship, not a filename convention, is the
  authority.
- Expand `CoreProperties` to preserve unmodelled XML. That is valuable but is
  a separate parser-model story and not required to correct part routing.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `core_properties_at_relationship_target_round_trip_in_place` | A custom package-level target loads, mutates, saves to the same part, retains the relationship, and does not create `/docProps/core.xml` |
| integration | `metadata_round_trip` | A new document creates conventional metadata with a valid package-level relationship |

The backlog test gate is
`core_properties_at_relationship_target_round_trip_in_place`, proving a
non-standard path round-trips with metadata intact.

## HLD impact

- `docs/hld/04-opc-and-packaging.md`

## Risk routing

- Public API of a published crate. The additive `CORE_PROPERTIES` constant is
  story-required. Run `cargo publish --workspace --dry-run` and assert every
  archive remains below 10 MiB.
- The parser and serializer row does not match because this diff changes OPC
  relationship routing only and leaves XML parsing and writing untouched.

## Hash harness

Expected to remain unchanged across all 28 selected entries. Metadata-bearing
sample packages will intentionally change `_rels/.rels`, which the harness does
not digest.

## Implementation checklist

- [x] Add the package-level core-properties relationship constant.
- [x] Resolve and retain the relationship target during document load.
- [x] Save metadata and its content type to the retained target.
- [x] Create the conventional part and relationship only when none exists.
- [x] Add the custom-target and conventional-target regression assertions.
- [x] Run focused OPC and rdocx integration tests plus the packaging rider.

## Open questions

None. Relationship routing is treated as distinct from changing an XML parser
or serializer, so unmodelled core-property preservation remains out of scope.
