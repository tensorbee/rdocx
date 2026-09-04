# F-232, correctness, pass 17

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 16 files and 8,335 changed lines, with 7,850 insertions and 485 deletions. All 55 focused `toc_` regression tests, the full `rdocx` regression binary with 241 passes and 2 ignored tests, all 373 `rdocx-oxml` unit tests and its doc test, all 247 `rdocx-layout` unit tests and its doc test, `cargo check -p rdocx --all-targets`, scoped Clippy with warnings denied, `cargo fmt --all --check`, the 49-entry hash harness, the prose check, the generated-skill drift check, and `git diff --check` pass.
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, block content controls still expose children that are invalid for their owner
`crates/rdocx-oxml/src/content_control.rs:578`
`crates/rdocx-oxml/src/content_control.rs:590`
`crates/rdocx-oxml/src/content_control.rs:596`
`crates/rdocx-oxml/src/content_control.rs:609`
`crates/rdocx/src/field.rs:6264`
`crates/rdocx/src/field.rs:6315`
`crates/rdocx-layout/src/engine.rs:99`
`crates/rdocx-layout/src/engine.rs:115`
`crates/rdocx-layout/src/table.rs:44`
`crates/rdocx-layout/src/table.rs:53`
`crates/rdocx-layout/src/table.rs:87`
`crates/rdocx-layout/src/table.rs:723`
`crates/rdocx/tests/regression_test.rs:2783`
`docs/hld/03-architecture.md:514`
`docs/hld/08-rendering-spec.md:1075`

The content-control parser distinguishes only inline from non-inline owners.
For every non-inline control it types paragraphs, tables, rows, and cells from
one unrestricted union. TOC source and public bookmark traversal then recurse
through every typed variant without retaining whether the control belongs to
the body, a table, a row, or a cell. Pagination correctly applies the narrower
owner grammar: a body control emits only paragraphs and tables, a table
control emits only rows, a row control emits only cells, and a cell control
emits only paragraphs and tables.

Consequently, a body control containing a row with a Heading 1 paragraph, a
table control containing a paragraph, or the equivalent mismatched row and
cell cases contributes a live TOC source even though the paragraph never
reaches layout. A TOC without hyperlinks and with that level's page number
omitted commits an entry sourced from the invalid control child. A TOC that
needs a PAGEREF instead rejects at the new unresolved-target boundary. Both
outcomes violate the owner-grammar and opacity contract. The pass-16 matrix
covers only each owner's expected child kind, so it does not distinguish
owner-aware discovery from the current context-free traversal.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-16 D1 is closed for valid body, table, row, and cell block controls.
  Their generated bookmarks resolve to exact pages 2, 3, 4, and 5, the public
  bookmark text agrees, and a second rebuild is stable.
- Unresolved-target diagnostics and atomicity: a PAGEREF whose valid bookmark
  does not reach pagination receives the stable layout diagnostic, loses its
  `TargetPage` classification, and makes rebuild reject before changing the
  live document. No cache mutation path was found because staging owns fresh
  caches and commit occurs only after final reopen.
- Correctness outside D1: no additional wrong-result, stale-result,
  target-association, source-order, bookmark-repair, repeat-build, or
  deterministic-substitution defect was found.
- Contract and public surface: the additive native rebuild operation and its
  compact report remain within the approved plan. Python, WASM, and CLI
  surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion,
  arithmetic, recursion-depth, allocation, or splice panic was found.
- OOXML namespace, ownership, and ordering outside D1: no additional
  expanded-name, direct-owner, wrapper-balance, schema-order, fixed-prefix,
  raw-slot, structural-prefix, or revision-depth defect was found.
- Verbatim preservation: content-control properties and end properties,
  paragraph properties, unowned field scaffolding, comments and processing
  instructions, raw XML, relationships, and untouched package parts remain
  preserved.
- Accepted and tracked projection: hyperlink, insertion, move-to,
  nested-control, and direct marker coordinates remain aligned across facade
  and layout. Deleted and move-from content stays excluded from the accepted
  view.
- Diagnostics outside D1: supported and unsupported complex TOCs plus direct
  and accepted revision simple TOCs retain stable counts. No additional
  diagnostic omission was found.
- Test gate outside D1: the pinned differential metadata and exact entry,
  hyperlink, level, page, raw target range, distinct-page, boundary-fragment,
  repair-policy, unresolved-target, and repeat-rebuild assertions remain
  mutation-sensitive.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
