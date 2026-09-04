# F-232, correctness, pass 18

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 18 files and 9,117 changed lines, with 8,428 insertions and 689 deletions. All 57 focused `toc_` regression tests, the full `rdocx` regression binary with 243 passes and 2 ignored tests, all 336 `rdocx` library tests with 4 ignored tests, all 374 `rdocx-oxml` unit tests and its doc test, all 247 `rdocx-layout` unit tests and its doc test, `cargo check -p rdocx --all-targets`, scoped Clippy with warnings denied, `cargo fmt --all --check`, the 49-entry hash harness, the prose check, the generated-skill drift check, and `git diff --check` pass.
**Verdict**: 2 defects, 0 smells, 0 nitpicks

## Defects

### D1, raw TOC ownership still erases the block content control placement
`crates/rdocx/src/field.rs:1310`
`crates/rdocx/src/field.rs:1325`
`crates/rdocx/src/field.rs:1333`
`crates/rdocx/src/field.rs:1355`
`crates/rdocx/src/field.rs:869`
`crates/rdocx/src/field.rs:3258`
`crates/rdocx/tests/regression_test.rs:2947`
`crates/rdocx/tests/regression_test.rs:2976`
`docs/hld/03-architecture.md:514`

The byte-oriented classifier records every first `w:sdtContent` as the same
generic `Content` owner. Under that owner it accepts paragraphs, tables, rows,
and cells as one union. It therefore does not apply the body, table, row, or
cell placement grammar that the remediated OXML parser, typed facade, and
layout now apply. Every raw scanner that uses this classifier can still count
or inspect a paragraph retained as opaque by the typed parser. This includes
complex TOC span discovery, simple TOC diagnostics, bookmark correlation, and
bookmark insertion offsets.

For example, place a table-owned content control with a direct paragraph before
a valid body TOC. The raw span scanner counts that invalid paragraph, while
typed source discovery does not. The raw TOC begin and end indexes are then
compared with a different typed paragraph order, and a generated bookmark can
be inserted into the wrong serialized paragraph. A complex or simple TOC
inside that invalid direct paragraph is also treated as a live field instead
of opaque XML. The new mismatch regressions place the complete invalid matrix
after the TOC and put only SEQ and TC fields inside it, so they do not exercise
either the pre-TOC coordinate shift or an invalidly owned TOC. Pass 17 D1 is
closed in the typed paths but remains open in the raw ownership paths.

### D2, the public standalone content control parser now hardcodes body ownership
`crates/rdocx-oxml/src/content_control.rs:375`
`crates/rdocx-oxml/src/content_control.rs:377`
`crates/rdocx-oxml/src/content_control.rs:380`
`crates/rdocx-oxml/src/content_control.rs:384`
`crates/rdocx-oxml/src/content_control.rs:388`
`crates/rdocx-oxml/src/content_control.rs:392`
`crates/rdocx-oxml/src/content_control.rs:396`
`crates/rdocx-oxml/src/content_control.rs:1067`

`CT_Sdt::from_xml` is an existing public parser whose documentation accepts a
content control without naming a placement, but it now always parses with
`SdtOwner::Body`. The placement-correct entry points for table, row, cell, and
inline controls are all crate-private. A downstream caller parsing a valid
standalone table-level control containing `w:tr`, or an inline control
containing `w:r`, now receives `SdtContent::RawXml` where the public parser
previously exposed the typed row or run, with no public way to request the
correct owner grammar. The new owner matrix invokes only the crate-private
helpers, so it does not protect this published low-level API. This is an
unapproved behavioral break outside the planned additive
`Document::rebuild_toc` surface.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass 17 typed-parser closure: body and cell controls type only paragraphs and
  tables, table controls type only rows, row controls type only cells, and
  inline controls type only runs. Invalid recognized Word children remain raw
  and serialize without loss.
- Typed facade and layout parity: owner-specific traversal agrees for valid
  body, table, row, and cell controls, including nested controls and controls
  moved in memory. The exact page 2 through 5 target matrix remains green.
- Public bookmark and field exposure: the tested owner-mismatch matrix exposes
  no invalid heading, SEQ, TC, or bookmark through typed APIs. D1 is the
  remaining raw-scanner path and D2 is the standalone parser contract.
- Correctness outside D1 and D2: no additional wrong-result, stale-result,
  target-association, source-order, bookmark-repair, repeat-build, diagnostic,
  or atomicity defect was found.
- Panics and bounds: no new reachable indexing, slicing, conversion,
  arithmetic, recursion-depth, allocation, or splice panic was found.
- OOXML ordering and preservation outside D1: no additional expanded-name,
  direct-owner, wrapper-balance, schema-order, fixed-prefix, raw-slot,
  structural-prefix, or revision-depth defect was found. Content-control
  properties, end properties, raw invalid children, field scaffolding,
  relationships, and untouched package parts remain preserved.
- Accepted and tracked projection: hyperlink, insertion, move-to,
  nested-control, and direct marker coordinates remain aligned across facade
  and layout. Deleted and move-from content stays excluded from the accepted
  view.
- Test gate: the pinned differential metadata and exact entry, hyperlink,
  level, page, raw target range, distinct-page, boundary-fragment,
  repair-policy, unresolved-target, and repeat-rebuild assertions remain
  mutation-sensitive outside the missing D1 and D2 trigger shapes.
- Structure and dependencies: no new trait, generic parameter, forwarding
  wrapper, module, feature flag, crate, runtime dependency, Python, WASM, or
  CLI surface was introduced. No separate structural smell was found.
