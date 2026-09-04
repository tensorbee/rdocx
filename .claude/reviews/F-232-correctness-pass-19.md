# F-232, correctness, pass 19

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 18 files and 9,238 changed lines, with 8,549 insertions and 689 deletions. Both focused pass 18 D1 regressions and the focused D2 public-parser regression pass. All 58 focused `toc_` regression tests, the full `rdocx` regression binary with 245 passes and 2 ignored tests, all 336 `rdocx` library tests with 4 ignored tests, all 375 `rdocx-oxml` unit tests and its doc test, and all 247 `rdocx-layout` unit tests and its doc test in both default and no-default-feature modes pass. `cargo check -p rdocx --all-targets`, scoped all-feature Clippy with warnings denied, `cargo fmt --all --check`, the 49-entry hash harness, the prose check, the generated-skill drift check, and `git diff --check` also pass.
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, a content control rejected by the typed parser still owns fields in the raw scanner
`crates/rdocx/src/field.rs:1342`
`crates/rdocx/src/field.rs:1345`
`crates/rdocx/src/field.rs:1691`
`crates/rdocx/src/field.rs:1695`
`crates/rdocx-oxml/src/content_control.rs:416`
`crates/rdocx-oxml/src/content_control.rs:450`
`crates/rdocx-oxml/src/content_control.rs:1086`
`crates/rdocx-oxml/src/document.rs:743`
`crates/rdocx-oxml/src/document.rs:746`
`crates/rdocx-oxml/src/text.rs:2652`
`crates/rdocx-oxml/src/text.rs:2673`
`docs/hld/03-architecture.md:510`

The remediated raw classifier now carries the correct placement once it sees a
Word `w:sdt`, but it decides that the control is typed from its parent and
local name alone. It never verifies that the complete control can be projected
by `CT_Sdt`. The typed body and paragraph parsers do make that check and retain
the entire control as raw XML when it fails. Existing OXML coverage demonstrates
this behavior with a nonnumeric modeled `w:id` property.

Put a complex TOC inside the `w:sdtContent` of that malformed control. The
typed parser preserves the complete control as opaque, while the byte scanner
accepts its content container, paragraph, runs, and field markers. Rebuild can
then replace the cached result inside producer XML that has no typed owner. A
simple TOC in the same control contributes a false diagnostic. The same split
exists for an inline control because `typed_inline_owner` also accepts `w:sdt`
and its first content child without validating the control parser result.
Pass 18 closes valid control placement and invalid child kinds, but a malformed
control shell still crosses the raw ownership boundary.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass 18 D1 direct closure: raw TOC, simple-field, bookmark, and insertion
  scanners now retain body, table, row, and cell placement through nested
  controls. Invalid child kinds before a valid TOC do not shift paragraph
  coordinates, and the tested invalidly owned complex and simple TOCs remain
  byte-identical with zero diagnostics.
- Pass 18 D2 closure: `CT_Sdt::from_xml` uses the context-free standalone union
  again. Paragraph, table, row, cell, and run children remain typed and each
  tested standalone control round-trips exactly, while document-owned parsers
  keep placement-specific entry points.
- Correctness outside D1: no additional wrong-result, stale-result,
  target-association, source-order, bookmark-repair, repeat-build, diagnostic,
  or atomicity defect was found.
- Contract and public surface: the approved additive
  `Document::rebuild_toc` operation and compact report remain intact, and the
  existing standalone parser behavior is restored. Python, WASM, and CLI
  surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion,
  arithmetic, recursion-depth, allocation, or splice panic was found.
- OOXML ordering and preservation outside D1: no additional expanded-name,
  direct-owner, wrapper-balance, schema-order, fixed-prefix, raw-slot,
  structural-prefix, or revision-depth defect was found. Valid content-control
  properties and end properties, invalid child XML, field scaffolding,
  relationships, and untouched package parts remain preserved.
- Facade and layout parity: valid body, table, row, cell, and inline controls
  retain the same accepted paragraph and marker order across source discovery,
  public bookmarks, deterministic layout, and repeated rebuild.
- Accepted and tracked projection: hyperlink, insertion, move-to,
  nested-control, and direct marker coordinates remain aligned. Deleted and
  move-from content stays excluded from the accepted view.
- Test gate: the pinned differential metadata and exact entry, hyperlink,
  level, page, raw target range, distinct-page, boundary-fragment,
  repair-policy, unresolved-target, and repeat-rebuild assertions remain
  mutation-sensitive. D1 identifies the remaining malformed-control-shell
  coverage gap.
- Structure and dependencies: no new trait, generic parameter, forwarding
  wrapper, module, feature flag, crate, runtime dependency, Python, WASM, or
  CLI surface was introduced. No separate structural smell was found.
