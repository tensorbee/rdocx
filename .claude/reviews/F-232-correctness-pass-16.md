# F-232, correctness, pass 16

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 15 files and 7,570 changed lines, with 7,236 insertions and 334 deletions. All 53 focused `toc_` regression tests, the full `rdocx` regression binary with 240 passes and 2 ignored tests, all 373 `rdocx-oxml` unit tests and its doc test, all 246 `rdocx-layout` unit tests and its doc test, `cargo check -p rdocx --all-targets`, scoped Clippy with warnings denied, `cargo fmt --all --check`, the 49-entry hash harness, the prose check, the generated-skill drift check, and `git diff --check` pass.
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, block-content-control TOC targets commit the unresolved page placeholder
`crates/rdocx/src/field.rs:6246`
`crates/rdocx/src/field.rs:6251`
`crates/rdocx-layout/src/engine.rs:1527`
`crates/rdocx-layout/src/engine.rs:1632`
`crates/rdocx-layout/src/table.rs:255`
`crates/rdocx-layout/src/table.rs:611`
`crates/rdocx-layout/src/engine.rs:5468`
`crates/rdocx-layout/src/engine.rs:4508`
`crates/rdocx/src/field.rs:3599`
`crates/rdocx/src/field.rs:3605`
`crates/rdocx/tests/regression_test.rs:2783`
`crates/rdocx/tests/regression_test.rs:2804`
`docs/hld/08-rendering-spec.md:1057`
`docs/hld/08-rendering-spec.md:1071`

TOC discovery recursively accepts paragraphs inside body, table, row, cell,
and block content controls as sources. The main layout transaction, however,
lays out only direct body paragraphs and tables. Its catch-all arm skips a
direct `BodyContent::ContentControl`, while table layout iterates only modeled
rows and explicitly drops `CellContent::ContentControl`. Table-level and
row-level control sidecars are likewise absent from the layout path. A heading
inside any such supported block control can therefore receive a generated
bookmark and a TOC PAGEREF entry without its target paragraph ever reaching a
page.

The failure is then hidden rather than rejected atomically. PAGEREF shaping
uses the fixed text `99` for every `TargetPage`. Post-pagination substitution
leaves that text unchanged when no zero-width target was laid out, and
`deterministic_toc_page_values` collects the surviving `TargetPage` run as if
it were a resolved page. Rebuild consequently writes `99` into the final TOC
instead of the control heading's actual page or an error. The regression
creates exactly a body-level control heading and calls deterministic layout,
but checks only bookmark exposure and entry count. It never checks the final
displayed page, so it passes with the unresolved placeholder. This contradicts
the documented main-story content-control correlation and the requirement
that page substitution use the page containing each target.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-15 D1 is closed. Bookmark projection retains the original in-scope Word
  aliases across direct-run, comment, and bookmark mutation refresh, including
  an alias declared only on the document root.
- Pass-15 D2 is closed. Facade and layout direction checks include marker
  encounter order when accepted coordinates tie. End-before-start is rejected,
  while start-before-end remains a valid empty range.
- Pass-15 D3 is closed. Paragraph-owned content controls preserve paragraph,
  table, row, and cell children as raw XML, including empty forms and nested
  inline controls. Accepted and tracked projections no longer flatten those
  invalid block shapes.
- Correctness outside D1: no additional wrong-result, atomicity,
  deterministic substitution, stale-result, source-order, bookmark-repair, or
  repeat-build defect was found.
- Contract and public surface: the additive native rebuild operation and its
  compact report remain within the approved plan. Python, WASM, and CLI
  surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion,
  arithmetic, recursion-depth, or allocation panic was found.
- OOXML namespace, ownership, and ordering: no additional expanded-name,
  direct-owner, wrapper-balance, schema-order, fixed-prefix, raw-slot,
  structural-prefix, or revision-depth defect was found.
- Verbatim preservation: content-control properties and end properties,
  paragraph properties, unowned field scaffolding, comments and processing
  instructions, raw XML, relationships, and untouched package parts remain
  preserved.
- Accepted and tracked inline projection: hyperlink, insertion, move-to,
  nested-control, and direct marker coordinates remain aligned across facade
  and layout. Deleted and move-from content stays excluded from the accepted
  view.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No additional diagnostic omission
  was found.
- Test gate outside D1: the pinned differential metadata and exact entry,
  hyperlink, level, page, raw target range, distinct-page, boundary-fragment,
  repair-policy, and repeat-rebuild assertions remain mutation-sensitive.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
