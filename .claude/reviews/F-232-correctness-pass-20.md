# F-232, correctness, pass 20

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 18 files and 9,547 changed lines, with 8,858 insertions and 689 deletions. All 59 focused `toc_` regression tests, the full `rdocx` regression binary with 248 passes and 2 ignored tests, all 336 `rdocx` library tests with 4 ignored tests, all 375 `rdocx-oxml` unit tests and its doc test, and all 247 `rdocx-layout` unit tests and its doc test in both default and no-default-feature modes pass. `cargo test -p oxml-layout --no-default-features` passes all 102 unit tests and 3 doc tests. `cargo check -p rdocx --all-targets`, scoped all-feature Clippy with warnings denied, the WASM target check, `cargo fmt --all --check`, the 49-entry hash harness, the prose check, the generated-skill drift check, the 94-test sprint-workflow gate with 1 skip, and `git diff --check` pass. The exact locally patched workspace packaging dry run from `/verify` step 10 passes with only `--allow-dirty` added, and every generated archive is below 10 MiB.
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass 19 D1 closure: each candidate start-tag content control is isolated at
  `crates/rdocx/src/field.rs:879`, rebuilt with its effective namespace scope,
  and accepted only when the placement-specific document, table, row, cell, or
  paragraph parser projects that exact direct control at
  `crates/rdocx/src/field.rs:913`. Rejection clears both block and inline
  ownership before descendants are scanned at
  `crates/rdocx/src/field.rs:947`.
- Malformed control behavior: the block and inline regressions at
  `crates/rdocx/tests/regression_test.rs:3072` and
  `crates/rdocx/tests/regression_test.rs:3088` prove that invalid modeled
  properties leave complex and simple TOCs byte-identical and produce no false
  report. The combined case at
  `crates/rdocx/tests/regression_test.rs:3102` also proves that opaque fields,
  bookmarks, paragraphs, and runs cannot shift the later valid TOC, leak public
  bookmarks or fields, add diagnostics, or destabilize the second rebuild.
- Typed and raw acceptance parity: block placement is retained through body,
  table, row, and cell owners, while paragraph-owned controls use only the
  inline grammar. A malformed or foreign control never acquires an accepted
  `sdtContent`, and nested rejected state cannot reconnect to a typed ancestor.
  No mismatch was found among complex-span discovery, simple-field counting,
  bookmark marker lookup, whole-range qualification, or paragraph insertion.
- Namespace handling and verbatim preservation: effective default and
  alternate-prefix bindings are decoded from the live ancestor stack and added
  only to the temporary validation copy at
  `crates/rdocx/src/field.rs:825`. Local declarations are not duplicated. The
  original control, field scaffolding, raw slots, property and end-property
  payloads, relationships, and untouched package parts are not rewritten by
  validation.
- Correctness and contract: no remaining wrong-result, stale-result,
  source-selection, ordering, page-association, bookmark-repair, repeat-build,
  diagnostic, or atomicity defect was found. The additive native
  `Document::rebuild_toc` operation and compact report remain within the
  approved plan, and Python, WASM, and CLI surfaces remain unchanged.
- Panics and bounds: no new reachable indexing, slicing, arithmetic,
  conversion, recursion-depth, allocation, or splice panic was found. The
  control fragment boundary scan reports unmatched or unclosed XML instead of
  underflowing at `crates/rdocx/src/field.rs:793`.
- OOXML ordering and ownership: no remaining expanded-name, direct-owner,
  wrapper-balance, schema-order, fixed-prefix, raw-slot, structural-prefix,
  revision-depth, or opaque-subtree defect was found. Accepted and tracked
  facade projections remain aligned with deterministic layout for body, table,
  row, cell, revision, hyperlink, and nested inline-control paths.
- Test gate: the pinned differential still compares exact ordered entry text,
  style, hyperlink, and PAGEREF tuples. Focused tests remain mutation-sensitive
  for collision-safe substitution, distinct pages, source boundaries, raw
  positions, malformed ownership, maximum ids, bookmark repair, unresolved
  targets, and repeated rebuild.
- Structure and dependencies: no unjustified trait, generic parameter,
  forwarding wrapper, module, feature flag, crate, runtime dependency, or
  published binding surface was introduced. No separate structural smell was
  found.
