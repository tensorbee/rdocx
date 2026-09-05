# F-237, integration, pass 17

**Reviewed**: staged F-237 squash integration against `b941c55`, 34 files,
10,260 additions and 215 deletions, including the `Document` conflict
reconciliation with integrated F-236, both approved plans, both clean final
feature reviews, the combined exports and HLD, and current focused check and
test evidence
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

- **D1, the staged diff hygiene gate fails.** The staged copies of the pass-15
  and pass-16 review artifacts each add a blank line after their final content
  (`.claude/reviews/F-237-all-pass-15.md:60`,
  `.claude/reviews/F-237-all-pass-16.md:49`). `git diff --cached --check`
  reports `new blank line at EOF` for both paths, so the current index cannot
  pass the required verification gate.

## Smells

None.

## Nitpicks

None.

## Not found

No conflict markers or integration semantic loss were found. The reconciled
`Document` retains both F-236 signature fields and all three F-237 glossary
fields (`crates/rdocx/src/document.rs:1401`,
`crates/rdocx/src/document.rs:1405`). Both state groups are initialized by
`Document::new`, copied by `clone_for_staging`, and reconstructed by
`from_package` (`crates/rdocx/src/document.rs:1792`,
`crates/rdocx/src/document.rs:1832`,
`crates/rdocx/src/document.rs:2032`,
`crates/rdocx/src/document.rs:2068`).

No staged-mutation signature interaction defect was found. Building-block
replacement edits a clone, marks the glossary dirty, serializes and reopens
the candidate, then commits only the validated reopened state
(`crates/rdocx/src/building_block.rs:370`,
`crates/rdocx/src/building_block.rs:377`,
`crates/rdocx/src/building_block.rs:392`). Legacy form mutation likewise uses
the shared clone, serialization, reopen, and commit path for both the main
document and package stories (`crates/rdocx/src/field.rs:108`,
`crates/rdocx/src/field.rs:144`, `crates/rdocx/src/field.rs:158`). The shared
serialization path flushes a dirty glossary before comparing the complete OPC
state and persists package-signature invalidation before writing
(`crates/rdocx/src/document.rs:2213`,
`crates/rdocx/src/document.rs:2365`,
`crates/rdocx/src/embedded.rs:688`). The final staged commit also preserves an
existing invalidation or derives one from a semantic package change
(`crates/rdocx/src/document.rs:1845`).

No export or HLD coexistence findings were found. The crate root retains both
the embedded and building-block modules and exports, plus the legacy form
types (`crates/rdocx/src/lib.rs:24`, `crates/rdocx/src/lib.rs:29`,
`crates/rdocx/src/lib.rs:46`, `crates/rdocx/src/lib.rs:57`,
`crates/rdocx/src/lib.rs:62`). The combined HLD retains the F-236 embedded and
signature contract alongside the F-237 form and glossary contract
(`docs/hld/02-scope-and-non-goals.md:158`,
`docs/hld/02-scope-and-non-goals.md:171`,
`docs/hld/04-opc-and-packaging.md:549`,
`docs/hld/04-opc-and-packaging.md:165`,
`docs/hld/10-bindings-spec.md:929`,
`docs/hld/10-bindings-spec.md:249`).

`cargo check -p rdocx --all-targets`, `cargo check -p rdocx --all-targets
--all-features`, and `cargo fmt --all --check` pass. The focused F-236
embedded regression command passes all 62 tests, and the focused F-237 facade
integration command passes all 38 tests. `python3 scripts/prose_check.py
--staged` passes. `git diff --cached --check` fails only for D1.
