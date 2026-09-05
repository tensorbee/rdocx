# F-237, integration, pass 18

**Reviewed**: staged F-237 squash integration against `b941c55`, 35 files,
10,331 additions and 215 deletions, including the pass-17 D1 correction, both
approved plans, the complete staged source and test delta, and the reconciled
F-236 and F-237 document state, mutation paths, exports, and HLD
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Not found

The pass-17 diff-hygiene defect is closed. The pass-15 and pass-16 review
artifacts now end immediately after their final content with one terminating
newline and no trailing blank line
(`.claude/reviews/F-237-all-pass-15.md:59`,
`.claude/reviews/F-237-all-pass-16.md:48`). `git diff --cached --check` passes.

No conflict marker, unmerged index entry, or integration semantic loss was
found. Every non-conflicting F-237 implementation and test file matches the
clean worker head. The shared `Document` retains both F-236 signature fields
and all three F-237 glossary fields (`crates/rdocx/src/document.rs:1401`,
`crates/rdocx/src/document.rs:1405`). Both state groups are initialized,
cloned for staging, and reconstructed on package reopen
(`crates/rdocx/src/document.rs:1792`,
`crates/rdocx/src/document.rs:1832`,
`crates/rdocx/src/document.rs:2032`,
`crates/rdocx/src/document.rs:2068`).

No staged-mutation interaction defect was found. Legacy form mutation flushes
the complete cloned state, serializes and reopens its candidate, then commits
only the validated result (`crates/rdocx/src/field.rs:108`,
`crates/rdocx/src/field.rs:144`, `crates/rdocx/src/field.rs:158`). Building-block
replacement follows the same clone, serialize, reopen, validate, and commit
sequence (`crates/rdocx/src/building_block.rs:370`,
`crates/rdocx/src/building_block.rs:378`,
`crates/rdocx/src/building_block.rs:392`). Glossary serialization participates
in the shared package flush before semantic signature comparison
(`crates/rdocx/src/document.rs:2365`, `crates/rdocx/src/embedded.rs:688`). The
shared commit path preserves prior invalidation and detects a changed signed
package (`crates/rdocx/src/document.rs:1845`).

No export or HLD coexistence defect was found. The crate root exports both
building-block and embedded APIs together with the legacy form values
(`crates/rdocx/src/lib.rs:24`, `crates/rdocx/src/lib.rs:29`,
`crates/rdocx/src/lib.rs:46`, `crates/rdocx/src/lib.rs:57`,
`crates/rdocx/src/lib.rs:62`). The combined HLD retains the F-236 executable
content and signature contract beside the F-237 form and glossary contract
(`docs/hld/02-scope-and-non-goals.md:157`,
`docs/hld/02-scope-and-non-goals.md:171`,
`docs/hld/04-opc-and-packaging.md:165`,
`docs/hld/04-opc-and-packaging.md:558`,
`docs/hld/10-bindings-spec.md:249`, `docs/hld/10-bindings-spec.md:925`).

`cargo check -p rdocx --all-targets` and `cargo fmt --all --check` pass. The
focused F-236 embedded regression command passes all 62 tests, and the focused
F-237 form and building-block integration command passes all 38 tests.
