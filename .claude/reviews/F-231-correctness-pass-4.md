# F-231, correctness, pass 4

**Reviewed**: complete revised working diff against `be9a49b`, 8 files, 2,145
additions and 117 deletions, plus the pass-1, pass-2, and pass-3 review records
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, the canonical packaging command cannot verify the reviewed working diff
`.claude/plans/F-231-design.md:104`

The revised risk rider requires the exact workspace dry run from `/verify` step
10, whose command starts at `.claude/commands/verify.md:70` without
`--allow-dirty`. Running that exact command packages the unchanged crates, then
exits 101 when it reaches `rdocx-oxml` because
`crates/rdocx-oxml/src/text.rs` contains the feature's uncommitted changes.
This is not a post-commit gate. The lifecycle runs `/verify` before
`/complete-feature` performs the commit at `.claude/WORKFLOW.md:16` and
`.claude/WORKFLOW.md:17`. The canonical command therefore cannot pass for this
working feature diff, or for any other feature that changes a publishable
crate. The same command with `--allow-dirty` verified all 22 local packages and
the archive-size assertion passed, so the remaining blocker is specifically
the missing dirty-worktree allowance rather than the patched dependency graph.

## Smells

None.

## Nitpicks

None.

## Not found

Pass-3 D1 is closed. `TocField` exposes the sequence identifier at
`crates/rdocx/src/field.rs:76`, and the evaluator validates and retains `\s` at
`crates/rdocx/src/field.rs:2359`. Pass-3 D2 is closed. Formula output now
normalizes non-integral finite values to 15 significant decimal digits at
`crates/rdocx/src/field.rs:3660`, with unit and pinned differential coverage
for `0.1 + 0.2` at `crates/rdocx/src/field.rs:4519` and
`crates/rdocx/tests/regression_test.rs:1071`. Pass-3 D3 remains open only as D1
above. All pass-1 and pass-2 findings remain closed.

No additional correctness, public-contract, panic, arithmetic-overflow, OOXML
child-order, namespace-prefix, significant-whitespace, unmodelled-XML,
dependency-family, oracle, test-coverage, trait, generic, wrapper-only type,
feature-flag, module, file-creation, or structural issue was found. The complete
`rdocx-oxml` suite passed with 367 unit tests and one doctest. All 20 focused
`rdocx` field unit tests, both extended-field regression tests, and
`cargo check -p rdocx --all-targets` passed. `git diff --check`, the tracked
prose check, and the package archive-size assertion also passed.
