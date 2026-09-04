# F-235, all, pass 7

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 3,845 insertions and 450 deletions
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Verified

The pass 6 defect is remediated. Direct significant inline-control raw children
enter full-slice attributed alignment only when their exact bytes and preceding
significant attributed-unit counts match
(`crates/rdocx/src/comparison.rs:3574`,
`crates/rdocx/src/comparison.rs:3706`). Text values are no longer part of the
raw-boundary key. The source interleaver retains the original non-run gaps while
replacing only correlated run-owner spans
(`crates/rdocx/src/comparison.rs:3895`).

The focused regression combines a preceding text replacement with one-run and
two-run physical segmentation under Word and Character policies, fixes the
minimal revision contents, and proves the raw subtree appears exactly once in
tracked, accepted, and rejected views
(`crates/rdocx/tests/regression_test.rs:13676`). Separate cases move the raw
boundary or change its bytes and require an error or a visible revision
(`crates/rdocx/tests/regression_test.rs:13724`). These cases make both fields of
the boundary key observable.

The additive public enum, options value, re-export, and option-taking method
retain the legacy default delegation and native-only surface
(`crates/rdocx/src/comparison.rs:43`,
`crates/rdocx/src/comparison.rs:223`, `crates/rdocx/src/lib.rs:44`). Story and
comment suppression occurs before related-story discovery and main-story
revision seeding (`crates/rdocx/src/comparison.rs:248`,
`crates/rdocx/src/comparison.rs:269`, `crates/rdocx/src/comparison.rs:396`).
Formatting, whitespace, field, comment, and granularity policy share the one
attributed-run path (`crates/rdocx/src/comparison.rs:2111`,
`crates/rdocx/src/comparison.rs:2548`). TextBox masking uses namespace-aware,
cross-input collision selection and exact restoration cardinality
(`crates/rdocx/src/comparison.rs:973`,
`crates/rdocx/src/comparison.rs:1037`,
`crates/rdocx/src/comparison.rs:1449`). Acceptance and rejection remain staged
package-wide postconditions before the single live commit
(`crates/rdocx/src/comparison.rs:316`,
`crates/rdocx/src/comparison.rs:341`, `crates/rdocx/src/comparison.rs:356`).

The exact record gate fixes kind, content, story, owner, and order for all three
granularities, every ignore flag, and each of the seven independently selected
story categories (`crates/rdocx/tests/regression_test.rs:13897`). No additional
correctness, contract, panic, OOXML, test, public-surface, dependency, or
structure issue was established.

The complete regression binary passed with 285 tests passed and 3 ignored. The
three comparison unit tests passed. The focused inline-control regression,
`cargo clippy -p rdocx --all-targets -- -D warnings`, `cargo fmt --all --check`,
and `git diff --check` passed.
