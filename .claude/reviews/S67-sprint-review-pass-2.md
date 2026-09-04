# S67 sprint review, pass 2

**Reviewed**: `sprint/s67` at `b31c903e12726ee25644baa3c54f0878b2b032fa`
against merge base `689849ff4a75bf0d4fe5afef7207a457cc256ba7`, 36 files,
12,148 insertions and 1,840 deletions, crates: `rdocx`
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

M22 remains open for F-236 through F-239
(`docs/sprints/BACKLOG.md:442`). The complete S67 boundary is the conjunction
of its three due gates. F-233 requires nested source-built records to generate
the expected rich output without stale fields
(`docs/hld/14-development-backlog.md:2143`). F-234 requires pinned document
pairs to match Word insertions, deletions, moves, and story placement
(`docs/hld/14-development-backlog.md:2151`). F-235 requires every policy to
change only declared records deterministically
(`docs/hld/14-development-backlog.md:2159`).

All three gates hold together. The rich merge gate at
`crates/rdocx/tests/regression_test.rs:12834`, the pinned Word gate at
`crates/rdocx/tests/regression_test.rs:10779`, and the exact policy gate at
`crates/rdocx/tests/regression_test.rs:13897` pass in the same integrated
regression binary. The Word test fixes build `16.104.25121423`, locale `en-US`,
and the normalized record comparison
(`crates/rdocx/tests/regression_test.rs:10779`). Its calibrated mutation test
covers kind, order, story, move pair, field owner, and formatting
(`crates/rdocx/tests/regression_test.rs:11080`). The AS_BUILT record identifies
the source-built 24-record capture and exact Word build
(`docs/sprints/AS_BUILT.md:11474`).

The complete integrated regression binary reports 285 passed and 3 ignored.
The F-233 lexical-scope unit test and all three F-235 comparison unit tests also
pass. The run state records a successful full verification with an unchanged
harness at the exact reviewed head
(`.claude/scratch/S67-run.json:82`). The independent hash check reports all 49
entries equal, consistent with all three AS_BUILT declarations
(`docs/sprints/AS_BUILT.md:11483`, `docs/sprints/AS_BUILT.md:11536`,
`docs/sprints/AS_BUILT.md:11595`). Clippy, formatting, prose, generated-skill,
and diff checks pass.

## Not found

- **Interaction**: F-233 stages and reopens each rich output before publication
  at `crates/rdocx/src/field.rs:668`, while F-234 and F-235 share comparison's
  single staged accept and reject boundary at
  `crates/rdocx/src/comparison.rs:235`. Both public families coexist in the
  facade exports at `crates/rdocx/src/lib.rs:45` and
  `crates/rdocx/src/lib.rs:56`. Their gates pass together with no field-owner,
  package-state, or public-name conflict.
- **Duplication**: comparison owns one relationship-resolved story traversal at
  `crates/rdocx/src/comparison.rs:396`, and revision resolution reuses its part
  inventory at `crates/rdocx/src/revision.rs:175`. F-235 extends that path
  rather than adding a second traversal.
- **Layering**: only `rdocx` source changed. No manifest or lockfile changed, so
  no new `oxml-*` dependency edge exists.
- **Harness**: the harness and manifest are unchanged. The exact reviewed-head
  verification and the independent 49-entry check agree with the three feature
  records cited above.
- **Gate**: the three due story gates are exercised by source-built fixtures,
  fixed normalized records, package-wide accept and reject checks, atomic
  failure cases, and mutation checks. They are not supported by prose alone.
- **Docs**: the rich merge contract at
  `docs/hld/12-testing-strategy.md:542` is retained before the complete
  comparison and policy contract at `docs/hld/12-testing-strategy.md:556`.
  Architecture, packaging, native facade, and test HLD text agree on staged
  publication, native-only additions, and byte preservation
  (`docs/hld/03-architecture.md:590`,
  `docs/hld/04-opc-and-packaging.md:670`,
  `docs/hld/10-bindings-spec.md:531`).
- **Deps**: no dependency declaration or lockfile changed.
- **Public surface**: the rich merge values and methods are the exact additive
  native surface requested by F-233 at `crates/rdocx/src/field.rs:56` and
  `crates/rdocx/src/field.rs:668`. The comparison enums, options, and method are
  the exact additive native surface requested by F-235 at
  `crates/rdocx/src/comparison.rs:43` and
  `crates/rdocx/src/comparison.rs:223`. Python, WASM, and CLI contain no new
  corresponding surface.
- **Ledgers and scope**: all three S67 rows are done with cleared owners in both
  live trackers (`docs/sprints/CURRENT_SPRINT.md:30`,
  `docs/sprints/BACKLOG.md:439`). The delivery rows and AS_BUILT entries match
  the approved sizes, dates, tests, HLD lists, and unchanged-harness evidence
  (`docs/sprints/SPRINT_TRACKER.md:384`,
  `docs/sprints/AS_BUILT.md:11442`). The only unrelated ledger-line changes are
  regenerated AUTOGEN summary counts, which now agree with the existing status
  rows (`docs/sprints/BACKLOG.md:20`). No unrelated source, HLD, plan, or review
  file appears in the integrated sprint delta.
