# S67 sprint review, pass 1

**Reviewed**: `sprint/s67` at `39934df01f78942277db826d859b5a2f2b626b55`
against merge base `689849ff4a75bf0d4fe5afef7207a457cc256ba7`, 28 files,
6,430 insertions and 314 deletions, crates: `rdocx`
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

This is the scheduled F-234 dependency-prefix boundary, not the completed S67
milestone. The applicable backlog gate is: "The pinned document pairs produce
the same insertions, deletions, moves, and story placement as Word at the
declared boundary" at `docs/hld/14-development-backlog.md:2151`.

The gate holds. The source-built differential
`full_story_comparison_matches_pinned_word_records` at
`crates/rdocx/tests/regression_test.rs:10664` passes, and its mutation matrix at
`crates/rdocx/tests/regression_test.rs:10965` rejects record kind, order, story,
move pair, field owner, and formatting perturbations. The package-wide accept
and reject proof starts at `crates/rdocx/tests/regression_test.rs:11029`, and
the exact unowned-byte proof starts at
`crates/rdocx/tests/regression_test.rs:11150`.

F-233 is already present in the integrated prefix. Its regression gate at
`crates/rdocx/tests/regression_test.rs:12719` passes beside the F-234 gate, and
the full integrated regression binary reports 267 passed and 3 ignored. The
49-entry hash harness also passes unchanged, consistent with the recorded
F-234 evidence at `docs/sprints/AS_BUILT.md:11483`.

F-235 is intentionally not due at this boundary. Its row remains pending at
`docs/sprints/CURRENT_SPRINT.md:32`, and the sequencing contract makes it follow
F-234 at `docs/sprints/CURRENT_SPRINT.md:38`. Character and word granularity and
ignore-policy acceptance therefore remain future F-235 evidence rather than a
failure of this dependency prefix.

## Not found

- **Interaction**: rich mail merge publishes reopened staged documents at
  `crates/rdocx/src/field.rs:668`, while comparison clones, validates, accepts,
  rejects, and commits through one package boundary at
  `crates/rdocx/src/comparison.rs:181`. Their named gates pass in one integrated
  test binary, with no shared-state or field-owner conflict found.
- **Duplication**: the full-story index is centralized at
  `crates/rdocx/src/comparison.rs:278` and is reused by package-wide revision
  resolution at `crates/rdocx/src/revision.rs:175`. No second F-234 traversal or
  source-map model was added.
- **Layering**: only `rdocx` source changed. No manifest changed, so no
  `oxml-*` to facade dependency edge was added.
- **Harness**: neither the harness nor its manifest changed. Check mode reports
  all 49 entries equal, matching both feature plans' unchanged expectation.
- **Gate**: both currently applicable named gates pass. The remaining policy
  gate belongs to pending F-235 as recorded above.
- **Docs**: the two contracts coexist in the four declared HLD files. The
  complete rich mail-merge test contract at
  `docs/hld/12-testing-strategy.md:542` is followed by the complete full-story
  comparison contract at `docs/hld/12-testing-strategy.md:556`.
- **Deps**: no dependency or lockfile changed.
- **Public surface**: the only new exports are the F-233 rich merge values at
  `crates/rdocx/src/lib.rs:54` and the two methods at
  `crates/rdocx/src/field.rs:668` and `crates/rdocx/src/field.rs:712`. F-234
  preserves the existing comparison signature at
  `crates/rdocx/src/comparison.rs:183`.
