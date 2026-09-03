# S66 sprint review, pass 1

**Reviewed**: `sprint/s66` against `5b93cadaa85a`, 18 files, 3,036 lines,
crates: `rdocx`, `rdocx-oxml`
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, F-232 requires a package check that cannot verify the integrated graph
`.claude/plans/F-232-design.md:103`

The approved pending plan requires `cargo publish --dry-run -p rdocx`. Running
that command at the clean reviewed HEAD exits 101 with 17 compile errors while
verifying the archive. Cargo resolves the published `rdocx-oxml` and
`rdocx-layout` versions, which do not contain APIs used by the current `rdocx`
source. The canonical gate avoids that false graph by packaging the workspace
with every internal crate patched to its reviewed local source at
`.claude/commands/verify.md:66`. F-232 depends on the newly completed F-231
boundary, so leaving the single-crate command in its mandatory rider makes the
next worker unable to satisfy the approved plan and fails to verify the source
that will actually be integrated. Before F-232 starts, its rider must use the
same dirty-worker and clean-integrated workspace gates already established for
F-231.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

The M22 end gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`). It does not yet hold, as expected
at this dependency-prefix boundary. F-232 and the later M22 stories remain
pending at `docs/sprints/BACKLOG.md:438`.

The completed F-231 slice does hold. The pinned Word differential
`extended_field_families_match_the_pinned_word_result` at
`crates/rdocx/tests/regression_test.rs:1046` passed, as did the field
scaffolding preservation regression and the shared-parser round-trip test.
The run state records full verification at clean HEAD `de4418a243b3`, the
independent clean patched-workspace publish dry run verified all 22 packages,
and its largest archive was 4,603,466 bytes. The hash harness independently
matched all 49 entries, consistent with the unchanged claim in
`docs/sprints/AS_BUILT.md:11394`.

## Not found

No other interaction issue was found. F-232 has no implementation in this
checkpoint, so no missing TOC rebuild behavior is attributed to the completed
F-231 slice. No duplication, layering, undeclared harness delta, HLD drift,
new dependency, or unrequested public surface was found. No `Cargo.toml` file
changed, and the only touched crates remain `rdocx` and `rdocx-oxml`. The
tracker, backlog, completion log, completed F-231 plan, and sprint run state
agree that F-231 is done while F-232 remains pending. `git diff --check`, the
prose check, and the generated-skill drift check passed.
