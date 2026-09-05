# S68 sprint review, pass 2

**Reviewed**: `sprint/s68` at `f9d9ee8879b9` against merge base
`2997915028a8`, 66 files and 20,173 changed lines, comprising 19,918 additions
and 255 deletions, crates: `oxml-opc`, `rdocx-oxml`, `rdocx`
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold, and S68 is not the milestone-ending
sprint. The embedded-content clause has executable evidence in
`word_embedded_inventory_reports_exact_hashes_relationship_paths_and_signature_state`
and the payload-preservation clause in
`ordinary_document_edits_preserve_every_embedded_payload_byte`
(`crates/rdocx/tests/regression_test.rs:14786`,
`crates/rdocx/tests/regression_test.rs:15020`). The modern package-variant
clause remains assigned to pending F-238
(`docs/sprints/BACKLOG.md:444`). It must remain open until the S69 round-trip
gate supplies that evidence.

## Not found

- **Pass-1 closure**: B1 is closed. The embedded Pack URI predicate now rejects
  `.` segments while retaining safe parent traversal
  (`crates/rdocx/src/embedded.rs:1016`). The focused regression rejects the
  exact `embeddings/./object1.bin` spelling before inventory or mutation and
  proves failure atomicity, then accepts a safe `../embeddings/object1.bin`
  target (`crates/rdocx/tests/regression_test.rs:17230`,
  `crates/rdocx/tests/regression_test.rs:17268`). The focused test and all 62
  `word_embedded_` regressions pass.
- **Duplication**: pass-1 S1 is validly filed as F-X077. Its backlog contract
  names all three consumers, the shared lowest-layer boundary, exclusions, and
  a mutation-sensitive regression gate while citing the originating review
  (`docs/hld/14-development-backlog.md:3698`). The live backlog records it as
  pending for S69, and the sprint plan places it in that wave
  (`docs/sprints/BACKLOG.md:532`, `docs/sprints/SPRINT_PLAN.md:1278`). No other
  sprint-local duplicate implementation was found.
- **Interaction**: no additional interaction finding was found. The shared
  document retains both signature and glossary state through construction,
  staging, reopen, and commit (`crates/rdocx/src/document.rs:1401`,
  `crates/rdocx/src/document.rs:1792`, `crates/rdocx/src/document.rs:1832`,
  `crates/rdocx/src/document.rs:2032`). Form and building-block mutation still
  flows through whole-package signature invalidation and validated staged
  commit (`crates/rdocx/src/field.rs:108`,
  `crates/rdocx/src/building_block.rs:370`,
  `crates/rdocx/src/document.rs:1845`,
  `crates/rdocx/src/document.rs:2365`). All 38 focused form and building-block
  integration tests pass.
- **Layering and dependencies**: no finding was found. The only new ordinary
  dependency remains the existing workspace `sha2` crate in `rdocx`, with the
  concrete inventory hash consumer in the private embedded module
  (`crates/rdocx/Cargo.toml:44`, `crates/rdocx/src/embedded.rs:796`). No
  `oxml-*` crate gained a format-crate edge, and
  `no_shared_crate_depends_on_a_format_crate` passes.
- **Harness and gate**: no sprint-scope finding was found. Both AS_BUILT entries
  declare the harness unchanged and record integrated archive-size and
  dependency evidence (`docs/sprints/AS_BUILT.md:11641`,
  `docs/sprints/AS_BUILT.md:11646`, `docs/sprints/AS_BUILT.md:11691`,
  `docs/sprints/AS_BUILT.md:11695`). The current independent harness check
  confirms all 49 entries match.
- **Docs and ledgers**: no finding was found. The HLD retains both feature
  contracts, the backlog and current sprint record F-236 and F-237 done, and
  the feature tracker records both delivery rows
  (`docs/sprints/BACKLOG.md:442`, `docs/sprints/CURRENT_SPRINT.md:32`,
  `docs/sprints/SPRINT_TRACKER.md:388`). The F-X077 additions are limited to
  the detailed backlog, live backlog, and scheduled sprint plan required to
  track the review disposition.
- **Public surface**: no finding was found. The exported embedded, legacy-form,
  and building-block vocabulary matches the two approved plans and remains
  native Rust only (`crates/rdocx/src/lib.rs:46`,
  `crates/rdocx/src/lib.rs:57`, `crates/rdocx/src/lib.rs:62`).

`cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `python3
scripts/prose_check.py`, `python3 scripts/sync_agent_skills.py --check`,
`python3 scripts/sprint_workflow.py status`, and `git diff --check main...HEAD`
pass.
