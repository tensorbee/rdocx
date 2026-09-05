# S68 sprint review, pass 1

**Reviewed**: `sprint/s68` at `b22ff7e1cb34` against merge base
`2997915028a8`, 63 files and 20,008 changed lines, comprising 19,756 additions
and 252 deletions, crates: `oxml-opc`, `rdocx-oxml`, `rdocx`
**Verdict**: 1 blocking, 1 should-fix, 0 nice-to-have

## Blocking

### B1, embedded relationships accept non-normalized dot segments
`crates/rdocx/src/embedded.rs:1016`
`crates/rdocx/src/building_block.rs:196`
`crates/oxml-opc/src/package.rs:344`
`docs/hld/04-opc-and-packaging.md:564`

The F-236 Pack URI predicate explicitly treats a `.` path segment as valid.
`OpcPackage::resolve_rel_target` then collapses that segment, so an embedded
relationship such as `embeddings/./object1.bin` resolves to an existing payload
and becomes actionable. The integrated F-237 predicate rejects the equivalent
glossary or story alias, and the shared HLD requires normalized internal Pack
URI references. One `Document` therefore applies two incompatible graph-safety
contracts, and the embedded path does not fail closed before inventory or
mutation. The embedded predicate must reject `.` segments while continuing to
allow necessary `..` traversal that remains inside the package root, with a
regression covering inventory and mutation atomicity for the `/./` spelling.

## Should-fix

### S1, XML lexical validation is triplicated across the sprint
`crates/rdocx/src/embedded.rs:1122`
`crates/rdocx-oxml/src/glossary.rs:574`
`crates/rdocx/src/field.rs:8057`

The sprint adds three independent copies of the same XML 1.0 literal-character,
event, namespace-declaration, QName, NCName, processing-instruction, expanded
attribute, and reference validation stack. The functions differ mainly in
error construction and owner labels, while security corrections must now be
applied and reviewed three times. The duplicate relationship-target predicates
have already diverged in B1, which demonstrates the same maintenance risk.
Move the format-neutral lexical checks to one lowest-common-layer utility and
keep only owner-specific root, schema-position, and error adaptation locally,
or file a bounded follow-up F-ID from this review before closing the sprint.

## Nice-to-have

None.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold, and S68 is not the milestone-ending
sprint. The embedded-content clause has direct evidence in
`word_embedded_inventory_reports_exact_hashes_relationship_paths_and_signature_state`
and the payload-preservation clause in
`ordinary_document_edits_preserve_every_embedded_payload_byte`
(`crates/rdocx/tests/regression_test.rs:14786`,
`crates/rdocx/tests/regression_test.rs:15020`). The modern package-variant
clause remains assigned to pending F-238
(`docs/sprints/BACKLOG.md:444`). It must not be marked met until that S69 gate
supplies its round-trip evidence.

## Not found

No additional interaction finding was found in the reconciled staged state.
The shared document retains, initializes, clones, and reloads both F-236
signature state and F-237 glossary state (`crates/rdocx/src/document.rs:1401`,
`crates/rdocx/src/document.rs:1792`, `crates/rdocx/src/document.rs:1832`,
`crates/rdocx/src/document.rs:2032`). Form and building-block mutations use the
same serialize, reopen, validate, and commit path, while dirty glossary bytes
participate in whole-package signature invalidation
(`crates/rdocx/src/field.rs:108`, `crates/rdocx/src/building_block.rs:370`,
`crates/rdocx/src/document.rs:1845`, `crates/rdocx/src/document.rs:2365`).

No layering, dependency, or public-surface finding was found. The only new
ordinary dependency is the existing workspace `sha2` crate in `rdocx`, with
the concrete SHA-256 inventory consumer in the private embedded module
(`crates/rdocx/Cargo.toml:44`, `crates/rdocx/src/embedded.rs:796`). No
`oxml-*` crate gained a format-crate dependency, and
`no_shared_crate_depends_on_a_format_crate` passes. The exported embedded,
legacy-form, and building-block vocabulary matches the two approved plans and
remains native Rust only (`crates/rdocx/src/lib.rs:46`,
`crates/rdocx/src/lib.rs:57`, `crates/rdocx/src/lib.rs:62`).

No harness, ledger, HLD-scope, archive-size, conflict-marker, or diff-hygiene
finding was found. Both AS_BUILT entries declare the hash harness unchanged and
record the integrated archive-size and dependency riders
(`docs/sprints/AS_BUILT.md:11641`, `docs/sprints/AS_BUILT.md:11646`,
`docs/sprints/AS_BUILT.md:11691`, `docs/sprints/AS_BUILT.md:11695`). The current
harness check confirms all 49 entries match. The backlog, current sprint, and
feature tracker consistently record both stories done
(`docs/sprints/BACKLOG.md:442`, `docs/sprints/CURRENT_SPRINT.md:32`,
`docs/sprints/SPRINT_TRACKER.md:388`).

The focused embedded regression command passes all 62 tests, the focused form
and building-block integration command passes all 38 tests, and the glossary
unit command passes all 41 tests. `cargo check -p rdocx --all-targets`,
`cargo fmt --all --check`, `python3 scripts/prose_check.py`,
`python3 scripts/sync_agent_skills.py --check`, `python3
scripts/sprint_workflow.py status`, and `git diff --check main...HEAD` pass.
