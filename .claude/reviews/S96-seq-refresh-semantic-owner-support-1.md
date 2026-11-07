# S96 bounded SEQ revision refresh source review

Verdict: 0 defects, 0 smells, 0 nitpicks in the assigned revision-refresh delta.
No source changes or Cargo commands were performed. This review excludes the
reviewer's separate current-bookmark provider and does not complete F272 or
accept the pending Native SEQ lifecycle regression.

## Reviewed boundary

The author patch is `/private/tmp/rdocx-s96-prepared-integration-readiness/msrv70-whole-story-fixture-repair-1/SEQ_REFRESH_SEMANTIC_OWNER.patch`,
SHA-256 `aea4ee46e3370e444cbe90aec8c252db59dc45a0b23a15450ec90b222ee35d37`. Its production change
is the twelve-line semantic handback at
`crates/rdocx-oxml/src/revision.rs:1240`. The added existing-module regression
starts at `crates/rdocx-oxml/src/revision.rs:3641`. The current whole revision
source SHA-256 is `1e8d4888a6ae9ca1981f9a41b408d799de9b7f8e90a6709bd33fb7fcba632c01`.

The F272 contract requires actual revision refresh serialization, assembly and
reparse to carry emitted facts, without origin recovery by text equality
(`.claude/plans/F-X272-design.md:5097`). Its scoped refresh must retain the mapped
writer result (`.claude/plans/F-X272-design.md:6524`). The reviewed repair restores
the already materialized semantic owner before the existing identity proof.

## Source and operation preservation

The scoped emitter still receives the caller's `WordOperationWork` through
`refresh_accepted_content_in` at `crates/rdocx-oxml/src/revision.rs:1120` and
`crates/rdocx-oxml/src/revision.rs:1153`. The existing emitted paragraph map and
original wrapper slices form the actual updated-domain map at
`crates/rdocx-oxml/src/revision.rs:1212`. The repair adds no operation constructor,
input registration, baseline, quota or origin lookup. The existing
`WordFieldInput::produced` sets `initial_delivery` to false at
`crates/rdocx-oxml/src/document.rs:2902`.

The new branch runs only when the previous edited owner was not deferred and
the actual refresh reader again deferred simple fields
(`crates/rdocx-oxml/src/revision.rs:1240`). It passes the same actual updated
revision bytes, namespace scope and produced input into the existing checked
semantic parser. The parser admits the qualified paragraph pieces before
payload growth, slices the actual body map and adds generated shell coordinates
at `crates/rdocx-oxml/src/revision.rs:1819`. It does not re-enroll those bytes as
original input. The refreshed read context remains attached at
`crates/rdocx-oxml/src/revision.rs:1237`.

## Guards, custody and evidence limits

The final accepted-path, content-slot and Field form/instruction/display proofs
remain in `crates/rdocx-oxml/src/text.rs:11927` and
`crates/rdocx-oxml/src/text.rs:11984`. Scoped calls still pass the caller's actual
registry at `crates/rdocx-oxml/src/revision.rs:1255`. Successful source-fact alias
synchronization remains after identity retention at
`crates/rdocx-oxml/src/revision.rs:1269`. The original wrapper bytes are retained
in the newly parsed revision. Installation into self remains last at
`crates/rdocx-oxml/src/revision.rs:1275`, so a failed semantic parse or identity
proof cannot replace the prior revision model.

The new test edits an actual deferred simple SEQ field, refreshes it, verifies
its typed form and published cache, then reparses the raw result and requests
its semantic view for both insertion and moveTo
(`crates/rdocx-oxml/src/revision.rs:3641`). It rejects the reported raw-only
refresh mismatch without changing the initial deferred-read expectation. Root
reports this lower test passed. This review performed no runtime execution.
The actual Native accepted-revision SEQ case and final affected batch remain
pending and must retain their real outcomes. No native, full-suite, performance
or release claim follows from this bounded source verdict.
