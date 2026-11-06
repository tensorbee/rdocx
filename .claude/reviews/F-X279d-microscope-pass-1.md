# F-X279d, microscope, pass 1

**Reviewed**: Working diff against `ff246c8b8b76ef573e8438b582259ba7045124ba`, 11 leased paths, 779 added and 17 removed lines. The diff includes the separately reviewed C2 owned-comment cleanup neighbor. All 11 source hashes match `REVIEW_SOURCE_FREEZE.json`.

**Contract**: Approved F-X279d design and Root split review. This pass covers crossing reply authoring and optional initials through Rust, Python and CLI, plus the stated lifecycle, imported-source and atomicity controls. Comparison and resolution composition remain F-X279b responsibilities. C1 provider closure is excluded.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. Remediation and another affected review are required.

## Defects

### D1, the reference-run boundary can truncate the parent text in a reply

`crates/rdocx/src/comments.rs:1927`
`crates/rdocx/src/comments.rs:1929`
`crates/rdocx/src/comments.rs:1933`

`anchor_reply_to_parent` finds any accepted run containing the parent reference and uses the boundary after that entire run as the reply end. The earlier unique-marker and same-story check at `crates/rdocx/src/comments.rs:1917` does not establish that this boundary follows the parent end. The helper then passes that endpoint to `anchor_story_range` without comparing it with `parent.end`.

An imported parent with the physical sequence start, Text `alpha `, reference run, Text `beta`, end has one start, end and reference in the same story. Its existing anchor covers `alpha beta`. Replying succeeds, but the new reply covers only `alpha `. This contradicts the required retained parent-text and crossing topology. The existing range guards validate the chosen endpoints, so they do not reject this sequence before publication.

The implementer reproduced this exact case without changing the frozen source. `/private/tmp/rdocx-s96-f279d-evidence/reference-order-probe.json` records `parent_anchor_before="alpha beta"`, `reply_result="success"` and `reply_anchor="alpha "`. Its extension hash is `7e90bb5bdc76a948ae9eba99d0a105828e2b2414d704f618b1364f84f73b9447`, independently verified against the frozen artifact in this pass.

The same endpoint selection also treats an entire mixed reference/Text run as the reference boundary. A suffix after the reference in that run can therefore be absorbed into the reply. This second representation is source-derived, not an executed result in this pass. Both representations require the same exact endpoint proof, rather than a check that merely finds the reference somewhere in a run.

Repair must preserve the parent text and required reference topology, or atomically refuse an input that the existing projection policy cannot represent. Add the actual early-reference discriminator and a mixed-run reference/suffix control alongside the existing dedicated-reference-run positives. Preserve supported ordinary, multi-paragraph and related-story replies.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Source and caller assessment

The legacy Rust reply shims delegate with absent initials. The additive Rust method stages the mutation, rejects invalid characters and dates, and commits only after related-story flush succeeds. The Python binding and installed stub agree on keyword-only `date=None, initials=None`. CLI `--initials` forwards omission, empty and nonempty values to the same Rust method and retains the existing atomic output publisher.

The reviewed controls exercise repeated replies, replying to a reply, selected-parent start adjacency, separate references, remove and resolve behavior, reopen, multi-paragraph and header owners, table cells and block controls, imported aliases, preserved opaque payload, missing or ambiguous markers and split-story refusal. The Python controls distinguish successful revision advancement and stale held views from refusal with unchanged bytes and live held views. The public Rust and CLI controls check saved/reopened initials and anchor text. These tests use conventional dedicated reference runs and do not discharge D1.

The separately reviewed C2 cleanup is not treated as acceptance of the D feature. Its namespace-qualified empty-owned-part policy remains bounded to the known library-owned MC metadata. Imported point, orphan and rangeless replies retain their existing saved representation. No comparison/import implementation or new C1 dependency is introduced by this authoring diff.

## Evidence applicability

Recorded evidence inspected includes the baseline receipt and actual final logs. The baseline records three failures through old public reply APIs, with first-failure scope explicitly retained. It does not prove later loop branches failed independently.

The final Native comments cohort records 30 passes and 1 ignored test. The public regression and CLI reply controls each record 1 pass. The fresh Python artifact records 24 passes, including overlapping C2 controls, and the typed caller reports success. The final extension hash and four Python build-source pins agree with the frozen sources. Strict all-targets, all-features Clippy and the recorded metadata checks succeeded. No reviewer build or test was run.

The historical self-hash test and constant remain unchanged and are not newly ignored. That test was excluded from the local cohort and still requires fresh final native execution. This pass makes no Word interoperability, hash-harness, final integrated gate, provider or comparison-composition completion claim.

## Not found

No additional correctness, panic, OOXML preservation, namespace, schema-order, public-binding or structural-rule finding was found in this bounded diff. The missing reference-order and mixed-run controls belong to D1 rather than a duplicate test finding. No independent smell or nitpick was found.
