# F-X282, Codex implementation, pass 1

**Reviewed**: Independent non-author review of the working diff on `work/f-x282-codex`, Base and HEAD `aff3f2b8744ed86755db8520c9bf6abeed5f2038`. Twelve tracked files, 950 insertions and 96 deletions. Read the complete dirty diff, approved plan, existing preparation/design finals, progress record and applicable HLD03/04/06/10/12 contracts.

**Verdict**: 2 defects, 0 smells, 0 nitpicks. Both defects concern required local regression evidence. No separate production correctness defect was established. This is not performance acceptance or permission to complete the story.

## Defects

### D1, the strict-unqualified acceptance discriminator is missing

`crates/rdocx/src/document.rs:49882`

The test named `issue437_unqualified_cell_candidate_matches_checked_legacy_policy` supplies only a missing relationship and an XML control character. It compares errors with the checked route, but supplies none of the four retained lexical inputs for which checked publication succeeds while strict validation fails. The separate foreign extension control at `crates/rdocx/src/document.rs:49842` is valuable, but its well-formed extension is strict-qualified and publishes directly.

Consequently these fixtures do not establish the acceptance half of C282-3. In particular, they do not exercise a successful strict-error fallback. A blanket refusal on strict failure would not be rejected by the missing acceptance controls. The approved A282-2R and C24 expressly require duplicate expanded attributes, a literal less-than attribute, an invalid qualified name and an unbound prefix to retain the actual checked publisher's acceptance. See `.claude/plans/F-X282-design.md:256`, `.claude/plans/F-X282-design.md:289` and the local regression row at `.claude/plans/F-X282-design.md:338`.

Add the promised retained-source controls on the implementation Base. For each, prove that the actual checked edit succeeds, that its emitted candidate fails the strict probe, and that the new bridge succeeds through checked publication with one callback and retained source after save. Keep the strict-qualified foreign positive control and its explicit namespace-placement byte difference. The historical Root existing-library probes explain the requirement but do not establish this current source's result. This finding does not assert that the present fallback implementation already rejects these inputs.

### D2, the physical-selection test omits the control-owned nested-table case

`crates/rdocx/src/document.rs:49758`

The physical-selection regression first tests a control containing a paragraph directly, then tests a table stored directly as `CellContent::Table`. It never places that nested table inside a cell content control. The existing public controlled-cell regression at `crates/rdocx-py/tests/test_core.py:12552` also contains a directly controlled paragraph, rather than a control-owned table.

These cases exercise different branches of the reader. A direct nested table is excluded from the outer cell paragraph sequence. A control-owned nested table can contribute paragraphs through the control reader and must fail the direct-edit gate. A wrong eligibility implementation that counted only direct control paragraphs could pass the present tests and address the wrong physical paragraph in this omitted state. A282-2R, C14 and the gate proof explicitly require this discriminator, at `.claude/plans/F-X282-design.md:177`, `.claude/plans/F-X282-design.md:256` and `.claude/plans/F-X282-design.md:279`.

Add the promised control-owned nested-table fixture with distinct outer and nested paragraph values. Exercise the actual bridge, assert checked-route selection and callback count, and prove the selected physical owner and untouched neighbors through live reads and save/reopen. Keep the existing direct nested-table control. No actual wrong-owner mutation was demonstrated in the current implementation.

## Smells

None found.

## Nitpicks

None recorded.

## Source and evidence authentication

The worker is in-progress in both shared status tables. The final diff exactly matches `FINAL_DIFF.patch`. All twelve current file hashes match `FINAL_SOURCE_MANIFEST.json`. The supplied handoff, source manifest and evidence-index hashes match the request: `a6fbfccbc5548f8bcccfe0e87a8258ae8fc9a89a7a81d2b3a90ac696ba2ea560`, `e8d033ab81f8ca914dd7baf90f46ff3e4e0f4a05d8ff2dbca6e08e13d7ce6b9c` and `838422ac6a4f2b4933fe8d96fa31f1b0779c636bfe19eee40214097f8d7224a7`.

All indexed log hashes match their receipts. Read the actual styles RED assertions, final twelve native test outcomes, OPC outcome, public Python RED traceback and GREEN results, focused 42-case pytest result, build and check outcomes. Child and wrapper exits agree, with no indexed timeout. The semantic REDs are distinct from earlier setup/fixture failures. Aggregate counts do not discharge D1 or D2.

The fresh extension and its build binary both hash to `bcd1116683d73c37241ed0d178e2ca55127db73d1901fd601b22797bf9efa74b`. Python receipts identify that runtime and interpreted test hash `823b57f953b6435c1fafc2a661336319ffa86d66be73f9e0a97f4335d317392e`. The compiled-input manifest differs from current source only in the declared native test carrier and interpreted Python test file. This is debug Python3.14 semantic evidence, not a release performance observation. The old extension and its build binary match `9bd8cc2ad5037562441719870e391d7e98a61ee97ea6cb5ecfce5737a30169bf`. The seventy recorded baseline applicability files match worker HEAD.

## Not found and retained limits

Correctness, panics, OOXML and structure produced no separate source finding in this pass. Traced byte-allocation identity before model identity, open proof, equal-value adoption, changed publication and cache refresh, staged clone/reopen, default-font and style/numbering writers, merge and field import publishers, and deep-copy HTML/layout consumers. COW conversions preserve the existing F-X278 font policy. The shared accessor preserves exact ownership and case-equivalent lookup. The generic has actual unit and run-index/path consumers.

Traced direct cell eligibility, candidate serialization, section refusal, strict qualification, relationship extraction, marker counts, retained-owner fallback, original-value reacquisition and invalidation after assignment. Checked fallback enters `replace_story_content` with actual source facts and the already-edited fragment. It does not rerun the callback or substitute the raw-owner replacement API. Failed closure and NUL admission tests preserve the observed live layout. No speculative F-X272 provider dependency or new raw-field refusal is warranted. Its actual future checked-candidate policy still needs integrated reconciliation.

The foreign-extension fixture honestly records equal models and retained opaque bytes with a paragraph namespace-placement difference. It does not prove exact saved-byte equivalence. The synthetic failed styles proof tests the guard, not discovery of a real parseable-open writer-error input. The malformed styles replacement policy is an intentional observed behavior delta, isolated for a later dedicated labelled commit. No commit has been made here.

All three original Linux release Python3.12 A282 ratio workloads and sampling boundaries remain pending. So do final F-X268 79/252/407, exact saved bytes and hash49, full native/workspace/oracle, installed runtime/typing, WASM, wheels, archives and hosted CI qualification. Local debug counters do not replace them. These final-union debts are not additional defects invented by this review.

No source changes, tests, builds, benchmarks, network access or Git mutations were performed. This review file is the only write. The review ends here and returns the sole-writer lease to Root for a distinct remediation phase.
