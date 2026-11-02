# F-X282, Codex implementation, pass 3

**Reviewed**: Third and last independent implementation peer pass, after the separate final1 test-contract repair. Worker `work/f-x282-codex`, Base `aff3f2b8744ed86755db8520c9bf6abeed5f2038`. Reviewed the exact change from `final1-before-test-repair`, the strengthened styles tests, test-only shadow counter handling, plan carrier reconciliation, progress record and actual receipts. This reviewer did not author the repair.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this affected repair. Fresh Root final review remains required. This does not reset a review counter, rewrite a prior finding or establish final performance acceptance.

## Defects

None found.

## Smells

None found.

## Nitpicks

None recorded.

## Actual-open and unchanged-edit control

`crates/rdocx/src/document.rs:50190`

The counter reset now precedes actual `Document::from_bytes`. The test requires one parse and one serialization proof, rather than inferring opening work from an already-open document. It retains both the styles model Arc and the OPC styles-part Arc. After a new reset, actual `try_replace_text`, `add_empty_table_row` and `to_bytes` must each preserve those allocations and leave primary styles work at zero. It inspects saved styles bytes and actual replacement text and row count. These assertions reject per-edit parse/serialization, loss of shared custody and edits that return success without changing their targets.

## Changed-style and retained-output control

`crates/rdocx/src/document.rs:50230`

The fixture imports a real styles part with producer root namespace/attribute metadata, two self-contained opaque children and an inherited-prefix opaque child. The public default-font setter changes the actual model. The test proves COW by retaining the original model and its Before font. Explicit flush must perform zero parses and one primary serialization and install a different part allocation.

The test decodes the returned saved ZIP and reads its actual styles part. It requires the producer namespace and attribute, exact single occurrences of the two self-contained opaque fragments, and ordered docDefaults, opaque children and modeled style. It parses that saved part to assert the Changed default font and then asserts the same value through a reopened Document. A changed-model-only assertion cannot substitute for these saved-output checks.

The inherited opaque fragment is handled honestly. The parser first makes its namespace self-contained. The fixture asserts the materialized fragment differs from the original inherited spelling, carries the expected binding/value and passes strict validation. Saved output must retain that materialized fragment exactly once and in order. It does not claim that parse-time namespace materialization leaves the original inherited fragment byte-identical.

## Counter boundary and production applicability

`crates/rdocx/src/document.rs:7879`

The existing debug wrapper still calls its closure. The only additions save and restore `FX282_STYLES_WORK` under `cfg(test)`. Inspected the wrapper consumers in debug reopen and the saved-output comparison. The ZIP shadow oracle still executes, serializes and parses. Its work is excluded from the explicitly primary counter, not removed from execution. This is not a zero-total-work or debug performance claim. The wrapper is not used to conceal primary styles work during initial open or the explicit changed-style flush.

Independently reconstructed the previous complete document source in memory from the current/frozen diff packets and verified its recorded full hash. Compared the entire prefix before the native test module. Removing exactly the added test-only save/restore and its comment recovers the prior prefix byte-for-byte, SHA-256 `c282062b435a09a8052545f2f5cecc8847d368f7d1c039a778c1cf9ad87fb86f`. All previous non-document source file hashes remain unchanged. There is no release-production or Python-extension behavior change in this repair.

The checklist adjustment at `.claude/plans/F-X282-design.md:426` names the actual existing rdocx and OPC test carriers. The diff changes only that carrier row. It does not delete a behavioral assertion or introduce an rdocx-oxml change to the deliberately unchanged table iterator.

## Evidence authentication and outcomes

Current handoff SHA-256 is `6ffcea49f9e1540f9eb2300bdb73d9a59cab0037458774e97ffb5f2cf58b9ab5`. Current source manifest is `70c1d6c2cc3d329767262d35bf36270ab01d2a12c6d06c3db932ee2862638f55`, evidence index is `ca938852c5571553b0d19712ec94b4c3c670aee9581d5e4f747a7cbfcf809d46`, and applicability record is `8cd3c718a6f9e8934eb09f0a5d38c416997c3cdbff39ea6b78aeecdfacb45689`. All thirteen current manifest file hashes and all indexed log hashes match.

Read the actual fifteen named test outcomes in `final1-test-native3.log`, including the strengthened actual-open test and new changed-style retained-output test. All pass. The receipt identifies the actual worker command and source build, with child/wrapper exits both zero and no timeout. Affected warnings-denied all-target/all-feature clippy, format and prose receipts also pass.

The two earlier test batches remain failed observations. The first exposed the inherited-namespace fixture assumption and shadow-counter contamination. The second exposed an attribute-bearing docDefaults opening rather than the assumed bare opening. The final fixture accounts for both. Neither failed batch is relabelled as a production RED or a new old-source comparison. Existing authenticated baseline RED, OPC and Python observations retain their separate scope and applicability. No unaffected extension or Python suite was rerun.

## History and retained limits

The prepare inspection caught promised styles assertions that the earlier local implementation and final1 review had missed. The historical reports remain unchanged. This pass accepts the newly supplied assertions after a distinct author phase. It does not retrospectively claim those assertions ran during pass2 or final1. Pass1 D1 and D2 remain closed by their retained lexical and physical-selection controls.

Correctness, contract, panics, OOXML, tests and structure produced no new finding in the affected repair. No new source file, module, crate, trait, feature or test binary is introduced. F-X278 font policy and both approved public additions are unchanged. The failed writer-proof fixture remains synthetic. The foreign-extension namespace-placement difference remains explicit. The malformed styles replacement refusal still requires its dedicated labelled behavior commit during actual preparation.

All exact Linux release Python3.12 A282-1/2/3 workloads, sizes, cold sampling rules and thresholds remain pending. Final saved bytes/hash49, F-X268 79/252/407, full workspace/native/oracle, installed runtime/typing, WASM, wheels, archives and hosted CI gates also remain pending. Actual typed-publication composition with the eventual integrated C1 source remains debt, without an invented provider edge or speculative raw-owner refusal. None of these final gates is discharged by the primary work counters or this clean review.

No tests, builds, benchmarks, network operations, Git mutations, source edits or plan edits were performed. This review file is the only write. Stop here and return the review-directory writer lease to Root for fresh final review.
