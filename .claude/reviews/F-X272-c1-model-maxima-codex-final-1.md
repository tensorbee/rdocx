# F-X272, bounded C1 model maxima, Codex final 1

**Reviewed**: canonical HEAD `4a220b6769fc2a5a2c90fc600f580478ea7e62f8` with preserved working carry. The frozen model-maxima increment has four existing paths, 722 added and 3 removed lines. Production changes are confined to `crates/rdocx-oxml/src/document.rs`. The text.rs changes add tests. Plan and progress changes describe this increment.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this bounded increment. This is an independent targeted audit, not a new full-source peer pass, cap reset, provider acceptance or full C1 acceptance. F-X283 authorship remains stopped and is outside this review.

## Defects

None found.

## Smells

None found.

## Nitpicks

None found.

## Correctness and contract

The retained reference differs from the frozen former guard only by its name and separate test instrumentation. The resource edge enumeration preserves its append order. Body, paragraph, run, control, table, row, cell, Field, revision, drawing and anchor branches were compared individually. Paragraph ruby text, stored control revisions and original runs, revision paragraph/run/nested storage, and both authoritative and legacy textbox models remain represented. The Field helper retains ordered paragraphs, all three cached run alternatives, cached fields, current and original instruction arguments and switches, authored original runs and physical span runs. Raw and scalar payloads do not acquire new modeled depth edges. See `crates/rdocx-oxml/src/document.rs:218`, `crates/rdocx-oxml/src/document.rs:441`, `crates/rdocx-oxml/src/text.rs:2577` and `crates/rdocx-oxml/src/content_control.rs:2380`.

Memo identity includes concrete kind and current borrowed allocation address. Immutable owners remain alive throughout this call, so separate current Fields sharing a source Arc do not collapse. Only entered nodes become active. Queued duplicate siblings cannot appear as active cycles. A completed summary combines each coordinate independently by maximum and then adds the local charge. Empty controls, tables and typed text bodies retain that charge. Field weight still requires an ordered result and Parsed or New Simple source form. Complex ordered results retain their old zero local Field charge. See `crates/rdocx-oxml/src/document.rs:183`, `crates/rdocx-oxml/src/document.rs:198`, `crates/rdocx-oxml/src/document.rs:351` and `crates/rdocx-oxml/src/text.rs:2311`.

Local entry refusal precedes child enumeration. Fitting completed summaries skip safely. Non-fitting summaries check the actual entry and descend in the same append/LIFO order as the former guard, rather than choosing weighted or table error by global coordinate priority. Postorder completion requires completed child summaries. The checked additions preserve inclusive 256 weighted and 32 table limits. No changed public model, allowance, signature, cutoff, XML serialization or callback path was found. See `crates/rdocx-oxml/src/document.rs:340`, `crates/rdocx-oxml/src/document.rs:375` and `crates/rdocx-oxml/src/document.rs:410`.

## Tests and evidence authentication

The three composed fixtures were inspected, not merely their pass count. They distinguish independent witnesses from sibling sums, 256/257 and 32/33 boundaries, fitting and stricter/incomparable shared paths, and exact old error text when one or both coordinates fail. They cover divergent current clones sharing source, all eight Field storage alternatives, retired original control runs, the actual Simple-versus-Complex predicate, authoritative and legacy textbox branches, and an empty typed body. The authored 2048-control tail asserts refusal after 256 successful local entries and dismantles the remaining tail iteratively. See `crates/rdocx-oxml/src/document.rs:9320`, `crates/rdocx-oxml/src/text.rs:32195` and `crates/rdocx-oxml/src/text.rs:32354`.

Read-only authentication confirmed all four canonical current hashes equal their frozen after snapshots, all before/after snapshot hashes match SOURCE_MANIFEST.json, and the exact unified patch reconstructs from those snapshots. The patch SHA-256 is `4631ddbfe7df8cd274d69c3371f5366122ea189a9c5099f53086427d5108a89a`. SOURCE_MANIFEST.json is `623f3c632b0fb73bb233b93dfcae17b393727e24d6633b46272589b963776caf`. AUTHOR_HANDOFF.md is `4cded41a776f126a171c8b93deb54e3385c4b62e9debc9b4b241c955326922a0`.

All 17 receipt log hashes match. The wrapper records actual child and wrapper exits separately. The final maxima log records 3 passes and actual comparable loop observations of 363 new popped entries versus 699 reference popped entries. Its 362 computed summaries remain a separate metric. Those counters have actual loop-site consumers at `crates/rdocx-oxml/src/document.rs:360` and `crates/rdocx-oxml/src/document.rs:453`. No release timing follows from them.

The recorded other batches show authored-model 2, table32 lifecycle 1, affine 2, original-capture 49 and field-filter 53 passes. Counts overlap and are not added. Those broader runs preceded the final test-only additions and loop diagnostics. Production behavior is unchanged by those later additions. The final maxima rerun and warnings-denied all-target/all-feature Clippy cover the final test source. Formatting exits 0. Prose exits 1 with the same nine pre-existing progress-file violations in the frozen before and after records. The initial fixture type error is recorded as a failed author check, not semantic RED.

Evidence is under `/private/tmp/rdocx-s96-root-c1-implementation/model-maxima-stage`. The packet authenticates source snapshots, commands, log bytes, outcomes and source-level applicability. It does not retain a hash of the executed test binary or a complete build-input manifest. Logs identify `target/debug/deps/rdocx_oxml-bd79d033e67a4a32` by pathname only. Exact executable identity is therefore not independently authenticated by this packet. This review neither upgrades that limitation into artifact proof nor reruns tests or builds.

## Not found

No correctness, bounded-contract, panic/overflow, OOXML-preservation, test-discriminator or structure finding was found in the reviewed increment. No new file, module, crate, trait, generic parameter, feature, public model or public API was introduced by it. Full graph equivalence here concerns the former stored-resource guard, not a claim that every parser or aggregate checked operation is now bounded.

## Retained debt and boundary

The operative borrowed-model clauses at `.claude/plans/F-X272-design.md:3815` remain partial. Parser-side constructor and positive/Option facts, populated source maxima and affine summaries, original/derived scheduling, standalone Result error consumption, projected aliases, unproved revision/ruby/drawing/field derivations, aggregate initial-field admission, selected/counterpart reservations and checked mutator resource-failure propagation remain open. Q-C1-01 exact raw fallback and Q-C1-02 direct legacy infallible reads retain their precise scope. Internal read helpers inside checked operations are not exempt.

Issue430, HLD reconciliation, native G02, schema/native qualification, exact performance and final integrated, installed, hash, CI, packaging and release gates remain pending. No supplied scoped result closes them. No source, test, build, network, Git state, plan, progress or delivery record was modified in this review. Only this report was written. Review stopped after publication of this report.
