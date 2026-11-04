# F-X279a, all aspects, implementation pass 2

**Review kind**: actual independent affected microscope after separate Root remediation. D1 and D2 retain their original identities. This is implementation pass2 within the original three-pass cap, not a new design peer or a cap reset.
**Reviewed**: frozen working diff on work/f-x279a-codex against Base6bea9093818d6b691bf88b2980e9cc17927be015. Eight tracked product/test/HLD paths,1051 insertions and24 deletions. Exact diff SHA25689a3fa69c1063eafa6088ce6661cb28c4527ee206332393b0de38c6c0f572439.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. D1 and D2 are closed for the bounded local implementation. Final integrated acceptance remains pending.

Read the approved contract, operative clean-Base correction, original review, remediation progress and frozen source/evidence. Earlier binding/CLI/stub/Python-test/HLD10 hashes are unchanged from pass1. Read the new shared context mechanism, both actual native producers, affected callers and exact regressions. No build, test, imported-extension execution, source fix or evidence modification was performed. Only this review file was written.

## Defects

Zero remaining findings.

## Smells

Zero findings.

## Nitpicks

Zero findings.

## D1 closure, shared context and bounded owner traversal

`crates/rdocx/src/revision.rs:805` prepares SnapshotFacts once for the actual story-part batch, outside the per-record iterator. `snapshot_facts` at:1215 uses the existing tree and paragraph-domain inventory. Its two source-order passes derive scalar per-element admission, wrapper/scope/control roles, hidden/removed state and field seeds. It does not serialize an owner, parse a second source tree, clone source strings or grant a new work allowance.

Persistent field records at:1178 retain predecessor indices and scalar instruction depth. A begin adds a record, a valid first separator adds its result-state successor, and an end selects the predecessor. Records are immutable after insertion. New records are bounded by observed markers. The source pass keeps one head per actual story/domain and expires completed lexical domains. Deleted/move-source branches use an undo checkpoint at:1295 so their marker changes remain available to records inside the selected local branch without contaminating the accepted context after it. Nested undo restores the prior head. Expiration and undo operate on actual recorded transitions, rather than copying a field stack into every node.

The same substantive `snapshot_child` grammar at:1368 drives source context and selected-owner traversal. Physical destination roles are cached from actual ancestry, canonical-w declarations are propagated as scalar facts and foreign/opaque paths remain closed. The independently modeled textbox root can prepare its own facts without authorizing an outer snapshot walk to cross its drawing. Property history is not an admitted text path.

`snapshot_text` at:1440 now walks only the selected current or revision-local subtree. It seeds each actual story on first encounter, uses the shared field arena plus a local append-only arena for transitions inside that selected owner, and checks scalar instruction depth per atom. Its arena indices either address immutable shared fields or the current local vector, with new indices appended before use. Existing owner-range lookup and resolver behavior remain separate and unchanged. This closes the avoidable whole-source and preceding-paragraph replay from D1. It does not claim that arbitrarily overlapping large selected owners have globally input-plus-output work, and no such new contract or quota is invented.

The new discriminator at `revision.rs:4213` uses1000 and2000 disjoint one-character insertions. It checks actual preparation-loop visits, exactly5N selected-owner traversal visits, no field records and equal output through the real story_part_revisions producer. The preparation counter measures its two explicit facts passes, not an assertion that unrelated existing helper loops vanished. The actual instrumented old candidate makes3503500 selected-owner/prefix loop visits for1000 outputs and fails the stated bound. Final fixed observations pass. This directly rejects the old mechanism without relying on elapsed-time thresholds or favorable timing selection.

The deletion-fork discriminator at:4264 checks records inside a removed branch, later accepted records, a selected insertion containing a deleted field marker, and a later independent insertion. Its exact empty/visible strings distinguish shared accepted-prefix state from selected revision-local traversal. Prior cross-paragraph field, nested instruction, control-domain and opaque tests remain in the actual passing batch. D1 is closed.

## D2 closure, actual qualified owner relationships

`crates/rdocx/src/revision.rs:1455` now selects the required properties/owner pair by revision kind, verifies both Word namespace identities and exact local names, and returns empty text when the chain is absent or unsupported. It does not change which records are counted or remove preserved source. A paragraph-mark rPrChange has no fabricated run owner and remains empty. A supported run, paragraph or table owner must also be admitted and, in current mode, visible under its actual ancestors.

The concrete imported-control discriminator at `revision.rs:4204` preserves all three direct misplaced property records plus the valid paragraph-mark revision. It asserts that all four remain inventoried and all texts are empty, instead of borrowing unrelated control text. Existing real run/paragraph/table property positives remain in the same passing batch and select current content rather than prior properties. Existing section and private package records remain empty. The actual old candidate returns unrelated control text for all three misplaced records and fails the exact assertion. No schema-wide rejection, count reduction or new refusal behavior substitutes for the fix. D2 is closed.

## Remaining consumer and preservation audit

Both actual StoryRevision producers still populate owned text. The private comparison-package producer remains explicitly empty. The story-part producer retains source order, existing owner matching and error behavior. Scalars still come from the original parse event stream through the existing core plain/entity decoders and CDATA decode. Namespace aliases, entity-split whitespace, unknown entities, tabs/breaks and foreign lookalikes retain their tested distinctions. No synthetic fragment or dirty-C1-only parser helper was introduced.

Current property text excludes deleted/move-source content and removed rows, including hidden ancestors. Revision-local text keeps its selected deletion/move-source payload. Nested atoms are not emitted through both aggregate and leaf walks. Table paragraph tabs and retained outer-row newlines use the agreed boundaries. Raw transparent wrappers retain the operative clean-Base canonical-w gate. Its final reconciliation against completed C1 remains explicit debt.

The unchanged Python source retains the additive optional str text constructor/default, visible empty-string signature, frozen property and derived equality. It still forwards native text exactly once, preserves the public export, required nullable timestamp and story omission, and refuses None conversion. Stub and public test sources remain aligned. No actual installed signature, runtime equality, typing or stubtest result is inferred from compilation or AST.

The unchanged CLI source still takes JSON text from the native getter, uses the exact eight snake_case revision labels through its two consumers, preserves independent story labels and retains seven human TSV columns. Actual CLI lifecycle/composition tests remain source-applicable and passed after remediation. HLD03 now describes shared preparation and qualified owner attribution. HLD10 remains the reviewed owned snapshot contract. Neither claims new comparison/comment/import support or final acceptance.

## Evidence authentication

All eight SOURCE_FREEZE_2.json file hashes and the exact current diff hash match. All seven inspected remediation/baseline-repeat receipt log hashes match their complete logs. The retained old source at review-baseline-source1 matches the original frozen revision.rs SHA25662d159d58046c67a17f39610aebe6167eaf0da27f9cf693717ae0a6297d3dce4. The tests-only instrumented source matches its recorded hash. The repeat script preserves the repaired bytes and restores them in finally, with the retained repair draft matching the current frozen revision source. The historical repeat was justified by a missing old executable identity, not selection among timing outcomes.

The preserved baseline executable matches SHA256b8c1b22bc56b6ef2344a63273bbf4da4804c4618e727f84958a698105e799458. The complete terminal repeat log has actual child exit101 and both semantic assertion failures: measured3503500 visits and three incorrect unrelated text strings. These observations are distinct from the initial old-CLI missing-text/kind RED and the earlier compile setup failures.

Final native receipt has actual child exit0 and16 passes,15 new regressions plus the existing carriage control. CLI remediation has actual child exit0 and two native lifecycle/JSON/TSV composition passes. The unchanged label implementation/test retains its earlier actual one-case observation. Three changed crates' all-target/all-feature Clippy has actual child exit0 with warnings denied. Format, prose and whitespace receipts each record actual success separately. Their logs are not masked pipeline statuses.

Current final executable hashes match the recorded receipts:

| Artifact | SHA256 |
|---|---|
| Native library test | 1c3252086b87b2b31ba7e542db2c548fccd5fd0742353e82eb5d510bc788090f |
| CLI integration test | 3cf5079a58f4d9a20c026c9a23f66221d1f911b5652b53fd3dda7c3a2ef97c24 |
| Product CLI | 4c226fd3fb6d7e2f0bf83ad1a9c66bb6f101511e28938dd14c9abc7521caa0f4 |

No new execution is attributed to this review. Earlier label policy evidence is reused for unchanged source rather than relabeled as a fresh repaired-producer test. Python evidence remains source/AST only.

## Not found

Correctness, contract, panic/index/arena safety, OOXML namespace/order/source retention, test discrimination, evidence interpretation and structural rules produced zero further findings. New private concrete facts and field-transition types have actual shared context/traversal consumers in the same existing file. No new source file/module/crate, generic parameter, trait, dependency, feature or public decoder/quota was introduced. Iterative source and owner walks avoid adding source-depth recursion.

## Acceptance boundary and handback

The reviewed local repair is clean. F-X279a remains in-progress with an approved plan until its final acceptance. Final C1 alias/wrapper coherence, installed Python3.9/3.12 runtime/signatures/typing/stubs, integrated full/hash, applicable SDK/native, WASM, archives/dry-runs and hosted CI remain mandatory pending gates. Parent/B/F272 barriers and the narrow Q-A430-28 exception are unchanged. No provider completion, sprint total, baseline, comment, push, tag or publication is supplied by this review.

This pass ends with this file. Return control to Root for adjudication and authorized preparation. No remediation or additional confirmation pass is started.
