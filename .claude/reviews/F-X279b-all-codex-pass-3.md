# F-X279b, all, implementation pass 3

**Reviewed**: Entire frozen working diff against Base and HEAD `391280577110c9b7b92ace0420a4a1d98c707889`, branch `work/f-x279b-codex`. Thirteen tracked changed paths including the operative plan, 3006 insertions and 103 deletions. This is the third and final independent implementation peer under the original cap of three, not a reset or fourth peer.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. D3 is closed across protected coalescing admission and the reply-copy publisher. Earlier D1/D2 closures remain valid. Root final and integrated acceptance remain required.

## Source and contract authentication

The reviewed immutable packet is `/private/tmp/rdocx-s96-f279b-evidence/peer3-complete-d3-frozen/source`. `SOURCE_PINS.json` SHA256 is `06aa69a48d7e2b7be5710b4e8d227949838d7100785028b54a6091877911b3e6`. All 16 source/configuration/plan snapshots authenticate against their pins and the worker files. The four context entries authenticate progress, plan, peer 2 and Root final 1. Both canonical D3 supplements authenticate against `CANONICAL_SUPPLEMENTS.json`. All source, context and supplement pins remained unchanged at the end of this review.

Read the canonical microscope/workflow requirements, current approved plan and progress, both earlier peers, Root final 1 and both D3 supplements. Revisited the complete production, test and HLD03/04/10 diff against Base and compared the successor with peer 2's exact snapshot. A/C/D authoring, markup and revision-vocabulary leases remain excluded. B retains A279-3/5/6 and P1/P2/C6 with the completed local F272 provider as an ancestor. No acceptance criterion or final rider was removed by remediation.

## D3 closure

### Protected admission no longer uses ordinary ignore masks

`crates/rdocx/src/comparison.rs:10759`, `crates/rdocx/src/comparison.rs:10789`, `crates/rdocx/src/comparison.rs:13242`

Protected paragraphs use a private default comparison view for the coalesced typed-content and formatting test. Caller ignore options remain in force for ordinary comparison only. The actual positive owning-source residual is still separately required through `HeldRevisions::same_owner_source` at `crates/rdocx/src/comparison.rs:3381`. Equal typed values cannot substitute for missing residual evidence.

The property-span signature now retains zero-width formatting facts. Empty text or ignored whitespace therefore cannot conceal a changed real run property. Actual nonempty whitespace remains significant in protected admission because the strict view does not discard it. Physical partitions of equal-property plain runs can still qualify. Structural tokens, Field form/source facts, raw/control content and source-residual boundaries remain independent barriers.

The regression at `crates/rdocx/tests/regression_test.rs:413347` checks all granularities, both formatting directions, empty/plain/all-space payloads, ignore masks, exact refusal bytes, the equal-property split positive and ordinary ignored-whitespace preservation. The installed Python controls additionally keep actual live paragraph/run handles through protected refusal and byte-preserving positives. These assertions reject the actual R2 formatting and whitespace bypasses, rather than merely checking that the old insertion text survived.

### Reply publication checks the entire containing held owner

`crates/rdocx/src/comparison.rs:907`, `crates/rdocx/src/comparison.rs:1108`, `crates/rdocx/src/comparison.rs:3336`

The actual original-plus-reply producer is checked before it is returned for copy publication. The guard reuses the real HeldRevisions source classifier, with its former Metadata dependency reduced to the boolean it actually used. All previous callers pass their unchanged `metadata.existing_revisions` value. The shared positive residual comparison has both ordinary admission and reply-qualification consumers.

Changed main and related story scopes are inspected using their actual source spans and inherited bindings. Main paragraphs, whole tables and block controls are compared as complete typed owners with positive owning residuals. Held section properties retain their separate typed/source check. Exact unchanged parts take the existing byte-equality shortcut. Intentional comment definitions are excluded from the body-story check and the existing revised-comment snapshot refusal remains upstream.

The index pairing is local to an actual authoring replay cloned from the original. That producer adds graph markers but does not insert, remove or reorder body owners. It is not post-mutation owner discovery on the arbitrary edited operand. A changed containing owner fails before publication with its physical location. An unrelated held owner can remain equal and pass, so the repair does not reject every document containing revisions.

The regression at `crates/rdocx/tests/regression_test.rs:413424` discriminates paragraph, table and block-control ownership. Table/control negatives place the old revision in another paragraph within the same whole owner, so an immediate-parent-only check fails the test. The corresponding unrelated-owner positives retain the old source through snapshot accept and reject. Python also checks refusal bytes, comment graph and live observer paragraph/run handles. Actual predecessor failures and the guard-only successor's still-failing publisher probe remain recorded in the canonical D3 supplements and private packets.

## Full-candidate review

D1 remains closed by using the actual original-plus-reply replay as the edited model and Source before relationship remapping and comparison Metadata extraction. D3 adds the previously absent whole-owner admission check to that same publisher. Ignore-formatting/whitespace body and related-story preservation, P1 graph acceptance/rejection and P2 selected resolution continue through the same concrete policies. Revised-comment snapshot refusal and unsupported protected-owner edits are not waived.

D2 remains closed through concrete row and cell helpers traversing direct and control-owned table/row collections with their real enclosing table-style context. The body-font frontier, style precedence, source-default selection, direct components, authoritative prepared bytes and save/reopen observations remain unchanged. Excluded inline-run controls, revised raw payload and conditional table formatting were not silently added.

Rechecked logical-word alignment, token and shell boundaries, default run behavior, character equality, protected and ordinary equality, positive Field/control source facts, paired nontext fallback, separate physical content and formatting owners, generated word revision envelopes and source interleaving. Per-owner Source publication remains separate from logical grouping. The paired leaf-capture branch does not acquire ordinary word formatting authority.

Rechecked native body-only insertion and Python's existing full companion/relationship closure, public insertion-policy constructors, omitted versus explicit list reuse, empty-source/clamped destinations and staged refusal. The binding signature and stub contract are unchanged. Deterministic font/list rendering assertions inspect actual source-effective positioned values or SVG as required, rather than only cached style text.

No new public API, quota, crate, module, trait, generic abstraction or feature flag is introduced by the D3 repair. No additional correctness, contract, panic/indexing, OOXML order/namespace, binding, test-structure or repository-structure finding was established.

## Actual evidence at report writing

`root-d3-publisher-green/RECEIPTS.json` records public Rust20, Native29, CLI4 plus9 and strict all-target/all-feature Clippy exit0. Retained binaries authenticate against their recorded hashes. Product, test, HLD and configuration pins match this final freeze. The plan-only difference is the subsequent truthful readiness update, not a production delta or relabeled execution.

The fresh installed successor completed during this review. `python-provider-gate/run-9le8r5bl/RESULT.json` records exactly32 passed nodes, no skipped/error/failed cases, stable source/configuration and successful build, installation and exact native/stub authentication. The wheel SHA256 is `7980d891149974421a169e35beee265b15d9f1c997996bf70c8c4e30d6de4592`, with native member `c3b13e62e03db0e885cc0b1fc7ee6efcc8a1f7de83cf96c5e1f47bbea8945134`. This reviewer independently verified the wheel, all three recorded archive members, exact JUnit membership/count and JUnit digest `30888dd8e4855bae421a11d1a02a315f7f326f8180dd577c54f64ecb43afa1c2`.

The32 nodes include the original B criteria, both peer-1 repairs, all protected formatting/whitespace cases, all three whole-owner reply publishers and four unchanged F272 provider/lifecycle nodes. The genuine byte-preserving no-op still asserts exact bytes plus live paragraph and run handles. The narrower B equal-paragraph oracle permits only its observed sectPr namespace-closure publication and applies the existing changed-publication handle policy. It does not redefine a genuine whole-operation no-op.

The run is local opt2 macOS CPython3.12.13 evidence, not hosted CI, a release wheel, the final installed matrix or full-suite qualification. Earlier failed wheels, fixture corrections, actual Source regressions, accepted-provider Base evidence and guard-only observations remain separate. Source6 reuse remains its qualified per-case branch applicability, not a new run on this source. This reviewer ran no Cargo, executable, network, native Word or oracle operation.

Final integrated verification, complete suites, hash/pixel measurements, installed/typing matrices, WASM, packaging, schema/SDK, performance and fresh native Word comment-candidate/comparison acceptance remain pending. Q28 stays unproved under its narrow existing exception. Optional deeper assessment stays deferred. No publication, sprint completion or whole-provider verdict follows from this report.

## Defects

0. D1/D2 remain closed. D3 is closed across both recorded admission and publisher mechanisms.

## Smells

0.

## Nitpicks

0.

## Not found

No additional correctness, contract, panics, OOXML, tests, structure or binding-compatibility finding in the complete frozen B candidate. Required final execution remains explicitly separate from the clean source review.

This pass ends here and returns the review-only lease. Only this review file was written. No source, test, progress, plan or HLD file was changed by the reviewer.
