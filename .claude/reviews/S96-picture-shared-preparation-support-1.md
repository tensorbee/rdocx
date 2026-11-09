# S96, shared picture preparation, bounded support pass 1

**Reviewed**: Frozen seven-path increment over the retained dirty H10 checkpoint, 313 additions and 18 deletions. Cold-owner, reply and CI test carry are excluded from this pass.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Identity

The immutable packet is `/private/tmp/rdocx-s96-picture-shared-preparation-implementation-1`. `SOURCE_BOUNDARY.json` authenticates as `0c48090ceb1e1368c2adba450b6cfc5b22fd27e427b08cb6f052250bb9076874`, and `increment.patch` as `b50307b88fd747e430be690fbd179b29bb2cf2972271fe3c119f2df5b45be5f1`. All seven before snapshots, seven final snapshots, current final source/progress pins, six protected carry/dependency pins and all 37 evidence hashes match. No clean-HEAD claim is substituted for this explicitly dirty candidate.

Final Native binary `61b70745d18dcc36ad9f788ebf70f7ca5c64d94c8f74d0f4d207b6f7baa6a1a4` and final binding binary `48ae8abdae5f34aa495695547bee8863b7e247cdc93a9eb4feab6509c99385d9` authenticate. The predecessor binary remains separately identified as `b2d8a830ba1b6e697441cc0ca4b79aa6a08630fbeb8d3f2732e99b4df7746b5d`.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Source assessment

- The existing borrowed and owned insertion-only signatures remain unchanged at `crates/rdocx/src/document.rs:32526` and `crates/rdocx/src/document.rs:32562`. Both pass `None` to the shared candidate implementation. The conditional at `crates/rdocx/src/document.rs:32853` short-circuits on this absence, so those methods acquire neither an extra inventory nor its possible errors.
- The additive doc-hidden owned bridge at `crates/rdocx/src/document.rs:32605` retains image-option validation, staged-origin capture, actual insertion, final staged state restoration and the actual inserted location. It returns existing `PictureOccurrence` values from the complete collector. It does not synthesize metadata from options, media equality, drawing IDs or guessed physical coordinates.
- Authoritative location reconstruction still runs after actual staged reopen at `crates/rdocx/src/document.rs:32803`. It resolves the refreshed story, actual direct insertion slot and projected physical paragraph. Only afterward may the inventory consume that same prepared main read and related-source vector. The existing phase2 guard ends before collection at `crates/rdocx/src/document.rs:32851`, preserving phase0 collector accounting without pausing work counters.
- Sharing requires the complete prepared main source and every related source to be Field-free at `crates/rdocx/src/document.rs:32853`. A decline leaves the inventory absent and invokes the original fresh collection after staged finish at `crates/rdocx/src/document.rs:32645`. An actual collector error propagates through `?` and is not converted into a decline or retried.
- The extracted substantive collector at `crates/rdocx/src/document.rs:33322` is shared by the ordinary `pictures` query and the new bridge. `PreparedMainRead::sources` at `crates/rdocx/src/document.rs:9230` preserves body-first and related-source order with the same actual part names and bytes. The complete owner, accepted-paragraph, raw occurrence, relationship/media, extent, ordinal, ambiguity and error logic remains in that collector. Unrelated-source failures are not skipped merely because an inserted location is already known.
- Collecting before staged finish is compatible with the actual finish implementation at `crates/rdocx/src/document.rs:26338`. It adjusts signature invalidation state and swaps layout-engine mutexes, without changing the package, model or XML used for picture inventory. Field-bearing sources retain their original fresh boundary.
- Both real binding consumers use the returned complete vector at `crates/rdocx-py/src/document.rs:16534` and `crates/rdocx-py/src/table.rs:2209`, then retain exact inserted-location selection, snapshot allocation, conservative preparation and commit/install ordering. Cell physical selection remains the original checked route. The rejected indexed-fingerprint shortcut is not present.
- The explicit callable and behavior contract is in `.claude/plans/F-X268-design.md:778`. Required current HLD descriptions are present in `docs/hld/03-architecture.md:3535` and `docs/hld/10-bindings-spec.md:5354`. No new type, parser, module, trait, generic, retained cache, Source input or resource allowance is added. The private collector has two substantive consumers.

## Evidence assessment

The retained predecessor test completes successful body and Cell insertion, full occurrence equality including private physical source, and save/reopen equivalence before failing the expected one-preparation assertion. Its actual work is `[2, 2]` for each binding sequence. The initial unqualified exact filter selected zero tests and is explicitly not the RED.

Final Native picture tests record 28 passes with zero failures or skips. Final binding unit tests record 28 passes with zero failures or skips. The new tests at `crates/rdocx/src/document.rs:57812`, `crates/rdocx/src/document.rs:57872` and `crates/rdocx-py/src/document.rs:21015` inspect full occurrence/package/legacy/reopen equality, one-preparation delivery for both actual consumers, Field fresh fallback, insertion-only success despite a later unrelated collector error, identical refusal for the complete inventory, held-bold Cell selection, invalid-image rollback and conservative held-view retirement. Existing affected picture controls retain repeated IDs, alternate content, namespace/raw admission, complete inventory, resource delivery and staged reopen refusal coverage.

The final strict Clippy, format, prose and diff receipts record exit0. Initial test fixture compilation failures remain distinct from product RED and final observations. These local checks establish the bounded mechanism and functional behavior. They do not establish the original picture timing law or native parity.

## Limits and handback

Correctness, contract, panic/index boundaries, OOXML preservation, tests and structure each produced zero findings within this increment. Current pins were rechecked before handback. No source edit, build, runtime, product import or native comparison was performed in this review.

Root final review, fresh shipping artifacts, applicable external callable/build qualification, original performance thresholds and final integrated native/memory/release riders remain pending. Existing feature caps and acceptance criteria are unchanged. Native long-note and restored plain-footer comparisons remain separate queued work.
