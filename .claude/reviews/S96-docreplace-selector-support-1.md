# S96 Document replacement selector support, pass 1

**Reviewed:** The frozen checked-selector increment against its retained pre-increment source, six files, 608 additions and two deletions. The 712-line increment patch includes Native and binding tests, the approved F268 contract, HLD03/HLD10 and progress. Accepted cold, reply, picture and CI-test carry is outside this pass.

**Verdict:** 0 defects, 0 smells, 0 nitpicks.

## Source and evidence identity

The canonical HEAD remains `39287e8cb9c4a014adf0b3191cbe3cdf99c20569`. This is a dirty reviewed candidate, not a clean HEAD or shipping artifact claim. The frozen packet is `/private/tmp/rdocx-s96-docreplace-selector-implementation-1`.

- `SOURCE_BOUNDARY.json`: `209f06c12aa6c1c12914c37ce3f49c50683cdd2e2e437985bb60133761ffa6c8`.
- `increment.patch`: `04641efc14f8e65d24d268a9a38b5df43019d71fd754c4408ee17f33fcd2921f`.
- Native document source: `5e120ed4970583b3b212b290771a5abd5f245f08f998fbd5e2a3b21a721d8af4`.
- Binding document source: `05c50e71735eb4debbe5e5eea05f2a7637594e925c58f49d58e2a53d2305fae3`.
- Final Native binary: `974513dd08d3f2fb4c9045c0f3107635a2f98c7ae78967698308b8b120e592d0`.
- Final binding binary: `674d602d7bd2c19ba7f96cb1c586bf29923bd983afca07a711bcea104ef2f0fe`.

Authenticated 72 declared file checks covering current final sources, protected carry/configuration, before/final snapshots and retained evidence, including all three relevant binaries. All matched. The final source and protected pins were checked again at report completion. No product source, tests, configuration or progress was changed by this review. No test or compiler was run.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Checked implementation and contract

The approved callable and its one binding consumer agree with `.claude/plans/F-X268-design.md:783`, `docs/hld/03-architecture.md:3538` and `docs/hld/10-bindings-spec.md:5357`. Existing public signatures remain unchanged. The added doc-hidden callable owns current discovery and qualification rather than accepting an external proof or stored fingerprint.

At `crates/rdocx-py/src/document.rs:17050`, the optimized route requires a bound item, one index and recognized supported kinds. `validate_owned_item` also validates the story at `crates/rdocx-py/src/document.rs:18310`, retaining foreign and causal-stale refusal before Native work. Constructed/numeric, unknown-kind and nested selectors use the existing resolver at `crates/rdocx-py/src/document.rs:17074`. Held formatting and retired text snapshots retain their distinct established behavior.

At `crates/rdocx/src/document.rs:47880`, the new operation obtains the actual complete spanned sources. `spanned_story_sources` always supplies the main entry before related sources at `crates/rdocx/src/document.rs:29960`, so its first-entry access has the existing producer invariant. Complete fallible `stories_in_sources` discovery finishes before selection. Missing-story refusal precedes replacement-character validation, followed by supported-kind and actual item-span checks at `crates/rdocx/src/document.rs:47916`. Unrelated malformed stories are not skipped by an early successful find.

The extracted `story_item_span` at `crates/rdocx/src/document.rs:31424` preserves the existing one-index, bounds and physical-kind checks. The ordinary `story_item_source` still obtains its checked owner and delegates those same checks. Nested paragraph handling in `story_range_paragraph_source` remains unchanged at `crates/rdocx/src/document.rs:31540`.

Only this owned operation constructs `CheckedItem`, immediately after actual qualification and without intervening model/package mutation, at `crates/rdocx/src/document.rs:47945`. The private target skips duplicate initial qualification at `crates/rdocx/src/document.rs:48027`. It does not skip the existing staged flush, prepared main/related read, canonical physical ownership comparison or subsequent Source/Field/copy/publication pipeline at `crates/rdocx/src/document.rs:48120`. The increment adds no parser, cache, authority catalog or input credit. Existing owned-candidate commit and error atomicity remain in use.

## Acceptance evidence examined

The retained predecessor binary `55d73e6c16c34c6d76a5d9de9107461d0d4246d9eec5d7c4c5ed992a01689293` actually failed the mechanism assertion with body/header preparation counts two and two. Both cases completed successful replacement, full saved/reopened package equality and unchanged cell assertions first. This is actual RED evidence, not a counterfactual claim.

The final Native log records three tests passed, zero failed, in 2.05 seconds. The source at `crates/rdocx/src/document.rs:58064` retains those semantics and requires one preflush preparation for body and header. The 42 fixture/action comparisons at `crates/rdocx/src/document.rs:58133` compare actual legacy and selector counts, mismatch details, exact errors and complete bytes. They include retained namespace aliases/raw content, simple and complex Fields, tables and controls, invalid replacement characters, bounds, kind errors, no-match and count mismatch. Malformed unrelated-story precedence and missing-story-before-invalid-literal are separate assertions. All seven related owner kinds actually reached successful replacement in the final log, rather than merely agreeing on refusal, using `crates/rdocx/src/document.rs:58271`.

The final binding log records one test passed, zero failed, in 1.44 seconds. Its 14 old/new pairs at `crates/rdocx-py/src/document.rs:21095` exercise bound/body/header success, held bold, existing text-snapshot retirement, absent/mismatch/invalid cases, explicit numeric construction, foreign/stale ownership, unknown kinds and nested fallback. They compare exception class/message or count and exact saved package bytes. Refusal/no-change cases also require unchanged revision and a usable held outside run. The live-text case is correctly retained as the existing stale refusal, not presented as success.

Five recorded existing controls passed for final-reopen rollback, preparation failure atomicity, owned/borrowed equivalence with real clone behavior, complete engine mutex retention and run-tab text policy. The authenticated gate receipt records strict Native/binding Clippy, formatting, prose and diff checks at exit zero. Separate initial unsupported direct nested-path and binding test compilation failures remain disclosed and are not relabeled as successful final evidence.

## Not found and remaining scope

Correctness: 0 findings. Contract/API: 0 findings. Panics/bounds: 0 findings. OOXML/raw preservation: 0 findings. Tests/evidence: 0 findings. Structure: 0 findings.

This is the bounded independent implementation review of the selector increment. Fresh ordinary artifact timings, final shipping artifacts, native application comparisons, memory/performance qualification and other final integrated riders remain pending. Recorded local functional results do not qualify those gates, reset earlier peer caps or certify the entire Sprint/provider.
