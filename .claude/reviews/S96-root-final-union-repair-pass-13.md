# S96, Root final picture preparation repair, pass 13

## Scope and verdict

Reviewed the frozen picture increment `b50307b88fd747e430be690fbd179b29bb2cf2972271fe3c119f2df5b45be5f1`, boundary `0c48090ceb1e1368c2adba450b6cfc5b22fd27e427b08cb6f052250bb9076874`, and independent support review `868f93a7326042f469411ee40e2e7d0dd09e544feea1895b4a0c7c3a4061d00e`. Prior cold, reply and CI-test increments are separately reviewed.

Defects: 0. Smells: 0. Nitpicks: 0.

## Assessment

Root inspected the actual insertion bridge, collector factoring, both binding consumers, tests, handoff and independent review. Existing insertion-only APIs at `crates/rdocx/src/document.rs:32526` and `crates/rdocx/src/document.rs:32562` retain their signatures and do not request an inventory. The additive support operation at `crates/rdocx/src/document.rs:32605` retains complete qualification. Actual location reconstruction precedes shared collection at `crates/rdocx/src/document.rs:32803`. The debug location phase ends before the collector, preserving work attribution rather than hiding work.

The complete collector at `crates/rdocx/src/document.rs:33322` has two substantive consumers. It retains source order, physical raw occurrence matching, accepted drawings, media relationships, extent and ordinal checks, and unrelated errors. The shared read belongs to the same immutable document. Staged finish at `crates/rdocx/src/document.rs:26338` does not change its source inputs. Complete main or related Field detection declines sharing and retains fresh collection after finish. Collector failures propagate without fallback retry.

Both bindings retain inserted-location selection, original missing-picture errors and conservative prepare/commit/install ordering. Cell selection uses the original actual-source reader. The proposed cached Cell shortcut was rejected before adoption after a real held-formatting fingerprint mismatch. It is not present in this increment.

The genuine predecessor completes successful body and Cell insertion, full private occurrence equality and save/reopen checks before observing two main and related preparations instead of one. Final Native and binding cohorts each record 28 passes, zero failures and zero skips. The new assertions compare complete occurrences, package bytes, legacy and post-finish reads, Field fallback, malformed unrelated source behavior, insertion-only behavior and held-view atomicity. Strict lint, format, prose and diff checks passed. Setup compilation and zero-test-selection attempts remain separately recorded.

## Pending

This accepts the bounded source mechanism only. A fresh shipping artifact must establish both original picture flat-cost thresholds. Other flat gates, cold407, native, memory, external callable qualification, final integrated verification and sprint review remain pending. No threshold or gate is waived, and this is not a feature completion or review-cap reset.
