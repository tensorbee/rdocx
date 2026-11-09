# F-X278 preserved-carry reconciliation, pass 1

**Reviewed**: canonical HEAD `aff3f2b8744ed86755db8520c9bf6abeed5f2038`, including F278 integration `52260431602d7b8e2b0c6b69d0af56c6f3c465d9` and F282 approval/claim, after restoration of the separately archived F272 carry. This is a bounded restoration/composition audit, not another implementation peer, counter reset or full feature acceptance.

**Verdict**: 0 defects, 0 smells, 0 nitpicks within the inspected reconciliation. The retained C1 exact raw-fallback failure and all final qualification debt remain blocking outside this narrow restoration decision.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Archive and difference accounting

Read `F278_CARRY_INTEGRATION_TRANSACTION.json` and independently authenticated all129 paths in `/private/tmp/rdocx-cross-review/canonical-carry-before-f278-20261102T115810Z/MANIFEST.json` against their archived files. No archive mismatch and no current missing path were found. The original recovery stash remains identified as `d8742d4fa021ff86e2b4e3bb1f56364e53c42b14`. This review did not apply or change it.

Current bytes differ from the archive at the six named restoration paths, plus the expected local coordination file `.claude/scratch/S96-run.json`. That seventh difference records F278 reviewed/consumed integration and the actual F282 running claim. It is not an additional source discrepancy or completion claim. All other archived paths, including the eight already committed F282 review records and the shared instruction-stage source, remain equal to the archive.

The restored paragraph-property and HLD10 differences against the archive have exactly the F278 changed lines from the reviewed integration. For the Python test carrier and facade document.rs, changed-line comparison differs only through hunk boundaries, blank lines and formatting. Ignoring those presentation differences, added and removed line inventories match the F278 integration. Coherent reads of the resulting functions and test boundaries found no lost carry material or extra product change.

## Source reconciliation

`crates/rdocx-oxml/src/paragraph_properties.rs:1876` retains F278's pair replacement: a more specific firstLine or hanging replaces both fields. F272 source/publication role handling and namespace-aware property writing remain the archived implementation. Both F278 merge/raw-preservation controls are present beside the carry tests. The property merge does not substitute for the checked field or source-publication guards.

`crates/rdocx/src/document.rs:25069` retains the reviewed FontPolicy helper and all diagnostic publishers. The F272 facade preflight at `:22706` still calls `preflight_current_field` on actual source fields, and preparation at `:25228` still calls `prepared_current_field` before publication. Neither is bypassed by font selection. Existing policy caches, deterministic omission, explicit system refusal and caller routes coexist with the restored candidate/source checks. This is source composition evidence, not a new proof that every pending C1 route is complete.

`crates/rdocx-py/tests/test_core.py:8261` retains the carry's whole-owner field grammar and atomic-refusal controls. The F278 policy test and expansion/contraction selection test also remain intact. The common carrier was merged, not replaced with one story's version.

HLD10 retains the carry's checked source/field publication requirements alongside the integrated bundled/system policy, live PowerPoint index, validation Display and Markdown behavior. The archive-to-current HLD differences are the intended F278 font/cache/WASM/CLI wording. No archived checked-owner section was dropped.

## F282 contract and status reconciliation

Current `.claude/plans/F-X282-design.md` exactly equals the committed approved HEAD version. Compared with the archived draft, product clauses are retained. Changes are approved status, the operative rolling approval/resource-wave section and historical labeling of old approval prerequisites. The original performance thresholds, cold sampling rules, at-most-once closure, legacy checked fallback, cache atomicity and final gates remain. The new approval explicitly requires preserving F278 policy and reconciling any later F272 checked-candidate contract. The old draft was not restored over this approval.

CURRENT retains the archived operational history and the newer takeover/integration/resource-wave checkpoints. Earlier running/draft statements are historical, with later dated checkpoints stating actual progression. F276 and F278 are reviewed/integrated locally while delivery remains in-progress. F282 is approved and running in its isolated worker. F272 remains incomplete, with no provider readiness inferred from the shared instruction prerequisite. No reviewed state is promoted to final acceptance or publication.

## Actual focused observations

The two supplied post-restoration checks completed during this audit. Read their final receipts and rehashed their logs:

- `f278-restored-carry-native-check.log`: actual `cargo check -p rdocx --all-targets --locked --offline`, child0 and wrapper0. It reports one unused-variable warning in a carried test. This is not a clean Clippy claim.
- `f278-restored-carry-layout-controls.log`: actual `cargo test -p rdocx-layout --lib --locked --offline issue433 -- --test-threads=1`, child0 and wrapper0, seven cases passed.

These are applicable focused composition observations. They neither rerun nor erase the known33capture passes and1exact raw-fallback failure. The independently clean instruction syntax/scalar stage remains only a prerequisite. Full owner/context qualification, populated-model and normalization summaries, scheduling and composed-operation admission remain due.

Correctness, contract, panics, OOXML, tests and structure have no new finding in this bounded reconciliation. The complete dirty carry and all unrelated giant source regions were not reaudited. No build, test, benchmark, network, native app, Git mutation, cleanup, delegation or source fix was performed by this reviewer. Only this report was written.

Issue430/provider work, native/schema/rendering, hash and reviewed deltas, installed bindings/typing/stubs, archives/wheels/WASM, hosted CI and release approval remain pending. This report ends the audit and returns the canonical writer lease to Root.
