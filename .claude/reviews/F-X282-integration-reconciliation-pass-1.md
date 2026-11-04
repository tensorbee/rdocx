# F-X282 integration reconciliation, pass1

Root adopts the actual independent Codex source reconciliation of immutable commit28a36b761eac87e446ac35d2556313aaa68c76c1. Verdict0defects,0smells,0nitpicks in this source reconciliation. The following original producer record is retained verbatim. Its temporary-source SHA256 is1436b84e4b888f828f7debcff53c986a1d5978870d8cb51ccdc096662d56430e. Later restored C1 source is separate, unfinished work.

# F-X282 integration reconciliation audit

Read-only audit of immutable canonical commit `28a36b761eac87e446ac35d2556313aaa68c76c1`. No product mismatch established. All reconciliation assertions below pass. The mutable canonical worktree is not the source of the product comparisons because Root is claiming the next story and may restore its preserved carry.

The prepared worker source is `a2a8bc7479606406e5b5faba0087b44a27912269`, after performance commit `f7f3a8559e56396f2f70d6907738dca4e784d308`. Its handoff-only tip is `1d7c4e007dd11fd84eb29a5759f7a0a1b4f1d9bf`. Base is `aff3f2b8744ed86755db8520c9bf6abeed5f2038`.

## Verified source and separation

All nineteen entries of PREPARE_FINAL_MANIFEST.json match worker tip bytes. All eighteen retained entries, including the actual prepared plan, all five reviews, source/tests and HLD files, match canonical commit bytes exactly. The nineteenth entry is the consumed structured handoff. It is absent from canonical HEAD and matches the preserved copy at `/private/tmp/rdocx-cross-review/F-X282-ready-for-integration.md` by SHA-256.

The complete worker performance commit diff is byte-identical to the complete canonical `e059f278` performance commit diff, 133070 bytes. The complete worker intentional behavior diff is byte-identical to the complete canonical `28a36b76` behavior diff, 1925 bytes. This proves the compatibility refusal stayed in a separate commit rather than being invisibly folded into performance work. There is no source or metadata omission hidden by a product-only comparison.

Compared all 326 compiled-input paths at canonical commit against worker source commit. Zero differences. Comparing the canonical bytes with the earlier build manifest yields only the already declared native test carrier and interpreted Python tests. Their test-only applicability and prior semantic outcomes remain separately recorded. No new build or execution is claimed.

Verified Base to performance to behavior to handoff-tip ancestry and canonical performance to behavior ancestry. The source-to-tip diff changes only `.claude/handoffs/F-X282-ready.md`. There are no crate changes after the worker source Head. At inspection the worker exists at the exact prepared tip with clean porcelain status. This permits clean worktree removal after recorded integration while retaining the worker branch. This audit did not remove either.

The local run-state F-X282 record identifies source Head `a2a8bc7479606406e5b5faba0087b44a27912269`, integration commit `28a36b761eac87e446ac35d2556313aaa68c76c1`, handoff `consumed`, state `reviewed`, the correct owner/branch/worktree and Base. Run state is a local scratch artifact, not a tracked HEAD blob. Reviewed is not completed.

## Failed bookkeeping assertion and correction

Root reported that its first integration bookkeeping assertion used the pre-prepare FINAL_SOURCE_MANIFEST.json for the subsequently prepared checklist-plan hash. That assertion failed before product verification. Root then continued making the integration commits before correcting the assertion. Continuing after the failed assertion was a process failure and is not suppressed by this successful reconciliation.

The correct comparison source is the actual PREPARE_FINAL_MANIFEST.json. This audit independently verified every prepared file, including the plan and reviews, both complete commit diffs, compiled-input equality and source ancestry. Those checks pass. No product mismatch was established by the original bookkeeping failure or by this audit. The failure must remain in the record rather than being relabelled as a successful first verification.

## Limits

No product tests, builds, network access, Git mutations, canonical writes or worker cleanup ran. Only this temporary report and the bounded companion JSON were written. Performance ratios, saved bytes/hash49, F-X268 79/252/407, full workspace/native/oracle, installed runtime/typing, WASM, wheels, archives, hosted CI and actual C1 typed-policy integration qualification remain pending. This source reconciliation does not discharge them or authorize publication.

Exact per-path results, diff hashes and captured run state are in `/private/tmp/rdocx-cross-review/F282_INTEGRATION_RECONCILIATION.json`. Stop and return control to Root.
