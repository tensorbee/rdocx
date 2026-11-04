# F-X272 bounded Source design, Codex final 1

AMENDMENT A REJECTED. C1 remains blocked. The C4 namespace mechanism is accepted within its existing approved contract, with implementation and evidence still due.

Actual Grok third and last bounded Source design peer is `/private/tmp/rdocx-cross-review/source-design-grok-peer-pass3.md`, SHA-256 `6ff9e87da63613428ff69584a93bebab90b18211775808d0580b4f94831dce2f`. Root read the complete report and DeepSeek V4, SHA-256 `667195e2c3f00c37232e7f635d602f4bbc065ceaf5c5b8f1366bb7d501a24a55`. All three Source design peers are used. The separate implementation peers remain exhausted. No fourth peer or counter reset follows.

## Blocking contract findings

1. `DEEPSEEK_DESIGN_V4.md:310` sends public from_inline_raw through a fresh budget and maps its error to None. This changes fallback and can hide a checked resource refusal. Lines 290 and 300 discard the parent budget, while the existing swallow sites at `crates/rdocx-oxml/src/text.rs:13461` and `crates/rdocx-oxml/src/content_control.rs:2213` remain. The proposal does not provide a closed propagation route through all owning constructors.

2. `DEEPSEEK_DESIGN_V4.md:350` compares the work only after successful content construction. Root independently checked `crates/rdocx-oxml/src/content_control.rs:2738`, `crates/rdocx-oxml/src/text.rs:10654` and `:692`: those constructors already allocate retained root records. No deferred expansion mechanism is specified. This does not establish a guard before the new retained-source expansion.

3. `DEEPSEEK_DESIGN_V4.md:356` consumes the interior in the initial planner without preserving its original fact contribution to both sides of the bound. `crates/rdocx-oxml/src/text.rs:19292` defines the existing allowance. A fresh nested allowance double counts input, and adding only needed work changes the threshold. The proposal must retain one complete original allowance and an exact work ledger. Discarding needed work alone also fails to restore declaration-pool state after fallback, as the peer shows at `crates/rdocx-oxml/src/document.rs:362`.

4. `DEEPSEEK_DESIGN_V4.md:352` charges projected owners without closing the rebuilt-paragraph and already-charged input distinction. `crates/rdocx-oxml/src/content_control.rs:2892` rebuilds a paragraph. The checked header and text-box routes also lose inherited prefixes at `crates/rdocx-oxml/src/placeholder.rs:399` and `:1691`. Fixing Result recursion alone is insufficient.

5. `DEEPSEEK_DESIGN_V4.md:230` claims that infallible projections create no capture records. The actual wrapper construction at `crates/rdocx-oxml/src/text.rs:10885` and root capture at `:692` contradict that allocation argument. No projection exemption is approved. A private capture policy or other signature-preserving mechanism must preserve actual public output and selection while closing the resource path.

6. The concrete callable contract is also incomplete: `DEEPSEEK_DESIGN_V4.md:273` declares CaptureBudget private in document.rs, but lines 303 to 307 use it in sibling content_control.rs. The internal type and helper visibility must match their real callers. This is an additional Root finding, not a claim that the peer checked it.

## Accepted bounded pieces and next work

The F2 rebuttal is valid: effective scope bytes are already in current_bytes and must not be charged twice. Property and revision entry predicates, ordered inherited reduction, exact-prefix matching and the C4 nearest-URI shadow charge are suitable design. Root agrees with the peer's zero C4 finding. Existing C2, C3 and C5 local dispositions are not reopened and are not provider acceptance.

C4 may receive a separately bounded repair assignment under the unchanged approved F272 contract. It needs actual RED and GREEN, preservation checks and Root code review. That permission does not authorize C1 changes, Amendment A, a weaker allowance or Source completion. Finish the current native checkpoint before granting a canonical source writer lease. No source writer is authorized by this report alone.

C1 returns to author remediation with no further peer pass. Any revised internal capture transaction must state its concrete admission, ownership, constructor-fallback, pool rollback, allocation and public compatibility paths. Do not weaken the original bound or silently contract fallback behavior. Root reviews the revision directly. No provider completion, formal design amendment, claim, commit, release or other acceptance follows.

The peer's 7467-character report exceeded its 6500-character bound. Preserve that protocol miss separately from its substantive correct findings. No test, native, schema or final-union evidence was executed by this review.
