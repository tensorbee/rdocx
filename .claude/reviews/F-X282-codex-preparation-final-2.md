# F-X282, Root preparation final 2

## Scope and identity

Affected final review of Grok V4 after the third and last independent design peer and Root final 1. This is not a fourth peer, implementation approval or a performance measurement. The separate implementation phase has its own three-pass cap.

Input: `/private/tmp/rdocx-grok-harness/f-x282-preparation/GROK_PROPOSAL_V4.md`, SHA-256 `50e340233d47d75dc39139cb5c7d1ab046c478788083297b9f6c960b719a599e`. Handoff SHA-256 `dcbc6a179e30ae7e4d9dfc01f43266dd801e5898f805fa367c3f6005aba98f4c`. Root compared V3 and V4, checked the changed validation recipe against canonical verify and wheel/CI workflows, and inspected the current styles flush. Prior final 1 remains immutable. No product source changed in this review.

## Prior dispositions

- D1 is closed for preparation. V4 lines 310 to 325 name the workspace test with both binding exclusions, release wheel producer and installed Python 3.12 tests, typing and stub checks, and the canonical local-source-patched packaging gate. These agree with `.claude/commands/verify.md:124`, `:183`, `.github/workflows/wheels.yml:134` and `:186`. The ordinary develop-profile CI job is not substituted for the release ratio gate.
- S1 is closed for preparation. V4 lines 362 to 364 distinguish the preserved serialization failure from malformed replacement refusal. The current `.ok()` path at `crates/rdocx/src/document.rs:26181` can treat a parse failure as changed and then write after successful root retention. The proposal requires rechecking the implementation Base and separately labelling a real behavior change. It correctly qualifies the package injection as crate-visible. This review did not execute a malformed fixture or certify a future Base.
- Peer N4 is closed. V4 line 94 requires current part bytes Arc identity before model pointer or value equality.
- Peer N5 and the hash account are closed as preparation. V4 lines 222 and 356 retain current source rejection and separate other stories' intentional baseline work. No new F-X272 product dependency is invented.

The preparation mechanisms accepted in final 1 remain applicable. Their runtime and output evidence is still pending.

## Defects

Zero new defects in the affected correction.

## Smells

Zero new smells in the affected correction.

## Nits

### N1, low, stale design-pass counter

`GROK_PROPOSAL_V4.md:10` and `:561` say the counter stays at 2 while the same paragraphs and line 18 acknowledge three completed design peers. The actual final handoff repeats 2. This contradicts the user's explicit separate phase caps and can cause an unintended fourth peer. Change these current counters to three used, zero remaining for design, with a separate future implementation cap of three. Preserve dated earlier pass counts and do not rerun a peer to make this mechanical correction.

The stale source-author-resumed sentence at `GROK_PROPOSAL_V4.md:14` should also be labelled as an earlier snapshot when correcting current coordination text. The current source is paused for design clarification as line 6 already states.

## Verdict and remaining gates

D1 and S1 remediation is accepted for preparation. V4 still requires the N1 metadata correction before canonical adoption. This is not an approved canonical design, a claim, provider completion or sprint acceptance. Root checks the affected correction after the author returns it. The three completed design peers are not reset.

All performance ratios at the original script boundaries, local implementation regressions, F-X268 79/252/407 evidence, saved bytes, hash 49, release Python runtime and typing/wheel checks, packaging, WASM, native evidence and hosted CI remain unexecuted for F-X282. No build, test, timer, publication or Git mutation ran in this review. Return control to the orchestrator.
