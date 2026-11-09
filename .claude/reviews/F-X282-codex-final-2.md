# F-X282, Codex final implementation microscope 2

**Reviewed**: Affected repair after final1 and actual implementation peer3 on work/f-x282-codex, Base aff3f2b8744ed86755db8520c9bf6abeed5f2038. Read current styles tests and shadow oracle call sites, approved checklist reconciliation, author handoff/applicability, peer3 and actual receipts. Final1 remains on record. Preparation caught its missed test-contract controls. All three implementation peers are now used, without a reset.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Local prepare may resume with explicit final debt. This is not performance, integrated acceptance or completion.

## Defects

None found.

## Smells

None found.

## Nitpicks

None recorded.

## Repair and evidence

The actual-open counter reset precedes from_bytes at crates/rdocx/src/document.rs:50196. It observes1parse/1proof and retains both model and part allocations. Real replacement, row addition and serialization preserve those Arcs, zero primary styles work and saved styles bytes. Actual changed text/row observations reject a no-op implementation. The separate changed-font fixture at crates/rdocx/src/document.rs:50230 invokes a real public publisher, preserves the original COW model and observes one changed serialization. It inspects saved and reopened font values, producer root metadata, exact self-contained opaque children and their order. Its inherited-prefix child explicitly verifies parse-time namespace materialization before asserting exact subsequent retention.

The existing debug shadow wrapper at crates/rdocx/src/document.rs:7880 saves/restores only the new cfg(test) primary styles counter, matching the existing distinction between primary work and the real ZIP-reference oracle. The oracle still executes and performs work. This is neither zero total debug work nor release timing evidence. Removing exactly this test instrumentation recovers the prior production prefix hash. The runtime extension remains applicable because this cfg(test) code is excluded from it. No release-production behavior changed in this repair.

Independently authenticated every current manifest file and exact full git diff. Checked final1-test-native3, final1-test-clippy, final1-test-fmt and final1-test-prose child/wrapper exits, timeout fields and log hashes. Read the native output: all15named issue437 tests pass. Previous OPC/Python evidence remains separate and applicable, not a newly executed combined suite. Initial two failed fixture/instrumentation batches remain failures and are not labelled production RED. Actual old-source performance-work and malformed replacement observations remain their original authenticated facts.

Checklist carrier wording now names the actual rdocx and OPC modules. CT_Body::tables was deliberately unchanged, so adding an unrelated rdocx-oxml test solely to tick the old carrier wording would not strengthen the behavioral evidence. No acceptance criterion or final gate was removed. This closes the hard prepare gap identified after final1.

## Not found and remaining gates

No new correctness, contract, panic, OOXML, test or structure defect was found in the affected repair. Final1's production trace and the prior lexical/physical-owner repairs remain applicable. The malformed replacement refusal still needs its isolated labelled behavior commit. Qualified foreign direct publication retains the explicitly documented namespace-placement distinction. Future C1 integration must reconcile the actual typed fallback, without imposing an unrelated raw-owner field policy.

All exact Linux release Python3.12 A282-1/2/3 ratios, cold sampling/original workloads, saved bytes/hash49, F26879/252/407, full workspace/native/oracle, installed runtime/typing, WASM, wheels/archives and hosted CI remain pending. Three implementation peers are exhausted, and this is Root final2 rather than a fourth peer. Separate reviewed-SHA release approval remains required. No tests/builds, source changes, network or Git mutation occurred in this review. This file is the only write. Review ends and the lease returns to the preparing orchestrator.
