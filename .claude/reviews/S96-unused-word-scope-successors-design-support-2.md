# S96, unused Word scope successors, design support pass 2

## Scope and verdict

Second bounded read-only assessment of the revised contract at `.claude/plans/F-X268-design.md:816`, retaining pass1 finding S1 and the existing phase counter. The source is still sprint/s96 at cec87f9501eb13d0d5d60572472b227483262620 with the preexisting dirty source preserved. No product implementation of this proposal is present.

Technical verdict: the revised substantive policy arrangement resolves S1. Zero defects, zero smells and zero nitpicks remain in the examined contract. Root contract adoption and implementation evidence remain separate. This is not a performance result, an implementation microscope or permission beyond this technical contract assessment.

## S1, resolved in the revised arrangement

The plan at `.claude/plans/F-X268-design.md:822` promotes and renames the existing substantive allocator at `crates/oxml-py-support/src/lib.rs:432` to `prepare_revision`, with the exact concrete `RevisionPreparation::{Global, Local, Replacement}` policy. It does not add a new replacement wrapper around that allocator. Existing `prepare_scopes`, `prepare_local_scopes` and `bump` compatibility entrypoints keep their existing contracts and dispatch to their actual existing policy.

This addresses `AGENTS.md:103` without granting an exception. The three modes describe actual cases and prevent an unsupported independent-boolean combination. The enum adds no trait, generic, source file or module. The shared support crate is unpublished at `crates/oxml-py-support/Cargo.toml:6`. Its two replacement consumers are explicitly the existing Word Conservative paths, not hypothetical consumers.

## State and consumer contract

The mode contract at `.claude/plans/F-X268-design.md:824` preserves document successor index zero for `document_scope`, `capture_prepared` and commit. Those real uses are at `crates/oxml-py-support/src/lib.rs:259`, 406 and 546. Replacement omits only the requested owner successor objects. Its owner `successor` queries return None, while the old entrypoints keep their current ordered results.

The exact virtual generation budget remains binding at `.claude/plans/F-X268-design.md:826`. For prior generation g, N supplied retired scopes and B births, Replacement must retain document generation g+1, birth generation g+N+2+j for zero-based j, and final generation g+N+1+B. Overflow checks must still reserve the skipped positions. No new generation can be published by a failed or dropped preparation.

The same contract retains the complete retired-state inventory, validation order, serial, operation label and first-cause writes. These are the substantive checks at `crates/oxml-py-support/src/lib.rs:439`, 463 and 511, followed by retirement at 528. No retirement may move into preparation, and no check may move after the first committed retirement.

The changed Word callers remain the Conservative branches at `crates/rdocx-py/src/document.rs:6816` and the publication selection at 11465. Their new owner states come from births at 6898 and 11483. Their document and future-handle capture remains required at 6941 and 6971. AppendBody and every other global/local publication policy remain on their existing behavior. PowerPoint retains complete successors, including its actual consumers at `crates/rpptx-py/src/presentation.rs:4150`, 4202 and 4569. The caller inventory and exclusions from pass1 remain applicable.

## Diagnostic refinement

The revised plan at `.claude/plans/F-X268-design.md:830` observes both actual ScopeGeneration creation sites, the constructor at `crates/oxml-py-support/src/lib.rs:293` and the `state` helper at 322. The total records actual object creations and `size_of::<ScopeGeneration>()` value payload bytes. It explicitly excludes Arc headers, allocator overhead and other heap allocations. The separately named allphase map row must not reinterpret the six-slot staging or eight-slot splice arrays at `crates/rdocx-py/src/document.rs:193`.

Root supplied one further finite refinement during this pass: a second diagnostic records the actual owner-successor subset immediately after each real successor construction in the allocator loop at `crates/oxml-py-support/src/lib.rs:480`. This is a valid installed-consumer discriminator. Exclude the implicit document entry by index zero, not by testing `RevisionScope::Document`, since a supplied scope can itself have that kind. The total must count each creation once, and the subset must not add its bytes to the total a second time.

With that positional definition, genuine old installed dispatch can complete all semantic checks, then fail a zero owner-successor assertion with a positive actual count. The new Conservative dispatch must produce zero actual owner successors while still allocating its document and births. AppendBody, local and PowerPoint controls with actual retired owners retain positive required successor observations. The existing Word debug inventory guards at `crates/rdocx-py/src/document.rs:5914`, 7167 and 7215 inspect owner facts and projections rather than preparing revision successors. No diagnostic pause or hidden successful allocation is justified.

Root must adopt the exact row names and this positional subset definition before implementation. The diagnostic remains debug-only, uses the existing private `_word_transaction_outputs` surface at `crates/rdocx-py/src/document.rs:206`, and creates no shipping counter or stable Python API.

## Evidence obligations and handback

The revised requirements at `.claude/plans/F-X268-design.md:832` include real predecessor execution, zero/one/many states and births, exact generation and retirement comparisons, dropped preparation, future capture, foreign/duplicate/retired/stale cases and both overflow boundaries. The two actual installed consumer RED/GREEN controls at 834 are required in addition to helper controls. Their semantic and atomicity assertions must pass before the mechanism assertion is evaluated. The old-dispatch refusal or missing-mode compilation cannot stand in for an observed allocation RED.

The existing-file and HLD scope is now explicit at `.claude/plans/F-X268-design.md:838`. Source/HLD adoption, scoped validation and an independent implementation microscope remain required. Current release timing, memory and other final gates must be reassessed on a new qualified artifact after any product change. This report neither infers timing causality from allocation counts nor changes any threshold, sample protocol, review cap or delivery state.

Correctness, contract consistency, panic/refusal paths, structural scope and diagnostic discrimination produced no remaining finding. This pass ran no compilation, tests, product import, benchmarks or source repair. It changed only this review and returns control to Root for explicit design closure.
