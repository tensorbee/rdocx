# S96, current modeled-guard scheduling, investigation21

## Disposition and actual result

A smaller safe storage candidate is available: keep the original child Vec local while using the standard extend iterator to enqueue Enter steps, then move that Vec into a previously queued empty Finish slot. This preserves the removed duplicate child buffer without retaining design19's per-child lookup of a Finish inside the growing pending Vec. Empty children can keep the direct original-vector Finish push and need no placeholder installation. This report does not adopt or implement the candidate.

The current valid ordinary result remains 75/79 flat passes, five linear passes and a primary407 pass. Reply, Paragraph.replace_text, Document.replace_text_at and Cell.add_picture fail their flat thresholds. Secondary cold insertion is 3.0192080012056977ms against 3 ms. Document picture and cold removal pass. The complete result is retained at `.claude/scratch/S96-final-evidence/f268-text-stack-loan-release-qualification-20261108-1/monitored-run-1/ACTUAL_TIMING_SUMMARY.json`. It is not invalidated, favorably retried or replaced by this investigation.

The timing comparison does not isolate design19 from 18/20, artifact code layout or the rest of the operation. It cannot establish that design19 caused any particular regression. The generated-code difference below supplies a concrete local mechanism worth testing. It does not establish elapsed causality, the cost of that mechanism or a promised repair.

Zero existing product defects, zero smells and zero nitpicks are established here. Root29 and Support11 remain correctness reviews with their original scope. This investigation writes only this report. Root remains the sole product/Cargo writer.

## Current contract and generated-code evidence

The current new-summary branch calls resource_children once, moves its returned Vec into Finish, reserves pending capacity and then retrieves every child by the saved Finish index. The completed-but-overbudget branch retains the original into_iter/map/extend path. Guards, map state and local checks precede both branches. References: `crates/rdocx-oxml/src/document.rs:1500`, `crates/rdocx-oxml/src/document.rs:1551`, `.claude/plans/F-X268-design.md:1142`.

The saved logical index is correct across reallocations, as Root29 and Support11 establish. Correctness does not require the optimizer to eliminate its repeated indexing or Vec push checks. Source-level removal of a child Vec clone is consequently insufficient to infer a lower total execution cost.

Read-only llvm-objdump inspection is available for two actual optimized-symbol payloads. The earlier Root28 image is `.claude/scratch/S96-final-evidence/f268-four-remaining-symbol-diagnostic-20261108-1/payload/rdocx/_rdocx.abi3.so`, SHA-256 d437857af1fbd225ee2b44568947e7cf0a5c99b7ce4be92fbc3b0b5c4aed4cfc. The current Root29 image is `.claude/scratch/S96-final-evidence/f268-five-remaining-symbol-diagnostic-20261108-1/payload/rdocx/_rdocx.abi3.so`, SHA-256 c0909f3d0b071c13f7f5b6ce768a2eadfa282e1d36333f73c83c9604aab42875. They are the separate symbol-preserving diagnostics, not the ordinary timed modules. The tool was `/Library/Developer/CommandLineTools/usr/bin/llvm-objdump`, selecting the actual word_model_children_maxima symbol.

The earlier function contains 709 decoded AArch64 instructions. Its child scheduling reserves once, then its loop at 0xeb4098 through 0xeb40b0 loads one borrowed owner, writes an Enter record, advances pointers and branches on the end pointer. The nonempty new-summary path also has the expected separate child-buffer allocation/copy before scheduling. That copy is precisely the work design19 removed.

The current function contains 781 decoded instructions. At0xeb4364 through 0xeb43b4 its new-summary child loop retains checks for the saved pending index, the Finish variant, the child index and the pending push capacity. A growth branch remains inside that loop despite the preceding reserve. Its completed-overbudget branch still has the compact standard-extend loop shape. A panic_bounds_check destination is present in the current function, consistent with the retained checked indexing. Those defensive checks are correct. Their generated existence is a performance tradeoff, not a product defect.

The counts describe these complete emitted functions, not isolated dynamic instructions per call or time. Different whole-source LTO artifacts are not a controlled19-only code-generation experiment. No candidate21 assembly exists yet, so it would be incorrect to claim that its borrowed iterator chain already compiles to the earlier tight loop or that the placeholder writes disappear.

The retained composed18/19/20 loan diagnostic also supplies a useful distribution. On its large reply state, Finish child lengths include 53565 zero, 10896 one, 5490 six, 3456 two and 49 of length 139. Large Document picture includes 122900 zero, 24422 one, 13860 six, 8712 two and 122 of length 141. Cold insertion1000 includes 480142 zero, 70000 six, 109 one and 30 each at 1001/1002. These are actual per-state source-level debug graph events, not allocation totals or counts from the new five-state optimized profile. They motivate preserving an efficient empty-child case rather than adding one unconditional replacement match to every leaf.

## Preferred finite candidate

Keep resource_children, ModelAdmissionStep, its Finish payload, the summaries map and the completed-overbudget branch unchanged. Change only new-summary scheduling after the original child enumeration.

For an empty children Vec, push Finish with that same original Vec directly. There is no Enter to enqueue or clone to eliminate. Keep all preceding enumeration, local checks and active-state insertion. Do not infer a model resource result from emptiness before the actual enumeration.

For a nonempty children Vec:

1. Save pending.len as the logical index for the Finish about to be appended.
2. Push the existing Finish variant with the same key/local tuple and Vec::new as its temporary child payload. This empty Vec creates no child-buffer allocation.
3. Keep children in its local variable and append the exact Enter sequence with the standard iterator expression below.
4. After extension completes, match that saved Finish slot once and move children into its payload. The same internal unchanged-variant assertion may protect the slot. Do not silently abandon installation or return a summary on a mismatch.

The substantive enqueue expression can be:

```rust
pending.extend(
    children
        .iter()
        .copied()
        .map(|child| ModelAdmissionStep::Enter(child, cost, tables)),
);
```

This retains a standard exact-length iterator/extend path and removes the per-child borrow of pending. It requests no custom trait, generic iterator wrapper, new frame variant, arena, child cache, raw pointer or unsafe access. Do not replace it with insertion/rotation of the queued records, a second model enumeration, a new child-step Vec or a cloned child Vec. No explicit per-child reserve or push is needed in that branch. Actual specialization, bounds elimination, descriptor writes and inlining remain generated-code observations to obtain from the implementation rather than promises made by the source expression.

## Exact transient and final-state safety

The placeholder exists only while one local scheduling expression appends Enter records. No step is popped or processed during that interval. The map closure constructs only an Enter from an already enumerated Copy WordOwner and the existing scalar costs. It invokes no model getter, Field projection, reader, callback or fallible Source operation. No other thread can inspect this call-local pending Vec.

Both the logical Finish index and its key/local fields remain unchanged through extension. Reallocation may move the buffer but preserves its existing logical element positions. After the iterator finishes, the saved slot is therefore still that same Finish. Install the original child Vec before the surrounding while loop can pop any step. Never retain a borrow into pending across the extension.

For children c0 through cn, the completed pending suffix is still Finish(parent, original children), Enter(c0), ..., Enter(cn). Its actual enum values and borrowed identities are the same as design19 and the original cloned-vector algorithm. LIFO still visits cn first, and the Finish fold still sees the original child order. Empty-child scheduling reaches the same final Finish value without transient replacement. References: `crates/rdocx-oxml/src/document.rs:1551`, `crates/rdocx-oxml/src/document.rs:1577`.

An allocator panic during push/extend unwinds the local pending and children values. A temporarily empty Finish is never interpreted as a completed summary. No user callback or modeled-resource error is introduced in this interval. The existing stored borrowed WordOwner values have no private model destructor or cloned AST to publish. This does not promise recovery from allocation failure, and no public error or new fallible-reservation policy is selected.

The original local weighted/table checks, completed-hit test, active-cycle state, completed-overbudget expansion, actual child enumeration, Finish maxima fold and root aggregation retain their positions. The candidate removes no Source, attribute, event, namespace, capture, constructor, Field or resource read. It alters no drawing/relationship reservation/remap, reader output, callback/refusal atomicity or publication. All proof and first-cause obligations from19 remain binding.

## Storage and Q22 attribution

The child buffer stays in a local variable until it moves into Finish, rather than entering Finish immediately. This does not extend its total lifetime beyond the existing pending Finish lifetime. Its actual capacity is the same original source-vector capacity used by design19. It does not regain the original clone's length-sized capacity or eliminate the previously disclosed spare-capacity cost.

The placeholder adds temporary descriptor stores and one installation match for a nonempty new summary. Its empty Vec has no child heap buffer. Standard extend can choose its own pending growth, whose actual capacity must be observed. A source-level empty placeholder must not be counted as a new retained vector or claimed as an optimized no-store operation. Empty leaves should retain the direct move case to avoid an unnecessary replacement match. Nothing survives the guard call and there is no new cache omission endpoint, model root, history, RSS ceiling or memory-policy relaxation.

Measure actual child-buffer clone/copy absence, local/Finish capacities, queued Enter copies, pending growth and descriptor/installation operations separately. Recheck final memory/equivalence/calibration applicability. Fewer source-level copies or fewer assembly checks do not establish lower peak/RSS or a faster public operation.

## Rollback alternative and attribution

A targeted rollback of19 is also technically safe: restore only its original new-summary children.clone plus original into_iter/map/extend scheduling, retaining18 and20 and their reader/ownership contracts. That would intentionally restore the actual duplicate child-vector allocation and shallow copies, as well as the earlier length-sized Finish capacity. The full guard/read/error contract is unchanged. The plan/HLD and memory attribution would have to describe that restored policy honestly.

Such a rollback is a source candidate, not permission to select an earlier favorable timing result. A controlled diagnostic comparison must bind otherwise identical18/20source, inputs, compiler/configuration and actual guard sequence. The current valid ordinary result stays retained. A subsequent reviewed rollback artifact requires its own ordinary build and unchanged complete qualification. A faster instrumented loop or a favorable subset cannot discharge failed public rows. A causal claim specifically about19requires a controlled isolated comparison, not the existing Root28-versus-Root29whole-artifact history.

The placeholder candidate is the narrower first investigation because it keeps19's removed duplicate buffer and addresses the observed local scheduling-code cost. This is a technical selection for possible adoption, not a guarantee that it beats rollback or meets any threshold.

## Required adoption and observations

Amend the existing F268 contract and HLD03/04 before product adoption. Preserve all18/20source and tests. Use the current reviewed manual-index implementation as the actual immediate predecessor, not a different Root28source that also lacks18/20.

Required local evidence includes:

- Exact full per-call normalized Enter/child/Finish and capacity traces, actual resource_children calls, summaries, weighted/table arguments, success/maxima and typed first refusal. Observe the installed full Finish payload before any processing. Do not treat its temporary empty payload as a successful resource fact or normalize away unexpected graph/capacity changes.
- Zero/one/many children, a wide body forcing pending reallocation during extension, repeated/shared roots, both prospective Field-summary parent-cost callers, completed fitting and completed-overbudget paths, independent limits, invalid-tail first cause and pre-writer output immutability. Keep existing private Field/read/refusal controls. No synthetic public cycle or allocator failure is invented.
- Actual current public state reproduction with unchanged setup and preceding actions. The newly declared large diagnostic states are replyk2, Paragraph.replace_textk1, Document.replace_text_atk1 and Cell.add_picturek1 at100, plus original407cold insertion1000. Record unvisited guard paths honestly. Semantic, held/fresh/refusal and save/reopen checks must precede mechanism assertions. Compare identical input ZIPs, complete outputs and every package member.
- Source-bound generated-code inspection of the actual immediate predecessor and candidate under the same optimized configuration, including the empty case, nonempty scheduling, retained checks, reserve/growth paths and placeholder installation. Static instruction counts are descriptive only. No assembly of a toy replacement function establishes the actual LTO consumer result.
- Scoped Native/OXML/binding checks, strict lint, affected installed observations and independent implementation review before a fresh ordinary release artifact. Only a new compiler-free admitted complete unchanged 79-flat/5-linear/primary/secondary407protocol can decide elapsed acceptance. Retain every prior failure and all final memory, Office, typing, hosted/archive and full-union riders.

No semantic blocker was found for the exact placeholder-local-buffer arrangement. Its actual generated code and elapsed sufficiency remain unproved. This pass ran only read-only source/evidence analysis and static disassembly, with no Cargo, compiler, product import, test, benchmark, elapsed qualification, source/plan edit or Git/network mutation. Only this report was written. Control returns to Root.
