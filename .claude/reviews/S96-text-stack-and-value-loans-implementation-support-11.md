# S96, text, model stack and value loans, support11

**Reviewed**: The seven-path, 895-line composed18/19/20 increment against the frozen Root28 plus F278 test/documentation and contributor-credit baseline. Scope is finite opening text materialization, modeled-resource child-vector scheduling and ephemeral standalone paragraph return ownership, with their local tests and F268/HLD contracts. Historical earlier sprint changes and the separately checked contributor credit are outside this increment.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

This is an independent implementation review. The reviewer investigated design19 but did not implement any of these changes. Root remains the sole product author. This pass changes only this review file and is neither a full gate nor performance acceptance.

## Authentication and scope

The packet is `.claude/scratch/S96-final-evidence/f268-text-and-model-stack-implementation-20261108-1`. FINAL_BOUNDARY.json SHA-256 is a00d5eaacb286d16e024e09db8c9080277488157c7af26774205bdedb625eca9. FINAL_INCREMENT.patch is 6ae27b51a4bd2dfc95fab207056181f979d004aea064b8cb6166d18211ee8eef. FINAL-inputs.json is d7f7e9d8b77cbb0cf7bd8fb82e4a743f83ac3b94b003d8fcd4d47daef8645563.

Independent read-only checks match every current file against all 4309 final input hashes, all seven changed paths against their frozen final-source copies and the boundary, and 4302 unchanged inputs against the before manifest. The ordinary debug producer's absolute-path manifest normalizes to the same complete final input map. Both complete materialization/guard and ownership mechanism comparisons were independently recomputed from retained ZIPs, every package member, actual input hashes and compressed graph records.

The final ordinary debug wheel is 5df81dc520c02261a33e06617c568c48b8ba6e79fb1fd327090e0e7a9466a818. Its Native module is 96a05103acf4321b338c602357f728e781e0e7c4999436094bbc827b04eb4f72. All ten wheel members match their receipt and installed payload. The temporary diagnostic entry point is absent from that ordinary module. The build, installation, test receipts and raw results were inspected without importing or executing the product during review.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Finite text and resource traversal

At `crates/rdocx/src/document.rs:35463`, direct text construction remains after successful qualification by the unchanged finite plain paragraph/property policy. Length accumulation uses checked byte addition, rejects an over-isize capacity request and retains the old accepted_text materializer for unavailable lengths or variants. The successful path requests one final String capacity and copies stored decoded Text values in run/content order. Empty text keeps zero requested capacity. It does not insert separators, normalize spaces or decode source XML again. accepted_run_paths still runs after text materialization. Existing opening/staged equality, writer-vector identity, late mismatch discard, model-loan alignment, lock and unsupported resource fallback order are unchanged.

The complete legacy-fact comparisons and new materialization cases discriminate empty paragraphs/runs/chunks, multiple Text chunks, many runs, spaces, character references and finite properties. Unicode, tab and unsupported-format cases retain finite misses. Existing private Field equality/resource and missing or changed opening/writer controls remain actual tested consumers. No general accepted-text or run-path traversal was replaced.

At `crates/rdocx-oxml/src/document.rs:1551`, resource_children still executes once after the original map and local limit checks. A newly summarized node moves that same child Vec into its existing Finish step. The saved logical index is valid throughout the short append-only scheduling loop. Each child is copied into a local before the next pending push can reallocate. The child count comes from that unchanged vector. Neither input data nor a stale backing pointer supplies an unchecked index. The documented invariant assertion at 1560 is therefore internal, not a new malformed-input panic.

Ascending child scheduling produces the same Finish/Enter stack suffix as before, preserving LIFO visitation. Completed fitting hits, active-cycle checks, completed-overbudget expansion and the Finish maxima fold retain their old expressions and order. Both prospective Field-summary parent-cost callers remain in the consumer set. No model equality, actual Field enumeration, XML read, attribute/capture, Source debit or resource check is skipped. The wide/shared-root case and independent weighted/table/Field controls exercise this boundary rather than asserting only a helper constant.

The storage tradeoff is disclosed correctly. The moved source Vec can have spare capacity, whereas the former Finish clone was length-sized. Large reply peak live Finish capacity changes from147 to 272 slots. Large Document picture changes 149 to 272. Those are temporary vector capacities, not RSS, retained-cache payload or universal peak savings. The exact source child capacities remain included in the graph observation and final memory calibration remains due.

## Paragraph ownership, private state and lifetime

The concrete private storage at `crates/rdocx/src/document.rs:9202` has the two actual forms required by contract20: an inline Owned reader result and a Model root/selector loan. The narrow large-enum allowance preserves the explicit no-extra-allocation miss contract. There is no Clone, PartialEq, DerefMut, mutable extraction or arbitrary-root public constructor. Its immutable Deref resolves only the selected leaf through checked access. It never compares or clones a complete table or its private siblings.

The Model constructor at `crates/rdocx/src/document.rs:9409` runs only after the unchanged candidate selector resolves, its complete key matches and a live Weak upgrade proves Arc::ptr_eq. Its Arc clone occurs after the metadata lock has dropped. The private immutable root and selector cannot change after construction, so the Deref invariant assertion at 9224 is justified. A concurrent certificate replacement or clear affects later hit opportunities, not the already-owned snapshot. No raw pointer or self-reference is retained.

Every miss still executes the actual fixed parser at 9431. The lexical Field guard, both unchanged finite qualifications and complete selected-leaf equality remain before key installation. The actual parsed value, including the first successful qualifying read, is moved into Owned at 9521. Reader errors remain uncached. Key replacement, selected-key removal, different-root eviction, poison recovery and source-child/reset/carry behavior are unchanged. Selected unsupported/private leaves retain their actual reader and identity behavior. A finite table leaf loan may own a root containing private siblings without reading, comparing or cloning those siblings.

The only product storage change in comments is the private immutable tuple at `crates/rdocx/src/comments.rs:3074`. Source ownership, unknown-parent priority, complete paragraph traversal and actual source-marker extraction at 3085 through 3106 remain before value consumption. Marked paragraph projections still use actual source XML. Other immutable text/run projections read the selected value through Deref. Reply explicitly drops CommentAnchorRead at 2034 before identifier reservation and anchored mutation. Snapshot listing at 3525 keeps its staged source and read through projection and returns only owned public result data. No binding result, Document or dormant certificate retains a returned loan.

The new lifetime tests distinguish actual first Owned values from later Model loans. They retain an old loan through replacement, cache clearing and factory/document drop, then prove the old root expires after its final loan drops. Two table selectors share one existing root and release it only after the last selector loan. Existing actual source replacement, COW, independent equal root, malformed/full-key, incidental hint, private selected/sibling, concurrent roots/selectors, poison and atomic refusal controls remain. Test adaptations explicitly compare selected CT_P values or references and do not add an owned cloning adapter.

## Evidence examined

The pinned replayed-owned predecessor build exits0. Its actual retained binary cfd8a60cf1e81c1345407d34770fd59baa707f652782f40e0be8730e37335986 executes the same discriminator and exits101. Complete selected values, one actual first read and one exact hit pass before the old owned-clone slot fails 1 against 0. The raw assertion, source and binary are retained in loan-owned-red-authenticated. This is an executed copy-work RED, not a missing-type compile error or a timing claim.

The final current-source observations cover 85 unique Native cases, five OXML cases and 74 fresh ordinary installed public cases. The latter comprise 64 and 10 disjoint JUnit successes, with zero failures, errors or skips. The Native union is 71 final risk cases plus 14 applicable current loan cases. The broad loan filter also encountered one preexisting ignored manual maximal-clone diagnostic outside these required cases. It is explicitly excluded from the 85 and is not counted as a pass. Existing first-cause, related malformed-source, Field, owner-scope, rollback and publication controls remain represented.

The first actual 8-state comparison binds 18/19 before and after. Every input, semantic/reopen result, whole ZIP, member hash and staging-work vector matches. Across those eight states, its total 7,691,845 normalized graph records compare exactly after zeroing only the declared Finish creation/fold capacity column. Source enumeration capacity and every other event, identity ordinal, cost, table depth, child index and call boundary remain exact. The comparison does not mask arbitrary capacity differences or omit traversal tails.

The separate 8-state ownership comparison binds the composed 18/19 owned-return source to 20. Every input, result, ZIP/member, text metric and staging vector matches again. Across those eight states, all 7,691,845 total graph records compare exactly, including every capacity field. These are two comparisons of eight states, not sixteen unique workload states. The comparison hashes are 8a62b0c1c7b3d608607ae510b331953a525a36c8dd53889b91221850e2a30fdf and dfeacea593fce43fea764602f6f78043a4852b9ef9affa786efe49125ef3b790 respectively.

The actual large cold query removes 6001 temporary run Strings and 137345 intermediate push bytes. Large reply removes 20682 positive duplicate child vectors and large Document picture removes 48668. The ownership comparison removes 154 actual large-reply hit CT_P clones, including 601 run values and 540 RPr boxes, and observes 154 real model loans. No loan remains live on public return. These are source-level debug operations and selected payload counts. They are not optimized allocation totals, full copied-model bytes, elapsed improvement or final memory acceptance.

On the measured host, actual CT_P, value-owner and storage sizes are each 2448 bytes with alignment 8, and the actual comment tuple remains 2624 bytes with alignment 8. The before measurement uses the real old CT_P representation and the after measurement uses the actual owner/storage types. This is observed layout equality on that configuration, not an assumed portable layout or a free table-root lifetime.

All-target checks, strict affected OXML and Native/binding all-target/all-feature Clippy with warnings denied, formatting, prose, adapter drift and diff checks have successful recorded outcomes. The zero-selected first text filter, early text fixture premises, five missing explicit test borrows, overwritten initial cold seed, incomplete initial graph observer and the extra ignored manual diagnostic remain disclosed. The complete later source-bound observations supply coverage without relabeling those earlier attempts. No unchanged product test was rerun for this review.

## Structure, limits and handback

Correctness, contract, panic/index invariants, OOXML preservation, ownership/private Field behavior, lifetime/concurrency, test discrimination, structure and evidence applicability produced no finding in this bounded increment. The owner is substantive two-form storage with an immutable-selection invariant, not a forwarding-only wrapper. No new trait, generic, module, source file, feature, dependency, stable API, Source permission or persistent cache is introduced. The adopted plan sections and HLD03/04/10 describe the implemented temporary storage and ownership costs.

All four latest ordinary elapsed misses remain open. A fresh ordinary release artifact and a newly admitted quiet window under the unchanged full performance protocol are still required. Final memory omission/equivalence/calibration/RSS/history, Office, whole installed typing, hosted/archive and full-union gates remain pending. The review does not complete an F-ID, reset a cap, alter a threshold/baseline, update a ledger or authorize publication.

Only this review file was written. This pass performed no compilation, product import, test, benchmark, timing, native app interaction, source repair, Git/network mutation or status change. Control returns to Root for its separate final assessment.
