# S96 admission frame and context, implementation support 9

**Reviewed**: Final composed frame12and context13increment against clean Root26/support8,3paths,474added and41removed lines. Canonical workflow/microscope, adopted plan amendments, HLD04, investigations12/13, full source changes and actual retained evidence were read. Unchanged integrated product paths consume the preceding applicable source reviews. This remains S96/F-X275 final-rider remediation, with no historical F-X268 lifecycle or pass-count reset.

**Verdict**:0defects,0smells,0nitpicks. Source and scoped correctness/mechanism evidence are clean for this finite union. No performance or final-rider acceptance is granted.

Only this report was written. No product edits, compilers, product imports, tests, instrumentation, benchmarks or confirmation executions occurred. Root remains sole product/Cargo writer.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Frozen boundary and contract

The final packet is `.claude/scratch/S96-final-evidence/f268-admission-context-weight-implementation-20261108-1`. FINAL_UNION_SOURCE_BOUNDARY.json:1 independently hashes to c55e5ba1e5bb0e3f2ff8304f1060489307d9d586439525866509a0855c574afd. FINAL_UNION_INCREMENT.patch independently hashes to86f32646619d8c5a86a08c6593caf3c7edec8dc3c844793b1a85c9aadc693e47. The context-only INCREMENT.patch separately matches eeaa86b0d35d153fdc8775785949770229a25cd94bdfc7d9c921a4f2bb45c6db. These are distinct component and union boundaries.

All4309final manifest entries match actual current files. Comparing the retained frame packet's Root26manifest with this final manifest gives exactly the3declared paths, leaving4306unchanged. All3Root26before copies, final-source copies and actual current files match their respective recorded hashes. The runtime document source is4c0e6adcf1f59ba28e23fa58dcb88bdae390ca1edf9c60f364195ebd67d09df7. Native/binding/Python product, test, configuration and protected union inputs outside these paths remain exact.

The approved frame contract is at .claude/plans/F-X268-design.md:1039and context contract at1057. They record actual predecessor work, unchanged ordinary failures and explicit composition before the next ordinary qualification. docs/hld/04-opc-and-packaging.md:5and7describe the implemented current intent without claiming performance acceptance. No ordinary frame12timing was performed and relabelled as a retry.

## Runtime source assessment

WordAdmissionFrame.local at `crates/rdocx-oxml/src/document.rs:1161` is the sole changed field representation. At the sole existing Start constructor, document.rs:9675selects the identical consumed-or-local bytes before matching. Exact p/r/rPr/t arms borrow static literal slices. Every other selected byte sequence owns value.to_vec. The debug consumed-owner precondition remains at9140. Empty events do not construct a frame. No reader-buffer/caller borrow escapes, no Unicode normalization is added and namespace/admission classification stays separate from local spelling.

Every frame-local read uses the original exact byte comparison or slice pattern with mechanical as_ref conversion. Parent dispatch, MC/VML/drawing admission, textbox identification and ordered-field backward charging preserve Boolean grouping and short-circuit order. The type remains private and gains no lifetime parameter, generic, forwarding helper, trait or public API. Both static and owned variants have actual consumers. Its stack representation is not assumed smaller.

The namespace map and copy_scope_weight at document.rs:9159remain call-local. The only actual map writes are the original attribute insert at9214and insert/remove arms in restore_namespace_depths at1170. URI bytes are owned immutable Vec values, with no mutable escape or in-place URI change. All such writes invalidate the scalar, including equal/redundant/depth-only replacements. Attribute invalidation at9213follows the original attribute/reserved-binding checks. Restoration invalidates before each actual undo write at1176. Empty undo does not write. PartWrapper undo clearing at9225still retains its already inserted declarations rather than restoring an old scalar.

The fold remains at its original unconditional Start/Empty position, now document.rs:9536. Some(bytes), including Some(0), reuses only a previously successful value. None executes the original iter/try_fold, raw URI lengths, decoded Word marker, checked-add order and Resource error conversion at9551. Some is assigned only after successful `?` at9567. No error or URI result is cached. No eager fold, incremental delta, unchecked/saturating aggregate or pre-normalized map replaces the original calculation.

The scalar's operands depend only on current prefix/raw URI bytes and the pure decoded_namespace_is projection. They have no actual XML reader, Source cursor, receipt, callback or mutable external policy. The depth-sensitive captured_namespace_is behavior remains untouched. Removing repeated pure projections neither grants Source authority nor changes actual namespace validation. False projection of invalid data remains the predecessor's false result, not a newly cached Input failure.

The existing next_copy_depth and Empty visits/nodes/bytes checks remain after the fold at9601. Start still stores that entry-scope value in its frame. End still charges frame.copy_scope_bytes before its undo restoration at9739. Namespace and owner/control/table guards, final duplicate-context arithmetic, actual capture/constructor/event/attribute reads, source-map offsets and fallible input length keep their original owners/order. A failed fold returns at the same point without installation. A warm scalar cannot suppress an earlier attribute or resource error.

The new permanent traces and reference fold are cfg(test) only at document.rs:1124and9570. The reference executes only when that test explicitly enables its trace. Ordinary Maturin builds do not compile cfg(test). Existing private Field XML work, Source accounting, constructor facts and error-cache behavior are unchanged. No AST, map snapshot, historical state, Source credit, global cache or retained key is introduced by the runtime changes. The scalar dies at the end of one classified admission.

## Actual tests and failure custody

The real frame test at document.rs:14518compares actual stored names at every push/pop after buffer growth/reuse, against independent event traversal and original virtual selection. It covers mixed Start/Empty, actual/synthetic roots, common/unknown/Unicode/case/prefix-sharing names, foreign/default/alias/shadow contexts and consumed p/r. It also compares actual complete outcomes and piecewise input. It rejects storing event-buffer slices or selecting the wrong virtual name.

The context test at14359executes actual zero/many/alias/default/encoded/foreign scopes, different and redundant shadows, Empty declarations, new-prefix removal, nested restores, invalid/duplicate/reserved attributes, malformed close/EOF and actual Comments PartWrapper facts. Its enabled cfgtest reference verifies the actual current-map scalar at every traced fold point. Per-call B/I/F/H/V tracing verifies valid reuse and cold folds across actual mutation epochs, including numerically unchanged writes. The test's predecessor path intentionally has no H branch. A successor that never reused would still be rejected by the separate actual public mechanism requirements, rather than being accepted because an epoch assertion was inapplicable.

Independent parsing found118unique complete successor tagged outcomes, exactly equal to the actual103retained frame12predecessor outcomes plus15fresh actual pre-context outcomes. The categories are84classified,12remaining-reader,7frame and15namespace outcomes. Separate stderr/stdout capture preserves complete result lines. Equality includes full work, constructor/control facts, projected aliases, complete source maps and typed underlying Input/Resource errors. It is not comparison of only success flags or error text. The old103execution is explicitly reused at its actual source identity, not relabelled as a fresh execution.

The successor local libtest summary reports19passed,0failed,0ignored. NATIVE_SCOPED.json:1 contains271additional distinct names. Parsing all271actual libtest receipts in native-tests.log gives271passed,0failed,0ignored. Their names are disjoint from the19already-passed admission names, yielding290distinct Rust controls. The retained binary independently hashes to5e7ef1210a9c972aabdf707a71a41a635bbba7809b3fdd90ed48e984f5174ce9. These are actual semantic passes, not child0or ignored-test counts. Relevant namespace/Source/Field/resource/remaining-reader risk selection remains explicit.

Independent JUnit parsing gives64distinct core installed cases and10distinct picture held/fresh/refusal/reopen companion cases, disjoint and totaling74, with0failures,0errors and0skips. The companions execute separately and do not insert additional getters before the exact original mechanism actions. Their runner selects the authenticated installed Native and current test source. No private diagnostic artifact substitutes for the uninstrumented public control run.

Retained failures remain honestly classified. The frame component LOCAL_COMPARISON_FAILED_1.json:1 records the wrapper's incorrect84-versus96tag assumption and its combined-shell continuation. LOCAL_COMPARISON_FAILED_2.json:1 preserves mixed stdout/stderr corruption and the failed parser. Its corrected serial separate streams use the same retained binary, without counting repeated captures as additional unique tests. TRACE_TYPE_ALIAS_APPLICABILITY.json:1 preserves the initial strict type-complexity failure and subsequent cfgtest alias correction. Direct source comparison confirms the only producer difference is that adjacent cfgtest type alias and its equivalent annotations. The current final artifact is freshly built over the exact final source, not relabelled through that older applicability claim.

The context packet NEW_PREPARE_FAILED_1.json:1 preserves the diagnostic text needle also matching the cfgtest reference fold. No failed preparation is called a build or runtime pass. The corrected meter is confined to the actual production cold-fold expression. No failing receipt was deleted, normalized into success or used for elapsed acceptance.

## Artifact and mechanism authentication

The fresh ordinary debug producer records child0and source_unchanged true in ordinary-debug/build-result.json:1. Its full4309producer input manifest still matches current source. Wheel SHA-2568dec0e50bc1ebda9be7dfa6a8c8bcc833b9bd9e18cc7887ec8d6e953aa4d5d93and Nativef5e5f4b2bfa6a741280f1aba43b98d67103ee2b584735122753a8e0ca203c157authenticate. All10wheel members match retained payloads and actual installed files. Default Maturin debug configuration has no test cfg or temporary diagnostic build flags.

Scoped all-target check and strict all-target/all-feature Clippy record child0. Their actual logs match SCOPED_SUMMARY.json:1 hashes and finished successfully. Final fmt/prose/adapters/diff receipts each report0. No product changes after those inputs are hidden behind metadata applicability.

Both context diagnostic builds record child0and restore all4309inputs exactly. Their10payload members and complete wheels independently authenticate. The old capsule's manifest/copies pin its actual pre-context source, and the new capsule pins the final current source. Both frame component diagnostic wheels and10payloads also independently authenticate. Their source/cfg differences and the test-only alias applicability remain separate from shipping artifacts.

The frame meters sit at actual selected-value Vec construction, the four actual Borrowed arms and the real Owned fallback. Rechecking all16observations confirms positive predecessor copies for each common name, exact replacement by corresponding static views, unchanged other-name count/copied bytes, equal constructed frame counts and actual classified/remaining reader attempts/consumed bytes. Borrowed payload lengths are not copied bytes. Other namespace/buffer/stack/output allocation and possible existing debug-reference work remain outside this narrow meter.

The context meters count actual production fold executions, closure entry visits and pure URI projection calls, then separately actual successful cold results, Some reuses, write attempts and invalidations. Every actual returned scalar is retained in sequence. The cfgtest reference source is guarded out of these Maturin builds. It does not secretly contribute a second fold to public work accounting. Prefix/raw URI payload lengths are inspected inputs, not exact decoder byte visits, heap, copied bytes, RSS or elapsed cost.

Independent comparison verified all16context action pairs: original reply action2, Cell picture action1and Document picture actions1/2, both sizes and two trials. Original setup/prior actions and exact predecessor input ZIPs are retained. Full semantics and reopen assertions precede publication of mechanism results. The API's picture-only paragraph behavior is preserved.

For every pair the complete actual scalar sequence is equal, map-write attempts equal new invalidation operations, old fold count equals new cold count plus reuses, and actual XML attempts/consumed bytes remain equal. Reuses are positive and iterator visits decrease on every measured state. For example large reply goes437396to9702actual entry visits, and large Document picture action2goes1830181to30497. Those are debug mechanism observations, not elapsed attribution or a total-operation complexity claim.

Independent byte comparisons authenticated all16complete context ZIP pairs and every member against their recorded maps/hashes. The16frame component ZIP pairs and all members likewise match. All32pair comparisons preserve exact inputs and full successful semantic/reopen receipts. Source scalar equality and package equality are separate evidence, not interchangeable assertions.

## Not found and pending acceptance

Correctness, contract, namespace/OOXML preservation, panic/overflow changes, successful-only storage, invalidation completeness, first-cause/error behavior, actual Source/private Field boundaries, consumer coverage, structure, artifact custody and semantic counts produced no findings.

This clean source/scoped review permits returning control to Root for the next workflow step. The last valid ordinary c2383cqualification still has77/79flat passes with reply and Document.add_picture failed. Its5linear and both407gates remain passed. There has been no ordinary frame12timing run or favorable retry hidden by composing the adopted increments.

A fresh normal final-source artifact and compiler-free unchanged complete performance protocol remain required. Final memory/lifetime/calibration, native Office, whole installed, typing, hosted, archive and full-union verification/review remain due. No acceptance row, threshold, workload, fixture, font, clock, aggregation, Source debit or memory policy is waived. No status, ledger, baseline, commit, push, tag or publication action was performed. This pass ends with this report.
