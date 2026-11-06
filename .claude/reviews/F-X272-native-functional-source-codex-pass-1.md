# F-X272, native functional source, pass 1

**Reviewed**: The sealed connected functional increment in `native-template-support/native-functional-source-batch`, not the moving worker or the whole F-X272 feature. The patch has 2,012 lines across eight before/after paths, including HLD03/04/10 and the worker progress snapshot. The approved F-X272 source, current-record, clone-lineage, checked-copy and concrete writer contracts govern this pass.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. The connected replay path requires remediation before bounded acceptance.

## Defects

### D1, namespace replay drops mapped Word Text facts at entity references

`crates/rdocx/src/document.rs:4217`
`crates/rdocx/src/document.rs:4235`
`crates/rdocx/src/document.rs:4482`

The new recipe sink classifies each write by whether its borrowed slice lies inside the actual input or a selected replacement event. Other writes become Generated. Quick XML 0.41 emits a `GeneralRef` as three writes: the static `&`, the borrowed reference name, and the static `;`. The replay loop forwards that event to this writer. Consequently, unchanged `A&amp;B` becomes copied `A`, generated `&`, copied `amp`, generated `;`, copied `B` in the source recipe.

This preserves the bytes but loses the existing Text-input fact. `crates/rdocx-oxml/src/document.rs:684` retains a Word event on slicing only when its complete encoded range lies in that slice. An actual model Text event covering `A&amp;B` lies in none of those fragments, so it is discarded during composition. `crates/rdocx-oxml/src/document.rs:5267` then cannot select that event at the grammar read. Its raw-range fallback also fails at `crates/rdocx-oxml/src/document.rs:968` because the entity delimiters have no Word image. An original/current raw Text interval spanning the entity has the same generated-gap problem.

This affects a normal mapped Text payload whenever namespace-changing replay is reached. The new paragraph/row grammar consumers rely on those facts to select actual input before charged copies. A later charged append or row-view copy can therefore refuse a legitimate input, even though the replay bytes match the ordinary output. This is a source-derived finding. No new reviewer runtime outcome is claimed.

The new test at `crates/rdocx/src/document.rs:84902` contains `A&amp;B`, but starts from an entirely Generated source and asserts replay bytes and repeat behavior. It cannot distinguish preservation of a real Text event or selected Word origin from the loss described above.

Preserve the actual complete encoded payload correspondence through this same replay, including the reference delimiters, without inferring origin from equal bytes or granting input to genuinely generated namespace syntax. Add a mapped-input replay control that reaches the actual Text selector and a charged consumer, plus the existing Generated/no-credit and known-Field refusal controls. Keep ordinary output and error behavior unchanged.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Connected source and caller assessment

The staged main writer retains actual emitted Sources beside weak handles to its concrete block records. A reused block can recover a preexisting ingress-frozen immutable record through the actual selected storage handle. The frozen catalog holds the real block or opening storage before cloning and does not populate itself from later emitted records. The two previously failing ordinary Native fixtures now exercise that restored normal path. This is CURRENT SOURCE eligibility, not an Original Field or JSON baseline.

The one-shot original-part selection is restricted to the initial actual clone route and the branch retaining unchanged typed document bytes. Later preparation and refreshed parts do not renew that role. Whole-model canonical reparses invalidate clone-path eligibility. Supported direct clone slots and operation-minted prepared Field roots reach the existing substantive model writers. The new lower prepared-root lookup reuses the operation anchor and registry key, while existing scalar and shared raw-record guards remain separate.

The notes-family w14 change reuses the already approved concrete source insertion producer and shared declaration planner. The scoped row wrapper uses checked composition before the same existing whole-table admission and parser. Direct/all view String copies reserve selected text-copy work before cloning. The public ordinary paths retain their existing optional-context behavior. No new public callable is implemented by this batch.

Namespace replay uses the same substantive event writer with an ordinary Vec sink and a scoped recipe sink. Checked composition materializes the latter output before its separately reserved legacy record. The new target-start branch preserves actual start-record slices and appends only its missing declaration bytes. Its error and conflict checks remain in the existing replay loop. D1 concerns the distinct unmodified entity event path through that sink.

## Source and evidence authentication

All eight before snapshots match `BEFORE.json`, whose hash matches the boundary manifest. All eight final snapshots match `SOURCE_BOUNDARY.json`. Reconstructing the complete unified patch from those pairs exactly reproduces `increment.patch`, with SHA256 `cefd987818d6e300028900552218f9965b0e3e650302c456f649e53fb578891f`.

All 60 individually pinned evidence files match their recorded hashes, including both archived executables. The final lower receipt and raw logs identify 89 distinct passing tests across disjoint filters. The final Native log identifies 52 distinct passing tests. The strict final Clippy child exit is 0 with a completed 16.44-second result. The four final metadata rows record formatting, explicit prose, skill drift and diff checks at exit 0. These records were inspected without rebuilding or running their executables.

The retained 50-pass, two-failure Native result is correctly classified as an introduced regression, followed by the 52-pass successor. The three publication fixture failures also occur on the retained predecessor and are not attributed to this batch. The reconciled tests now distinguish classification from actual Field selection before paid copy, and preserve the zero-input unpaid-copy refusal. Earlier fixture, compiler and lint attempts remain separately labeled. The final lower executable covers the final two assertion spellings, while the Native result is qualified to identical production source rather than relabeled as a different whole-file build.

The quick-xml entity behavior supporting D1 was inspected in the pinned local 0.41.0 reader and writer sources. The writer handles `GeneralRef` with `write_wrapped(b"&", &e, b";")`. The actual model parser retains the complete `read_text` encoded range, and the actual selector requires a whole matching event or contiguous raw Word image. This establishes the missing join without a new parser or probe.

## Explicit exclusions and remaining work

`PENDING_HANDOFFS.json` remains operative. Returned-reader Text/raw companions, optimized and full canonical reparse handbacks, related notes/glossary refresh, private stored-control graphs, Revision/Ruby contexts, refreshed textbox/import sources and the canonical-neighbor join remain pending. The proposed returned-read methods and note-reader methods are not implementation accepted here. Partial maps do not certify ordinary gaps or relevant-consumer completeness.

This pass does not qualify comments authoring, full-suite execution, native applications, oracles, hashes or baselines, packaging, final SemVer, provider acceptance or the integrated union. It adds no security investigation, quota, conservative fallback refusal or new product scope.

## Not found

No additional correctness, panic, OOXML namespace/schema-order, preservation or structural-rule finding was found in the sealed delta. No independent smell or nitpick was found. The target-replay input-selection test gap is part of D1, not a duplicate finding.
