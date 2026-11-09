# F-X281, implementation, pass 1

**Reviewed**: Frozen dirty increment on `work/f-x281-codex` relative to continuation HEAD `514a076e592517ef128e555f52cef1085fa702de`, 30 paths, 5208 insertions and 368 deletions. Existing progress is the 31st manifest entry. Previously reviewed F283 is part of that HEAD, not a new F281 change. Original claim Base is `4a220b6769fc2a5a2c90fc600f580478ea7e62f8`.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. This is actual implementation peer 1 of 3. It is not Root final review, preparation, integration or final acceptance.

## Defects

None found.

## Smells

None found.

## Nitpicks

None found.

## Authentication and scope

Read the repository guidance, workflow, microscope command, approved operative plan, progress and cited HLD. The sprint records retain F-X281 as in progress (`docs/sprints/CURRENT_SPRINT.md:31`, `docs/sprints/BACKLOG.md:852`). Inspected the complete changed production graph, test carriers and four HLD changes. No product, test, build, oracle or network execution occurred in this pass. Only this review was written after the author STOP and exclusive report lease.

Independently authenticated the Phase B handoff SHA256 `11df71cf729e085a1422becee6e6d58f33595a5f7d869c43ba91c9532df27add`, source manifest `b70b3e2f6d8760f636fc47bedaf30bc916c94630e7f6e5149d79a014876c9513`, all 31 listed file hashes, HEAD and ordinary diff `0a0c1d5c55edb0fba29c60ee1d252fb8fab7bb39cd41a73ebfc92dc85c2dcc82`. The plan patch, applicability, artifact and receipt manifest hashes match. All Phase B listed retained artifact hashes and receipt/log pairs match their manifests. Phase A handoff, manifest, artifact manifest and receipt manifest match their declared hashes, and its seven listed retained copies authenticate. Nine unchanged low-level source hashes match the final source manifest. Phase A layout source reuse is narrower because Phase B adds the chart diagnostic caller.

## Not found

| Area | Source and consumer inspection | Result |
|---|---|---|
| A281-1, image classification and EMF | `crates/oxml-pdf/src/image.rs:31`, `crates/oxml-pdf/src/image.rs:406`, `crates/oxml-pdf/src/image.rs:576`, `crates/oxml-pdf/src/image.rs:1067`, `crates/oxml-pdf/src/image.rs:1355`, `crates/oxml-media/src/lib.rs:339` | No defect found in sniff precedence, inclusive bounds/frame fallback, integer scale, checked source/output admission, negative RestoreDC, map modes, transforms, solid object selection, paths, current position, DIB depths/masks/opaque RGB, crop/sign/orientation or source/pattern/destination ROP dependencies. Supported paint and unsupported-only outcomes remain distinct. |
| Picture diagnostic custody and routes | `crates/rdocx-layout/src/input.rs:126`, `crates/rdocx-layout/src/engine.rs:12331`, `crates/rdocx-layout/src/engine.rs:14157`, `crates/rdocx-layout/src/engine.rs:14988`, `crates/rdocx-layout/src/engine.rs:15373`, `crates/rdocx/src/document.rs:49907`, `crates/rpptx/src/lib.rs:29343` | No defect found in physical relationship/story identity, scoped lookup, transaction replacement, VML cache snapshot/replay/accounting, missing VML treatment, orphan exclusion or deduplication. All seven fresh Word facade publishers annotate before return/cache. Presentation failures retain the owning relationship. |
| A281-2, Adam7 | `crates/oxml-pdf/src/image.rs:96`, `crates/oxml-pdf/src/image.rs:351`, `crates/rpptx/src/lib.rs:29431` | No defect found in seven-pass geometry, empty-pass handling, per-pass filter history, scattering, existing channel expansion or preview pass-byte accounting. Existing plain-PNG compatibility and named preview residuals are retained. |
| A281-3, raw extra axes | `crates/oxml-chart/src/lib.rs:10515`, `crates/oxml-chart/src/lib.rs:10756`, `crates/oxml-chart/src/lib.rs:10947`, `crates/oxml-chart/src/lib.rs:19001` | No defect found in first-two typed ownership or later raw arrival order for bar, line and area, including repeated declared ids. Other plot-family refusal rules and unknown-sibling handling remain separate. Preservation is not described as schema conformity or rendering a third axis. |
| A281-4, numeric labels | `crates/oxml-chart/src/lib.rs:2337`, `crates/oxml-chart/src/lib.rs:2356`, `crates/oxml-chart/src/lib.rs:9017`, `crates/oxml-chart/src/lib.rs:9065`, `crates/oxml-chart/src/lib.rs:9240` | No defect found in the approved section/token grammar, exact General exemption, literal/grouping distinction, reserved refusals, finite-only painter fallback or deduplicated format diagnostics. Non-finite errors remain errors, and stored format spellings are not rewritten. |
| A281-5, chart semantics and public APIs | `crates/oxml-chart/src/lib.rs:32`, `crates/rdocx-html/src/lib.rs:51`, `crates/rdocx-html/src/lib.rs:138`, `crates/rdocx-html/src/emitter.rs:647`, `crates/rdocx-html/src/markdown.rs:1221`, `crates/rdocx/src/document.rs:46606`, `crates/rpptx/src/html.rs:704`, `crates/rdocx-cli/src/commands.rs:894` | No defect found in complete-byte title qualification, alias/shadow/path controls, empty-rich precedence, relationship loading, direct/selected AlternateContent routing, sparse/non-finite cells, finite Display values, escaping or diagnostic separation. Word tuples, existing Python strings, save/CLI callers, unchanged HtmlInput fields and empty-map legacy/EPUB paths match the contract. |
| A281-6, hidden paint | `crates/rpptx-oxml/src/namespace.rs:183`, `crates/rpptx-oxml/src/shape_tree.rs:54`, `crates/rpptx-layout/src/context.rs:4097`, `crates/rpptx/src/html.rs:304` | No defect found in qualified physical-owner dispatch, foreign/local-shadow controls, group pruning, later visible siblings or selected chart-choice handling. Slide/layout/master and notes/handout surfaces share the walker. Semantic text and chart traversal remain present. |
| A281-7, ActualText | `crates/oxml-pdf/src/writer.rs:1252`, `crates/oxml-pdf/src/writer.rs:1393`, `crates/oxml-pdf/src/writer.rs:3899` | No defect found in own-span text, space spans, reordered paint, inherited marked-content recursion, unchanged glyph matrices or unchanged ToUnicode. Legacy Text remains outside the change. |
| A281-8, master family | `crates/rpptx-layout/src/context.rs:4407`, `crates/rpptx-layout/src/context.rs:10081` | No defect found in idx-free equivalent-family qualification followed by actual idx, exact type, Body, tree order and old idx fallback. Transform/body/list consumers share the chain. Layout lookup, occupancy and picture-placeholder behavior retain their separate policies. |

No additional panic/overflow defect, child-order regression, public API/default change, ownership/retirement defect or structural abstraction problem was found. Checked record/bitmap slices and admission precede the relevant allocations. The private owner recorder retains Send and Sync. No new source module, file, crate, trait, generic, feature or dependency is introduced. HLD05/07/08/10 describe current behavior and the explicit residuals (`docs/hld/05-drawingml-model.md:502`, `docs/hld/07-inheritance-and-resolution.md:387`, `docs/hld/08-rendering-spec.md:2751`, `docs/hld/10-bindings-spec.md:4684`).

## Evidence and remaining limits

Read actual final native log outcomes, not only totals: Word 5, owner custody 1 and presentation 4 pass. HTML injection has 8 passing cases, including the new escaping/projection case. Each CLI has its actual chart export case passing. The local Python receipt has six passing cases on refreshed debug CPython 3.14.7 modules: Word EMF 1800 red/0 other, presentation explicit and width-only EMF 4050/0, same-run plain/Adam7 equality, hidden-versus-removed PNG/PDF equality across slide/notes/handout and Two Content tops 109/109. Their assertions are in `crates/rdocx-py/tests/test_core.py:16240` and `crates/rpptx-py/tests/test_documented_examples.py:12648`. These are local extension observations, not installed-wheel, release or native Office qualification.

The final header/footer correction at `crates/rdocx-layout/src/engine.rs:20385` selects either valid Text or MultilingualText, grows charged advance capacity while preserving rich-span validity, and keeps the original oversized rejection and empty-cache assertions. The final four-case receipt passes and records capacity 524297. The scoped 514a layout control fails at the same old Text-only premise. This establishes a fixture problem in that scoped control, not a whole-Base build or a product RED. The removed Word width-only premise likewise is not relabeled a product failure.

Phase A current image 21 and chart 6 observations, unchanged writer tests, layout/XML compatibility observations and public consumer evidence have their stated applicability. Earlier original RED executable bytes were overwritten before retention. Their recorded execution hashes, source/input snapshots and logs remain, but this review does not claim an independent rehash of those lost original binaries. Later retained RED copies are separate. Setup, link/filter errors and failed aggregate commands remain failures rather than being counted as successful semantic evidence.

All eight acceptance rows remain required at `.claude/plans/F-X281-design.md:274`. Local source/unit controls do not discharge every public issue fixture, native appearance or extraction outcome. Final Office, pinned LibreOffice, Poppler bbox, schema, full installed families and strict typing/stubs/API, MSRV/WASM/docs, deterministic hash/pixel corpus, archive and nine-crate publish dry-run/size, hosted CI and integrated-union gates remain pending. A moved pdf/pages, pdf/resources or pdf/bytes key still stops the gate without baseline update (`.claude/plans/F-X281-design.md:383`). Explicit bounded EMF approximations and export/preview/master residuals are not promoted to passed features or manufactured as defects.

Report complete. Return the lease to Root. No remediation or confirmation pass was performed.
