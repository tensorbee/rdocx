# F-X273, all, pass 1

**Reviewed**: Full working result against immutable claim Base `564ad7f5d1d57a7828c721eba58cbe79da10aae3`, including committed variant refactor HEAD `3510ec5ee2c12f8aee9bc52e4434624532f160d8` and dirty successor. Branch `work/f-x273-codex`. Ten tracked files,3150 additions and269 deletions.

**Verdict**: 2 defects,0 smells,0 nitpicks. Local integration/provider readiness is blocked by D1/D2. Final native parity is not recorded as passed.

## Frozen identity and review boundary

Actual `git diff Base --binary` SHA256 is `6b045669366589cba5375f08f4bf4b0ee379f1917eab5a2675061c055cd1069e`, matching the frozen diff. Independently recomputed hashes match every changed-file and progress-file entry in `fx273-local-ready-identity.json`. Product227 manifest is `be5dff87779bb572e84a861abd786058fc74d4ec9c98a1a26f7ef3e23f936607`. Actor/root stopped writes/builds before review. Only this review file is written. No tests, builds, product/native oracle invocation or remediation occurred. The initial default-sandbox report write was refused without changing a file, then the authorized report write used scoped escalation.

Read AGENTS, CLAUDE, WORKFLOW, microscope, approved canonical plan through81cc44a0, worker annotations/progress, relevant HLD03/04/08/10/12 intent, changed production paths and regression carriers. The scope includes the committed refactor. Approved doc-hidden omission query and fallible bounded furniture preparation are part of the contract. Separate geometry/navigation commit labels remain a preparation action after review.

## Defects

### D1, Furniture floating tables bypass existing non-overlap placement

`crates/rdocx-layout/src/paginator.rs:477`

The new furniture publisher calls `floating_wrap` for each floating table and immediately records its rectangle and keep-out band at lines478 to493. It never uses `FloatingTable.overlap_allowed` or carries already settled tables. Two header/footer tables with the same page-relative origin and intersecting dimensions therefore keep identical rectangles even if either has `w:tblOverlap w:val="never"`. All rows paint from these unresolved positions through `render_hf_blocks` at `crates/rdocx-layout/src/paginator.rs:6055`. Repeated preparation cannot fix identical unresolved rectangles.

This is source-derived, not executed semantic RED. Concrete existing authority: `crates/rdocx-layout/src/table.rs:154` carries the policy, line213 derives it from actual properties, and `crates/rdocx-layout/src/paginator.rs:1136` settles a later float below an earlier non-overlappable float. Body publication invokes it at line2495. HLD08 specifies this within-page rule at `docs/hld/08-rendering-spec.md:1011`. Approved F273 scope traces floating positions through existing owners at `.claude/plans/F-X273-design.md:25`, with ordinary floating-table A3 at line63. This does not request facing-page or section-scoped resolution.

The furniture wrap control inserts one floating table at a time at `crates/rdocx-layout/src/engine.rs:28527`. Its successful assertions cannot discriminate two competing publishers.

Smallest correction: settle each physical furniture table once, in story order, against already settled floating tables for that selected story/page. Reuse the whole-table rectangle for every row, reservation and paint. Do not settle each row separately or mix drawings into the table-only overlap relation. Keep original cache/source blocks immutable. Obtain actual RED for coincident tables where either forbids overlap, then GREEN with both-allows/disjoint controls, multirow ownership, selected header/footer and matching flow/paint. Native positions remain final debt.

### D2, Repeated preparation can certify overlapping chained clearance

`crates/rdocx-layout/src/paginator.rs:4264`

Top-and-bottom clearance tests every obstacle against the original `paragraph_top`. Lines4266 to4269 update an offset only for obstacles starting at or above that original origin plus1pt. Both internal passes repeat the same predicate. The adjusted paragraph returns without checking its final displaced line origin at `crates/rdocx-layout/src/paginator.rs:4349`.

Concrete source-derived trigger: two distinct page-anchored ordinary frames using `wrap="notBeside"`, followed by ordinary header flow. With header origin18pt, frame A spans18 to78pt and B spans70 to130pt. Different y values keep them separate groups. A moves flow to78pt. B is ignored because70 exceeds the original18+1, yet text now starts inside B. These are actual TopAndBottom reservations at `crates/rdocx-layout/src/paginator.rs:462`. This uses the provider's own geometry and does not assume a native position.

The new outer check misses it. Each candidate reflows the original paragraph at `crates/rdocx-layout/src/paginator.rs:579` through the same two-pass helper at line586. Fixed page-relative obstacles stay fixed. Identical incorrectly displaced output repeats, so `state == previous.state()` returns success at line604. Deterministic output from incomplete reflow does not prove final reservations satisfied. It affects fixed-origin headers and can affect footer preparation. D1 is separate, since frames do not participate in floating-table overlap policy.

The canonical requirement includes stable line breaks and obstacle reservations at `.claude/plans/F-X273-design.md:176`, forbidding guessed paint on failure at line178. Existing body pagination's documented two-pass approximation grants no stronger furniture consistency proof. The current one-frame notBeside test at `crates/rdocx-layout/src/engine.rs:28522` proves only one displacement. Footer and cycle controls at lines28556 and28610 do not test another obstacle reached after displacement.

Smallest correction: establish clearance against the final moved text band and validate final reservations/line state before accepting prepared furniture. Clearance can transitively reach another obstacle, so one special extra check is insufficient. Keep original-source rebreak authority and bounded cycle/failure handling. Preserve the separately approved legacy body approximation/error policy. Obtain actual chained-frame RED then GREEN, including reversed obstacle order, touching/disjoint controls and both selected furniture sides. Also challenge narrowing that grows text into a later obstacle, since reservation calculation at `crates/rdocx-layout/src/paginator.rs:4290` precedes the final line break. Required supported cases must succeed. Real cycles/failures must still refuse with section/page context and preserve source/cache. No executed failure for these new inputs is claimed.

## Smells

0 findings.

## Nitpicks

0 findings.

## Full aspects and consumer coverage

Correctness/contract findings are D1/D2. No additional findings in signed/zero margins, missing/blank selected variants, physical/displayed page identity, written/effective geometry, ordinary gaps/borders, table flow pricing, navigation or API scope.

Panics:0 additional findings. Reviewed frame/table indexing and expects through private paired construction/grouping/preparation, shared failures and continuation. New indexes arise from internal block/semantics inventory, not public arrays.

OOXML:0 findings. No parser/model writer expansion. Geometry reads preserved properties without rewriting imported frame/drop-cap state. Existing TOC staging, source qualification and atomic publication remain authoritative. No generic XML equality or raw preservation bypass is introduced.

Tests: D1/D2 identify uncovered equivalent consumers/combinations. Existing assertions are useful for their executed cases. No additional test defect. Source counterfactuals above are not claimed actual RED.

Structure:0 findings. Private concrete carriers hold actual placement, byte-accounting and failure state. No new trait, generic parameter, crate, module, source file, flag or public authoring operation. The one additive doc-hidden query is approved.

Furniture tracing covers variant selection, source shaping/cache safety, grouping and documented Word defaults, drop-cap exclusion, gaps, preparation, clearance, frames/tables/drawings, page/text borders and body/cell anchors. Prepared bundles otherwise serve clearance/paint at `crates/rdocx-layout/src/paginator.rs:1925`, line1982 and line3040. Retained metadata byte accounting is at `crates/rdocx-layout/src/engine.rs:7595`. Rich tables retain the existing conservative cache gate at line6258. D1/D2 concern incomplete preparation, not absence of bundle reuse.

Body tracing includes fit, columns, line numbering, vertical alignment, rotated fragment displacement and final extraction after flush at `crates/rdocx-layout/src/paginator.rs:1502`. Fresh endnote pages establish upright single-column top flow before effective geometry at line3568, reuse selected furniture and propagate errors. Engine checks failures before returning at `crates/rdocx-layout/src/engine.rs:4616` and line4757. Legacy low-level signatures remain unchanged.

Navigation tracing includes physical registration/remapping, bounded subtree omission, accepted removed rows/whole tables, actual continuation decisions, nested/control paths, cache replay and placed navigation. `crates/rdocx-layout/src/engine.rs:619` marks registered descendants of an actually omitted subtree. Placement publishes at `crates/rdocx-layout/src/table.rs:506`, line750, line1764 and line1895. Cache replay uses returned cell decisions at `crates/rdocx-layout/src/engine.rs:9330`, under existing safe-table qualification. Query identity and None/Some distinctions remain at `crates/rdocx-layout/src/lib.rs:497`. TOC checks omission before every selector at `crates/rdocx/src/field.rs:10003`, retaining missing/stale refusal. STYLEREF and numbered-outline assertions were inspected beyond counts.

## Acceptance and evidence

Receipts are under `/private/tmp/rdocx-s96-run-refresh`. `fx273-root-auth-pre-review.json` authenticates files/logs/binaries with empty mismatch inventories. It is authentication, not a test rerun or native parity claim.

| Row | Local observation and limit |
|---|---|
| A1 | Tall/fitting header, exact five-Normal8pt gap delta, original logo/table, CLI position/PDF production, Python layout. Immutable-input CLI five-Normal top148.4538167pt differs from Base56.65pt. Native positions pending. |
| A2 | First/even/default, selected blank stories, zero/negative margins and tall footer. Same selected variant drives clearance and paint. |
| A3 | Existing flow/floating/frame/default/drop-cap/anchor/border positives remain useful. D1/D2 block full local closure. Word positions remain pending. |
| A4 | Fitting/overflow notes, continuation furniture/PAGE/NUMPAGES/restart27, upright continuation under vertical/center source, columns/line numbers/aligned fragments. D2 affects wrapping combinations. |
| A5 | Merged/unmerged, deleted row, nested/control/orphan and cache cases. Visible neighboring/nested headings survive. |
| A6 | Plain/styled20-context native selector matrix plus nested controls, stale source, STYLEREF and numbered outline. CLI5contexts and Python20contexts retain original controls. |
| A7 | Warm/cold same-owner spacing/text/variant edits, source-cache replay and held Python header owner. Logo fixture is complete before image relationship publication. |
| A8 | Focused integrated composition remains due after F274, with final native/PDF/PNG/text evidence. Not claimed from isolated F273 cases. |

`fx273-final-local.json` binds lower32/native5/CLI2 and check/clippy/fmt/prose/adapters/whitespace success to product manifestbe5dff87. Lower membership retains original cache/reflow/border/note controls. Native cases inspect exact entries, adjacent field consumers and reopened results. CLI PDF signature proves production, not paint/native parity.

Python's historical batch is21pass/1fail, not22pass. The short-header expectation was corrected after actual default8pt after-spacing was traced. Current assertions retain that overflow positive and add explicit zero-after fitting at `crates/rdocx-py/tests/test_core.py:12245`. `fx273-provider-corrected-consumer-auth.json` binds unchanged product, wheel `e48be80e7e666b38341f0abd25ae268f05378d08fec5036534ae4dfbf5a9c175`, actual native `b236a165dadd44330fca9c5058899ea658a66300b51a25d0f35f91d77cffbba8`, unchanged other21ASTs and corrected1pass. Consume applicable21 plus corrected1, not a fictional single22pass command.

Actual Base RED and subsequent geometry/navigation/caller/continuation outcomes remain separately attributed. Seven immutable Base CLI inputs now pass in `fx273-current-cli-original-green.json`. Compile/setup failures and stopped broad-filter attempt receive no semantic credit. D1/D2 have not yet been run as regressions.

## Pending final gates

Final Word version/input/output/font qualification and0.5pt geometry, pinned LO26.2.5.2/Poppler26.01.0, exact pages/entries, complete installed Python3.9/3.12, strict typing/stubtest, hosted F271 jobs, workspace/full suites,49hashes/7pixels with reviewed attribution, deterministic exports, WASM, README/current measurements, genuine archive verification/10MiB ceiling, supply chain, union composition and final metadata remain pending. Preparation is not native acceptance. No baseline update, integration/provider completion or publication is approved.

## Handback

Return lease/control to root. Preserve frozen source and prior evidence. Reproduce D1/D2 before distinct remediation, refresh affected local observations, then obtain the next independent full pass. This pass ends here without fixes.
