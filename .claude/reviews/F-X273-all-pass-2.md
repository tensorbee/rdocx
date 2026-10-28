# F-X273, all, pass 2

**Reviewed**: Full working result against immutable claim Base `564ad7f5d1d57a7828c721eba58cbe79da10aae3`, including committed variant refactor HEAD `3510ec5ee2c12f8aee9bc52e4434624532f160d8` and dirty successor on `work/f-x273-codex`. Eleven tracked files,4088 additions and338 deletions. Prior pass1 report is retained separately.

**Verdict**: 0 defects,0 smells,0 nitpicks. D1 and D2 are addressed for the approved local provider gate. Scoped integration/preparation readiness is approved. Final differential/native and integrated sprint acceptance remain pending.

## Frozen identity and scope

Actual full Base binary diff SHA256 is `e903e5081b3ae0c5ed94b473b3f5b397ced95348915d867aca75e74cc9f16f7f`, matching `fx273-review1-local-ready.diff`. Independently recomputed diff and all changed/progress/prior-review file hashes match `fx273-review1-local-ready-identity.json`. Product227 manifest is `377e3fc0d520b9bd37102053ead47947b223af4a2ba6fab873390ac5affcb7ef`.

Actor and root stopped all worker writes/builds before the lease. This pass writes only this report. No product tests, builds, oracle/native runs, temporary helpers, source fixes, plan/progress changes or remediation occurred. Scoped tool escalation is solely for this authorized report in the sibling worker directory.

Read canonical microscope/workflow/repository rules, approved F273 contract through81cc44a0, worker annotations and current HLD08 intent, cited HLD source/ownership/binding/testing constraints, progress and prior full/focused advisories. Canonical F274-only clarification changes no F273 contract. The review covers full Base scope, with unchanged-path evidence from pass1 consumed explicitly. An in-memory comparison of pass1 and current production content confirms the later engine changes are test additions, with navigation/table/lib/field production paths unchanged. New shared painter/block-view paths receive fresh analysis.

## Defects

0 open findings.

### D1, Floating furniture table settlement, addressed

`crates/rdocx-layout/src/paginator.rs:485`

At physical row0 the locator now resolves the whole table through existing `settle_float`, records only table peers and emits one keep-out rectangle. Later rows borrow the same settled rectangle at line504. Row y offsets derive from actual preceding row heights. The existing overlap_allowed policy is retained, not inferred from drawing/frame overlap. Drawings/ordinary frames stay outside the table-only relation. Actual row painting uses these prepared positions, so reservation and paint cannot independently choose unsettled origins.

Inspected `issue424_furniture_tables_settle_whole_owners_once_in_story_order` at `crates/rdocx-layout/src/engine.rs:28508`. It covers both selected sides, one/two rows, each forbidden-overlap combination, both-allows and disjoint controls. It asserts once-only emission, whole-table displacement, row offsets and input purity. Actual prior RED contains12forbidden-overlap failures in `fx273-review1-d1-d2-oracle-corrected-red.log`. Current source-bound lower43 executes the repaired case successfully. This closes the existing within-page publisher omission, not facing-page or section-scoped resolution or native parity.

### D2, Final line reservations and occupied-band transport, addressed

`crates/rdocx-layout/src/paginator.rs:4261`

The owning policy now separates established Body approximation from strict Furniture. Strict reservations traverse raw Word-restored line bands, repeatedly clearing every newly intersected top-and-bottom obstacle per line. Earlier lines remain above later obstacles. The first gap becomes content_offset_top, while later gaps remain private before_lines. Square reservations use each actually displaced band and the per-side union envelope, preserving both-allows table controls.

At `crates/rdocx-layout/src/paginator.rs:4471`, reservations recomputed from the actual rebreak result must equal the supplied reservations before success. Equality of outer prepared output alone no longer supplies that proof. Each rebreak uses original items, while restored metrics are recomputed without synthetic gap inflation. Inner32rebreak and outer32candidate guards remain distinct termination boundaries, with explicit cycle/failure errors. A bound is not convergence or Word evidence.

The private transport is coherent across every affected consumer:

- `WrapReflow` owns paragraph plus before_lines at `crates/rdocx-layout/src/paginator.rs:4362`. Returned raw line metrics receive border padding once, rather than encoding later clearance into ascent/height at line4496.
- PreparedFurniture retains gaps in its state and prices block height at `crates/rdocx-layout/src/paginator.rs:402`. Each outer candidate resets original blocks and empty gap vectors at line608. Locator/following origins add the same totals at line554 and line632.
- Private ParagraphView borrows gaps at `crates/rdocx-layout/src/block.rs:177`. All legacy body/note/cell/frame constructors carry empty gaps. Body reflow explicitly chooses Body at `crates/rdocx-layout/src/paginator.rs:4534`, retaining two passes, additive side reservations and error fallback.
- Furniture painter borrows the prepared gap vector at `crates/rdocx-layout/src/paginator.rs:6258`. Common painter moves y before all glyph/decorated/image/rich publishers at line5082. Raw content height/ascent remain ordinary font metrics, so highlight/link/rich boxes do not span skipped clearance.
- Paragraph border height includes later gaps once at `crates/rdocx-layout/src/paginator.rs:5073`. First clearance already shifts its origin. Change bars use prepared block_height at line6275. Following flow, footer origin and body clearance consume the same gap-inclusive height.
- Logical reflow consumes already positioned clones and preserves field/source/note metadata at `crates/rdocx-layout/src/paginator.rs:5593`. Its early-continue branch and fallback retain the common displaced origin. Final bottom-padding advance does not suppress gap consumption or alter provenance.

The actual escalation is retained, not reconstructed as immediate success. First full-paragraph repair passed initial3cases, then actual later-line4context RED moved an early line from26.25 to138.25. The per-line ascent/height attempt then produced actual146occupied-band mismatches, including highlight y66,height76 crossing object70to130, in `fx273-review1-d2-decorated-red.log`. Root designer and independent reviewer explicitly investigated the shared publishers before the private-gap successor. The paused candidate compile correction was mechanical, not semantic evidence. The border-height omission identified during that investigation is now repaired in the shared painter, not bypassed by weaker assertions.

Inspected actual successor assertions at `crates/rdocx-layout/src/engine.rs:28563`, `crates/rdocx-layout/src/paginator.rs:10084`, line10250 and the later-square case. They discriminate chained/reversed/touching/disjoint obstacles, earlier/later text, source-less/sourced ordinary and multilingual text, natural highlight/link boxes, border extent, first-gap treatment,8pt following spacing, immutable input and unmodified raw metrics. Square growth reaches later obstacles after narrowing, rather than checking only initial lines. Both sides are exercised.

The real Engine control at `crates/rdocx-layout/src/engine.rs:28624` goes beyond raw placeholder assertions. Actual PAGE27 and NUMPAGES1 substitution retains header/footer source path[2], placement page1, ordinary/rich text and natural occupied boxes after clearance. Existing footer bottom-origin, failure/recovery, overflow-note and cache tests remain in current lower43. This closes D2's internal/source-contract invariant with actual local GREEN, while final Word rendering remains due.

## Smells

0 findings.

## Nitpicks

0 findings.

## Full independent aspect and acceptance assessment

Correctness:0 additional findings. Rechecked shared selected preparation/paint, signed/zero margins, selected missing/blank first/even, spacing/borders, table flow exclusion, ordinary frame groups/documented Word defaults/drop-cap fallback, furniture/written anchors, notes, columns, line numbers, vertical alignment and rotated fragment publication. The shared prepared result reaches every applicable body page and fresh endnote continuation. The final-page fragment lifecycle and upright top-flow continuation are unchanged from the reviewed corrected pass1 result.

Contract:0 findings. The single additive doc-hidden paragraph_omitted_by_table query remains the approved implementation-facing exception. Existing authoring signatures and public navigation tuple remain unchanged. Separate geometry/navigation commit labels are pending preparation, as expressly recorded, not falsely completed. Donor archive measurements remain excluded. The local provider gate does not claim the final differential gate.

Panics:0 findings. Private frame/table ranges arise from paired construction/grouping. The new row0 inventory guarantees the later-row lookup, and before_lines vectors are initialized to original block cardinality for every candidate. Gap lookup in painting is bounded. Cycles and reflow failures propagate through existing LayoutError before publication. No public array/index admission or unbounded retry was added.

OOXML:0 findings. No changed parser/model serializer drops unmodelled content. Geometry reads preserved imported properties and does not rewrite source. Existing staged TOC generation, freshness/refusal and publication remain authoritative. Namespace/schema and raw-preservation policies are unchanged.

Tests:0 findings. Actual semantic failures are retained separately from compile/setup/oracle errors. Assertions challenge supplied cases and shared publishers. No product/native suite was rerun during review. Native parity is not inferred from self-produced inputs, PDF signatures or internal agreement.

Structure:0 findings. New concrete private gap/reservation/result carriers have actual owners and consumers. The two policy variants have two real callers today. No new trait, generic parameter, crate, module, source file, feature flag or public model field is introduced. Private ParagraphView borrowing avoids public font-metric/model expansion.

| A-ID | Observation and remaining boundary |
|---|---|
| A1 | Original five Normal8pt/header logo-table inputs, exact spacing deltas, positioned body geometry, fitting and negative controls. Current CLI/Python consumers qualify internally. Native coordinates pending. |
| A2 | First/even/default, selected blank/missing, zero/negative sides and continuation variants. Source selection and preparation/paint share the variant authority. |
| A3 | Flow/floating/frame/default/drop-cap, table settlement, wrap chains, painted occupied bands, margin anchors and page/text borders. D1/D2 closed locally. Primary/actual Word positions remain final debt. |
| A4 | Effective footer band, fitting/overflow footnotes/endnotes, continuation furniture, PAGE/NUMPAGES/restart, upright continuation, columns/alignment/line numbering and fragment agreement. Final combined native positions pending. |
| A5 | Actual placement-qualified merge/deleted-row omission, nested/control/orphan and cache cases, while placed neighbors remain eligible. Provider and test source unchanged since pass1. |
| A6 | All TOC selectors test qualified omission before selection, with source-stamp refusal, STYLEREF/numbered outline and reopened results. Five native navigation executions are carried with bounded applicability, not current executions. |
| A7 | Source-bound warm/cold same-owner spacing/text/variant and held Python owner controls, actual cache replay and nonmutation on refusal. Fresh Python22 includes current layout consumers. |
| A8 | Focused integrated F272/F273/F274 composition after providers and actual native/PDF/PNG/text union remains pending. This review grants no premature composition acceptance. |

Navigation source tracing remains sound: actual table-placement decisions publish bounded physical descendants, cache-safe returned cell decisions replay on current paths, query None differs from Some(false/true), TOC excludes only proven omissions before every selector and existing freshness remains mandatory. The table/lib/field source and native test assertions are byte-identical to pass1. Current shared Body branch uses empty gaps and keeps its former two-pass behavior, with current body-controls in lower43. No stale result or lexical vMerge alone becomes omission authority.

## Actual evidence and applicability

Frozen receipts under `/private/tmp/rdocx-s96-run-refresh`:

- `fx273-review1-final-local.json`: current lower43, CLI2, scoped all-target check, strict three-crate clippy, fmt, prose, adapter drift and whitespace success. Tests and actual CLI application hashes are authenticated by `fx273-review1-root-auth-complete-local.json`, with empty source/log/binary mismatch inventories.
- `fx273-review1-cli-original-green.json`: same seven immutable Base CLI inputs pass against the current authenticated application. Original four Issue426 combinations and styled no-outline control remain, alongside both original header reproductions.
- `fx273-review1-provider-python.json`: actual NEW fresh dev wheel build/install/test all return0. Entire current selection is22passed,647deselected,45.81seconds. This is an actual22pass batch, unlike the prior21plus corrected1 applicability composition.
- Wheel SHA256 `70e5dd68aa4a08f38d4e4d9f0c8a527b17ac0d0dd578900934706b039c5abf77`, wheel native and actual isolated imported native SHA256 `1d84a2b936243efd7c20bdc2a3e826b4c046c96cc1d3933bd9ffe33b0bf26156`. Independently read/hashed the wheel, native member and installed module during this pass, all match. Product227 and selected test source remain unchanged before build, after build and after tests. Dev qualification is not release/archive measurement.

Five native navigation cases are expressly predecessor executions on manifestbe5dff87, binary `41ed749e6be619af5660242f08b4c708d5a0c696575aa81f775101e5423915cf`. They are not relabelled current. Applicability is bounded by unchanged navigation/query/table/field/test source, no furniture or wrapping inputs and unchanged empty-gap Body semantics. The root authenticates original receipt/artifact identities. Current lower navigation/cache/refusal controls and current CLI/Python exercise the successor. Final workspace/native execution is still due at the integrated boundary.

Old Base failures, first-pass failures, later-lineRED4, decoratedRED146, failed compiles/setup/oracle assertions and all intermediate successes remain attributed to their actual producers. None is silently replaced by the final GREEN or counted as native evidence.

## Explicit final debt and handback

Remaining: actual Word tool/input/output/font qualification and0.5pt geometry, exact page counts/entry sets, pinned LibreOffice26.2.5.2/Poppler26.01.0, native furniture/frames/anchors/borders/notes, full integrated/native/workspace suites, hosted F271 jobs, installed Python3.9/3.12 complete suites, typing/stubtest, WASM,49hashes/7pixels with reviewed attribution, deterministic exports, README examples/current measurements, genuine archive compilation/10MiB ceiling, supply chain, full risk union, A8 composition and final metadata checks. Prepared tools/helpers are not native product acceptance. No baseline update or final acceptance is granted.

End pass and return sole-writer lease/control to root. D1/D2 are closed for scoped readiness. Record intentional commits and structured preparation through the normal workflow. No confirmation microscope is requested after this clean pass. Material reconciliation or source changes still require their applicable checks/review.
