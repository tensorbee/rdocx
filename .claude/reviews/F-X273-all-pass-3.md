# F-X273, all, pass 3

**Reviewed**: Complete F-X273 scope from immutable Base `564ad7f5d1d57a7828c721eba58cbe79da10aae3`, including integrated provider `fca298cc`, subsequent approved native reconciliation and working remediation at HEAD `2fcb5d86176e7c1a5199337c9623f76dcdd391fa` on `sprint/s96`. The complete Base comparison contains55files,9124additions and1042deletions, including separately attributed F-X271/F-X274 and sprint metadata. The working remediation contains11files,1999additions and399deletions. The integrated F-X273 provider commit contains13files,4285additions and347deletions.

**Verdict**: 1 defect,0 smells,0 nitpicks. Local remediation readiness is blocked by D3. Earlier D1/D2 retain their pass2 disposition. Final differential/native and integrated sprint acceptance remain pending.

## Frozen identity and review method

Independently recomputed working binary diff SHA256 is `ef6a44bc505c891d52f883013cce29fea3a51db6bfcbc8b039dce0e1f19f6224`, matching `f273-native-final-frozen-working.diff` and `f273-native-final-frozen-receipt.json`. All227current product hashes match `f273-native-remediation-product-after19.json`, logical manifest `0fd46ecbf81019b65e30d60aa2abe3e9dddfbad2adb69a71287f69645d3904f9`. Every scoped19 recorded log and artifact hash checked matches. No builds, tests, external applications, native exports or source changes occurred during this review. Only this report was written.

Read repository/user guidance, canonical workflow and microscope, current plan including final note/frame/rotation contracts, progress, prior independent passes and HLD ownership/source/rendering/binding/testing intent. BACKLOG and CURRENT_SPRINT show done solely for the explicitly approved provider-prefix dependency exception. The plan remains approved and final acceptance debt remains due. That qualified state is preserved rather than interpreted as final completion.

This is a full feature review with prior independent unchanged-scope analysis consumed explicitly. It is not another confirmation run of pass2's unchanged navigation and original furniture scope. Fresh inspection covers the repaired invariants, their consumers and new tests, while tracing existing table omission/TOC source qualification and selected furniture through the integrated result. Later F-X274 page formatting/parser changes are separate scope. Its inherited endnote PAGE test reconciliation is inspected as test-only reference-owner correction, not a new F-X273 field policy.

## Defects

### D3, Continuation notice can prevent every note line from making progress

`crates/rdocx-layout/src/paginator.rs:2233`

The shared fragment planner computes `fresh_capacity` before reserving the continuation notice at line2201, then subtracts that notice from actual capacity at lines2216to2218. Its oversized-line escape compares against the larger `fresh_capacity` at line2236. A line that fits that larger area but exceeds the actual notice-reduced area therefore places zero lines even on a note-only page.

A concrete supported source discriminator is an effective100pt page band, absent separator/continuation-separator records, an authored12pt continuation notice and a note with exact90pt then12pt lines. The existing6pt separator offset leaves `fresh_capacity=94`. The102pt remaining note requires continuation, reducing actual capacity to82. Its first90pt line exceeds82but does not exceed94, so neither normal admission nor the oversized escape consumes it. After the body reference is admitted on an empty page by the existing no-progress reference exception, the complete note is carried. The following note-only page repeats this exact state with `body_height=0`. No line is published and the carried ordinal remains unchanged.

The downstream guard at `crates/rdocx-layout/src/paginator.rs:3277` then reaches its debug assertion at line3282. In a release build it breaks and returns pages with the remaining note text omitted. The guard prevents an infinite loop but does not preserve source content. This violates the approved shared fragment contract's applicable notice-cost and oversized-single-line progress requirements at `.claude/plans/F-X273-design.md:260`.

This is a source-derived finding, not an executed RED. The reviewer was explicitly prohibited from running tests. Fix the planner's capacity/progress authority while retaining carried-before-fresh ordering, short whole-note movement and the actual continuation notice. Add a current-source RED/GREEN using a line between notice-reduced and pre-notice capacities, plus a notice at least as tall as the available note area. Assert complete text/order, forward carried ordinals and no panic or silent truncation. Exercise prediction and committed publication through the shared plan, rather than only testing the planner arithmetic.

The existing `carried_footnote_uses_authored_separator_and_notice` at `crates/rdocx/tests/integration_test.rs:25003` uses many ordinary short lines and checks selected text presence. It does not discriminate this line-capacity gap. The130-line native reproduction has no continuation notice and cannot reject it either.

## Smells

0 findings.

## Nitpicks

0 findings.

## Aspect and consumer assessment

Correctness and panics: D3 is the only additional finding. Inspected shared note prediction/reservation/admission/claim/commit/flush, finite short-note versus long-note decisions, carried-before-fresh ordering, note source width and paragraph indents. Source-event context is restored on success/error and repeated widths. D3 leaves the promised termination/content invariant incomplete despite the ordinary long-note success.

Contract: No additional finding. The new constructor output explicitly authors the old add_footnote display marker, while imported marker absence stays absent. Existing create-footnote/create-endnote source markers move after rPr. No authoring API or navigation tuple changes. Rotated eligibility positively checks exact source tbRl, with private section ownership and same-angle tbRlV exclusion. The bounded policy and raw metric preservation agree with the amended contract. Final installed Python constructor coverage remains due.

OOXML: No additional finding. Inspected raw note event extraction through the existing run writer's namespace authority, direct-child depth, empty/start event equivalence, source run boundaries and terminal boundary. Aliased Word marks are admitted, foreign lookalikes and opaque descendants are excluded. Rendering does not rewrite imported events. The authoring save/reopen regression inspects namespace-qualified marker sequence after matching rPr, automatic/custom mark counts and imported absence. Marker runs retain their actual source style and source events do not become self-referencing body note events.

Furniture: No additional finding. Ordinary paragraph frames now publish a distinct private body obstacle inventory. Same-story flow consumes table/drawing obstacles only. Selected frame paint and body wrap share the prepared physical rectangle. Omitted vertical anchor resolves on the bounded written-margin basis, explicit text uses following story flow, and general body/drop-cap behavior remains separate. Inspected header/footer flow plus frame plus body assertions for omitted/text/margin/page and crossing/touch/disjoint inputs, grouped frames and notBeside controls. The eligible drawing replacements retain the strict D2 transitive, growth, highlight/link/border and raw metric tests. Stable footer preparation still resets original blocks and prices the same final placements/gaps before acceptance. Existing table-only overlap settlement and cycle/bound refusal remain intact.

Rotated paint and cache: No additional finding. The baseline is a pure private projection of exact line height plus qualified actual face metrics. Cached/raw/transient line advances and metrics remain unchanged. Traced ordinary/rich/Group/MarkedContent text, tab leaders, decorations, fields, highlight/link container bands, fragments and numbered-line recording. Numbered lines use the shared policy and add first-fragment top border padding once. Graphics-only/failed face qualification returns the old raw baseline. Upright notes, cells, frames and furniture pass no main-body policy. Tests distinguish exact line12/24and font11/16, omitted/auto/center/baseline, top/center/bottom, selected furniture contexts, inherited alignment, equal-width orientation changes and actual unchanged-paragraph cache hits. Repeated reflow and painting preserve input/raw bytes.

Navigation: No additional finding. Consumed prior full scope analysis and inspected current actual table omission publishers, private navigation query and TOC check before all selectors. Deleted rows, actual merge continuation and accepted removed nested tables/controls publish bounded physical descendants. Cache replay retains current physical paths. Some(true), Some(false) and None remain distinct, with source freshness and atomic refusal preserved. Placed nested/orphan controls and STYLEREF/numbered-outline coverage remain. Unrelated F-X274 formatter changes are not credited as F-X273 repair.

Tests/evidence: D3 identifies a missing discriminator. Inspected actual assertions rather than treating aggregate counts as proof. The note-specific observer now traverses actual Group/MarkedContent paint order and transforms. Its independent120unique paragraph inventory proves text/order/page coverage separately from separator detection, while marker and separator widths remain asserted. The F-X274 endnote test finds the actual body reference owner and asserts that owner differs from continuation printed page numbering. Warm/fresh and displayed continuation sequence checks remain. No new field policy is inferred from that correction.

Structure: No additional finding. New private concrete note-plan and baseline metadata have actual consumers. Existing files carry the implementation. No new trait, generic, public model field, source file, module, crate or feature flag appears in the remediation. The approved doc-hidden omission query remains the existing bounded API exception.

## Acceptance and evidence limits

The authenticated `f273-native-remediation-local19.json` records one actual final scoped sequence:92lower,8nativeAPI and2CLI tests pass, followed by all-target three-crate check, strict clippy, fmt, prose, adapter drift and whitespace checks. The eight nativeAPI tests are native Rust executions, not Word oracle observations. All checked artifact/log/source identities match. None of those executed cases disproves D3.

| A-ID | Review assessment |
|---|---|
| A1 | Original tall header/logo/table and five spaced Normal paragraphs remain represented through actual positioned layout and CLI PDF checks. Final native numeric parity is pending. |
| A2 | Selected first/even/default/blank/missing stories, zero and negative margins retain one preparation/paint authority. |
| A3 | Frame eligibility/default-axis repair is coherent with the bounded observed native contract. Table/drawing strict wrapping and prior D1/D2 controls remain. Final native/pinned-LO geometry is pending. |
| A4 | Actual130-line49/51/30local source reproduction, full-width marker-absent note, authored separators, continuation inventory and rotated controls are useful. D3 blocks complete fragment/progress closure. |
| A5 | Physical source-qualified table omission and visible neighboring/nested headings remain intact. Original issue combinations remain required final evidence. |
| A6 | All TOC selectors retain qualified omission and freshness. STYLEREF/numbered-outline and saved/reopened controls remain. Final native navigation/page targets are pending. |
| A7 | Same-owner edits and orientation/alignment transitions compare warm/cold output with real cache reuse. Fresh installed binding/held-owner qualification after remediation remains due. |
| A8 | Actual integrated F-X272/F-X273/F-X274 native/PDF/text/PNG union remains pending. No isolated local success grants this row. |

Original native mismatches, failed intermediate86/2observations, shallow-observer diagnoses, line-number RED, graphics-only RED and setup/compiler failures remain historical evidence with their actual producers. They are not overwritten by scoped19 or relabelled current passes. This source review does not independently rerun or close external/native observations.

Remaining final debt includes actual Word and pinned LO/Poppler measurements at unchanged0.5pt/exact-page/entry thresholds, immutable source/font/output provenance, deterministic PDF/PNG/text and reviewed49hash/7pixel deltas, full workspace/native/risk union, installed Python3.9/3.12runtime/typing/stubs, WASM, hosted gates, genuine archives/10MiB ceiling, README measurements and final metadata reconciliation. No baseline recording or final acceptance is granted.

End review and return control to root. Remediation belongs to a separate phase. D3 requires affected local evidence and another independent microscope before local readiness.
