# S96 final output delta, Codex pass 1

Reviewed producer: `15424ace0d3397684293e776da61ae84f9013c98`.
Scope: actual frozen deterministic outputs, attribution and proposed baseline adoption.

## Findings

Defects: 0. Smells: 0. No source or baseline was changed during this review.

## Attribution

F-X278 declares the exact fourteen changed keys at
`.claude/plans/F-X278-design.md:529`. They are PDF bytes and page-content
fingerprints for contract, feature_showcase, invoice, proposal, quote and report,
plus invoice and quote page1 PNG fingerprints. Automatic text contrast is
implemented at `crates/rdocx-layout/src/engine.rs:16659`.

F-X277b owns the two additional document XML keys, feature_showcase and report.
Its no-heading TOC construction separates the result paragraph at
`crates/rdocx/src/document.rs:47264`. The sample calls occur before their first
heading. The historical donor five-key forecast is not a measured acceptance
count. The actual integrated union is sixteen, with two XML keys assigned to
F-X277b and fourteen output keys assigned to F-X278. No unnamed fifth key is
invented or required to pad this actual inventory.

The retained actual old and new feature/proposal PDFs have the same page counts,
MediaBoxes, glyphs, fonts and matrices. Eighteen black-to-white fill operators
are their complete inflated content change. All seven new PDF resource
fingerprints match baseline. Neither old nor new retained sample PDFs contain
ActualText, so F-X281 is not a contributor here. The actual feature XML change
is only the paragraph boundary between the TOC separator and end. Other missing
old counterparts are explicitly source-supported, not claimed as byte diffs.

## Independent checks

Root independently recomputed all forty-nine fingerprints from the twenty-one
retained files and matched the measured sixteen deltas. Root also decoded all
seven actual pinned-Poppler rasters and matched their recorded dimensions and
fingerprints. Poppler 26.01.0 was genuinely built from the repository-pinned
source and authenticated through the existing installer receipt. All seven
raster children exited zero at 150 DPI. Five golden entries match exactly,
while invoice and quote change at unchanged 1275 by 1650 dimensions.

Root inspected the seven-sample native contact sheet, four old/new diagnostic
PDF page pairs and the actual pinned invoice/quote rasters. The intended white
text appears on dark headers. Diagnostic Poppler 26.09 page pairs are not used
for golden values. No independent layout change is attributed to the colour
repair.

Evidence is retained under
`/private/tmp/rdocx-ci-triage-42293933089-20261106/H2-15424ace-output-attribution`,
including `FINAL_ATTRIBUTION.json`, `EXACT_COLOR_ONLY_PROOF.json`,
`ROOT_INDEPENDENT_CHECK.json` and
`pinned-poppler-26.01-measurement/MEASUREMENT.json`.

## Disposition

Accept the measured sixteen hash entries and two golden entries for a separate
labelled behaviour-delta commit. Preserve all other entries and the actual
Poppler identity. This is output attribution and baseline review only. Fresh
final-source harness checks, remaining product/native/hosted gates and sprint
closure remain pending.
