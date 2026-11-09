# F-X272 C4, Codex final 1

REPAIR REQUIRED. The C4 production mechanism is accepted locally, but the incremental test diff loses an existing regression and one named control does not prove its claim. This is Root's direct affected final review after the exhausted peer cycle, not a fourth peer. C1 and Amendment A remain blocked.

## Identity and evidence

DeepSeek returned the sole canonical write lease and exited 0. Source SHA-256 is 524ca83fd08a828b5aaa02ae2d503dcffd122b4a93cfa8479b774a87be2ffaa6. Root authenticated the independently pinned before manifest edab0d2d1e324ef47f593609a1f9439d80df162810fb8739f1cedc240071d6f4 and checked all 3945 protected paths. Exactly document.rs and the append-only F272 progress changed, with no new repository path. The progress prefix remains byte-identical. The frozen before document.rs independently matches 968089d3d3f54cd538fa7b0cba5734c5737979e03c04b50a4f798b710dfea38b. The complete incremental diff is retained in /private/tmp/rdocx-cross-review/SOURCE_C4_ROOT_INCREMENTAL.diff.

Root inspected the final test bodies, reversal and reapplication commands and actual logs. Final old-source RED has three failures and three passing controls. Final GREEN has six passing tests, the capture family has 32 passes, affected namespace and capture cases have six passes, and the scoped all-targets check finishes successfully. Two existing table.rs warnings remain. This consumes scoped producer evidence without rerunning it. No full gate, provider acceptance or native qualification follows.

## Findings

1. Defect, regression coverage loss. At crates/rdocx-oxml/src/document.rs:9577 the six new tests replace the previous final test rather than preceding it. The authenticated before file at /private/tmp/rdocx-cross-review/fx272-grok-after-root3-frozen/crates/rdocx-oxml/src/document.rs:9576 contains ct_note_properties_close_the_actual_supplied_tag, which is absent from the after file. Restore that existing test unchanged and run it. Passing the new selected family cannot detect a deleted test outside that family.

2. Smell, misleading discrimination. At crates/rdocx-oxml/src/document.rs:9683 the different-prefix control assigns different URIs to x and y, and line 9691 supplies y as an inherited binding. Its successful admission proves its own y inheritance, but does not distinguish a correct exact-prefix lookup from an incorrect URI-only lookup. Retain a genuine same-URI, different-prefix control against the changed resolver or charge behavior. It must fail for the plausible URI-only borrowing implementation, without inventing an admission refusal that the unchanged allowance does not require.

## Accepted mechanism and exclusions

The inherited reducer at document.rs:567 follows ordered exact-prefix last-wins semantics. The binding resolver checks the nearest frame before comparing URI at :798 and :846, so a historical equal pair behind a nearer shadow no longer supplies the run frame. The shadow refusal tests also inspect actual saved raw declarations on admitted small documents. The near-prefix control retains its genuine positive case. Declaration pooling and the factor124 bound remain unchanged. No additional production correctness, panic, public API, schema-order or structural finding was found in this bounded diff. Formatting-only changes within the permitted file are distinguished from behavior changes.

C2, C3 and C5 dispositions are unchanged. C1, D8, all-pass-12, Issue430 and final schema, native, hash, bindings, packaging, WASM, hosted and release evidence remain pending. Return to a separate author repair phase, with no new peer or counter reset. No source edits or runtime checks were performed by this review. STOP.
