# F-X277c, all aspects, independent pass1

**Review kind**: actual independent implementation microscope after Root's identified self-review. This is the first independent implementation pass, within the existing three-pass cap. The exhausted design peers are not reset.
**Reviewed**: work/f-x277c-codex, working candidate against bd48016f2ef48ed77c63505b2db21ba9e8100382. Seven tracked product/test/HLD paths, 794 insertions and 23 deletions. Exact diff SHA256 94b5a06871292a3e67cda0a935ba91d78396b34b088bd0d4d652ff710189c771.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. The local implementation is consistent with the approved four-row contract. Final integrated qualifications remain pending.

Read the approved design, Issue432 source observations, current progress, Root self-review, actual changed source/test/HLD files, relevant property parser/writer and native/Python/CLI/WASM consumers, and retained implementation evidence. No builds, tests, runtime imports, source fixes, formatting or evidence edits were performed. This review writes only this file.

## Defects

Zero findings.

## Smells

Zero findings.

## Nitpicks

Zero findings.

## A277-4, merge content and atomicity

At `crates/rdocx/src/table.rs:1214`, only exactly one paragraph can qualify as lone-empty. The comparison clones it, clears properties and recognized root-attribute records, then compares the complete remaining CT_P with CT_P::new(). It does not infer emptiness from displayed text or enumerate a fragile subset of sidecars. Comment ranges, bookmarks, revisions, equations, ruby, controls and opaque raw children therefore remain content. Multiple paragraphs and nested tables are retained. Row-major concatenation remains unchanged for substantive cells. When every selected cell qualifies, the original first cell's paragraph, including properties and root attributes, is retained at `table.rs:1232`.

Existing topology and metadata refusal precede the candidate clone and final assignment. Indexing uses the validated selected ranges. No error or new panic boundary is introduced. Tests at `table.rs:3625` cover horizontal/vertical and empty placement/all-empty/nonempty combinations, and at `:3724` cover semantic/opaque zero-text payloads, multiple paragraphs and nested tables. Existing owner/opaque merge controls are also retained and actually observed green. Python Cell.merge continues to call the native operation through its existing staged owner publication, with a new public regression source in the existing test carrier.

## A277-5, exact property preservation and exclusion

`crates/rdocx/src/table.rs:737` constructs a fresh empty paragraph and chooses the source cell's last direct paragraph. It clones pPr and preserves its paragraph-mark rPr intact. Only if that mark is absent does it clone the last direct run's rPr, including the no-pPr fallback case. Source paragraph identity and text are not copied. The implementation follows the approved explicit exclusion list: sectPr, all typed/raw numPr data, numbering-change carriers and pPrChange are cleared. It does not invent a broader policy that would strip unrelated mark formatting or arbitrary rPr payloads.

The raw revision carrier cleanup at `table.rs:779` examines the first qualified element. The actual property parser supplies externally bound fragments when needed. Explicit Word bindings are recognized, existing implicit unbound canonical w fragments follow the model's established convention, and explicitly foreign bindings are retained. Only the named numPr/sectPr/pPrChange roots are excluded. Unknown other children keep their original bytes and positions. Parser failure conservatively retains the carrier. Typed property emission still uses the existing schema-order writer. No new namespace admission, public error or resource allowance is introduced.

Both real add_row and insert_column consumers install the substantive inherited paragraph at `table.rs:1040` and `:1136`. The original formatted_empty_cell remains cell-only. Its split_cell caller still uses that original path, so row/column inheritance does not silently change split behavior. The helper has two actual callers and performs the real transformation rather than forwarding.

Tests at `table.rs:3662` cover row/before/after column insertion, last-of-two paragraph selection, last-run bold fallback, removed typed ownership, unknown pPr bytes, unchanged source cell and atomic opaque-tcPr refusal. At `:3774`, mark priority, zero-paragraph source, direct-only fallback, source identity exclusion and unchanged split defaults are asserted. At `:3845`, duplicate/malformed raw Word ownership is removed while foreign pPrChange/numPr and explicitly rebound w:sectPr survive. No unsupported universal imported-property claim is inferred from these selected controls. The actual Python row/column wrappers are staged native consumers, with explicit saved-XML public test source and runtime debt retained.

## A277-7, authored TOC and foreign control

At `crates/rdocx/src/document.rs:44524`, the extra typed paragraph is added only when selected headings are empty. Begin/instruction/separate remain on the existing title paragraph, and the existing last-paragraph field-end logic chooses the new paragraph. Existing heading-bearing insertion, max-level clamp, bookmarks and entry hyperlink formatting remain unchanged. No rebuild parser broadening is present.

The native test at `crates/rdocx/tests/integration_test.rs:36765` exercises both authored insertion orders, two rebuilt entries and save/reopen/rebuild. At `:36854`, an actual title-sharing foreign one-paragraph field still produces the original paragraph-boundary refusal and unchanged serialized document. This preserves R-A277-7 instead of turning imported support into new scope. Python test source calls the existing insert_toc/rebuild_toc methods in both orders. F273 is not treated as a provider.

## A277-8, actual consumers and chart preservation

The only replacement production-policy change is removal of the chart traversal from `crates/rdocx/src/document.rs:46064`. Typed body/header/footer work, raw text-box traversal, note handling, existing reparse/count behavior and error policy remain in their existing paths. Regex already excludes chart caches and remains unchanged. Template mail merge and chart-data mutation remain separate consumers. WASM replacePlaceholder still calls native replace_text through its existing wrapper at `crates/rdocx-wasm/src/lib.rs:108`, so the shared chart exclusion reaches that API without a new signature. Runtime/target qualification remains due, not claimed by source tracing.

At `crates/rdocx/tests/integration_test.rs:36796`, real chart, relationship and embedded workbook entries are retained and byte-compared after both literal and regex replacement, while the body count becomes exactly one. This is a real companion-preservation check, including the series/cache bytes, rather than merely searching an aggregate XML string. At `:36874`, native report/literal/regex agree at six logical occurrences across body/header/footer and their text boxes. Saved raw XML proves both physical branches were rewritten while counted once.

The actual CLI test at `crates/rdocx-cli/tests/integration.rs:8532` invokes plain text, JSON text, replace --expect and reopened text on the same chart and multi-story fixtures. JSON counts only the paragraph/story text projections, avoiding double counting run-detail data. The implementation uses the actual text_report route, not the superseded body-only inference. Existing three CLI story/control/note/tracked expectation cases are retained and observed green. Python uses real try_replace_text and replace_all_regex APIs in its new source test. No installed Python result or all-imported/all-revision semantic proof is claimed from these fixtures.

## Evidence authentication

Authenticated all seven FINAL_SOURCE.json file hashes and the exact current Git diff hash. All seven named final-check logs match their recorded SHA256 values. All three retained final Rust executable hashes match current bytes. FINAL_SOURCE applicability is explicit: final native/CLI evidence covers the same production and integration/CLI source, with subsequent table changes limited to cfg(test) controls. The later unit/control run authenticates that final unit source. This is not a promotion of an older green result across a production fix.

Reconstructed baseline-tests.patch against the exact Base entirely in memory and reproduced all three BASELINE_SOURCE.json source hashes. The baseline patch is tests-only. Its SHA256 is d4c962ab522cd3bc42b446304c4e650b7c975b10dd8216aa64d9e844b9cc2878. Actual baseline unit failures distinguish merge and absent inheritance, native TOC-before-heading fails the existing boundary refusal, and corrected baseline-chart-2 fails because literal replacement changes word/charts/chart1.xml. The earlier chart companion lookup and compile/setup mistakes are preserved and not counted as product RED.

Actual retained final receipts show five unit cases, four native cases and one CLI case passed, plus four existing native owner/merge and three existing CLI controls. Warnings-denied all-target/all-feature native/CLI Clippy is checked exit0. Later failed fixture/setup attempts are explicitly classified and superseded by affected successful checks. Format/prose/whitespace/AST observations remain local checks. No installed binding or oracle evidence is substituted by AST or source inspection.

Artifact identities consumed without execution:

| Artifact | SHA256 |
|---|---|
| Final unit executable rdocx-4b490f76e4e2a26d | 2ba35e7bb002ef48146ae2c0a8110fa166220781fa6855f6e47b0f06bcb10dfe |
| Final native executable integration_test-5254c5efb4197021 | f788e5c4d73bc2d6a35468c213f281ed759cb836935c0a1d33401580da292ad7 |
| Final CLI test executable integration-c9e904dddc7fee67 | 086bd595a9d4dd560d9bb621a6760cf59b6a0499d3f84ccd74466e5877fb419d |

## Structure, HLD and handoff

No new product source file/module, trait, generic, dependency, feature flag or public API was introduced. The new property helper is substantive and scoped to its two authorized publishers. Full CT_P equality reduces the merge cases rather than growing a sidecar-specific content taxonomy. New indexing derives from the same source vectors and preserves existing atomicity boundaries. HLD03/HLD10 edits stay within table, authored TOC and search semantics, with no math/shape/notes ownership takeover.

Final installed Python3.9/3.12 and typing/stubs, WASM target/runtime, SDK/schema and raw-preservation riders, native residuals, full integrated union/hash, archive/dry-run and hosted CI remain mandatory pending debt. Source inspection and local reopen are not those external qualifications. Parent/all-child/F272 completion and F280 barriers remain unchanged. No provider, baseline, commit, prepare, completion, reporter comment or publication is claimed here.

This review ends with this file. Return control to Root for normal adjudication and subsequent workflow, preserving the original caps. No remediation or confirmation pass is started here.
