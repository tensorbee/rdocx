# F-X273, all aspects, pass 9

**Reviewed**: Focused audit of the current F-X273 Owned paragraph companion, structure-assignment publishers and paint consumers at `b7b7f7c4b8050dcfe8387fa58bda073e064974a1` on `sprint/s96`, using newly acquired source-bound sample evidence. The three source files belong to the committed carrier increment after `cab3f8873bb2ecfc95a24d4c1579200351a8e426`, which contains7,604additions and2,211deletions across block.rs, engine.rs and paginator.rs. This pass audits the identified structure lifecycle within that increment. It does not repeat an unchanged clean pass or review the unrelated F-X275 working diff.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. D2 requires remediation before accepting the affected PDF deltas or final qualification. The prior pass7D1tab finding remains closed by pass8.

## Defects

### D2, Owned paragraphs paint with the pre-assignment structure identity

`crates/rdocx-layout/src/engine.rs:10602`

The Owned paragraph arm passes only its boxed ParagraphBlock to `assign_owned_paragraph_structure`. That function assigns the current paragraph/list identity to the block at `engine.rs:10653` and `engine.rs:10662`, but the arm never updates its ParagraphSemantics companion. Both production Owned constructors captured the block's earlier structure identity at `engine.rs:6414` and `engine.rs:6523`. The ordinary paragraph builder initializes that identity to None at `crates/rdocx-layout/src/block.rs:2131`.

The stale copy is authoritative during painting. Owned `LayoutBlockLike::paragraph` installs the companion on the view at `block.rs:1780`, and `ParagraphView::structure_id` selects the companion's None instead of the newly assigned block identity at `block.rs:1674`. Consequently `crates/rdocx-layout/src/paginator.rs:6462` skips the paragraph's structured-content wrapping. Final `mark_remaining_artifacts` then wraps the unmarked body text with no structure identity at `engine.rs:10525`. Body prose therefore becomes a PDF Artifact despite its paragraph or list node having been allocated in the document structure.

This reaches real admitted inputs. `engine.rs:6391` selects Owned for non-cache-safe or excluded-source paragraphs, and cache admission excludes numbering, fields, content controls and other existing content at `engine.rs:6929`. The independent fallback after an unretained font trace also creates Owned at `engine.rs:6518`. The companion addition makes both paths susceptible. This is not a justified removal of furniture from reading order.

Actual retained evidence reproduces the failure on the deciding sample path. In feature_showcase page2, the visible `Table of Contents` paragraph changes from3`/P` MCID starts in the retained1431PDF to0MCIDs and3`/Artifact` starts in both the fresh current generator PDF and the reopened current PDF. The fresh PDF digest `fa2fdbe790a8cb26888755c3aaacb299c2dae5a0fe6da6ed4b60d4f66d8804a8` is exactly the current CI observation. Thus the fault is present in the fresh deciding output, independently of the separately observed fresh/reopened differences.

Preserve the authoritative companion contract while synchronizing the actual structure assignment and its consumers. A blanket fallback from companion None would conceal the broken publisher and change override semantics. Regression coverage must distinguish an Owned body paragraph from a normal shared paragraph, verify its allocated structure identity reaches the actual painted text, and retain list/continuation and warm/fresh ownership where applicable. The existing `shared_cached_blocks_keep_result_local_semantics_exact` test at `engine.rs:22109` exercises shared payloads and an aggregate presence check, while `word_blocks_build_document_order_semantics_before_pagination` at `engine.rs:26641` uses plain LayoutBlock. Neither currently catches this Owned-companion publication failure.

## Smells

0 additional findings.

## Nitpicks

0 findings.

## Contract and publisher inspection

Read the current approved plan, HLD08 and pass8, then traced the actual constructors, result-local assignment, view selection, continuation copy and final artifact wrapping. The approved plan at `.claude/plans/F-X273-design.md:549` requires the Owned paragraph view to expose its exact companion, and line606requires preservation of structure ownership. HLD08 describes result-local overlay ownership and distinguishes actual page furniture from prose artifacts. D2 violates that retained contract.

The ordinary shared Paragraph arm updates ParagraphSemantics at `engine.rs:10592`. Shared table recursion updates each cell paragraph companion at `engine.rs:10718`. Owned tables use their plain block view and do not install a stale paragraph companion. The tbRl Owned-to-shared conversion occurs before document structure assignment and subsequently enters the companion-updating shared arm. Plain compatibility structure assignment and plain paragraph views use their block identity. No second defect was established in those routes.

The continuation builder reads `para.structure_id()` at `paginator.rs:5574`, so it consumes the same authority and can carry D2 forward. The correction must therefore fix the producing invariant, rather than only the final PDF wrapper. Source rebinding, paint-owner presence validation and border-group mutation do not refresh this structure field. No evidence warrants changing their unrelated behavior in this pass.

## Evidence and limits

The current source hashes match the diagnostic build's recorded inputs: engine.rs `bdf3bcf4ea5d1650a778eb2c2b145f603c457e7eda7e44e1bb3e520050e53bc6`, block.rs `de2d9df3265c252557241972e65c7f0caac97e849c8fcc3aceaf2162190e32cc`, and paginator.rs `42fd843b44dc62b84430e983a050dab026065c7a9dfaa93fee8063eaba0da4d2`. Independently recomputed the frozen diagnostic CLI digest `c20c48edb2bb0fb450ec7b124c0476d2c8160e43b3a79be0e2b4bdc1f5692f04` and the fresh feature_showcase PDF digest above.

Inspected `/private/tmp/rdocx-s96-run-refresh/fx273-b7-render-diagnostic-v1`, including the build receipt, unchanged generator source identity, generator-to-CI comparison and actual sample streams. All22changed sample hashes match the fresh generator observations. The reopened CLI outputs have a separately retained scope and are not substituted for the fresh generator's differing PDF bytes.

The detailed sample packet is `s96-ci-current-monitor-v1/b7-sample-analysis`. Its fresh/reopened comparison identifies five differing pages across feature_showcase, invoice and report. The successor `fresh-reopened-semantic-only-proof-v2.json` accounts for those stream differences using only structure/Artifact begin tags and consequent MCID renumbering, with no glyph, position or drawing-number normalization. The conservative v1 result remains retained. This supports the semantic lifecycle diagnosis without claiming that every observed sample delta has the same cause.

Correctness, contract and test coverage produce the single D2finding above. Panics, OOXML preservation and structural complexity produce0additional findings in this bounded audit. The approved kerning policy, proposed oracle-fixture amendments, native Word parity and unrelated CI repairs are outside this pass. No native, oracle, hash, raster or full gate is newly accepted. No baseline update is authorized while D2remains open.

No source, test, configuration or temporary file was modified. No product import, build, test, network or native action ran during this review. Only this report was written. End the pass and return the canonical review-file lease to root for a separate remediation phase.
