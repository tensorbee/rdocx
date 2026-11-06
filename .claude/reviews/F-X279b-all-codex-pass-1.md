# F-X279b, all, implementation pass 1

**Reviewed**: Frozen working diff on `work/f-x279b-codex`, Base and HEAD `391280577110c9b7b92ace0420a4a1d98c707889`. Twelve product, test and HLD paths, 2219 insertions and 75 deletions. This is pass 1 of the original implementation cap of three.

**Verdict**: 2 defects, 0 smells, 0 nitpicks. Return for remediation. This is not whole-story or final-union acceptance.

## Source and contract

The reviewed snapshot is `/private/tmp/rdocx-s96-f279b-evidence/final-local-9/source`. Its `SOURCE_PINS.json` SHA256 is `d7e5232812f5215ce88193fd6588362dae0ded0cd42bdd9155279a952f242374`. All 15 product/configuration snapshots match their pins and the working files. The plan and progress match `REVIEW_CONTEXT_PINS.json`, SHA256 `e1365f7ea8ff2016a46dec5c3ac59a5789620ce8db2c6993b561087666161221`. All 17 source/context pins were checked again after reading and were unchanged.

Read CLAUDE, WORKFLOW, the microscope command, the approved B plan, its risk routing and HLD03/04/10 contracts. Current sprint and backlog both mark B in progress. A owns revision text/vocabulary, C owns comment markup/styles and D owns crossing-reply authoring. Their integrated code is composition context, not a new B implementation claim. B owns A279-3/5/6 and P1/P2/C6. The genuine F272 provider is an ancestor of this Base. Its exclusions and final riders remain.

## Defects

### D1, crossing-reply shortcut publishes edits that the caller asked to ignore

`crates/rdocx/src/comparison.rs:967`, `crates/rdocx/src/comparison.rs:1222`, `crates/rdocx/src/comparison.rs:1301`

`crossing_reply_additions` proves replay-versus-edited equality using the caller's comparison options. Those normalized operands intentionally omit ignored formatting or whitespace and can mask ignored textbox content. A successful proof then copies the edited body and related stories directly instead of retaining original ignored values or running their policy-aware publishers.

A concrete source-derived discriminator is an original bold body paragraph with a parent comment, an edited clone adding the supported crossing reply and changing that paragraph to italic, and `ignore_formatting=true`. Replaying the reply on the original keeps bold. The normalized check ignores the bold/italic difference, so the reply shortcut qualifies, but the direct edited-body copy installs italic without a formatting revision. This violates the existing left-biased ignore contract in `docs/hld/03-architecture.md:1462`. The same predicate/publication mismatch applies to ignored whitespace. These examples were not executed during this read-only review.

The P1 shortcut must establish graph-only eligibility without erasing differences it later copies, or publish the body through the policy-preserving route. Add a combined reply-plus-ignored-body-change control, including acceptance and rejection of the comment snapshot. Existing P1 tests exercise a graph-only change and therefore do not discriminate this leak.

### D2, source-default preparation skips typed row and cell control owners

`crates/rdocx/src/field.rs:5306`, `crates/rdocx/src/field.rs:5339`, `crates/rdocx/src/field.rs:5385`

The new table preparation traverses only `table.rows`, `row.cells` and each cell's content. It omits `CT_Tbl.content_controls` and `CT_Row.content_controls`. The Row branch inside `prepare_control` repeats the omission. Those are actual typed row and cell controls, not the excluded inline-run SDTs. Their ordinary paragraphs are retained and emitted. The existing physical paragraph traversal explicitly visits both collections at `crates/rdocx/src/document.rs:2338` and `crates/rdocx/src/document.rs:2359`.

For example, a source with Georgia/28 document defaults and an ordinary unstyled paragraph in a row SDT or cell SDT will skip both explicit source-default style selection and direct font-component preparation. The helper nevertheless clears paragraph-default flags on all source styles. After insertion into a Calibri/22 destination, the unprepared paragraph can inherit the destination defaults instead of the promised source pair. This is a source-derived missing branch, not an executed runtime claim.

A279-5 includes ordinary table-cell paragraphs and the actual enclosing table style context. Its exclusions name inline run controls, revised raw payload and conditional table formatting, not these typed row/cell owners. Route those collections through the same preparation with the actual enclosing table context and cover a saved/reopened control-owned cell paragraph. Existing font-priority tests cover direct rows/cells only.

## Evidence and limits

Inspected public Rust assertions for all three equal-split granularities, structural-token neighbors, logical-word edits across physical owners, formatting policy, hyperlink placement, protected-owner preservation/refusal, P1/P2 resolution, source-default/style-chain font priority, saved body bytes, list instances and deterministic bundled-font SVG. Inspected the Python insertion binding, its unchanged signature, staged publication, companion closure and tests for notes, comment graphs, chart/image closure, held handles and omitted/explicit policy. CLI tests inspect real command output and revision text.

`PUBLIC16_APPLICABILITY.json` preserves the actual failed 15-PASS/1-FAIL predecessor command. Its 15 passing test bodies and all production/HLD/configuration files are unchanged, and the corrected optional-part assertion has a separate genuine 1-PASS execution. This supports 16 individually green public controls, not a relabeled successful whole-16 command. The prior full-ZIP ignore-comments assertion was narrowed to original paragraph/anchor graph and optional comment parts. The retained package diff identifies only inherited namespace declarations on sectPr, not a changed comment graph.

Current receipts show CLI 4 and 9 controls, Native issue434 29 controls, strict all-target/all-feature Clippy exit 0, and four completed formerly failing Source controls including the mapped-content case. Eight retained executable paths were independently hashed against their receipts. The remaining slow Source controls were still being collected during this pass. Fresh Python extension/runtime evidence is explicitly pending. No build or executable was run by this reviewer.

Historical REDs and unsuccessful intermediate repairs remain distinct. In particular, accepted-provider evidence establishes actual Base GREEN for the mapped-content case. The new ordinary Word aligner now retains paired nontext replacement instead of splitting that barrier into deletion/insertion. This review does not relabel earlier failures or infer a fresh Base execution.

Final integrated verification, full suites, hash measurements, installed Python/typing matrices, WASM, packaging and native Word candidate acceptance remain unperformed final riders. Pending final execution is not presented as a source defect, a local PASS or completed acceptance.

## Smells

0.

## Nitpicks

0.

## Not found

No additional findings in panic/indexing and checked arithmetic, OOXML order and namespace preservation, API/stub compatibility, or structure. The diff adds no crate, module, trait, generic abstraction or feature flag. Owner-local publication and physical Source selections remain separate from logical equality and generated revision wrapping. The actual same-word grouping retains physical owner consumption, token and shell barriers, and the paired leaf-capture path. Default fragment policy, native body-only insertion and Python full-closure insertion remain distinct as approved. The two defects above are correctness and contract findings, with missing discriminating combinations in otherwise useful local tests.

The review ends here. No product, test, progress, plan or HLD file was changed.
