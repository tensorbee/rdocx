# F-X278, Codex implementation review, pass 3

**Reviewed**: working tree on `work/f-x278-codex`, unchanged Base and HEAD `5820539cd277b417ae060e15010eb73bd1d00b01`. The implementation diff contains 21 tracked files, 2970 insertions and 348 deletions. This is the third and final available independent implementation peer, performed by Codex after the user's authorized takeover. Two earlier DeepSeek reports remain actual historical reviews. This report neither impersonates DeepSeek nor resets the counter. Root final review follows this phase.

**Contract**: approved `.claude/plans/F-X278-design.md`, SHA-256 `cacd4844683b63f7c9fb2fbcde658d7b3b03fcdc36eb747768ac46fe9af02bae`, including its operative rolling approval, seven A-IDs, consumer inventory and residuals. Read the full plan, current progress, prior two implementation reports and applicable repository review rules. Earlier draft and approval-prerequisite prose is historical under the operative approval.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the inspected scope. All surviving prior findings have a supported disposition below. Local implementation review is clean. Final feature acceptance, Root final review, integration and the pending union riders are not claimed.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Source and evidence authentication

The author was stopped before this report-only lease. Handoff SHA-256 is `d983a7419d8f860dee2d1db0d162970a685ca3870b786f3aa6a643ffe2eb4304`. `FINAL_PINS.json` SHA-256 is `5d92f35ab1d5ef80996f11f0e511328660deab230020d921e1937defdf66931d`. Independently rehashed all128 pinned files with zero missing files or mismatches. Those pins bind source, plan, progress, previous reviews, scripts, staged package and named evidence. The handoff is an author claim, not acceptance.

Independently compared the326 build inputs in `SOURCE_INPUTS_BUILD.json` with current source. The sole difference is the interpreted `rdocx-py/tests/test_core.py` fixture. Production inputs match the successful extension build. Its current hash is `dcda9c18bd904053fd26cb2651df2b3f3d892705a6c88a28b25bf32eb3b07b3e`, which the deciding runtime prints. The post-build change avoids duplicate page0 in the one-page selection and does not alter binding or production code. `BUILD_APPLICABILITY.json` records this distinction.

The exact staged artifact is `/private/tmp/rdocx-s96-root-f278-repair/pyext/rdocx/_rdocx.so`, SHA-256 `cadb117962e6d93dfeee048493dedae6e8d3da77b93388cd65e032789e87dbf7`. Its identity is bound by the manifest, extension receipt and actual import line. This is Python3.14.7 local-extension evidence, not an installed wheel or final Python matrix result.

Read the actual `run_python_selected.py` harness. It extracts exactly the two named public test definitions from current `test_core.py`, preserves their real `__file__`, compiles and executes them, supplies a temporary path to the first and calls the second directly. Their imports occur inside the definitions. There is no pytest shim, decorator replacement, skip or external oracle. The deciding log reports both definitions passed on the authenticated artifact. This is direct execution of two actual test definitions, not a claim that pytest collected or ran the complete module.

Read applicable logs and adjacent receipts for seven engine cases, two existing tbRl controls, two integration cases, the named dense-form regression, two shared-font cases, two paragraph-property cases, sequential CLI feature-on/off cases, six-crate Clippy, extension build and Python runs. Inspected log hashes match their receipts, with child0 and wrapper0 on the deciding greens. Clippy includes all targets, all features and warnings denied for the six changed crates. Failed attempts remain failures. Counts overlap and are not summed into unique coverage. Prior public/EPUB/ODT, cache-nine, Rust feature-off and nested-table observations are consumed only on their unchanged applicable paths, with the supplied reuse authentication.

## Inspected implementation and consumer composition

Bounded coherent reads covered all21 changed-file diffs and the corresponding production functions and discriminating assertions. Particular attention was given to:

- Shared font mapping, contribution recording, diagnostic reset and deduplication in `oxml-layout/src/font.rs`.
- TOC run and stored-segment style filtering, both first-line merge providers, automatic-colour call chains, Word numbering and symbol consumers, and the tbRl publication arm in `rdocx-layout/src/engine.rs`.
- Table shading through direct and recursive control paragraphs, nested-table isolation, base table-border ownership and the retained row barriers in `table.rs` and `paginator.rs`.
- Rust policy selection, accepted caches, tracked uncached layout, diagnostic publishers and explicit ignore-list methods in `rdocx/src/document.rs`.
- Every changed Python publisher, keyword and stub, plus public SVG re-export and the unchanged deterministic PDF/A exclusion.
- CLI option plumbing and separate PDF/image/render/layout policy branches, including nonempty system feature-off probes and publication refusal.
- Public, EPUB and ODT indent projections, generated TOC, caller-over-embedded metrics, dense-form geometry and PNG assertions, and all seven focused engine regressions.
- HLD08, HLD10 and HLD15 replacement policy, row barriers, inheritance wording, symbol boundaries and current public API inventory.

The Rust helper returns the cache Arc for empty accepted input, an uncached result for tracked input, and an owned caller-font bundle wrapped in Arc for nonempty input. Derived Default is bundled. Explicit system requests refuse before an engine on feature-off builds. Explicit deterministic, caller-only, bundled-fallback, alias and additional-font methods retain their named engines. Python and CLI select the documented public routes rather than calling the private helper or treating the bool as effective on an ignore-list method. `inspect` gains no option. Save/PDF/page and raster omission paths still flow through the common selection.

The Wingdings2 map and unknown-symbol channel are Word-only. The existing PowerPoint methods keep their false Word-mode branch and existing outputs. New substitution facts remain separate from `current_font_diagnostics`, reset per layout, deduplicate and replay through contribution recording. The public shared API includes the three approved doc-hidden methods. No extra version bump or publication is inferred.

## Prior finding dispositions

| Prior item | Disposition and inspected discriminator |
|---|---|
| D1 and N6, stale selected-page layout | Fixed at `rdocx-py/src/document.rs:13848`. `render_pages` delegates to its diagnostic sibling. That sibling derives omitted indices from the exact layout it renders. The new test at `test_core.py:16004` uses real same-family Sans/Mono metrics to force expansion and contraction, both tiers, omitted and explicit selections, exact sibling bytes and diagnostic equality. Old artifact logs observe1 instead of3 and the inverse out-of-range failure. Fresh logs observe3 and1. |
| D2 and D3, stale public policy docs | Fixed in Rust PDF diagnostics and Python stub documentation. Bundled default, caller priority and explicit system selection agree with implementation. |
| D4, TOC matcher documentation | Fixed. HLD08 describes the implemented trimmed numeric id suffix and exact trimmed `toc N` name rule. Matcher behavior was not broadened by this repair. |
| S1, silent64-fact truncation | Fixed. The cap is removed. `font.rs:6753` contributes70 distinct real requests, checks all messages and tier, repeated-request deduplication, reset and PowerPoint-channel separation. Actual contributed-text RED is64 versus70. The earlier resolution-only fixture returning0 is not semantic RED for the cap. |
| S2 and N9, dense-form evidence/comment | Fixed. The named regression actually passed. Existing reviewed expected checksum5302818789495566545, non-white32119 and foreground1682 remain in this repair. The comment now identifies table-border ownership and inside-band geometry. This is a local executed regression, not a recorded hash baseline or native observation. |
| S3, diagnostic tuple content | Fixed. The actual Python definition pins the complete missing-family to Carlito bundled fact across diagnostic siblings, and caller-satisfied tuples are empty. |
| S4, system omission selected pages | Fixed. The Python test checks no-caller system page count, omitted raster count, sibling bytes and diagnostic agreement. |
| S5, CLI image branch | Fixed. Feature-on PNG convert covers bundled/system with and without callers, comparing output bytes with the selected layout. Feature-off covers caller-only PNG and empty/nonempty explicit-system PNG refusals with no output. Variants ran sequentially. |
| S6, cantSplit | Fixed. The engine test preserves page1 and the original[23,17] within-one-line bounds for each alignment, while cantSplit must produce[0,40]. |
| S7, caller over same-family embedded/bundled | Fixed. `integration_test.rs:1990` uses embedded/bundled Liberation Sans and caller Mono renamed to that same physical requested family. Caller output differs from both, wins both supported tiers and survives reopen. |
| S8, default SVG omission | Fixed. The actual Python definition checks default page0, last valid page and out-of-range omission, alongside caller and system routes. |
| S10, tbRl merge adequacy | Fixed, including the source defect exposed by the stronger fixture. Direct and numbering overrides must produce520.8pt baseline. Removing either merge independently fails at526.5865234375. Those are mutation-adequacy checks, distinct from the genuine earlier list-ownership panic RED. |
| N1 and N2 | Fixed. Both row-barrier lists include nested tables and clipping/exact height. Effective inherited explicit/theme colour remains ahead of automatic contrast. |
| N3, N4 and N5 | Fixed. TableGrid id, zero underline lines of any colour plus zero blue lines, and actual same-element firstLine/hanging layout at72pt are asserted. |
| N7, error precedence | Rebuttal supported. The plan requires explicit-system refusal before either engine. It does not specify precedence between two simultaneously invalid options. No new precedence rule is required. |
| N8, cache-test naming | Rebuttal supported. The actual existing test asserts Arc pointer identity and one omission layout invocation, followed by two uncached caller bundles. Its name matches those observations. |

Peer1 F1 through F16 remain resolved through the above repairs and retained controls. In particular, stored TOC segment and generated TOC, Word `w:sym`, absolute indent, exact grid, nested diagnostic identity, donor luminance attribution and lower-byte Wingdings wording are present. No old clause was silently deleted to close them.

## tbRl Owned repair

At `rdocx-layout/src/engine.rs:4831`, the Owned arm now updates its existing paragraph semantics and retains Owned representation. The shared structure dispatcher at `:10570` assigns lists and anchored figures through the Owned path. Its ordinary shared Paragraph path explicitly assumes no list or anchors. Retaining Owned therefore repairs the actual assertion failure without weakening that invariant or changing broad structure assignment.

The baseline is consumed from paragraph semantics in pagination at `paginator.rs:5160`, including Owned paragraphs. The strengthened test observes painted transformed baseline, ListItem and Figure ownership, rather than merely calling a merge helper. Existing exact-baseline and cache controls remain green. The change does not replace the independently required first-line indent assertions, which remain in the same test.

## Seven acceptance rows

| A-ID | Local review outcome | Final boundary |
|---|---|---|
| A278-1 | Stored/run TOC paths and generated TOC are black without underline. Body Hyperlink control retains colour and underline. | Native and pinned oracle observations. |
| A278-2 | Pair replacement occurs in both providers. Six absolute72pt controls, same-element hanging, public/EPUB/ODT and real tbRl baseline/list/anchor behavior are discriminated. Raw unknown-child preservation remains. | Integrated and native calibration riders. |
| A278-3 | Base table borders remain table-owned. Local style TableGrid is3x3 and direct all-none0x0. Dense-form geometry executes. | Deterministic hash/pixel and oracle checks. Partial/conditional edges retain their declared residual boundary. |
| A278-4 | All three alignments match the base split on page1 within the original tolerance. cantSplit remains a whole-row control. | Native placement of short-cell content. |
| A278-5 | Paragraph shading precedes this cell fill. Direct and control paragraphs receive fill, nested tables resolve independently, and body None remains black. Explicit/inherited colours keep precedence. | Word threshold calibration. |
| A278-6 | Numbering and explicit symbol consumers map F0A1 and preserve F0A7 with one named unknown fact. Shared PowerPoint path remains separate. | Final oracle and coverage qualification, with no measured Wingdings coverage claim. |
| A278-7 | Rust, Python and CLI use bundled omission, explicit system opt-in and caller priority. Diagnostic twins agree, page selection uses the rendered layout, caller metrics discriminate and feature-off requests refuse before publication. | Linux/CLI agreement, installed versions and final union. |

## Zero-finding aspects and limits

Correctness, contract, panics, OOXML, tests and structure produced zero new findings. No new unchecked input slice, schema child-order change, raw-subtree loss, forbidden dependency, trait, generic, feature, source file or test binary was found. The additive public control and SVG visibility remain in the approved Word0.17.0 inventory, and the three FontManager additions remain shared0.15.0. The review does not substitute source-level stub inspection for installed typing/runtime agreement.

No source or test was edited. No build, test, benchmark, oracle, native app, network, Git mutation, cleanup or delegation was executed by this reviewer. Giant unrelated files and author JSONL histories were not exhaustively replayed. Recorded evidence was inspected for actual assertions, artifact applicability and child outcomes.

Root final, full integrated suites, deterministic hash49/pixels7 and reviewed actual14-key deltas, native Word, pinned LibreOffice/pdftoppm, installed Python3.9/3.12 typing/stubs, WASM, wheels, archives, hosted CI, publication dry-run and size checks remain pending. Historical integration462pass/10fail/8ignored and library391pass/17fail/1ignored remain failing observations, not green or waived. No deterministic baseline was recorded and no feature completion or integration is asserted.

This report ends implementation peer3. The reviewer stops and returns the writer lease to Root.
