# F-X277c, Preserve cell formatting, authored TOC rebuilding and chart-excluding replacement

**Status**: approved
**Sprint**: S96
**Size**: M
**Depends on**: none
**Test gate**: regression

## Problem

Complete original Issue432 A277-4, A277-5, A277-7 and A277-8 as one bounded leaf. The parent remains the completion umbrella. F-X277b owns A277-2, A277-3 and A277-6 with F-X272. F-X277a owns A277-1. This child introduces no public signature or source file/module/crate.

## Review lineage and readiness

All three original design peers remain exhausted. This split is engineering reconciliation, not a fourth peer or approval. Root affected final audit remains required. The original implementation cap remains three passes. The design remains draft and cannot be claimed until approved and registered. Separate independent implementation review precedes prepare. A same-session self-review is not that independent review.

No F-X272 provider is claimed. Existing checked source readers and mutation bounds stay unchanged. This leaf changes typed table construction, the number of authored TOC paragraphs and the set of parts visited by literal replacement. It does not add a namespace context, raw constructor, parser allowance, note ownership rule or style scan. Source call paths already exist at product Base. Final union validation remains necessary after F272 integration.

## Spec reference

- docs/hld/14-development-backlog.md, original F-X277 Issue432 rows A277-4,5,7,8 and proposed child registration.
- docs/hld/10-bindings-spec.md, Native Word facade stability, table mutation and both literal replacement paragraphs. Adjacent regex behavior and separate template/chart-data behavior remain.
- docs/hld/03-architecture.md, table authoring, TOC and story/text search architecture.
- Original Issue432 authenticated body and PR441 provenance retained by F-X277b. Source reconciliation below describes actual current source, not donor implementation or Word proof.
- OOXML paragraph-property order and complex-field begin/instruction/separate/end apply to generated output. Unknown raw children retain bytes and position. Existing schema validation rider remains mandatory, with exact validator version and generated candidate identity recorded at final union. No unverified primary-schema claim is made by this preparation.

## Approach

A277-4: In Table::merge_cells classify a cell only when content is exactly one paragraph. Clone that paragraph, clear properties and only root-attribute extra_xml records identified by CT_P::raw_is_root_attributes, then compare to CT_P::new(). All remaining payloads, including runs, hyperlinks, equations, controls, bookmarks, comments, revisions, ruby, unknown raw children and nested tables, count as content. Two empty paragraphs count as content. If all selected cells are lone-empty, preserve the origin paragraph including its properties and root attributes. Otherwise concatenate substantive cells in row-major order. Keep existing topology, opaque metadata refusal and candidate assign-back boundary.

A277-5: formatted_empty_cell keeps its existing cell-only tcPr clone and opaque-tcPr refusal unchanged, including the split_cell caller. Add one substantive private helper in existing table.rs, inherited_empty_paragraph(source: &CT_Tc) -> CT_P, used only by add_row and insert_column. It constructs the empty paragraph and transforms copied properties, rather than forwarding a call. Copy pPr from the source cell's last paragraph. Preserve its mark rPr. Only when the mark lacks rPr, use the last run's direct rPr. Clear sectPr, numPr, numbering-change records and pPrChange. Preserve other unknown pPr children. New cell has one empty paragraph. No list numbering or paragraph/root identity is copied. Both row and column routes call formatted_empty_cell for tcPr and use inherited_empty_paragraph to set the single paragraph. split_cell keeps its original formatted_empty_cell-only route and existing default paragraph formatting. No split behavior change is authorized. A zero-paragraph source remains a default empty paragraph. Existing source cell stays unchanged. Last paragraph is the inherited contract, not a native Word observation.

A277-7: If insert_toc sees no headings, append one empty paragraph before attaching field characters. Begin, instruction and separate stay on title paragraph. End belongs to following paragraph. Both heading-before and heading-after authored construction orders rebuild to Introduction, Methods. Existing foreign one-paragraph TOC refusal and message remain. R-A277-7 stays separately recorded. TOC hyperlinks remain bare. No F-X273 provider is needed.

A277-8: Remove only the chart-part literal-replacement loop. Preserve body, tables, controls, headers, footers, normal notes and text-box replacement, existing duplicate text-box counting rules, candidate atomicity, error behavior and regex behavior. Chart cache and embedded workbook remain unchanged. Template mail merge and chart-data APIs retain their separate chart behavior. No direct Document::text widening is implied.

The entire original replacement/text agreement obligation belongs to this child. Current CLI plain text calls text_report, which includes other supported stories. The historical assertion that CLI necessarily misses a text box because Document::text is body-only is superseded and cannot justify a waiver. The current generated body plus authored DrawingML/VML text-box discriminator is authenticated in /private/tmp/rdocx-s96-a2778-consumer-probe/PROBE_RESULT.json and ACTUAL_CLI_RESULT.json. Reviewed source84534f382a2b261b7a0e5c8d67471f746ef6ad37 has unchanged text/report/replacement and CLI source versus product ed588. Native body-only text count1, report count2 with no diagnostics, literal replacement count2 and reopened report count2 agree. Actual CLI text count2, replace --expect2 exit0 and reopened CLI text count2 agree. Serialized input has two physical text-box branches plus body, and both branches are rewritten while counted once. This closes the specific stale CLI mismatch premise. It does not qualify all imports, revisions, headers/footers, JSON, regex or final integrated source. Those remain regression/rider obligations below. No current broader implementation change is inferred from this single fixture. Never silently drop text-box coverage or mark a known mismatch passed.

## Consumer map

| Scope | Existing entry points | Actual source anchor |
|---|---|---|
| Cell merge | Table::merge_cells, Python Cell.merge | crates/rdocx/src/table.rs:1058, crates/rdocx-py/src/table.rs:2287 |
| Inherited formatting | add_row and insert_column use new private substantive paragraph helper after existing formatted_empty_cell, Python twins | table.rs:713,925,973, Python table.rs:1186,1233 |
| Unchanged split consumer | split_cell keeps original cell-only formatted_empty_cell behavior | table.rs:1234 |
| Authored TOC | insert_toc, rebuild_toc and existing Python wrappers | document.rs:44474, field.rs:3388 |
| Literal search | try_replace_text, checked batch/expected routes, CLI replace, Python and WASM replacePlaceholder | document.rs:44767,45319,46208, WASM lib.rs:107 |
| Comparison observations | CLI text plain/JSON, text_report, literal and regex count | CLI commands.rs:169,226, document.rs:82833,45860 |

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A277-4 | Merge horizontal and vertical, second empty, first empty, and both nonempty. Also comment range, revision, ruby, bookmark, and unknown raw child | Paragraph lists are `['Total']`, `['Sum']`, and `['Total', 'Sum']`. Both-empty remains the origin paragraph. Each non-empty payload keeps both paragraphs | `merge_cells`, `Cell.merge` | `issue432_merge_drops_only_a_lone_empty_paragraph` | Same Rust test. No pixel gate |
| A277-5 | `insert_column(1, before=False)` and `add_row` beside a shaded, centered, one-run bold cell. Also a two-paragraph cell whose last paragraph is left italic, and a source `w:numPr` | One-run cell: `w:jc val="center"` and `w:b` under `w:pPr/w:rPr`. Two-paragraph cell: left and italic. Source `w:numPr` is absent on the new cell. Unknown `pPr` children survive | Both table methods, Python wrappers | `issue432_new_cells_inherit_paragraph_mark_not_numbering` | Same. Q-S96-09 is the recorded last-paragraph choice, not a user question |
| A277-7 | `insert_toc` before headings, and `insert_toc` after headings, then `rebuild_toc` | Entry text is Introduction, Methods both times. A foreign one-paragraph TOC still errors with the existing message | `insert_toc`, `rebuild_toc`, Python wrappers | `issue432_rebuild_toc_both_insert_orders` | R-A277-7 stays a separate residual, not a pass. F-X273 is not evidence for this row |
| A277-8 | Body North once, chart categories North/South and series Sales, plus separate body/header/footer text-box fixtures and duplicate branch controls | Chart case literal/regex count1, CLI replace --expect1 succeeds, CLI text agrees, cache/workbook unchanged North and Sales. Actual supported-story text-box counts agree with replacement under existing branch-count convention | Native checked literal/batch/expected/regex, CLI text/replace, Python, WASM replacePlaceholder | issue432_replacement_skips_chart_cache_and_workbook plus actual text-box agreement controls | Installed bindings, actual CLI and integrated rider evidence mandatory. No text-box waiver |

## Behavior contracts and evidence

| Boundary | Contract and evidence status |
|---|---|
| Raw root versus content | Existing CT_P::raw_is_root_attributes is available at clean Base. Clear root records only for emptiness comparison, preserve origin paragraph and all opaque child bytes |
| Property ownership | Only row/column routes use the substantive inherited_empty_paragraph transformation, split_cell keeps cell-only helper behavior. Clone only existing paragraph property data into new empty paragraph. No source roots, original-byte parser credit, public ownership type or namespace context introduced |
| Numeric/ordering | Existing table indices, widths, topology errors and TOC max-level clamp unchanged. Concatenation is row-major. No unit conversion or new bounds |
| Search agreement | Current CLI/report and actual replacement observations mandatory. No inference from Document::text alone and no residual waiver |
| Atomic refusal | Existing candidate commit boundaries and opaque-table refusal retained. Error messages and quota formulas unchanged |

## Test matrix

| Case | Required discrimination |
|---|---|
| Merge | Horizontal/vertical, first/second empty, both text, all empty origin formatting, two empty paragraphs, nested table, comments/revisions/ruby/bookmarks/equations/controls and opaque raw child retained. Properties/root attributes alone remain lone-empty. Existing refusal leaves original unchanged |
| New cells | Before/after column and add_row, one/two/zero source paragraphs, mark rPr present versus absent/last-run fallback, source numPr/sectPr/pPrChange/numbering-change omitted, unknown pPr child retained, source unchanged. Existing opaque tcPr refusal atomic. split_cell beside paragraph-formatted merged source still produces default empty added paragraphs, matching pre-change bytes/semantics |
| TOC | Both authored orders, clamp/bounds existing behavior, original field layout and foreign single-paragraph error unchanged, saved/reopened rebuilt entry labels Introduction/Methods |
| Replace | Real chart/cache/workbook bytes before/after, Sales unchanged, literal/regex/native/CLI expected count, table/control/header/footer/normal-note/text-box supported story controls, duplicate branch count once, Python and WASM route agreement. Separator notes excluded as before |

## Test plan

Use existing table tests, crates/rdocx/tests/integration_test.rs, document.rs/field.rs unit carriers, existing CLI integration entrypoint, crates/rdocx-py/tests/test_core.py and existing WASM tests. Do not add a new test binary or product source file. Batch meaningful RED/GREEN cases by existing binary. Authenticate pre-repair failures separately from compile/setup errors. Scoped cargo check/test, fmt, clippy and prose checks are local. Python AST alone is not installed runtime evidence. Actual installed Python supported-version suites, WASM target check and runtime path, actual CLI observations, schema/raw preservation riders and exclusive hash gate are final integrated obligations. Do not report any as passed without actual receipts.

Full workspace/oracle/hash verification runs once on the integrated union per workflow. No worker records a baseline. Inspect actual integrated intentional deltas and label their commit, without inheriting an unverified historical five-key allocation. Deterministic fonts apply to any final rendering baseline. Native visual/pixel evidence is not newly promised for table paragraph count or TOC labels, and the original applicable rider union is not waived.

## HLD impact

- docs/hld/03-architecture.md
- docs/hld/10-bindings-spec.md

Update table merge/last-paragraph inheritance, authored no-heading TOC layout, and literal search/chart-exclusion/current story-count semantics only. Leave notes/style/numbering sections, OfficeMath and GLM shape sections to their owners. HLD04 notes stays F-X277b. HLD14 and sprint metadata are integrator-owned registration, not worker HLD implementation.

## Risk routing

| Matched canonical trigger | Reference | Additional check |
|---|---|---|
| Any parser or serialiser | .claude/skills/risk-routing.md:25 and HLD04/06 | Exact schema child order for generated pPr/rPr/TOC runs. Prefix-tolerant import, fixed-prefix modeled write and byte-preserved opaque cell/pPr child controls |
| Layout, pagination, line breaking, text shaping | .claude/skills/risk-routing.md:24 and HLD08 | Inherited paragraph formatting and split TOC can alter rendering. Use deterministic fonts for any baseline, declare actual integrated output deltas |
| Public API of a published crate | .claude/skills/risk-routing.md:28 and HLD10 | No signatures/public fields added. Declare intentional behavior at Word0.17.0. Final cargo publish --dry-run and crate-size gate |
| WASM or PyO3 bindings | .claude/skills/risk-routing.md:29 and HLD10 | Actual Python table/TOC/replacement consumers. WASM replacement check/runtime. Final wasm32 target check, workspace tests exclude both extension-module Python crates |
| Compatibility, numeric conversion or collection ordering | .claude/skills/risk-routing.md:33 and differential-testing compatibility facts | Row-major merge, retained origin properties, last-paragraph/last-run rule and unchanged width/index/clamp/error rules. Exact primary implementation provenance required for any external parity claim |
| XML creation or relocation in multiple parent contexts | .claude/skills/risk-routing.md:36 and design validation | New empty cell paragraph holds pPr with rPr in proper sequence, no sectPr/numPr/revision-property ownership. Authored TOC field end in following paragraph. Body/header/footer text-box controls retain legal placement |
| Shared semantic behavior with multiple consumers | .claude/skills/risk-routing.md:40 and differential-testing composition | Same fixtures across native/CLI/Python/WASM replacement and actual report/reopen observations, inherited table native/Python agreement |
| Grammar, scope, compatibility mode or empty collection | .claude/skills/risk-routing.md:41 and differential-testing composition | Zero/one/many source paragraphs, all-empty/lone-empty/contentful merge, zero-heading TOC, foreign TOC refusal and atomicity, duplicate text-box branch count |
| Completion status, capability ownership or evidence adoption | .claude/skills/risk-routing.md:42 and WORKFLOW acceptance contracts | Exact guard hashes, current metadata input revalidation, no provider completion and final policy checks after adopted ownership changes |
| Contributed PR or new issue acceptance | .claude/skills/risk-routing.md:43 and design acceptance criteria | Preserve Issue432 A4,5,7,8 in full, exact authenticated issue-body and latest PR-head pin revalidation before adoption. R-A277-7 preserved, no stale text-box waiver |

No new public types, parser quotas, traits, generics, source modules/files, feature flags or release-script changes. No new external parity claim is made by preparation. Any oracle used by final rider must have actual pinned version and candidate identity. All riders are additive and final union consumes them once.

## Dependency readiness

None for C product implementation. Existing clean Base already provides raw root-attribute classification, typed cell edits, authored TOC construction, native text_report, CLI text and checked replacement. The authenticated body/text-box fixture closes a specific historical false inference and supports this independence conclusion, not universal semantic readiness. F272 is unfinished and remains B and parent prerequisite. Parent completion and F280 gates cannot be unblocked by this child alone. Actual imported/tracked/header/footer/JSON/regex and final integrated acceptance remain mandatory test work. If implementation reveals a genuine new checked-source context requirement, stop and reconcile that exact contract with Root rather than borrowing unfinished C1 credit.

## Rejected alternatives

- Keep independent typed table/TOC/chart changes waiting for notes topology: no actual source/provider edge supports that delay.
- Use paragraph.text().is_empty() or a hand-maintained payload list: either discards raw or semantic zero-text content.
- Copy first paragraph, all numbering, section properties or revision ownership into new cells: violates inherited last-paragraph contract and new owner semantics.
- Broaden formatted_empty_cell paragraph behavior for its split_cell consumer: A277-5 authorizes row/column only. The separate substantive helper isolates transformation with two real callers today.
- Broaden rebuild_toc to imported one-paragraph fields: changes retained foreign refusal outside authored scope.
- Rewrite chart cache or workbook to restore agreement: category mutation belongs to chart-data API.
- Treat body-only Document::text as current CLI output, waive text boxes or generalize one fixture to all imported/revision variants: actual consumer evidence and complete acceptance forbid each shortcut.
- Reset parent design peers or count model self-review as independent microscope: inherited cap and actual independent review remain.

## Hash harness

Exclusive final integrated hash owner only. This worker neither records nor edits baselines. Table paragraph pruning/inheritance and empty-heading TOC can intentionally alter XML or rendering. Chart exclusion intentionally preserves charts. Label each intentional behavior change and review actual integrated deltas at final union. Historical aggregate five-entry claim is not a worker allocation or evidence. Baseline fonts remain deterministic. An unexplained delta blocks integration/release under the normal workflow.

## Implementation checklist

- [x] Root affected final audit and approved canonical design/registration, authenticated Issue432/PR441 current identities.
- [x] Clean metadata-only Base exception, normal claim and /start-feature --claimed. Exclude all unrelated C1 carries.
- [x] Baseline RED fixtures in existing carriers for merge, inherited formatting, both TOC orders and chart exclusion. Separate expected GREEN preservation controls from semantic RED.
- [x] Implement only four C-owned clauses, retain existing atomicity/errors/raw preservation and no new source files/modules/public surface. Keep formatted_empty_cell cell-only for split_cell, add substantive inherited_empty_paragraph only to row/column routes and lock split unchanged.
- [x] Actual local CLI JSON/plain, native report/literal/regex and body/header/footer text-box paired-branch controls pass. Native/Python/WASM source consumers are audited and public Python regressions pass AST. Installed runtime is not claimed.
- [ ] Final installed Python3.9/3.12 and typing/stubs, WASM target/runtime, SDK/raw preservation, native residuals, integrated full/hash, archives/dry-runs and hosted CI remain pending.
- [x] HLD03/HLD10 exact C sections updated, math/shape/notes/style/numbering ownership respected.
- [x] Scoped checks and authenticated GREEN receipts, self-review then actual independent microscope within original implementation cap.
- [x] /complete-feature --prepare only after independent review and scoped gate, with truthful pending final riders. Integrator owns totals, hash and final union.

## Coordination

Separate claimed worktree on clean metadata-only descendant of product ed588. Never transfer unfinished canonical C1 source or Python owner-test carry. Existing integration_test.rs, test_core.py, document.rs and HLD03/HLD10 are shared files in other worktrees, not simultaneous writers in this one. Integrator resolves separate semantic hunks explicitly. No math.rs/equation conversion or shape-conversion source lease. Root must authorize precise metadata-only Base exception and claim through normal lifecycle. Parent waits all children and F272. F280 barriers remain unchanged.

## Open questions and claim boundary

No new product-policy question is manufactured. The bounded A8 current native/CLI text-box discriminator is authenticated as preparation evidence. The broader acceptance matrix and final integrated consumers remain unperformed. This is a draft proposal, not final acceptance. Root affected final audit, canonical adoption/approval and normal claim are still due. No source, provider, model or runtime completion is recorded here.

## Root affected final approval and claim

Codex affected final1 accepts this bounded contract with zero findings. Original three design peers remain consumed. Current Issue432 body and PR441 head are unchanged in the authenticated12:48UTC GET-only refresh. The prior draft/preparation statements above are historical. Root approves product implementation on a guarded clean metadata-only Base, with no F272 provider completion, no imported/story waiver and all final riders pending. Actual claim owner is codex in work/f-x277c-codex at /private/tmp/rdocx-s96-f277c-codex. No unfinished canonical C1 source enters that Base.

## Local preparation evidence

Root implemented this child and recorded source-pinned old-producer semantic RED and fixed-source GREEN in the existing progress record and /private/tmp/rdocx-s96-f277c-implementation. Identified Codex self-review1and actual independent microscope1are clean. Five unit, four native and one CLI new cases plus four native and three CLI existing controls pass. All-target/all-feature native/CLI Clippy, fmt, prose, whitespace and Python AST pass. This is local preparation only. The plan stays approved, delivery stays in-progress and every named final rider remains pending. Four labelled source commits retain distinct intentional behavioral changes, with final cumulative source equal to the reviewed candidate. No provider completion or sprint totals are recorded.
