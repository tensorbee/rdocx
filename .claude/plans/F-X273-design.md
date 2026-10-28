# F-X273, Repair header body clearance and complete TOC navigation

**Status**: approved
**Sprint**: S96
**Size**: L
**Depends on**: F-283, F-X264

## Problem

`crates/rdocx-layout/src/paginator.rs`, `Pager::begin_page`, starts body flow at written margins even when header/footer flow occupies that band. PR428 head `87f101bc8e688bd93d358b72bc69c28732e9dad6` partly repairs this, but discloses omitted paragraph spacing, incorrect floating-owner treatment and absent native Word observations. Its no-`\u` styled-heading TOC case also remains broken. In `engine.rs:561`, source registration traverses table cells whose content table layout may never place, leaving qualified navigation facts incomplete.

## Spec reference

- `docs/hld/08-rendering-spec.md`, "Word section geometry and page numbering", "Word bookmark field pagination", "Word table topology and physical pictures".
- `docs/hld/03-architecture.md`, physical source owner and cached layout authority.
- `docs/hld/04-opc-and-packaging.md`, "Producer read and writer authority".
- `docs/hld/10-bindings-spec.md`, Python layout/TOC and current-source qualification.
- `docs/hld/12-testing-strategy.md`, deterministic rendering, Word live-story and page-authoring controls.
- `docs/hld/14-development-backlog.md`, F-X273.

## Approach

Adapt only PR428's own variant-selection refactor, geometry and navigation increments, preserving separate labelled behavior commits and current main tests. Do not copy archive measurements. Reconcile after F-X272 because shared regression and Python carriers are exclusive resources, though these are not formal dependencies.

One private selected-story variant controls actual header, footer and watermark selection. A selected missing/blank first/even story must not accidentally measure the default. Measure and paint the same paragraph spacing, border-group gaps, table rows and selected fonts. Keep section/furniture geometry separate from the effective body band. For nonnegative top/bottom margins, including zero, clear against actual selected flow extents. Negative margins use their absolute magnitude and retain exact placement. Do not count out-of-flow floating tables/frames/anchored drawings as ordinary flow. Trace their wrapping and margin anchor positions through the existing owners rather than blanket-dropping them. Qualify frame/table containment, page/text border geometry and margin-relative anchors against actual Word before final acceptance.

Carry the resulting band through fit, columns, line numbers, vertical alignment, notes, endnote continuation and page fields. Invalidate all affected header/footer and restart/substitution cache inputs. Warm incremental and cold fresh layouts must agree after a same-owner fact-changing spacing/text/variant edit. No change to font discovery or numeric unit constructors.

Give every physical table paragraph source-qualified navigation facts even if a merge continuation or accepted-view deleted row is unplaced. Preserve source identity and freshness checks. Exclude only proven unplaced continuation/deleted-row content from TOC selection regardless of whether selection came from outline level, style or explicit switch. Preserve eligible placed cell/nested-table headings. Keep STYLEREF and numbered-outline consumers consistent. Do not solve the donor gap by dropping all unplaced/hidden headings or bypassing stale-fact validation.

No new public authoring operation. Existing Rust `Document::rebuild_toc(&mut self) -> Result<TocRebuildReport>` and Python `rebuild_toc(self) -> TocRebuildReport` keep signatures and atomicity. CLI `rdocx toc rebuild` and existing layout/render APIs use the same native providers. Any proposed helper stays concrete in existing files. No new source file/module/crate/trait/generic/feature flag.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Selected furniture flow | `layout_header_footer_variant`, paginator variant selection, `render_hf_blocks` | Same selected first/even/default source and spacing policy | Measure, paginate, paint, saved/reopened source |
| Body band | `sect_pr_to_geometry`, `Pager::begin_page`, note/endnote pagination | Section geometry remains furniture reference, effective band controls flow | First/last body positions, note/anchor/border positions and page count |
| Table navigation | `SourceRegistry::collect_table`, paragraph navigation, `toc_paragraph_level` | Physical source identity plus proven placement, same accepted view | Registration, layout, TOC rebuild, STYLEREF, numbered outlines |
| External consumers | Rust deterministic layout, Python fragments/rebuild/render, CLI layout/toc/convert | Shared native source-qualified bundle | Same input before/after rebuild, update, save/reopen, PDF/PNG |
| Cached layouts | Native and Python held/fresh queries | Existing source stamps and cache key owners | Same-owner text/spacing/variant edits, warm/cold comparison |

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| Margins | Nonnegative top clears flow, negative top uses exact absolute distance. Bottom is symmetric. Twips to points remains /20 | ECMA 376 Part 1, fifth edition, clause 17.6.11, printed pages 579 to 583, source pinned below. Word fixture qualification remains due. Microsoft OpenXML3.0.1 `PageMargin` identifies `w:pgMar`, https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.pagemargin?view=openxml-3.0.1 | Positive/zero/negative with tall/short selected story, compare positioned coordinates |
| Header spacing | Painted and measured flow include actual effective paragraph spacing and borders, tables price cells once | Existing paragraph/table measure and paint paths, native Word remains deciding | Five Normal paragraphs with 8 ptafter cannot remain the donor 108.5 ptresult merely because overlap disappeared |
| Floating owners, anchors and borders | Separate flow extent from positioned paint, written furniture geometry and body clearance | Applicable OOXML primary clauses and actual Word controls required before parity claim | Flow/floating table/frame pair, margin/page anchor and text/page border pair |
| Navigation | Continuation/deleted-row paragraphs retain current qualified identity but contribute no heading/page target | Existing `WordLayoutResult::paragraph_navigation` contract, PR428 partial discriminator | Styled and plain omitted rows, with/without outline switch, placed neighboring heading |
| TOC transactions | Missing/stale source still refuses atomically, valid omitted cell is not a missing source | Current checked generated-table contract | Unplaced positive paired with stale-fact negative |

## Dependency readiness

F-283a nd F-X264 are done. Final native Word and pinned LO observations are pending, not source-qualified acceptance. The F-X274 formal barrier requires this story's genuine local gate, integrated source ancestry and zero-finding microscope. For that checkpoint, this design approves a LOCAL regression implementation gate consisting of A1 toA6 below. The backlog's differential final gate remains pending until the final integrated union. Do not use pending native evidence as a passed provider gate.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Full Issue 424 header/logo/table reproduction and five spaced Normal header paragraphs | Body begins below actual flow, short header leaves band unchanged, measured/painted extents agree | Rust fragments, Python, CLI layout/PDF | Adapt donor tall-header case plus spacing/table discriminators | Native Word and pinned LO coordinates/page count |
| A2 | Positive/negative top and bottom, first/even/default, absent/blank selected story | Positive margins clear selected flow only, negative sides retain absolute distance | Paginator/render | Selected-variant and signed-margin local cases | Word default/first/even physical page controls |
| A3 | Flow/floating table and frame, behind/front/margin-relative drawing, text/page border | Only true flow consumes band, positioned furniture and anchors follow qualified owner policy | Layout, paint and notes | Paired flow/floating/anchor/border cases | Primary clauses plus Word positions, no guessed parity |
| A4 | Tall footer with footnote/endnote, overflow, line numbers, columns/vertical alignment | Notes and last body marks fit effective band without moving unrelated furniture | Native layout/PDF | Donor endnote offset plus bottom-clearance cases | Combined Word/pinned-LO page and note positions |
| A5 | Original Issue 426 merged/unmerged, deleted row, nested table, plain/styled omitted content | All qualified source paragraphs represented, omitted content never selected, placed headings retained | Rust/Python/CLI TOC | Extend `toc_outline_switch_rebuilds_past_cells_table_layout_drops` | Original issue scripts plus installed runtime |
| A6 | `\o`, `\u`, `\t`, `\h`, `\z` risk-based combinations, STYLEREF and numbered outline | Correct visible entry set/targets and source freshness, no blanket exclusion | TOC and adjacent field consumers | Styled no-`\u` failure repaired, stale-source refusal control | Saved/reopened TOC/native output and combined page targets |
| A7 | Same-owner spacing/text edit and first/even variant edit with held/fresh queries | Warm and cold layout agree in band, page count, navigation and text | Rust/Python caches | Real fact-changing publication controls | Installed source-qualified live-handle suite |
| A8 | Revised merged-table document with tall header, roman/letter footer | Compare then rebuild/update/save/reopen/render yields coherent physical pages and TOC targets | F-X272/F-X273/F-X274 union | Focused integration after providers | Actual integrated native/PDF/text/PNG union |

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | Tall/short, positive/negative, first blank/default tall | Only selected positive-margin flow changes band | Always measuring default or using negative signed coordinate |
| C2 | Five spaced paragraphs versus Header style, table cells versus floating table/frame | Price effective flow exactly once | Clearing painted lines while omitting spacing, counting floating paint as rows |
| C3 | Styled continuation/deleted row versus visible nested-cell heading, `\u` present/absent | Omitted owner skipped for every selector, visible heading retained | Fixing outline-only or dropping all table headings |
| C4 | Fresh facts versus changed source stamp | Valid omission succeeds, stale bundle refuses | Disabling freshness to avoid TOC failure |
| C5 | Warm layout after same-owner edit versus cold layout | All geometry/navigation/field outputs equal | Cache keyed only by identity/revision or source count |

The matrix is risk-based. It does not promise every Cartesian switch/owner combination. All original issue reproduction combinations remain required.

## Rejected alternatives

- Clear only the header lines currently painted. The donor discloses the spacing defect this leaves.
- Skip all unplaced headings globally. It changes unrelated hidden-heading behavior and weakens source qualification.
- Treat a different LibreOffice release as Word evidence. It proves neither required oracle.

## Test plan

**Test gate**: differential, verbatim from backlog. Final native/pinned-oracle debt is retained.

Approved genuine LOCAL provider implementation gate: current-Base RED/fixed GREEN for A1 toA6 and the applicable A7 cache regressions, through selected `cargo test -p rdocx-layout --lib <case>`, `cargo test -p rdocx --test regression_test toc_outline_switch_rebuilds_past_cells_table_layout_drops`, selected CLI TOC tests and freshly built selected Python Issue 424/426 nodes. Batch related cases in existing binaries. Scoped check/clippy/fmt and independent microscope remain required. This approved local gate does not claim differential completion and is required before F-X274 can start.

Final sprint: native Word with actual version, input/output hashes, production steps and positions, LibreOffice 26.2.5.2 and Poppler 26.01.0, deterministic library PDF/PNG/text, installed 3.9/3.12 and strict typing/stubs, full risk union and 49 hashes/7 pixels. Predeclared native geometry tolerance is 0.5 pt per measured text/furniture edge with identical qualified font files, and exact page counts and entry sets. Use extracted PDF coordinates where available. At 144 DPI this allows one raster pixel. Record font hashes and native layout inputs. A mismatch beyond tolerance remains a failed criterion, without post-measurement tolerance changes. Unsupported primary/native cases remain pending rather than silently dropping them.

## HLD impact

- `docs/hld/08-rendering-spec.md`

## Risk routing

Layout/pagination adds deterministic fonts and explicit output deltas. Units adds unchanged truncation plus signed-margin boundary controls. Shared semantics/caches adds warm/cold and evaluate/rebuild/update/reopen/paint comparisons. Contributed acceptance retains omitted issue rows. External oracle requires exact source/tool/artifact provenance. Binding consumers add final WASM compile and installed typing/runtime. Affected public behavior adds current archive verification under 10 MiB. Parser/serialized source preservation controls remain where navigation/section inputs are touched.

## Hash harness

Variant refactor must be unchanged. Geometry and navigation are separate labelled intentional changes. Geometry may alter tall/spaced furniture pages and corpus pagination, including the donor's disclosed `redlined_no_footer.docx` and `form_footnotes.docx`. Donor claimed 49 hashes unchanged, but this does not predict the complete spacing repair. Inspect all 49/7 actual deltas before changing any baseline. Every changed case needs a reviewed expected-output account. No incidental re-recording authorized.

## Implementation checklist

- [ ] Close primary semantics and predeclare native numeric tolerances.
- [ ] Reconcile PR428 own commits after F-X272 without overwriting tests/measurements.
- [ ] Keep variant refactor and each intentional behavior change separately labelled.
- [ ] Price and paint the same complete selected flow.
- [ ] Qualify floating owners, anchors, borders and notes without scope omissions.
- [ ] Repair every source-qualified TOC selector and preserve freshness refusals.
- [ ] Run local regressions/scoped gate/microscope and preserve final oracle debt.
- [ ] Update listed HLD intent and authenticate provider ancestry before F-X274.

## Open questions

No new material product choice is proposed. Native Word observations and primary source qualification are factual work still due. Approval of this plan explicitly authorizes the named genuine LOCAL provider gate while retaining the differential final gate.

## Contract closure

No public signature change or new source carrier planned. Consumer coverage, callable APIs and local/final evidence are mapped. Primary clauses and predeclared native tolerance are recorded. Native observations remain required final evidence rather than unmade design choices. F-X274 must not interpret provider design approval as completed implementation.

## Intake identity

Read-only GitHub refresh on 2026-10-28 found no body or contribution-head changes from the S96 intake. Original payloads and comments are retained under `/private/tmp/rdocx-s96-run-refresh`.

- Record 424, updated `2026-10-27T20:08:03Z`, body SHA-256 `c66665fffd68be5e4469d1557532f3cc30caadabf10f7b03b665a5683c4d9f83`.
- Record 426, updated `2026-10-27T20:08:06Z`, body SHA-256 `36dff32b59ca3099e25d2b4e6a5a0e8d03624504fd47f2110aed4f5e1ab6bee7`.
- Record 428, updated `2026-10-27T21:22:09Z`, body SHA-256 `635213f1f93940b66c2856d18a2d60caaf246d94419a40db90a1ec3b4e5aa5ba`.
  Contribution head `87f101bc8e688bd93d358b72bc69c28732e9dad6`.

## Primary specification pin

Official ECMA376 Part 1, fifth edition December 2016, retrieved 2026-10-28 from https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip. ZIP SHA-256 `9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129`,43501721 bytes. PDF and extracted text retained in `/private/tmp/rdocx-s96-run-refresh`.

Clause 17.6.11, printed pages 579 to 583, directly confirms the nonnegative top/bottom clearance rule and exact negative-margin distance. Zero needs an explicit control. This closes the basic signed-margin source fact, not the still-pending floating/frame/anchor/border native qualification.

Primary context clauses: 17.3.1.11, printed page 208, distinguishes framed paragraphs from ordinary flow and groups adjacent equal frames. Clause 17.4.57, printed pages 442 to 448, defines floating tables as outside main flow and gives page/margin/text position bases. Clause 20.4.2.11, printed page 3122, makes `positionV` use its declared base. Clause 17.6.10, printed pages 577 to 579, distinguishes page borders measured from page edges versus text margins. Preserve the written section/furniture reference for margin-relative geometry, with the effective body band controlling body flow. Actual Word positions must qualify that composed policy. A contradictory native result requires revising affected semantics, not recording a pass.

Design closure on 2026-10-28: consumer map, observable A-IDs, discriminating cases, public signatures, imported states, HLD impact and local/final evidence were checked against current source. Q-S96-01 is answered in F-X272. Approvals follow provider contract order. No material user decision remains. Pending native/hosted/installed observations are retained final acceptance debt, not recorded passes.

## Implementation-time contribution refresh

Read-only refresh on 2026-10-28 found PR 428 rebased onto released main `24af95d3d94c25e95f61a75e7e21f18fa95b6363`, with current head `d3f71b6dcb0163793c7df0e40c0fb19ef12bf236`, updated `2026-10-28T06:43:09Z`. Its body SHA-256 remains `635213f1f93940b66c2856d18a2d60caaf246d94419a40db90a1ec3b4e5aa5ba`. Historical initial pins above remain attributed to that intake.

Its own commits are variant refactor `6ee32977d937a0f9187dd15ad8b2d1933a76daa9`, header geometry `143a1b016419a109dfad168c336d3c2a6d418413`, navigation `78a737ecdc2cb7b8170d3386c198149ee38a2cff`, Python tests `050e222b5a5bcbb02a210f1ac422ed6e9ed99de2` and excluded archive measurements `d3f71b6dcb0163793c7df0e40c0fb19ef12bf236`. Per-file added/removed payload comparison confirms identical product source, Python tests and HLD changes. The native regression payload adds only one blank line. README/archive measurements differ and remain excluded. The approved scope and disclosed omissions are unchanged.

Exact refreshed metadata, commits, diff and comparison receipt are retained under `/private/tmp/rdocx-s96-run-refresh/428-next-wave*`. Reconcile these own changes after F-X272 without replacing current tests or copying measurements. No acceptance or completion follows from the rebase.
