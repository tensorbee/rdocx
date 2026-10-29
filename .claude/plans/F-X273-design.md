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

- [x] Close primary source facts and predeclare native numeric tolerances.
- [x] Reconcile PR428 own commits after F-X272 without overwriting tests or measurements.
- [x] Keep unchanged variant refactor and intentional navigation and geometry separately labelled.
- [x] Price and paint the same complete selected flow, with actual local controls.
- [x] Implement floating owners, anchors, borders and notes with the named local controls.
- [x] Repair every source-qualified TOC selector and preserve freshness refusals.
- [x] Run local regressions and scoped checks, obtain clean independent pass2 and preserve final debt.
- [x] Update the listed HLD08 implementation intent.
- [ ] Complete native floating, anchor, border, note and numeric qualification in the final sprint union.
- [ ] Authenticate integrated provider ancestry and execute the approved local provider checkpoint before F-X274 claim.
- [ ] Pass every final acceptance rider in the approved test matrix and sprint gate.

## Open questions

No new material product choice is proposed. Native Word observations and primary source qualification are factual work still due. Approval of this plan explicitly authorizes the named genuine LOCAL provider gate while retaining the differential final gate.

## Contract closure

The approved doc-hidden native omission query is implemented without changing the existing navigation tuple or authoring signatures. Concrete private carriers serve the mapped consumers. Consumer coverage, callable APIs and local/final evidence are mapped. Primary clauses and predeclared native tolerance are recorded. Native observations remain required final evidence rather than unmade design choices. F-X274 must not interpret provider design approval as completed implementation.

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

## Qualified table omission provider

Read-only consumer tracing confirms that the existing navigation tuple cannot distinguish a proven table omission from other unplaced paragraphs. Preserve its public tuple shape. Carry the proven omission as a concrete private navigation fact through the existing registry, engine and WordLayoutResult providers. Add one implementation-facing doc-hidden query in the existing lib.rs: `pub fn paragraph_omitted_by_table(&self, node: SourceNodeId) -> Option<bool>`. Some(true) means this exact result has qualified evidence that table layout omits that physical paragraph. Some(false) means qualified current navigation exists without that omission. None means the source has no qualified navigation fact. Unknown identity must never become permission to bypass freshness validation.

This is an additive native query for the existing TOC caller, not a new authoring operation, Python/WASM method or replacement navigation API. No existing signature or public tuple changes. Its semver disposition is additive and implementation-facing. A private fact can use the existing ParagraphNavigation and WordLayoutResult carriers, which already have private fields. No new file, trait, generic or crate is needed.

Publish omission only from the actual table-placement decision and its revision view, retaining physical identity through nested cells and controls. A lexical vMerge=continue alone does not prove an orphan continuation was dropped, and absence from paragraph_pages alone does not prove a table omission. Preserve actual placed navigation and neighboring visible/nested headings. The TOC caller checks this qualified fact before every style, outline or explicit selector while retaining its existing source freshness and atomicity requirements. Existing STYLEREF and numbered-outline consumers must observe the same underlying placement policy. Add an applicable orphan/placed control if table layout admits it, alongside the already required continuation/deleted-row and stale-source cases. This concretizes A5/A6 without claiming native parity or altering the approved final gates.


## Ordinary frame boundary

The A3 floating-frame repair covers ordinary positioned paragraph frames. ECMA376 Part1 clause17.3.1.11, printed page211, gives dropCap a distinct next-non-frame-paragraph positioning role. Absent dropCap, or its explicit none value, is an ordinary frame. Drop and margin drop-cap modes require relative glyph/line treatment that the existing renderer does not implement. HLD02, paragraph following the Word formatting capability overview, retains those existing layout boundaries.

Do not expand F-X273 into general drop-cap rendering or authoring. Preserve imported source and the existing unsupported drop-cap layout fallback. An explicit drop-cap mode must not accidentally enter the new ordinary positioned-frame path. Record this limitation in the existing HLD08 intent update and progress, with a discriminating drop-cap versus ordinary-frame local control. This does not waive any named A3 ordinary floating-frame, grouping, wrap, anchor or native acceptance case. All existing final gates remain required. The implementing owner acknowledges this canonical clarification without changing Base or replacing in-flight work.


## Word frame default compatibility

The ordinary-frame provider must distinguish ECMA defaults from documented Word behavior. Microsoft's [MS-OE376 framePr implementation note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6ff920ba-f035-407d-aa15-c4e2ad916cb6), section2.1.48, paragraphs e to h, updated2022-08-16 and independently checked2026-10-28, documents text as Word's omitted horizontal and vertical anchor default. It also documents atLeast for an omitted height rule with nonzero height, and automatic height for zero height. ECMA's corresponding omitted defaults are page anchors and auto height rule.

Use the documented Word-facing defaults in this compatibility repair, retaining the imported model and raw attributes unchanged. Explicit anchors/rules retain their existing meanings. Add current-source default-form RED/GREEN controls for omitted anchors, absent rule with nonzero height and zero height, alongside explicit page/text controls. Record the ECMA divergence in HLD08. Preserve every earlier explicit-context observation with its actual source applicability. This factual reconciliation does not authorize a general frame or drop-cap authoring API. Actual Word16.114 positions and the pinned LO comparison remain pending final qualification, and a contradictory native observation requires reconciliation rather than a parity claim.


## Footer preparation consistency and bounded failure

A focused independent source advisory found no reusable bottom-anchored furniture solver. Prepare one private selected per-page furniture result before body clearance, then reuse that exact result for paint. Keep original cached source blocks authoritative. Footer preparation must satisfy origin equals page height minus footer distance minus prepared flow height, with stable prepared line breaks, obstacle reservations and float placement. Equal height alone does not establish a consistent result. A failed reflow or an unchanged-result sentinel is not proof of successful preparation.

Use the existing fallible Engine layout route and LayoutError for a preparation that cannot establish consistency. Detect repeated states and stop after at most32 candidate preparations as an internal termination guard. The bound is not a native Word maximum, a performance acceptance budget or evidence of convergence. Each candidate starts from original blocks. A repeated state, exhausted bound or actual preparation failure returns an explicit section/page-qualified layout error. Never paint a guessed last candidate, silently drop wrapping or report a cycle as successful qualification. Preserve the public low-level pagination signatures and their existing positioned-furniture failure policy, since they carry no new private positioned-furniture metadata. The approved ordinary-flow clearance and spacing repair still reaches every applicable caller.

Local controls must distinguish fixed-origin header preparation, footer height changed by page-relative wrapping, paragraph-relative placement, stable measured/painted reuse and repeated-state refusal without source mutation. Every named supported A3/A4 example must actually stabilize and pass. Refusal cannot satisfy those examples or waive final Word/pinned-LO qualification. This is a bounded implementation decision using the existing fallible layout boundary, without a new authoring operation or public API.


## Pre-review local implementation evidence

Base remains `564ad7f5d1d57a7828c721eba58cbe79da10aae3`. The unchanged private variant refactor is separately labelled at `3510ec5ee2c12f8aee9bc52e4434624532f160d8`, qualified by the same three before/after cache/variant controls and scoped check. Intentional geometry and navigation increments remain separate labels. Product input changes exclude the fifth archive-measurement donor commit.

Actual Base RED is preserved under `/private/tmp/rdocx-s96-run-refresh`: initial lower7 gives4failures/3controls, native20selector contexts give16failures/4controls, authenticated predecessor Python7 gives3failures/4controls, and authenticated Base CLI7 gives4failed criteria/3controls. The CLI generator retains original five Normal paragraphs with8point after spacing and full descriptive table/1.5cm logo input at2cm top margin. All four original Issue426 merged/unmerged and outline/no-outline combinations remain, with a separate styled no-outline discriminator. Predecessor artifact identity is verified against Base product files, not represented as a current-source wheel.

Pre-review local lower qualification is32/32GREEN in1.09seconds in `fx273-complete-lower-green.log`. It includes actual selected spacing, nonzero border groups, multi-paragraph cells, frames/defaults/drop-cap fallback/wrapping, footer stable-state/refusal, signed margins, written anchors/page-text borders, columns/alignment/line numbers, footnote/endnote overflow and physical/displayed field provenance, navigation/cache/source-stamp controls. Intermediate failed compiles, setup aborts, test-oracle mismatches and source repair failures remain attributed in the existing progress record. They are not relabelled as semantic passes.

Pre-review local native5 and CLI2 pass, including original selector controls, styled omission, full logo/table layout/PDF and saved/reopened TOC. The same seven original source-built Base CLI inputs also pass against the authenticated current CLI, with unchanged inputs. Scoped all-target check, strict clippy, fmt, prose, adapter drift and whitespace pass. Receipts are `fx273-final-local.json` and `fx273-current-cli-original-green.json`.

Fresh dev wheel SHA256 `e48be80e7e666b38341f0abd25ae268f05378d08fec5036534ae4dfbf5a9c175` installs into an isolated Python3.12 environment. Actual imported native SHA256 `b236a165dadd44330fca9c5058899ea658a66300b51a25d0f35f91d77cffbba8` matches its wheel payload. All227product paths match before build, after build and after tests, manifest `be5dff87779bb572e84a861abd786058fc74d4ec9c98a1a26f7ef3e23f936607`. Initial fresh22selection gives21passes and one incorrect short-fit assertion. Saved source confirms document defaults8point after spacing, so one Normal paragraph actually clears to58.490763rather than fitting within2cm. The corrected carrier retains that nonzero clearance and adds an explicit zero-after fitting control, with existing tall/negative controls unchanged. Only that test function changes and passes1/1on the same authenticated product. Applicable unchanged21plus corrected1qualify all22selected cases, not one22pass execution. `fx273-provider-corrected-consumer-auth.json` proves the prior source reconstruction hash and unchanged other test functions. The failed full attempt remains intact.

HLD08 describes current implementation intent. This paragraph records the pre-review evidence stage. Subsequent clean independent pass2 and separate intentional labels are recorded in the preparation checkpoint below. Every final native Word/pinned LO/Poppler, full workspace/native, installed3.9/3.12, typing/WASM,49hash/7pixel, archive/README, risk and F-X272/F-X273/F-X274 union rider remains pending. This evidence is local provider readiness work and does not grant final acceptance.


## Independent pass1 and investigated successor

Full independent pass1 reports D1/D2, zero smells/nitpicks. D1 bypasses existing table-only overlap settlement. D2 repeats incomplete reservation output. Actual current-source RED contains12forbidden table overlaps, partial78point clearance inside a second70to130point object, and6later-line square reservations missing108points. Both-allows/disjoint and valid touch controls remain positive. A first whole-paragraph strict repair passes initial cases but later-line source-contract4fails. A second metric-encoding approach gives actual146highlight/link/rich occupied-band mismatches. Root designer and independent reviewer investigate before the successor, preserving both failures and actual source identities.

The successor settles each physical table once and transports private line gaps separately from unchanged Word metrics. Every occupied-band consumer sees the moved origin, while border/changebar/following flow price gaps once. Strict final-band reservations are positively checked against line-break inputs, with a separate32rebreak work guard. Original page32candidate guard and cycle refusal remain. No body approximation rewrite or public model/API change occurs. Initial investigated5cases pass0.25seconds, followed by bounded43lower controls passing1.10seconds, including all previous32, all new D1/D2 cases, existing body/variant controls and actual Engine PAGE/NUMPAGES substitution with story placement. These are internal/source-contract observations, not native Word parity. The formatted investigated successor passes43/43lower controls in1.09seconds and current CLI2/2in5.67seconds, plus scoped all-target check, strict clippy, fmt, prose, adapter and whitespace gates. The seven immutable original Base CLI inputs also pass against the current authenticated application. Native5navigation observations are consumed as predecessor executions only: query/provider/test source is unchanged, inputs have no furniture or wraps, and body policy/empty-gap view behavior remains equivalent. Their original source/binary identities are preserved in `fx273-review1-final-local.json`, not relabelled as current execution.

A NEW dev artifact passes the full current22Python selection in45.81seconds,647deselected. Wheel SHA256 `70e5dd68aa4a08f38d4e4d9f0c8a527b17ac0d0dd578900934706b039c5abf77` and isolated native SHA256 `1d84a2b936243efd7c20bdc2a3e826b4c046c96cc1d3933bd9ffe33b0bf26156` match. All227product paths remain exact before build, after build and after tests, manifest `377e3fc0d520b9bd37102053ead47947b223af4a2ba6fab873390ac5affcb7ef`. Selected Python source is unchanged during build/test. Receipts are `fx273-review1-provider-python.json` and `fx273-review1-final-local.json`. Root independently authenticates current source/local artifacts in `fx273-review1-root-auth-before-python.json`.

At this historical successor stage, D1/D2 implementation and local observations were ready for independent full pass2. That later pass and the recorded labels are described in the prepared checkpoint below. Previous frozen receipts and failed attempts remain historical and applicable only to their named scope. All final sprint debt remains pending.


## Prepared local handoff checkpoint

Independent `.claude/reviews/F-X273-all-pass-2.md` reports0defects,0smells,0nitpicks against the complete frozen Base diff. Root read the full report and restored a preparation-only lease. Unchanged refactor `3510ec5ee2c12f8aee9bc52e4434624532f160d8`, intentional navigation `af193d6d` and intentional geometry `468f3f41` are separate coherent increments. Commit decomposition changes no reviewed product, test or HLD bytes. No passing gate or confirmation microscope is repeated.

The actual current lower43, CLI2 and fresh Python22 observations remain qualified by `fx273-review1-final-local.json`, `fx273-review1-provider-python.json` and independent complete root authentication `fx273-review1-root-auth-complete-local.json`. Product227 manifest remains `377e3fc0d520b9bd37102053ead47947b223af4a2ba6fab873390ac5affcb7ef`. The five native navigation successes remain carried predecessor observations with original manifest `be5dff87779bb572e84a861abd786058fc74d4ec9c98a1a26f7ef3e23f936607` and binary `41ed749e6be619af5660242f08b4c708d5a0c696575aa81f775101e5423915cf`. Query, navigation provider and test source are unchanged, those inputs have no furniture or wraps, and body policy with empty-gap ParagraphView remains equivalent. They are not called current-source executions. Fresh wheel20TOC contexts provide current consumer coverage. Exact names and commands remain in the receipts.

Preparation uses `sprint-final-acceptance`, keeps this plan approved and retains Base `564ad7f5d1d57a7828c721eba58cbe79da10aae3`. Actual native Word16.114 and pinned LibreOffice26.2.5.2/Poppler26.01.0, numeric/page/entry parity, full workspace/native and hosted gates, installed Python3.9/3.12, typing/stubs/WASM,49hashes/7pixels, genuine archives and10MiB ceiling, README measurements, supply chain/risk and F-X272/F-X273/F-X274 union remain pending. Geometry/navigation intentional deltas require actual final harness review. No baseline recording, final oracle pass or delivery completion is claimed. Root owns integration and the approved provider checkpoint before dependent claim. Progress remains live until root copies it.


## Integrated local provider checkpoint

Reviewed worker tip `711e7d2c3f83b0fbb38b1f3b4f757d405be62753` is integrated at `fca298cc1a90dd77682ebbe6b652c56dd74bd228`. All 227 product files and every reviewed product, test and HLD byte match the qualified worker. Four plan conflicts were blank-line or annotation-only differences. Canonical contracts are retained and historical readiness wording is explicit. The structured handoff was validated and consumed. No source reconciliation or duplicate product gate was needed.

The approved genuine LOCAL provider gate and independent microscope pass2 have actually passed. Existing run state and BACKLOG/CURRENT_SPRINT record provider implementation completion, clearing its owner solely to satisfy the formal F-X274 dependency. This plan stays approved. Every named final native/full/risk/A8 rider remains pending, and AS_BUILT/SPRINT_TRACKER entries wait for actual final union acceptance. The consumer Base must descend from the integrated source, with ancestry checked before claim. The clean worker worktree is removed after preserving progress, and its branch is retained.

## Final native note contract reconciliation

Final A4 measurements on 2026-10-29 expose inherited note policies that violate the retained source and numeric contract. This is bounded remediation within A4, without a new authoring API, model variant, crate, module or file. The exact native input is footnote-130.docx SHA2567d502ab5610e2573fb815fbfb1b196b46482dafe4aec922339145e9e6caa4daa. Actual Word16.114 and pinned LO both place its reference on page1 and emit three pages. The current library strands the reference on page2 and emits four. Root and the separate designer/independent investigation agree on the source cause before another repair. This is an inherited defect exposed by the approved fixture, not a newly introduced regression.

The source has130Normal paragraphs with exact12pt lines, no indent and no footnoteRef event. Current unconditional12pt note-width reduction/paint offset and synthesized marker do not represent that source. ECMA376Part1 clause17.11.13 makes footnoteRef the presence of the note-body mark. Clauses17.11.23 and17.11.1 define authored separator and continuation-separator marks. Existing run raw-child preservation retains these events and their positions. Use a private source-qualified rendering projection through existing paragraph/run/style and note-label owners. Marker absence must remain absence. A present note-body marker uses its actual run boundary and effective style without becoming a self-referencing body note event. Full section width plus resolved paragraph indents/tabs controls note measurement and paint together. Keep package ID, displayed label and physical source identity independent. Inspect existing native-created note source and preserve its documented authoring behavior through explicit source construction where required, without inventing an imported marker.

Authored separator-only stories must not be discarded by text-emptiness filtering. Price and paint their qualified events once using resolved paragraph metrics. Preserve literal authored separator content, selected continuation separator/notice and absent-story fallback. These are render facts, with the original XML left unchanged. The existing synthetic-marker/12pt tests require explicit source-authority reconciliation, never deletion solely to make the new case pass.

Fit prediction and committed note splitting must share fragment accounting. Preserve the move-whole policy for a finite fresh note that fits a fresh effective page. When the note itself exceeds that capacity, admit its reference with a nonempty first fragment on the current page when space exists. Preserve carried/fresh ordering, multiple references, separator/continuation-notice cost and no-progress termination. Do not claim notes before deciding where the reference is placed. Use selected effective furniture geometry without moving unrelated furniture.

Local discriminators cover short whole-note movement, long note after prior body content and on an empty page, multiple/carried notes, authored and absent marker, styled indent/tab marker, separator-only/text/absent stories, continuation and source-preserving warm/cold placement. Batch in existing test binaries, retaining current-source RED and repaired GREEN. Current-source native PDF/CLI mismatch is already an actual RED observation with exact artifacts, not an inferred predecessor execution. Reconcile source-produced note fixtures explicitly and retain body-reference numbering, revision projection and source placement tests.

Final acceptance uses the immutable original and measures reference ownership, all note text/order and selected first/last anchors at the unchanged0.5pt and exact-page thresholds. Word note ranges are49/51/30lines across pages1/2/3, NOTE_000 visible atx72,y125.6088 and NOTE_129 atx72,y701.6088 on page3. Those positions follow the authored12pt separator,6pt existing offset and12pt note lines, not a new hardcoded native split constant. Actual Word and LO split differences remain separately recorded and require the existing cited/ asserted divergence process. Update only the existing HLD08 impact file, including source-event authority and actual inherited-policy correction. Final installed/native/hosted/full/risk gates remain pending.

Independent contract review and source-authoring compatibility inspection are required before this remediation starts. No implementation is authorized by this unreviewed reconciliation paragraph alone.

Native authoring compatibility refinement: existing Document.add_footnote currently writes only note text and relies on the synthetic renderer marker. Preserve its existing displayed-label behavior by making that existing constructor publish the same qualified note-body marker and reference-style run already used by create_footnote_at. This changes newly generated source intentionally, without rewriting imported marker-absent source. Existing create_footnote_at/create_endnote_at automatic marks and custom-mark literal text remain authoritative, with exactly one visible mark. There is no new native add_endnote operation. Include generated-source/save/reopen and installed Python add_footnote coverage in this amendment. The affected existing facade source/test file joins the implementation work list, with HLD08 remaining the only HLD impact.

## Final native furniture-frame eligibility reconciliation

Independent designer and source reviewer identify producer/consumer eligibility as the shared fault before another D2 repair. The existing strict line-gap transport and bounded stable-state solver are not disproved by these observations. Retain their mathematics and actual table/drawing controls. Correct which obstacles reach each consumer instead of changing line metrics or disabling all wrapping.

Actual native single-axis controls preserve every package payload except the one declared frame attribute. Four inputs pass official Word-root XSD admission and Word16.114 export, with unchanged source hashes. Horizontal text alone preserves the omitted result. Vertical text alone selects following header-flow origin. Vertical margin equals omission. Vertical page uses the page origin. Three separate footer controls, also schema-admitted with actual successful exports, show omission equals margin at visibley41.6088 while explicit text uses following footer flow atvisibley767.6088. In both stories, ordinary flow does not move around the frame. Header explicit-text y100 is disjoint from the body line and leaves BODYx72. The same frame aty140 crosses that line and moves BODYx216, while header FLOW staysx72. Receipts frame-anchor-axis-native-v2.json, frame-anchor-axis-measured-v2.json, native-layout-axis-controls-v1.json and native-layout-axis-measured-v1.json preserve native commands, hashes and coordinates.

For ordinary positioned header/footer paragraph frames in this compatibility repair, retain authored anchor presence through resolution. Resolve omitted horizontal anchor on the text/column basis. In this qualified furniture context, resolve omitted vertical anchor on the written margin basis observed in Word16.114, distinct from explicit text, margin and page. Do not change general body-frame, drop-cap or imported-model default semantics. This observed compatibility rule conflicts with the broader text default in MS-OE3762.1.48, so HLD08 must state the exact bounded native context and documentary difference instead of silently claiming that note predicts these coordinates. Preserve raw/model attributes unchanged.

Keep producing class and consumer eligibility private in the existing furniture/paginator carriers. Same-story ordinary header/footer flow ignores paragraph-frame obstacles, matching the actual native overlaps. Tables and drawings retain their own existing same-story wrap and table-only overlap-settlement policies. The body receives the selected prepared header/footer frame rectangles with their actual wrap/distances on the written page basis. Apply the existing body wrapping policy to actual intersections, including the strict touch/disjoint boundary. Do not move floating paint into flow clearance. Use one physical prepared frame rectangle for paint and body obstacles so footer convergence cannot mask displaced frames. Preserve selected first/even/default/blank ownership, stable prepared-result reuse, bounded/cycle refusal and source-qualified cache behavior.

Local discriminators must inspect header/footer FLOW, frame paint and BODY together, including omitted/text/margin/page, crossing/touch/disjoint, multi-paragraph equal-frame groups, notBeside chains and selected variants. Keep strict D2 transitive/growth/highlight/link/border/changebar tests on genuinely eligible table/drawing obstacles, rather than deleting reservation coverage. Actual original native BODY-only successes are partial evidence and do not excuse the masked112pt footer frame error. Source edits require current-source RED/GREEN and a full independent microscope. Final Word/pinned-LO/font/geometry/raster/full/risk debt and unchanged thresholds remain mandatory. No new public type, method, source file, module, crate or feature flag is authorized.

Independent contract review remains required before source implementation. Native rotation reconciliation is still pending its additional axis observations and is not authorized for a guessed correction.

The note fragment authority must define fresh-page capacity after the selected separator and applicable continuation-notice costs. It returns ordered line ranges for carried notes before fresh notes, preserving an explicit oversized-single-line no-progress policy. Prediction and commit consume that same plan, rather than clamping an aggregate height independently. A present empty authored separator remains distinct from an absent record. Qualified foreign-namespace marker lookalikes remain opaque. Source inspection confirms create_footnote_at and create_endnote_at already author raw styled markers, add_footnote does not, and custom-mark creation uses literal text. Imported older marker-absent packages stay absent under the source-driven rule, even if originally created by an older API. Document this source-compatibility boundary explicitly.

Root approves these two bounded remediation contracts after actual independent review on2026-10-29. The reviewer read the final sections including constructor compatibility and fragment accounting, found no concrete contract blockers and expressly excluded rotation. The designer independently inspected actual native authoring constructors and source-event preservation. Root read both reports and the isolated-axis native observations. This supersedes the earlier implementation hold for notes and ordinary furniture frames only. Rotation still requires its separate fact-based contract. Every source repair requires independent microscope and affected local/final evidence before completion.
