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

Final sprint: native Word with actual version, input/output hashes, production steps and positions, LibreOffice 26.2.5.2 and Poppler 26.01.0, deterministic library PDF/PNG/text, installed 3.9/3.12 and strict typing/stubs, full risk union and 49 hashes/7 pixels. Predeclared native geometry tolerance is 0.5 pt per measured text/furniture edge with identical qualified font files, and exact page counts and entry sets. Use extracted PDF coordinates where available. At 144 DPI this allows one raster pixel. Record font hashes and native layout inputs. A mismatch beyond tolerance remains a failed comparison, without post-measurement tolerance changes. It blocks acceptance until repaired or given a source-supported intentional oracle disposition under the canonical differential-testing skill. A disposition must preserve the failed measurement, cite the already approved target behavior and primary/native authority, and assert our side. It cannot waive an explicitly promised matching consumer, an unexplained mismatch or missing evidence. Unsupported primary/native cases remain pending rather than silently dropping them.

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

## Final native rotated-body contract reconciliation

Root selects a bounded compatibility repair for the observed ordinary main-body tbRl context, with exact paragraph line spacing. This closes the remaining A4 rotated-body observation without changing the0.5pt threshold or inventing a constant3.49pt translation. The original body-band-tbRl-center input and all native outputs remain immutable. Independent source investigation and actual native axis/metric controls provide the facts below. Other vertical orientations, general table-cell rotation, upright related stories and unqualified mixed-font behavior retain their existing boundaries, without a new authoring surface.

Actual Word16.114 top/center/bottom main-body controls, with five-line furniture, no selected furniture and asymmetric header-only selection, retain the written inline originy36. Their cross-axis body placement remains independent of that ordinary furniture selection. The current library instead substitutes clearedy78 for this rotated body's written basis. Keep the ordinary rotated body on its written physical band before transposition and section vertical alignment. Preserve upright furniture paint, physical notes and selected positioned-frame eligibility as distinct owners. Do not reuse the BiDi TextDirection flag for section rotation. Do not apply this rule to every vertical or cell context without qualification. Horizontal selected-flow clearance remains unchanged.

ECMA376Part1 clause17.3.1.39, printed255to256, delegates an omitted textAlignment to consumer determination. Clause17.18.91, printed1503to1504, distinguishes automatic, character-center and baseline alignment. Actual isolated controls establish that absent/auto and explicit center match in this tbRl source, while explicit baseline differs. Removing docGrid, line numbering or both, and specifying all four paragraph-end font slots explicitly, leaves the native baseline unchanged. These controls exclude those candidate causes for this case. Preserve authored alignment through effective style/direct resolution and private layout metadata, rather than treating automatic as explicit baseline.

The two final paired native controls isolate the numeric owners. For exact line12/font11, line24/font11 and line12/font16, explicit-baseline offsets from the logical right edge are9.6,19.2,9.6pt, or0.8times the exact LINE height. The automatic/center offsets add3.84,3.84,5.76pt respectively. Qualified actual Liberation Sans ascent1854and signed descent-434atUPEM2048 give the central offset as half the ascent plus signed descent, scaled by the actual run size. The resulting candidate baselines fit all six observations within0.214pt. Retain those actual residuals. They are not proof of a specific native rounding rule. The existing multilingual0.8times FONT size helper has a different authority and remains unchanged.

Resolve the exact-line baseline and central-font term from their separate authorities. Use one private policy and existing raw face metrics, not an already adjusted line ascent or guessed font name. Preserve original font IDs, requested sizes, authored baseline offsets, raw line advance/metrics and source XML. Normal paint, decorated/rich/Group/Marked text, hit boxes, occupied bands, field substitution and fragments must consume the same policy. Apply it once per fresh line break, including strict wrapping/reflow after raw line-height restoration and before reservation agreement. Repeated preparation must not accumulate the offset. Keep unrelated multilingual shaping and exact-line policy intact.

Use existing private ParagraphSemantics and ParagraphView or equivalent existing private carriers. No field is added to public ParagraphBlock, ParagraphReflow or neutral LayoutLine solely to carry this policy. If cached blocks depend on orientation or resolved baseline policy, add that private fact to the existing cache identity, including equal-width orientation changes. Otherwise preserve raw cached blocks and project the policy consistently at every consumer. Warm/cold source and actual section ownership must agree. Upright notes/header/footer paragraphs must not inherit main-body rotation merely from using section margins.

Local discriminators retain the original42pt origin RED, original3.491pt baseline RED, explicit default/center/baseline, font11/16 and exact line12/24, top/center/bottom, symmetric/asymmetric furniture, inherited/direct alignment, equal-width orientation cache changes, repeated reflow and rich/decorated/field occupied bands. Test both physical geometry and the resulting baseline with actual metric-derived expectations. Final acceptance reuses the immutable native outputs and fixed font inputs, requires original and paired controls within0.5pt with exact page counts, and retains raster/field/source/full-risk debt. Unsupported combinations remain explicitly unqualified rather than gaining a parity claim.

Receipts are rotation-axis-measure-v2.json, rotation-baseline-native-measure-v1.json, native-rotation-metric-controls-v1.json, native-rotation-baseline-pairs-v1.json and rotation-baseline-numeric-diagnostic-v1.json. They distinguish actual Word outputs from frozen pre-remediation CLI observations and localized raw-operator diagnostics. Original PDF bytes and parser limits remain unchanged. HLD08 remains the sole HLD impact. This contract awaits independent review before source implementation.

## Repeated continuation-observer investigation

The lower88 successor retains86passes and two failures in separator-based page inventory. Root and the independent reviewer investigated before another repair. The selected authored continuationSeparator is emitted as a translated Group containing a0.5pt Line. The old compatibility helper unwraps only MarkedContent, so its shallow separator detector sees the first page's fallback Line and misses the actual authored continuation group. These logged failures do not establish dropped note content. Correct the note-specific test observer to traverse actual paint order and accumulated transforms, preserving short/full-width and marker assertions. Independently inventory all120note paragraph texts across physical pages and prove at least two pages contain note text, with one first-fragment marker and no continuation repeat. Preserve authored-empty versus absent separator coverage. Do not change product separator output to satisfy a shallow test observer. No general compatibility-helper rewrite or unrelated test weakening is authorized.

Independent review closes the remaining provenance detail: eligibility requires the exact source section text_direction equal to tbRl, carried through an existing private section carrier such as SharedSection and the private paragraph policy. PageGeometry.body_rotation equal to90is insufficient because tbRlV maps to the same angle. Do not add eligibility fields to public PageGeometry or Section. Include a same-angle tbRlV negative control and upright related-story controls. Root accepts this clarification and approves the rotated-body and continuation-observer contracts after actual independent review on2026-10-29. The numerical route remains the bounded separately evidenced LINE-height/face-metric policy above, with every final threshold and unqualified boundary intact. Implementation and full independent microscope are still required.

Implementation closure preserves raw cached and transient metrics. Baseline evaluation is a pure function of the current freshly broken line, qualified actual text-face metrics and private source policy. Painter and line-number consumers use it without storing an accumulated adjustment. Inspected highlight/link/fragment and wrapping equations price the unchanged container band, so they do not replace raw metrics with a text offset. A graphics-only line or failed face qualification retains its previous baseline. Actual graphics-only RED showed an8.4pt unintended displacement before this qualifier, with its original source/binary retained. This enforces the existing bounded text contract rather than expanding it to unqualified graphics. Native residuals and final thresholds remain unchanged.


## Source-supported pinned LibreOffice disposition

The collection at source 1431cb40 contains 98 CLI output pairs. All 42 Word BODY_MARK edge and exact-page checks pass at the original 0.5 pt threshold. Pinned LibreOffice passes 31 of 42 checks. All 11 LibreOffice mismatches remain recorded numeric failures. The disposition below retains existing Word-target policy under the required oracle triage. It changes no output, tolerance or supported input. Root inspected the exact ECMA clauses and independent per-case map. The next actual microscope must review this clarification. Furniture, font, raster, operation and composition acceptance remains pending.

- Selected missing, blank and populated first/even stories use their qualified source owner and section start parity. ECMA 17.10.1, 17.10.5 and 17.10.6 prohibit borrowing another type and specify start-based parity. Existing issue424_selected_blank_stories_and_zero_top_use_only_selected_flow asserts missing-first and missing-even/start2 BODY y0 with no default header. Explicit empty and populated cases distinguish their own flow height. LibreOffice differs on variants-first-missing, variants-even-missing, variants-even-blank and variants-even-populated. Its import or selection cause is unproven.
- Floating tables stay outside ordinary furniture flow, with whole-table overlap settlement under ECMA 17.4.56 and 17.4.57. The table-floating local control already asserts the written 36 pt body band. The paired-table matrix retains its original unformatted FLOW_CONTROL and compares body placement against that same source without tables, alongside overlap and never controls. Do not transplant the exact-spacing native control's numeric band into this different source. Word agrees for table-floating, tables-never and tables-overlap. LibreOffice differs. BODY alone cannot qualify actual table settlement or paint, which remain separate criteria.
- Ordinary furniture frames with omitted vertical anchor retain the approved native Word margin-context behavior. Generic MS-OE376 text-default prose does not predict this context. Retain the independent header/footer axis controls and existing exact frame, FLOW and BODY assertions. The frame-default-anchors-omitted mismatch does not establish its LibreOffice cause.
- Main-story Center and Bottom use the selected effective band under ECMA 17.6.23 and signed-clearance intent. Existing Center coverage and the source-equivalent Bottom assertion retain band 78..714 and a 12 pt bottom line at y702, distinct from the native text bbox 701.6088. LibreOffice positions both controls near the effective top, for an unproven internal reason. Exact-source tbRl retains its approved written band, private baseline and complete alignment matrix. The body-band-tbRl-center mismatch remains explicit, without extending vertical-text support.

Raw positions, artifact/font/input hashes and clauses are in release-prep/f273-lo-dispositions-03/per-case-dispositions.json and final-f273-body-anchor-comparisons-03.json. Word BODY acceptance covers only the measured text anchor and exact physical pages. Complete font identity and other native criteria remain due. The two additional local assertions lock already correct native-observed behavior. They are positive coverage, with no invented old-source RED or product repair.


## K1, effective Word run kerning for required same-font edges

The final same-font investigation identifies50required text right-edge failures whose
retained library widths equal pinned HarfBuzz14.4default kerning and whose kern=0widths
are within0.041783pt of Word. Preserve these original measurements and failures. This
amendment intentionally corrects Word run kerning for the existing required cases. It
does not relax0.5pt tolerance, drop labels, substitute fonts or claim that a feature
counterfactual is repaired-product evidence. Existing table/endnote/frame, source,
raster and final acceptance obligations remain unchanged.

Reuse existing CT_RPr.kern and sz through docDefaults, table, paragraph/character style
chains and direct run/result formatting. No new public authoring field or raw kern
bridge is added. Under ECMA376Part1,17.3.2.19, absence inherits and never-applied means
no kerning. A present admitted nonnegative half-point threshold enables kerning iff it
is less than or equal to the effective nominal sz half-point value. Use the existing
nominal11pt fallback when sz is absent. Evaluate before automatic superscript/subscript
or caps display-size scaling. Zero is an explicit enabling threshold. Preserve raw,
foreign and malformed source according to existing parser/serializer rules, without
new source normalization or a broader parser repair.

Keep Word policy private in existing Word run/style/annotation carriers. Shared generic
FontManager APIs and PowerPoint callers retain their current default shaping, caps,
language, fallback, script, bidi and metric behavior. Do not change an ambient font
manager default or encode policy in synthetic FontIds. Preserve font data, face indices,
variation coordinates, font IDs and existing trace/embedding semantics. Thread only
an explicit optional kern override through shared private shapers, composed with the
existing caps features. Cache identity distinguishes generic default, explicit off and
explicit on for equal font/text/size requests, with existing collision/poison/size
bounds retained.

The following six additive doc-hidden neutral methods, the one neutral origin
function approved in the stored-field annotation section below, and the seventh
context-preserving method approved in the mixed legacy producer closure are the
sole additional public API. These seven methods are the public-method exception
to the earlier no-new-public-method restriction for this amendment. Their
signatures use only existing public types. No new public type, field, source file,
module, crate, trait, generic parameter, feature flag or authoring operation is added.
Existing constructors and public signatures retain their behavior.

On FontManager:

    #[doc(hidden)]
    pub fn shape_text_with_kerning(
        &self, font_id: FontId, text: &str, size_pt: f64, enabled: bool,
    ) -> Result<ShapedText>;

    #[doc(hidden)]
    pub fn shape_paragraph_with_kerning(
        &mut self,
        segments: Vec<(TextSegment, Option<String>)>,
        base_direction: TextDirection,
        no_wrap: bool,
        enabled: &[bool],
        legacy_shaping: bool,
    ) -> Result<Vec<MultilingualTextSegment>>;

On MultilingualTextSegment:

    #[doc(hidden)]
    pub fn with_base_and_advances(
        self, base: TextSegment, x_advances: Vec<f64>,
    ) -> Result<Self>;

    #[doc(hidden)]
    pub fn with_hyphenation_language(self, language: Option<String>) -> Self;

    #[doc(hidden)]
    pub fn uses_legacy_shaping(&self) -> bool;

    #[doc(hidden)]
    pub fn kerning_override(&self) -> Option<bool>;

The paragraph request vector is input-occurrence-aligned, including equal text/face/size
runs. A count mismatch refuses before shaping. legacy_shaping preserves the existing
plain Word segment-property, line-break and paint conventions instead of silently
turning the kerning fix into a language, bidi, offset or output-variant change. Existing
true multilingual and PowerPoint caps behavior remains intact. Context-preserving base
replacement retains all private shaping/break metadata and requires compatible logical
text, font identity, size, direction, glyph identity and cluster ownership. Its new
advance vectors must have the actual glyph count and finite values, and its retained
position/cluster vectors must remain valid. Invalid requests use the existing neutral
LayoutError/Result path and publish no replacement. Existing public constructors keep
their current validation and defaults. Private span metadata retains
Word kerning, legacy mode and existing automatic-hyphenation language through splitting,
wrapping and repeated reflow. Do not convert policy-bearing Word hyphenation back into
an unqualified public TextSegment. Atomic markers, symbols, note marks, annotation/ruby
text and plain/rich field fallbacks use their own source owner or explicit generated
Word fallback, never the surrounding body's policy. Structural Target/SequenceContext
markers retain their current synthetic glyph/text/zero-advance representation while
acquiring explicit Word (false,true) context at the known producer. They are not shaped
as ordinary visible text or classified later from content.

Keep all current public TextSegment, ParagraphBlock, ParagraphReflow, PageFrame,
GlyphRun and MultilingualGlyphRun models unchanged. Word spacing, bidi reconstruction,
legacy paint, decoration/link bounds, character spacing, hanging spaces, right/centre/
decimal tabs and automatic hyphenation must consume the same per-occurrence shaping
context. In particular, never infer uniform character spacing by subtracting a default
kerned width from an explicitly unkerned segment. Preserve exact-line, rotation, OMML,
VML and other unrelated policies. Word annotation kerning is applied at its existing
source shaping step before construction of an opaque Group.

Use contract A for late replacement: every positioned Text or MultilingualText leaf
with field_kind OR field_source has exactly one private companion entry keyed by its
exact page-local positioned-element tree path. Include empty placeholders and hidden
structural markers. A Word-owned member stores Some((kern,legacy)). An identified
actual neutral producer stores explicit None. A missing entry is an internal layout
failure, not a default-shaping fallback. SourceSpan/FieldSource presence, text, font,
size, coordinates, addresses and provenance mode never determine shaping-policy
ownership or lookup. Non-field ruby/EastAsian/OMML/chart Groups are outside membership,
not neutral-labelled members. Unexpected opaque field-bearing subtrees without an
identified producer/companion fail closed.

Produce the companion during the same checked paint operation that emits the actual
field-bearing glyph leaf, reading its private source context before it is erased by
public positioned output. Private PaintedElements operations move element/policy
subtrees transactionally through append, split, Group/Marked wrapping, section/cell
rotation, clipping, logical ordering/permutations, header/footer behind/front splices,
notes/special separators/endnote appends and engine artifact wrapping. Coordinate-only
edits leave paths unchanged. Check unique keys, exact field-leaf cardinality, target
node kinds and page/map vector cardinality. Existing public neutral wrappers keep their
signatures. Internal Word entry points use the checked producer and propagate failures.

Late PAGE/NUMPAGES/SECTION/SECTIONPAGES/PAGEREF/SEQ and STYLEREF replacements look up the
current exact traversal path. When one original leaf expands into a Group, build a
replacement companion for each new child with that original occurrence's context. If
a replacement actually introduces separately resolved source fragments, record each
fragment's own context at its producer. Recursive Group/Marked rewrites carry detached
local subtree maps. Stable removal rebases all following sibling paths. STYLEREF sees
the map after earlier substitutions/removals. Do not change field syntax, source/cache
publication, fragment ownership, or F272's field_source_run_fragments_at API.

Retain raw and substituted companions with their respective private page/cache artifacts,
including restarted prefixes/tails and endnote same-page reuse. Include exact private
policy identity in field-substitution/restart cache equality. Charge every entry, path
and private metadata allocation under existing cache ceilings. A missing/inconsistent or
over-budget required map refuses the checked layout/cache admission. It is never
truncated or silently defaulted. Keep warm/cold and provenance-on/off results equal.

Required discriminators:

- Inherited/absent/direct thresholds with inclusive sz ties and superscript nominal-size control.
- Identical font/text/size cache alternation between generic/default, Word off and Word on.
- Mixed-policy equal-run occurrences, plain/multilingual fallback and existing PPT caps.
- Forced wrap, hyphenation and repeated float reflow.
- Decimal/right/centre tabs, spacing, decorations and source ranges.
- Nested Group/Marked paths with equal-field fragments and provenance off.
- Structural empty markers, logical permutations and header/footer layer splices.
- Rotations/clips and note/special/endnote placement.
- Leaf-to-subtree expansion followed by structural removal and STYLEREF.
- Raw/final cache prefix/tail reuse after same-owner kern-only edits.

Assert actual glyph advances, positioned edges and policy ownership, including missing-policy failure,
not only initial ShapedText widths. All existing generic/PPT controls remain unchanged.

HLD08 remains the sole HLD impact. Any work-list expansion is limited to the existing
font.rs/line.rs and Word engine/paginator/notes/block/convert/style-resolver files needed
for this closure, with tests in existing entrypoints/modules. F272 returned its writer lease after the occurrence-preservation checkpoint. Root
controls the canonical writer lease and will reassign implementation explicitly. Root approves this bounded technical amendment on 2026-10-29 after reading the
independent policy, 94-site companion and clarified contract-A design investigations.
The exact draft and independent verdict remain retained with SHA-256 identities
77aa7f7bc78d5e98c974b9897e72d1ab21598f2cb2afd71417fcbfb4c72efd22 and
83ebaf3c7f9540c581321ab54bb9bc1aee82d9d21f0003fe322fe0985bd7d5c4. The executing owner must read and acknowledge
the current full plan content identity before implementation and hold the sole writer
lease. This approval adds no unlisted public methods or public models. After reviewed
source implementation, microscope and scoped evidence, rebuild
a freshly authenticated final product artifact and rerun every original same-font edge
case using retained native PDFs and fixed font profiles. The integrated full/native/
raster/risk and source/field/furniture union must pass before final acceptance. Earlier
1431/test-only6506 and feature-counterfactual evidence remain historical, not qualification
of the future repaired artifact.


K1 closes the already required same-font edge consumers in A1/A3/A4, the same-owner
cache contract in A7 and the compare/rebuild/update/render union in A8. It is not a
replacement for any original issue reproduction. HLD08 must describe the effective
Word threshold, unchanged neutral/PPT default, private legacy/rich carrier, field-only
positioned companion, exact cache identity and checked lifetime/accounting before
verification. Existing source inheritance is unchanged, so the HLD impact remains
exactly the existing rendering-spec file.

This is a separately labelled intentional Word shaping bug fix in the hash harness.
Expect rendered PDF/PNG differences in the seven Word samples where absent or inherited
thresholds change advances or wrapping. XML parts and neutral/PowerPoint default output
remain unchanged unless another separately reviewed story already declares a delta.
The historical1431 nineteen-delta account is retained as historical evidence. Inspect
every actual49-hash/7-pixel successor delta and obtain a concrete reviewed per-case
account before any baseline update. No current-output baseline is approved here.

Table logical ownership, floating anchors, separator metrics, page-border inclusion,
cell-relative drawings, front drawing order and the unresolved framed-image axis remain
separate native remediation work. This K1 approval does not approve guessed precedence,
separator constants, frame translations, tolerance changes or omitted source cases.
At this K1 approval Word was closed and native controls were paused. The user later
confirmed restored documents and authorized guarded tests. Word is running, but the
read-only custody helper needs requalification before any new export. Existing raw
failures remain failed and additional native observations are still pending.

## R4, exact per-page front drawing composition with K1 ownership

The retained native short noncell drawing control has matching rectangles but a
paint-order failure. behindDoc=false must paint above intersecting document text,
including text after its owner. Preserve actual native/operator/mask failures and
fixed thresholds. Root's technical correction retains current geometry, wrapping,
source identity, relative encounter order, selected furniture and public models.
It adds no relativeHeight sorting or new note/nested-anchor support.

Use the approved private K1 PaintedElements transaction for behind, regular and front
subtrees. Place anchors at their current cursor/column point, but route false-behindDoc
paint to the exact owning physical page front destination. Propagate explicit layers
through body, split paragraphs, flowed/floating/repeated/nested/rotated/clipped cells
and furniture tables. Keep change bars separate from rotated-cell anchor furniture.
Apply existing transforms/clips to each owned slice with the same equations. Preserve
one anchor occurrence on the first split fragment and one physical repetition per page.
Do not drain on a column-only finish_page transition.

Use this private existing-file representation, integrated with K1's field-only map:

    WordParagraphPaint { anchors: Vec<WordAnchorPaint>, rows: Vec<WordRowPaint> }
    WordAnchorPaint::Image
    WordAnchorPaint::Group { policies: PaintPolicies }
    WordAnchorPaint::Shape { paragraphs: Vec<WordParagraphPaint> }
    WordRowPaint {
        line: usize, item: usize, behind_end: usize, flow_end: usize,
        children_len: usize, policies: PaintPolicies
    }

Anchor contexts align exactly with the same producer's anchored vector, and shape
paragraph contexts with its text vector. Row keys identify only exact prepainted row
Group occurrences created at the known row adapter. Empty vectors positively mean no
such content for that owner. They do not request a search elsewhere. Validate counts,
variants, unique line/item keys, monotone child boundaries, actual Group kind and exact
field-leaf map cardinality. Never infer context from text, media, coordinates, source
provenance presence or pointer identity.

The existing public row Group remains sole owner of all glyph/image bytes. Its direct
children are concatenated behind, regular, front in the same checked construction that
records the interval endpoints. Retain its current identity/opacity1/no-effects outer
wrapper contract. At placement, detach actual slices and maps, wrapping each with the
same original placement/clip/transform. Refuse a nontrivial outer opacity/effect group
rather than silently splitting its isolation. Do not paint both the compatibility
Group and the extracted layers. This closes already painted note/textbox table rows
without enabling currently unpainted ordinary note/nested paragraph anchors.

ParagraphSemantics gains paint: Arc<WordParagraphPaint> and uses Clone rather than Copy.
SharedLayoutBlock::Owned gains semantics: Option<ParagraphSemantics>, required Some iff
its actual owned block is a paragraph. Its paragraph view exposes that exact value.
CellBlockSemantics already carries ParagraphSemantics. HeaderFooterBlockSemantics and
NoteRenderParagraph carry the same owned Arc. ParagraphCacheEntry retains it only when
its existing admission permits. Arc has actual cached and repeated-placement consumers
now. No global map, public block field or source-address service is introduced.

Existing private paragraph core returns (ParagraphBlock, WordParagraphPaint). Directed
and in-table variants return (ParagraphBlock, TextDirection, WordParagraphPaint).
collect_anchored_drawings returns (Vec<AnchoredDrawing>, Vec<WordAnchorPaint>).
story_table_rows returns Vec<(ParagraphBlock, WordParagraphPaint)>. Existing public
compatibility entry points keep signatures and behavior. Internal checked Word calls
must consume the context-returning route, including shape children and special notes.
No forwarding-only wrapper is introduced. Known neutral producers are explicitly
admitted at their existing public boundaries and do not impersonate Word ownership.

Private anchored_elements/render_shape_text receive exact finite Word or Neutral owner
context and return checked PaintedElements. Shape child layers compose inside the
outer anchor before its transform/Marked wrapper. A child never escapes the outer
anchor's layer, clip or coordinate owner. Split/slice/reflow/rebinding operations
transport contexts with their owning vectors. If a generated row is ever structurally
reflowed, its exact keys move transactionally or the unsupported operation refuses.

Final order preserves current endpoint priorities: back page border, watermark, body
back, furniture back, regular header/body/note/footer text, body/table/note front,
existing furniture front, front page border. Preserve current order within every layer.
This is a bounded composition decision and no claim of newly measured native front-front
priority between arbitrary stories. Existing geometry and original source controls stay
unchanged. Native raster/operator qualification on the repaired artifact remains due.

Same-page late endnotes must insert before front layers. RecordedPagination,
SharedPagination and private PassResult carry page-aligned checked
PagePaintBoundary { behind_end, regular_end, body_front_end, furniture_front_end,
elements_len }. Validate monotone endpoints and actual top-level length. Late note
behind/regular/front slices insert at their exact respective boundary, updating every
later endpoint and K1 path transactionally. No output scan discovers a boundary.
Retire descriptors after actual late append/checked flattening and before artifact
wrapping/raw cache capture. Finalized prefix/tail pages retain final K1 maps and do not
receive a duplicate endnote append. Public neutral append_endnote_pages is unchanged.
Any altered descriptor lifetime requires exact cache equality and charged storage.

Preserve existing rich/anchor/Group cache exclusions. Charge all newly retained Arc
allocations/control blocks, vector capacities, recursive shape projections and field
entry/path capacities under existing paragraph/table/furniture/restart ceilings. Use
exact value equality, never Arc addresses or fingerprints as authority. Overflow or a
missing required map refuses admission/publication, never truncates or defaults context.
Existing public signatures remain unchanged. Apart from the seven approved methods and
the single neutral origin function below, no additional public API, public type/field,
source file/module/crate, trait/generic or feature flag is added.
Existing engine, paginator, block, notes, table/convert/style carriers required by K1
are the work list. HLD08 remains the sole HLD impact.

Keep original short noncell native input and capture actual current-source local REDs
before implementation. Discriminate later intersecting body text, two front owners,
inline images, column-only transitions, real page changes, first-fragment single paint,
clipped/rotated/nested/repeated cells, furniture-table bypass, prepainted note/shape rows,
same-page/overflow endnotes and warm/cold prefix/tail reuse. Assert exact unchanged
rectangles, source/media/structure ownership, no drop/duplication, marker/page partition
and field-only policy paths/cardinality with equal text and provenance off. Preserve
ordinary/frame furniture controls and current unsupported row-splitting boundaries.
Ordinary unpainted note/nested anchors remain evidence-backed exclusions, with no new
placement or parity claim. R1/R2/R3/R5 geometry and T1/T2/T3 remain separate work.

Independent design sources: f273-r4-layer-design-v1/diagnosis.md SHA256
1ce9ebcd116e715394d5ed99d6232845c2a37278ade50dab91ae6f6a96642ceb,
and concrete-carrier-addendum.md SHA256
7a0c21ccebe6c51c6efd9cadad32cc4349c2ec24f48b3874044fd46d0891ad4c.
Root has read both fully and the exact amendment received an independent zero-omission,
zero-contradiction design review, retained in amendment-review-v1.txt. After F272 explicitly
returned its sole writer lease, root approves this bounded technical amendment on
2026-10-29. It closes A3 painted ordering, A4 late-note composition and A7/A8 cached
field/source lifetime obligations, while retaining all pending native qualification. Implementation, scoped proof, independent microscope,
fresh native/raster artifact and final full/risk/hash union remain required. No baseline
update or acceptance waiver is authorized.

The assigned implementing owner must read and acknowledge this full canonical plan
content identity before affected work. No implementation lease follows from approval
alone. Root assigns each implementation separately and preserves one writer/build owner.


## Current F272 ordered-owner preparation for K1 and R4

K1/R4 consume F-X272's approved current ordered COW result owner through its two specified native Field bridges. Layout/evaluation/source operations prepare each required current result once, select the existing RevisionView, establish complete injective association to actual story/paragraph/field/result paths and content slots, then freeze operation-local storage before borrowing spans. A missing contiguous cached_result_runs slice does not imply plain or empty content. Prepared clone addresses may index existing local maps only after this association. The read projection bridge does not normalize the live receiver. Checked mutation/publication already supplies normalized current owners.

Resolve shaping policy from each actual projected run/result owner and capture the field-only positioned companion at the existing checked Word emitter. Source occurrence identity and page-local positioned tree paths remain distinct. Qualified inline-control virtual fields retain F272's physical-slot/source-writer authority and atomic outer boundary. This clarification changes no approved K1/R4 signature, membership rule, layer policy, cache eligibility, line-number font policy or unsupported anchor boundary. Source implementation starts only under root's explicit sole-writer assignment after provider handback.

Root approved this contract clarification on 2026-10-29 after authenticating the seven sealed reconciliation members and reading the complete report. The reconciliation reports zero required redesign or signature changes. It does not qualify the still-changing provider implementation. Sole-writer ownership and the existing dependency readiness rules remain binding.


## Generated Word line numbers, source style and counter correction

Root technical amendment to F-X273 A4/A7/A8. Independent contract closure and sole-writer assignment precede source implementation. The old21 and five additional native controls are bounded source/used-visible-sample observations. They establish neither universal Word font identity nor repaired-product parity. Source nominal11pt/16pt remains authoritative. Exported effective11.04/16.08pt does not become a layout constant. Native0.5pt and exact-page thresholds remain unchanged.

Resolve generated line-number style once per Word layout from the normative styles part. Recognize the unique correctly typed primary name `Normal` for the paragraph contribution and `line number` for the character contribution, after trimming leading/trailing whitespace. Do not use serialized style ID or default flag as a substitute for those names. The observed ID/name controls and MS-OI29500 sections2.1.237(c,d) support that bounded recognition. Exact case matching, no localized/alias expansion, and omission of an ambiguous duplicated role are conservative technical choices for otherwise unqualified names. Preserve all imported name/type/default/ID/XML bytes and emit a bounded diagnostic for ambiguous/wrong-type role declarations rather than guessing, renaming source or refusing the whole layout. An absent role contributes no properties. The default paragraph selected for ordinary BODY content is not an implicit replacement for absent Normal.

Use document-default run properties, the named Normal paragraph style and its paragraph basedOn ancestors, then the named Line Number character style and its character basedOn ancestors. Resolve basedOn by actual ID. Each role follows only correctly typed ancestors. A missing ancestor ends that chain. A repeated ID ends the finite chain before applying it again, using the existing cycle prevention and a bounded diagnostic. Each property sees each owner once. Ignore child properties of the exact IDs DefaultParagraphFont, NoList and TableNormal for Word interpretation, as required by MS-OI29500 section2.1.242(a), while preserving source. Because basedOn is itself a child, stop that ignored-ID branch rather than following its ignored edge. A wrong-type edge ends the relevant chain rather than importing another style level. These malformed-chain choices preserve imported data and deterministic rendering without expanding public authoring APIs.

For non-toggle family/size/kern properties, the nearest explicit property within each chain wins, and a later applicable level overrides an earlier one. Use the existing Word ASCII slot/theme resolver and unchanged FontManager fallback/substitution for generated ASCII decimal labels. Missing family is passed as missing to that existing resolver, not replaced by literal serif. Use K1's existing nominal11pt fallback only when effective sz is absent. Preserve explicitly supplied zero/unsupported sizes in the imported model and XML. The private generated-Word producer requires a finite positive resolved nominal size before shaping. Otherwise it emits a bounded diagnostic and paints no number glyph, while retaining line tracking and counter progression. This is an explicit private policy, since the generic shaper does not reject zero sizes. Do not change that generic shaper or clamp source. Bold/italic follow ECMA376Part1 section17.7.3: nearest explicit value in each basedOn chain determines that level. If document-default value is true, the effective value is true. Otherwise combine paragraph and character level values by XOR, with missing values false. There is no table-style or direct-run contribution for these generated section labels. This is a bounded normative technical context decision, not new native proof of every combined toggle. Do not change the generic run resolver or shared/PPT default policy.

Carry one resolved face/nominal size/bold/italic/K1 kern decision through the existing private Word section, pass and Pager owners, keeping public exhaustive LineNumbering/PageGeometry/Section and public paginate signatures unchanged. The same values drive shaping, right-alignment width and GlyphRun publication. K1 supplies explicit Word kern on/off, while public neutral routes keep current default shaping and9pt compatibility. Keep the existing body baseline/track recording, transforms, artifact color/ownership and geometry. Number glyphs remain outside K1's field-only membership and move through R4 as ordinary known non-field artifact subtrees. At continuous sections move the companion exactly when existing page section ownership moves, including the empty-page member switch. Do not create a new per-line mixed-section font authority or change grouping/page breaks. Normal/LineNumber styles are document-wide facts, not section-specific overrides.

For Word section counters, absent lnNumType or omitted countBy produces no numbered-line records or glyphs. Raw XML omissions remain omitted. Omitted start is zero skipped numbers. Explicit start is the skipped count, so first logical value is start+1, including0->1,3->4 and32767->32768. Apply countBy selection to that displayed logical value, then preserve existing restart/page/section ownership. Use a checked/wide private counter so all existing parser-admitted u32 start values have a deterministic start+1 without authoring clamps or wraparound. This wide-counter behavior beyond Word's documented start0..32767 range is a compatibility technical decision, not Word parity. Retain existing nonzero countBy modulo behavior beyond Word's documented1..100 range and existing explicit countBy0 rendering compatibility as1, without rewriting source. Explicit0 handling is not inferred from the omitted-countBy observation. Keep public neutral pagination start semantics unchanged. A compact private Word counter policy may accompany the font carrier rather than reinterpreting public geometry globally.

Use only existing source/test carriers. No new public fields or methods, files/modules/crates, assets, traits/generics, feature flags or parser-domain restrictions. Private helpers in the existing carriers are permitted. HLD08 remains the sole HLD impact. Preserve existing style/theme/alias/font equality invalidation, raw prefix/tail/unchanged-body eligibility and capacity accounting. An entirely transient scalar companion adds no retained allocation. Actual retained glyph storage remains fully charged. If implementation retains an additional companion, charge it and compare exact values including finite size bits. A numbering-present Word companion must be explicit even when face resolution failed. It must not silently enter neutral9pt compatibility.

This is a separately labelled intentional Word line-number rendering/counter bug fix. Expected deltas are Word PDF/PNG samples containing enabled numbering whose source family/size/weight/start/countBy omission differs from the old serif/regular9pt/value-as-start behavior. Preserve XML, unnumbered Word output and neutral/PPT output. Inspect every actual49hash/7pixel successor delta before any baseline update. Current local/source RED, finite affected GREEN, independent microscope, repaired-artifact native comparison using existing outputs, warm/cold/cache and final integrated full/risk/installed riders remain mandatory. Existing raw native failures, diagnostic stops and successor measurements remain immutable.


Root approves this bounded technical amendment on 2026-10-29 after the exact corrected candidate received independent closure with zero remaining contradictions or required clarifications. Independent report line-number-contract-independent-closure-v1/REPORT.md has SHA256 25ed9266297d43869fe59445d12828f6f6b0dbb15e147dd6d80bbcf53d722fa4. Root read that complete report, the sealed proposal and the pinned toggle text. This approval supersedes earlier progress statements that no line-number font hierarchy, counter or transport correction is approved. Those statements describe the earlier evidence stage. No historical result is promoted to repaired-product acceptance. No user answer is invented: these are root technical decisions within the rendering repair. The earlier Q-S96-02 complex-field answer remains separate.

The actual new five observation union is line-number-five-qualified-observation-union-root-v1.json, SHA256 261331cb02ff031ee07d26227c88326497c551794f14d4df88ab3a3ce82f77c4. Root authenticated all four successor output maps and 935 unchanged guards each. The original normal-id complete observation is retained without remeasurement. Original four diagnostic stops remain immutable. The exact primary records and source joins are retained by the closure packet. Microsoft built-in name rules are https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/716e4d7c-3b23-4210-9617-2b8324486a3b and ignored-ID rules are https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/bb6afca9-88a7-4b34-911f-7110e8062bb1. Pinned ECMA17.7.3 supplies toggle semantics. The existing primary start/countBy record supplies counter semantics.

### Finite local and final evidence matrix

1. Reuse source constructors for old21 plus five new style graphs. Assert resolved source nominal family/size/b/i/kern and resulting actual GlyphRun face/size/flags/advances, not PDF11.04 constants. Retain direct BODY formatting and DefaultParagraphFont variations as negative authorities.
2. Test ID/name pairs, absent Normal with unrelated default paragraph, absent Line Number, duplicate trimmed name, wrong type, missing basedOn, cycle and ordinary parent. Assert deterministic bounded result and unchanged imported XML. No local alias test is described as new Word parity.
3. Toggle table: same-level child false overrides parent true, missing inherits, Normal true plus LineNumber true yields false when docDefaults false, and docDefaults true remains true. Cross bold/italic separately, include old Normal-only and LineNumber-only positives. Do not change generic resolve_run_properties.
4. Test ASCII direct/theme family slots, absent family/size, unavailable face substitution, source sz22/32, and inherited/zero/absent kern. Assert shape/publication use one tuple, width determines right alignment, baseline unchanged. Public neutral struct literals and outputs remain callable/unchanged.
5. Test lnNumType absent, countBy omitted/0/1/3/100/101, start omitted/0/1/3/32767/u32::MAX, and NewPage/NewSection/Continuous at actual existing ownership transitions. Source omission and explicit0 remain different. Assert start+1 before count modulo, no overflow/panic/clamp/source rewrite. Beyond-Word-range cases are deterministic compatibility tests.
6. Test body-unchanged style/theme/font/alias edits with warm versus cold results, existing raw prefix/tail and unchanged-body paths. Assert no stale9pt/serif glyphs and exact full glyph state. Transient carrier requires no invented allocation budget test, but any retained storage must be charged.
7. Test separate sections, columns, continuous occupied page versus empty-page ownership, late notes and center/rotated body baseline where they cross actual transport boundaries. Document-wide font does not invent section-specific source styling. Known numbers contribute zero field-policy map members and survive R4 composition once.

Reuse existing retained native outputs for final repaired-artifact comparisons at unchanged0.5pt/exact-page criteria. Local source/consumer agreement is not native proof. No additional style variant is made a mandatory new oracle obligation by this report. Every actual hash delta must have this labelled change's concrete account before baseline movement.


## Word-legacy allocation closure, technical clarification

On2026-10-30root approved exact-capacity normalization only at the explicit
Word-legacy factory and context-preserving replacements/setters carrying those
spans. Generic/PPT public constructors and default shaping keep their behavior.
Vec-to-boxed-slice-to-Vec and String-to-boxed-str-to-String provide exact reported
capacities, rather than assuming shrink_to_fit succeeds. Audit clone, split,
inherit_shaping_context, with_hyphenation_language and with_base_and_advances.
Public base allocations remain measured through their actual capacities.

Legacy cache admission additionally needs positive actual private Word-producer
association for otherwise inaccessible private language bytes. Preserve that
association through splitting/reflow/clones in existing private block/semantics
ownership and charge any retained scalar/vector facts with exact equality. Do
not infer it from text, exposed language, source provenance or an arbitrary
MultilingualTextSegment flag. Missing proof refuses cache admission rather than
undercounting, while existing ordinary cache positives remain required. True
multilingual, anchor and opaque Group exclusions remain. No seventh public getter
or new public model is authorized. Spare-capacity factory/replacement, split,
hyphenation and warm/cold discriminators belong to the existing carriers.


## Legacy ordinary shaping-chunk preservation

The two authenticated finite15route packets distinguish old Unicode shaping chunks
from extra script-only partitions. With the original em-dash witness, normal wrapped
old/candidate glyphs, advances and line widths agree. With the same witness minus
only the em dash, the actual old plain pipeline emits one chunk. The legacy factory
adds a script-only boundary and changes the final fi substitution and width by
-0.05908203125pt at468pt/wraptrue, both kern states. These are prospective low-level
ordinary pipeline facts, not current main-Word adoption or native parity. Retain both
original packets and the earlier18provider observations without promoting their
producer scope.

For legacy_shaping=true, suppress shaping splits caused only by the coarse script
classifier. Retain actual bidi-level changes and explicit-direction resolution and
barriers. Shape each complete old per-segment Unicode chunk with the existing guessed
segment properties and that occurrence's explicit kern. Do not impose the entire
pre-break occurrence's guessed properties across old independently reshaped chunks.
Select these Unicode SHAPING cuts even when no_wrap=true. Publish all soft
break_after=false in no-wrap mode, including nonempty segment ends. Empty occurrences
retain their explicit context and false break. Preserve source scalar ranges, base
metrics, glyph/advance/cluster ownership, request-count refusal and occurrence alignment.

Use the existing private level-helper break_offsets argument as a pair of existing
HashSet references, shaping and published, destructured into explicit local names.
Nonlegacy callers pass the same existing offsets for both. Legacy no-wrap callers
pass the per-segment shaping cuts plus an empty published set. This adds no eighth
argument, public method/type/field, helper framework, module, crate or feature flag.
Generic/default shaping, nonlegacy rich/PPT language/fallback/bidi/positions and caps
remain unchanged. All six approved doc-hidden public signatures remain unchanged.
Legacy raw explicit-direction/mixed-bidi arguments retain their current resolution,
actual level barriers and validation, without claiming equality to the ordinary Auto
plain pipeline. Do not reject accepted arguments or erase a level barrier to obtain
ordinary equivalence.

Reuse the existing line breaker, actual hyphenation-prefix/remainder reshaping,
context-qualified spacing/tab widths and private-context-preserving transformations.
Keep the generic per-Unicode-chunk reshape convention, source/cache ownership and
structural-marker representation. The current Word rich path remains nonlegacy.
Future ordinary K1 adoption must not insert new script/language slicing before this
provider. Existing private exact-capacity normalization and retained allocation
accounting remain mandatory, with no new full-text owner or cache authority inferred
from legacy metadata.

Extend existing font and line test bodies with the two exact witnessed inputs,
both kern states,468pt fitting and20pt actual-wrap/overflow controls, public no-wrap,
empty/equal-occurrence alignment and source/context assertions. The no-dash20pt case
is unbreakable overflow, not genuine wrap. Preserve existing explicit-direction,
rich/PPT/caps, tabs, hanging-space, hyphenation/reflow and constructor controls. Batch
one affected lib artifact acquisition and execute the finite named children, with
current source/artifact/profile provenance. The two actual15route packets supply
finite pre-repair evidence. A newly built successor and fresh independent microscope
are required after source work. Full K1 paint/R4/cache and native/full/risk/hash gates
remain pending. No baseline, status or completion change follows from this amendment.

Root adopts this bounded technical amendment on2026-10-30after reading the complete independent proposal, both actual15route packets and the existing caller/test worklist. Local source, native and final acceptance remain distinct.


## Main Word finite producer closure

Root approves this private-carrier clarification before main implementation on
2026-10-30after inspecting actual marker construction, shared line breaking and
the independent concrete callable answer. The additional seventh method is only
the precise operation adopted in the mixed legacy producer closure below.

Keep atomic list labels as actual InlineItem::Marker, whole-shaped with explicit
kern from their own effective marker_rpr before display scaling. They retain the
old grouping, item-metric and no-internal-break behavior. The existing shared
breaker measures/clones Marker, including rich bidi movement, without reshaping
it. Do not run labels through the ordinary legacy Unicode-chunk factory or change
them into MTS. Numbering suffix text has its own actual producer style and follows
the normal policy-bearing text route.

Structural Target/SequenceContext producers retain their exact synthetic text,
glyph0andzeroadvances. At construction the existing private per-input style
sidecar explicitly records finite origin Structural versus AtomicMarker with its
own kern. Every other Word Text/HyphenatedText becomes policy-bearing MTS before
source-stage closure, including empty and field-owned inputs. Remaining plain
field members must therefore be actual known structural producers. Plain nonfield
atomic list markers remain outside field-only map membership.

WordParagraphPaint may additionally own private plain_field_policy:Option<(bool,bool)>
and an exact structural producer count. Some(false,true)is established only by the
actual Word producer after verifying the complete conversion invariant, never by
inspecting FieldKind/text/font/source/provenance to guess ownership later. Fragment,
reflow and permutation operations transport/validate that proof and its actual
selected count. The checked emitter reads the positive owner fact while producing
each remaining plain field leaf and captures its exact positioned path. Missing
proof or an unexpected plain field refuses. Prepainted Group/row/shape subtrees
still require their exact separate companions, not this paragraph fallback.

The same private Arc may carry an optional positive legacy-cache producer proof
with exact retained private-language/allocation facts, established only at the
known Word construction route. Source-supplied/neutral blocks cannot acquire this
proof from MTS flags or equal public values. Charge all new private scalar/vector/
Arc storage under existing ceilings and preserve exact value equality. These
additional private facts refine the earlier anchors/rows-only sketch. No public
model/field/signature, source file, generic, trait, flag or API is added.

Discriminate whole atomic labels and synthetic markers, marker metrics/grouping,
rich bidi transport, empty fields, unexpected plain-field refusal, known-neutral
fields and positive/missing cache producer proof within the existing carriers.
This closes callable transport before code, not runtime/native acceptance.

## Stored-field annotation owners and opaque-origin bridge

The six approved shaping methods remain unchanged. The mixed legacy producer
closure adds only its precise seventh method. Their public API restriction has
one explicit additive exception, a neutral doc-hidden function in the existing public
oxml_layout::line module:

```rust
#[doc(hidden)]
pub fn break_into_lines_with_opaque_origins(
    items: &[InlineItem],
    params: &LineBreakParams,
    fm: &FontManager,
    base_direction: Option<TextDirection>,
) -> Result<(Vec<LayoutLine>, Vec<(usize, usize, usize)>)>
```

Each origin is (output line, output item, original top-level input slot). None selects
the existing plain breaker. Some(direction), including Auto, selects the existing rich
breaker for that direction. Existing entrypoints retain their signatures, defaults
and exact output. No new public type, field, trait, module, feature flag, dependency or
lib.rs re-export is required. HLD08 remains the only HLD impact.

The actual breakable-segment constructor captures every top-level Group/Figure slot.
Private paired storage carries it with the same item through segment queues, LineState
push and actual flush. Rich bindings move by the same actual bidi visual_order as the
items. No Group value, text, coordinates, source span, address or encounter matching
recovers an origin afterward. An outer Figure has one origin even when its child is
a Group/Figure. Nested Group/Marked children do not invent input slots. Text, Marker,
MTS, tabs and images have no origin entry. Empty input retains its existing empty line
and empty origins. Zero-width opaque roots retain their actual occurrence.

Origins are ordered by output line/item. Check all bounds, unique output paths, unique
input slots, exact Group/Figure kind alignment and complete cardinality. Every admitted
opaque input has one binding and no other input has one. Missing, duplicate, wrong-kind
or over-budget metadata refuses through LayoutError without partial publication.
Use checked count/byte arithmetic and fallible metadata reservations. Existing layout
and shaping failures propagate. Old untracked calls need no retained origin allocation.

At a known stored Field owner, before StoredFieldRun splitting, mint one private
operation-local occurrence token shared by that owner's stored runs. Child recursion
mints a distinct token. Do not reconstruct it from emitted FieldSource, source spans,
Group ordinals or addresses. A token identifies the known current owner, not its
shaping policy. Preserve the actual source-qualified cached-child kind set: unlocked
PAGE, NUMPAGES, SECTION, SECTIONPAGES and PAGEREF/TargetPage. SequenceRepeat is not
produced by this stored-child route and gains no new support here. Structural
Target/SequenceContext remain synthetic and never acquire visible annotation recipes.
No new source=None annotation fallback is introduced.

For a known stored-field EastAsian/emphasis producer, carry its resolved source
annotation context before opaque Group construction. Base-result glyphs acquire their
actual field roles and Word shaping policy. Brackets, emphasis marks and decoration
remain non-field. Non-field ruby/EastAsian/emphasis/OMML/chart Groups remain outside
field-map membership. Do not exclude a known field result merely because old annotation
code erased its roles.

The private per-input context records each actual emitted Group slot. Word shaping
transports that slot when ordinary text expands into MTS. WordParagraphPaint retains
these input contexts and the bridge's exact line/item association. Float reflow uses
the same bridge on the retained input vector and updates private selected-line contexts.
Retain the known producer's actual plain/rich decision in this owner. Newly introduced
legacy MTS alone does not select rich layout. Existing post-break row projections keep
their exact keys and unsupported structural-reflow boundary.

Checked emission records annotation roots at their actual positioned paths, beside the
field-only leaf map. Their separate descriptor is not a field leaf/cardinality entry.
It owns the finite annotation variant and actual resolved font/metrics, scaling,
spacing, decoration, explicit kern and required source owner facts. Group/Marked wrap,
splice, split, removal, permutation, rotation and clip transport its exact root path
in the same checked operation as the leaf policies. Validate expected root kind and
base-member association. Missing or inconsistent context refuses.

Late substitution preserves CURRENT traversal and seen order. Ordinary and annotated
fragments of one actual result share the existing first-fragment transaction. If an
earlier ordinary fragment consumed a STYLEREF value, remove the complete continuation
annotation root and its decorations instead of regenerating it. Otherwise regenerate
the complete annotated value once through its saved actual annotation producer and
remove continuation chunks and marks. Admitted cached-child page kinds likewise use
one transaction per known annotated owner. Preserve clear_kind/unplaced handling and
structural removal before STYLEREF. Do not change unrelated ordinary fragment behavior.
An empty final STYLEREF removes its whole annotation root, recipe and leaf map.

An empty cache may emit zero StoredFieldRun. Capture its known owner and inherited
style before flattening loses them, creating an actual empty field leaf and recipe,
or refuse explicitly under the empty-field contract. A known empty-leaf root descriptor
permits later nonempty annotated replacement while retaining initial source/caret
metrics. Do not manufacture visible brackets/emphasis for an empty value.

Keep Group/rich/anchor cache exclusions. Compare all added producer/root context by
value and charge every retained Arc/control block, vector/path capacity, owned string
and annotation-property allocation under the existing ceilings. Never truncate or
infer a policy from source identity alone. Native and full risk/hash qualification
remain pending.

Local cases use existing test bodies and one coherent affected artifact per binary:
byte-equal Groups with opposite owners through plain/rich bidi and wrap, exact parity
with old entrypoints, tabs/leaders/hanging text, empty/zero-width roots, Figure(Image),
Figure(Group), nested children, stored annotation fields through repeated float reflow,
shared ordinary/annotated traversal, actual admitted cached-child kinds, true empty
owners, missing-context refusal and unchanged non-field annotation controls.

### Known repeated table-header paint scope

At the actual repeat_table_headers paint operation, mint a private repetition scope
for known annotated source owners. A column-only transition can repeat a header on
the same physical page. Preserve each such paint, while sharing continuation dedup
only within that repeat. Carry its scope through all relevant ordinary and annotated
fragments of the same known owner, including nested private carriers. Do not infer
a scope from FieldSource, text, coordinates or addresses. Unrelated neutral and
ordinary field behavior remains unchanged. Reset the repetition namespace only at an
actual physical-page boundary. Context identity and allocation charges remain exact.

The local discriminator includes a multi-column same-page repeated header, a split
continuation that must not become a new repeat, and a new-page repeat. This private
clarification adds no public API/model and preserves the single neutral origin-function
exception and all final native/full/hash gates.


## Mixed legacy producer context closure

Root adopts the exact technical contract retained at
/private/tmp/rdocx-s96-run-refresh/fx273-mixed-legacy-producer-closure-v1/REPORT.md,
SHA256 c7afbb162624a1cc23f8859a589128fd11e08aace1deec7297aa60c15f34b6ec.

In the existing MultilingualTextSegment implementation only:

```rust
#[doc(hidden)]
pub fn with_logical_index_and_legacy_shaping(
    self,
    logical_index: usize,
    legacy_shaping: bool,
) -> Self;
```

It changes only the logical span ordinal and legacy reflow/paint convention. It preserves every base field, actual glyph ID, advance, position, cluster, script, direction, bidi level, break-after value, kerning override, language and hyphenation-language value. When the resulting legacy flag is true, reuse the existing exact-capacity private-storage normalization. Base public allocation capacities remain directly measured. False does not reshape or reset anything. It establishes no Word ownership or cache proof. No new public type, field, flag, dependency, old signature change or further API is added. This is exactly a seventh method plus the previously adopted one neutral origin function.

## Actual producer and assembly rules

The original rich candidate gate excluded empty text and field_kind inputs (retained before engine15395). The old rich batch and its actual order remain intact. Do not include excluded plain computed-field fallbacks merely because their surrounding paragraph is rich. The finite plain-field fact is assigned at the known field expansion/stored-child producer before shaping, in the existing private per-input style carrier. It is not inferred from erased output FieldSource, text, geometry or addresses.

For a rich Word paragraph, shape exactly the old eligible rich-input batch once, with explicit per-input kerning. Each actually excluded plain-field occurrence goes through its own legacy factory context, with that source occurrence's actual base, language, explicit direction, kerning and no-wrap facts. The already-approved legacy factory supplies Unicode shaping cuts, guessed properties, explicit bidi barriers and no-wrap publication behavior. Empty nonstructural Word text uses its actual explicit legacy context. Structural and atomic Marker inputs retain their previously adopted routes.

During actual source-order assembly, consume the actual rich queue or the actual known plain owner output. Assign checked monotone global span ordinals, one per real emitted MTS, with the new operation. Overflow refuses before publication. Never manufacture dummy prefix inputs, match output values, infer an owner from ordinal, restart a colliding local ordinal domain or shape duplicate whole paragraphs. Record occurrence companions using the final assigned ordinal before public output erases context.

An actual original HyphenatedText source whose rich-shaped span satisfied the existing `!needs_word_multilingual_layout(span.text())` gate used to become HyphenatedText (before engine15509). Retain that exact producer condition. Preserve its actual initial rich glyph/position/cluster result and use the new operation to attach the old legacy split/reflow/paint convention. Retain its actual hyphenation language and explicit kerning. Genuine rich siblings remain rich. No general script-classification or ambient policy is introduced.

Shared line and hyphenation transforms preserve the global ordinal and context. Existing positive Word cache association remains required independently of flags. Rich/Group/anchor cache exclusions stay unchanged. Charge any new private producer fact or retained allocation under the existing ceilings and compare its exact value. No configured cache bound changes.

## Mixed hanging-space rule

Legacy-carried MTS already participates in the plain physical-end space test. Exclude such spans from BOTH rich-only `last_content` ordinal selection and rich-only `hangs` ordinal comparisons (shared line312/319). This reproduces their former Text participation while genuine rich spans retain the existing logical ordinal rule. Do not globally disable rich hanging spaces or infer legacy behavior from text or FieldKind during painting.

## Rich-line permutation of legacy carriers

The rich line breaker currently overrides the computed level with every MTS shaped
level (line1247) and rebuilds every moved MTS through multilingual_segment_with_level
(line1349). A former Text item did neither. For legacy-carried MTS, use the same
actual text and base.direction path as former Text without a shaped-level override.
The actual bidi slot permutation moves that MTS unchanged, preserving its glyph,
position and cluster payload. Do not rebuild or reverse its glyphs/clusters through
the rich level helper. Genuine rich MTS retains both existing shaped-level override
and level-adjusted glyph rebuilding. The operation-local origin bindings follow the
same actual permutation as before.

Include a finite mixed RTL/neutral/plain-field/trailing-space parity control against
the former Text representation, checking visual slots, unchanged actual glyph and
advance payloads, plain hanging-space counts and repeated reflow. This shares the
existing named test bodies and adds no public operation.

## Finite evidence

Retain the current old observers in `field_only_directional_paragraph_keeps_bidi_through_drawing_reflow` and `rich_mixed_script_paragraph_retains_conditional_hyphenation`. Their existing actual failures become the local REDs for the producer correction.

In existing bodies only, discriminate equal plain field spans with opposite kerning policies next to genuine rich siblings, unique monotone source-assembly ordinals, original Text publication and rich-sibling publication, Latin conditional hyphens with explicit off/on context, RTL/trailing spaces using the exact plain-versus-rich hanging rules, and repeated drawing/width reflow with held/warm/fresh equality. Exercise the new method's unchanged glyph/base/cluster/language/kern values and exact legacy private capacities on spare-capacity inputs. Keep current rich language/spacing, old 18 provider children, origin bridge, annotation/repetition, note/layer/cache and original public RED controls in the same coherent affected artifact batch. The ignored native/HarfBuzz caps oracle remains final debt.

The other failures retain source-based corrections: explicit-off reference plus the same 2pt spacing assertion, both unchanged cache ceilings plus exact actual FIFO suffix identity, and new cache discriminator selection by the actual current paragraph key. Preserve all failed raw observations. Native/full/hash/raster/installed/WASM/hosted/independent-review gates remain pending.


## Known legacy glyph cluster order clarification

Root adopts the exact technical clarification retained at
/private/tmp/rdocx-s96-run-refresh/fx273-legacy-cluster-order-closure-v1/REPORT.md,
SHA256 3e58a5a150ba782d9235767f32d09dd7640917886302f9c99212f4533fb79440.

Actual green-children-2 retains 45 PASS and 2 FAIL. One new fixture assumption concerns the existing fixed99 PAGE placeholder. The shared failure is an actual factory refusal before line layout at line.rs3256 for the finite numeric/trailing-space control. Source analysis identifies independently guessed legacy glyph order versus layout bidi parity as the cause candidate. No extra runtime experiment or implementation has followed it.

Keep every actual glyph, advance, cluster, source range, layout bidi level, direction, script and break boundary unchanged. Preserve the actual bidi/explicit-direction partition barriers. Never reverse clusters to pretend a guessed glyph buffer followed layout bidi direction, change stored layout levels to satisfy validation, or silently use rich explicit shaping for legacy content.

Use one private context-aware constructor in the existing MultilingualTextSegment implementation. Its parameters are the existing constructor's existing types plus the existing legacy_shaping bool. Public new retains its exact signature and strict default by selecting false. For a known legacy constructor only, a cluster map must satisfy the existing complete coverage/range/monotonic validator in either actual LTR or RTL glyph order, independently of layout bidi parity. Reject when neither order is valid. Keep all other glyph counts, finite positions, source text lengths and valid bidi-level checks exactly. No new public operation, field, model, hidden scalar or retained allocation is required.

Use that private constructor at the existing explicit-context factory empty and shaped-range branches (font2645/2983), context-preserving reshape (font3118), and known legacy hanging-space slice (line718). Preserve the original context flags during construction rather than only inheriting them after a strict constructor has already refused. Genuine rich/default callers retain parity-coupled validation. The existing with_base_and_advances method uses the same either-valid-order rule only when its owner already carries explicit legacy context. Its glyph identity, finite-position and source/base consistency checks remain unchanged.

The seventh method remains exactly its approved signature and payload-preservation contract. Production assembly preserves false for already rich results, preserves true for actual legacy results, or sets true on the actual old Latin-hyphen producer. It does not request false on an oppositely ordered legacy result or manufacture a rich reshaping. Generic construction and unmodified existing callers do not acquire the private legacy exception.

Retain all immutable failed observations. In the same existing bodies, require the numeric/trailing-space factory and mixed old-Text parity control to pass, including the rich-initial Latin span with whitespace and repeated width reflow. Retain strict public invalid-level/map refusals, and add an explicit genuine-rich parity mismatch refusal beside an accepted fully covered opposite-order legacy control. Both malformed directions remain rejected. No test may change source bidi levels, drop the control or tolerate a refusal. Reacquire both affected binaries together and rerun the same finite shared/Word batch after source correction. Native/full/hash/raster/installed/WASM/hosted/independent-review acceptance remains pending.
