# F-X274, all aspects, pass 1

**Reviewed**: Full working tree against immutable Base and HEAD `8cb33ef53717159bcf73f655552b99589272a274`, branch `work/f-x274-codex`. Eleven tracked files, 1844 additions and 244 deletions. No staged-only or untracked implementation omitted.

**Frozen identity**: Binary diff SHA-256 `9c2f8b8b402b0c1b9ed9444ad829e9067eb1fe60950ea1f447d69e17101648a4`. Independently recomputed the diff hash and every changed-file hash in `fx274-local-freeze.json`, with zero mismatches. Product227 manifest `d3e63f4d80da582ab315a2714fcddbff0f42422eb5fc800ab562d69f2a2e29ba`. Final Python test SHA-256 `0ed2fa14c3e86470c4730f1c2087ecde30dfd6509def91d6a1167faf19e58728`.

**Verdict**: 1 defect in test evidence, 0 smells, 0 nitpicks. No separate production correctness defect established. Return for bounded remediation and a distinct successor microscope. This is not final acceptance.

## Defects

### D1, the direct multilingual visual-order assertion is unreachable

`crates/rdocx-layout/src/engine.rs:30247`, `crates/rdocx-layout/src/engine.rs:30418`, `crates/rdocx-layout/src/engine.rs:30440`.

The direct-consumer matrix supplies `STORED`, an empty string and `Latin العربية`. Its explicit visual-order assertion runs only when `stored == "אב العربية"`, which is never supplied. The remaining assertion sorts spans by logical index before reconstructing expected text. The geometric assertion proves contiguous advances, but does not distinguish visual permutations of the spans. Consequently, an implementation that publishes the right logical text and metadata in the wrong visual order can pass this test.

This matters for the newly introduced multi-span publisher at `crates/rdocx-layout/src/engine.rs:9241`, which delegates visual ordering to the existing line breaker before assigning successive x positions. The committed rendering intent promises existing bidi visual order at `docs/hld/08-rendering-spec.md:2026`. The in-flight evidence statement also specifically claims adjacent Hebrew/Arabic coverage at `.claude/scratch/F-X274-progress.md:104`. That particular observation is not supported by the current matrix. The standalone Word mixed-script text assertion at `crates/rdocx-layout/src/engine.rs:30230` does not replace this discriminator, since dynamic field segments are excluded from the multilingual producer at `crates/rdocx-layout/src/engine.rs:14770`.

Use an actually admitted multi-script direct-carrier case whose expected visual order is asserted, retaining Group and MarkedContent, logical-text reconstruction, zero/one/many and source/field/note metadata checks. Record its actual focused outcome and correct the evidence attribution. This is a source-proven test gap, not an observed product visual-order RED. No test or product code was changed during review.

## Smells

Zero.

## Nitpicks

Zero.

## Contract and consumer review

Read canonical CLAUDE, WORKFLOW and microscope instructions, the approved annotated design including namespace-capture and original-input qualifications, progress, and the affected HLD08/HLD10 sections. BACKLOG and CURRENT_SPRINT identify F-X274 as in-progress. The sole-writer review boundary was respected. No builds, tests, helper generation, product imports or native applications ran during this pass.

- **Shared formatting and bounds**: `engine.rs:955` validates page-family instruction shape before formatting. `engine.rs:1000` owns supported section qualification, numeric override versus character-only processing and shared staged formatting. `engine.rs:931` preserves ordered general switches and selects repeated-letter page formatting rather than ordinary spreadsheet letters. `engine.rs:1149` bounds repeated output before allocation. The only new exported operations are the two approved concrete doc-hidden helpers. No model field, authoring enum, binding method, trait, generic, crate or source file was added.
- **Section source authority**: `document.rs:1542` merges effective URI-bearing scope with owner overrides, preserving the verified canonical Word binding policy. Both Empty at `document.rs:1574` and Start at `document.rs:1930` use that scope before first-owner selection, so later duplicates receive the same closure and remain opaque. `engine.rs:1034` reads a temporary clone through the existing writer, retaining root-attribute records while excluding opaque child candidates. It resolves expanded names for the direct modeled owner and its attributes. A malformed/unqualified modeled carrier cannot borrow a later opaque duplicate. Writer/reader failure remains unsupported. The original section is not mutated.
- **Absence and implicit final section**: absence of a page-number model is qualified decimal, while unsupported source remains distinct. Both layout at `engine.rs:5346` and facade aggregation at `field.rs:16747` append the paginator's implicit final decimal section only when body sectPr is absent. They do not turn arbitrary missing indexes into supported section state.
- **Owner and target context**: `engine.rs:705` records effective physical instructions and stored displays, with inherited lock context retained separately. `engine.rs:4040` supplies a source registry for standalone page-family consumers, rather than relying solely on provenance layout. `engine.rs:5396` computes values from each actual placement, using the PAGE source or PAGEREF physical target section. Existing note-reference ownership is consumed through those placements. NUMPAGES, SECTION and SECTIONPAGES remain decimal by default.
- **Classification**: `engine.rs:3679` separates stored/formatted display from actual missing target status. Both Text at `engine.rs:9076` and MultilingualText at `engine.rs:9144` clear only an unplaced TargetPage. Unsupported formatting does not erase Page family/placement metadata. Locked owners and nested display owners retain the established shaping suppression and physical identities at `engine.rs:10680` and `engine.rs:10808`. The unchanged CardText placement and unresolved-target regressions are applicable discriminators for the repaired distinction.
- **Publication and reuse**: ordinary text, direct multilingual zero/one/many, Group and MarkedContent consumers were traced through `engine.rs:9050`. The direct-rich publisher retains baseline, field/source/note metadata and uses the existing bidi authority. D1 limits the claimed proof of its visual ordering. Per-page substitution identity at `engine.rs:5462` includes instruction, relevant owner/target format, lock, value and stored display. Retained context capacity is charged at `engine.rs:7525` and shrunk before retention. The shared format map is prepared before reuse decisions.
- **Cache lifecycle**: `field.rs:16580` keeps staged source identities, lock checks, unplaced refusals and dynamic header/footer cache exclusions. It invokes the shared provider using the PAGE placement or PAGEREF target section at `field.rs:16700`. Same-owner format/switch changes, held-handle invalidation and save/reopen are checked by the affected native/Python cases. No new public defaults or stub changes are required for the unchanged methods.
- **Generated labels**: `field.rs:12076` qualifies TOC physical targets and exact shared formatted labels through current source instructions, rather than parsing labels as physical indexes. `field.rs:25901` formats generated INDEX, caption TOC/TOF and TOA target labels, keeping range endpoints, ordering and sequence prefixes separate. All callers pass per-operation section facts. Unsupported or unresolved targets still propagate atomic refusal through the existing staged publishers.
- **Dependency wiring**: Cargo changes add only the approved direct workspace quick-xml edge to rdocx-layout. No version upgrade or forbidden lower-to-facade dependency was introduced. HLD impact is limited to the two approved files.

## A-ID evidence and limits

| ID | Evidence inspected and scope |
|---|---|
| A1 | The exact original and separately labelled intended numbering control remain distinct. Design lines155 to161 preserve singleton/order and first-owner facts. Local probes are not Word recovery evidence. Exact original native output and interpretation remain pending. |
| A2 | Lower standalone ordinary/bidi source cases, actual PAGE/PAGEREF distinct owners, target-format warm/cold mutation and selected first/even/default headers discriminate owner and target section context. Five formats are exercised in fresh Python render/cache/save/reopen cases at `test_core.py:15606`. Native variants and all promised final consumers remain pending. |
| A3 | Shared stage tests at `engine.rs:29966`, Python character/numeric-switch cases and CLI update cases at `integration.rs:7840` distinguish numeric override from section-preserving character switches and unsupported instructions. Final native switch observations remain pending. |
| A4 | Lower limits assert 26/27/28/52/53, Roman3999/4000 and the exact4096-character alphabetic boundary. Python covers supported and refused saved/render states. These are explicit library limits, not established native maxima. |
| A5 | Pure lower capture matrix covers fragment/full import, part/section root, default/alias/local-shadow, Empty/Start, duplicates, opaque payload/value identity, unchanged canonical bytes and start mutation/reopen. Helper tests exclude nested prior sections and malformed modeled fallback. Actual native source/reopen qualification remains pending. |
| A6 | Native existing imported complex-footer controls remain unchanged. Python paired render/cache/reopen checks exercise page-only, layout-backed and combined updates, all four footer placements, eligible body cache and stale held handle. The exact combined simple-field dirty/closure delta is attributed to predecessor/current probes, not waived generally. |
| A7 | `engine.rs:29932` and `engine.rs:30094` compare warm/cold format/instruction and unchanged-instruction target-format mutations. Cache capacity controls include the new context. Final full installed/runtime/typing union remains pending. |
| A8 | Existing qualified generated-target/navigation and provider regressions are consumed locally. Temporary external fixture preparation is not product evidence. Actual integrated prior-redline, tall furniture/merged heading, section label, PDF/PNG/text/native composition remains pending. |

The tests inspect actual labels, physical target/owner facts, caches, source bytes, metadata and save/reopen, not merely success totals. D1 is the exception in explicit visual-order qualification. No existing product assertion was removed by this diff. The corrected new footer oracle is supported by actual predecessor/current operation-specific receipts rather than a hypothetical normalization allowance.

## Evidence provenance consumed

All temporary evidence names below are under `/private/tmp/rdocx-s96-run-refresh`.

- `fx274-local-freeze.json` binds the current changed files, progress and full diff. Read-only independent hash verification found zero mismatches.
- `fx274-root-native19-auth.json` authenticates actual Rust native19 child exit0, membership, source and executed binary. `fx274-final-native.log` records19 passed in70.67s. This is a Rust binary, not Microsoft Word.
- `fx274-final-lower.log` records16 passed in3.81s. `fx274-frozen-oxml-page-number.log` retains the four applicable pure lower results. `fx274-final-cli.log` records4 passed in4.79s. The cfg-test alias-only successor is separately authenticated by `fx274-root-test-only-successor-auth.json` and does not invent a rerun.
- `fx274-final-scoped-successor.json` records actual strict four-crate clippy, fmt, prose, adapter and whitespace exits and hashes, with unchanged product manifest. Scoped check evidence is retained in the local receipts.
- Fresh wheel SHA-256 `7fd7c78096da1c34db3371a8a21eaffc1a779a868a3c165374ba817bbd4235d9`, actual isolated native `36dc3825120653caa756767ee0ac782d92b4e2d8c292d94a0210906b084f5e06`. `fx274-investigated-provider-python.json` and root producer authentication bind actual package/import/source. Python evidence is51 unchanged observed passes from the original51pass/1fail run plus one corrected-node actual pass in3.90s on the same product artifact. There was no single52pass run.
- Retained helper failures, pure capture RED, missing-operand RED, implicit-section failures, unplaced-target classification RED, CardText classification RED and direct-rich zero/many RED remain historical with their original producers. They are not relabelled current executions. The predecessor252-field observation is carried under its explicit applicability, not claimed current or used as a performance measurement.

## Not found and final debt

**Correctness**: zero additional production findings after owner/target, default/unsupported, stage, label, source/cache and publisher tracing.

**Contract**: zero additional findings. Original A1 remains explicitly unresolved at its final native boundary, not silently replaced by the intended control. Public signatures, resource bounds and dependency/HLD impact match the approved scope.

**Panics**: zero findings. New indexing uses qualified lookups or established identity invariants. The one-span pop is guarded by exact length. Numeric output bounds and checked depth decrement are preserved.

**OOXML**: zero findings. Scoped source closure uses existing carriers without namespace guesses, preserves unknown values/payload and first-owner source order, and does not expand general serializer/model policy.

**Tests**: D1 above. Zero other test defects established.

**Structure**: zero findings. Concrete private carriers and existing helpers serve named callers. No unsupported abstraction or public authoring surface was introduced.

All final gates remain due: actual source-qualified Word and pinned LO/Poppler observations, original-input native recovery and saved output, exact font and geometry proof, full workspace/native suites, installed3.9/3.12 runtime and strict typing/stubs, WASM, hash49/pixel7 with reviewed intentional deltas, archive/size/supply-chain, README/current measurements, hosted CI and integrated A8/risk union. No gate is waived or declared green by this review. No remediation occurred. Return sole write control to the orchestrator after this report.
