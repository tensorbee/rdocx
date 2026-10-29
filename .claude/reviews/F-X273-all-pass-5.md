# F-X273, all, pass 5

**Reviewed**: Frozen working result at HEAD `1431cb408dc57a24353bf826933408b0ae2ab8ec` on `sprint/s96`,7files,89additions and9deletions. Prior full pass3 and affected pass4 production audits, plus F-X275 pass2 preparation review, are explicitly consumed for unchanged scope. Fresh review covers the test observers, added positive geometry assertions, source-supported oracle disposition and affected archive measurements.

**Verdict**: 0 defects,0 smells,0 nitpicks. Local review exit condition is met. Final integrated/native/differential and release acceptance remain pending.

## Identity and review boundary

All seven file hashes match `f273-observer-contract-freeze29.json`. Independently recomputed working binary diff SHA256 is `4724caa80614b2493c18f0ae7a3130bde049c9d14f1e4a1bdf552152dc48af0e`. Scoped26 receipt SHA256 is `fc907caa1bbadc2cb569a7b8b7bad0b2498922b16c734e40bd9f6d87fa0b27f3`. Archive validation28 receipt SHA256 is `cef3773985fdcce36fdcb0cc8bd2215264104d6b75f7d9114a70c8545ec4c68f`. Every scoped26 recorded artifact/log hash checked matches.

No builds, tests, native apps or source remediation occurred during review. Only this report was written. Current production-prefix comparison proves engine.rs is byte-identical before its existing cfg-test module. All changed Rust hunks are within that module or the integration test. Existing runtime, note planner, painter, navigation and constructor behavior remain unchanged from the reviewed production candidate. No confirmation audit of unaffected runtime scope was repeated.

## Defects

0 findings. Prior D3 retains its pass4 addressed disposition.

## Smells

0 findings.

## Nitpicks

0 findings.

## Test assertions and affected consumers

At `crates/rdocx/tests/integration_test.rs:24738`, the restart fixture retains exactly two physical pages and both body-reference B labels. Exact per-page nonempty trimmed line inventories additionally require one standalone B, the correct unique note text and correct order. Page2 retains Page boundary before its body. A C restart, duplicated marker, moved/duplicated note or wrong-page text fails these inventories. Trimming removes extraction whitespace, not labels or page ownership. The original source authors no paragraph indent/tab, and create_footnote publishes an inline styled raw footnoteRef followed by text. The former four-space expectation was a stale synthetic-indent observation, as independently diagnosed before this correction. No source indentation was added merely to restore that output.

The parent test still calls this same child. Actual full06 remains470pass/2fail/8ignored, with both failures at the old child assertion. The corrected scoped execution does not retroactively turn that full run green. Actual local26 includes both affected tests through the all-feature nativeAPI selection.

At `crates/rdocx-layout/src/engine.rs:29476`, the paired-table matrix preserves its original unformatted FLOW_CONTROL XML and all table overlap, separation, row-offset, emission and source-immutability assertions. Its new baseline contains that identical paragraph with no tables, separately for header/footer. Body x/y and physical page count must match that baseline across one/two rows, either-never, both-overlap and disjoint controls. This rejects accidental table flow clearance without replacing inherited auto259/8pt-after source spacing with an exact12pt fixture or fitting a measured constant. Actual scoped25's literal36failure remains a test-authority failure. The independent v2 diagnosis correctly distinguishes that source from the exact-spacing native table controls.

At `crates/rdocx-layout/src/engine.rs:30134`, the new horizontal Bottom control explicitly authors five exact12pt, zero-spacing paragraphs on both selected furniture sides and an exact12pt body. Written geometry and furniture derive effective band78to714. The assertions require one page, x72, body band y702 and height12. This locks the approved effective-band Bottom policy without confusing Word's text bbox top701.6088with the layout line band. Existing Center, line-number and furniture-position checks remain. Both additions are positive coverage of existing behavior, not invented historical product REDs.

## Oracle contract assessment

The wording at `.claude/plans/F-X273-design.md:94` and `docs/hld/08-rendering-spec.md:1939` follows existing Word-target authority and the canonical differential-testing rule5. It does not require a new product design decision for these bounded dispositions. The original contract already makes native Word deciding for header metrics, requires selected story identity, excludes floating ownership from ordinary flow and specifies effective-band alignment. The approved frame and exact-source tbRl amendments already select their bounded Word compatibility behavior. Canonical oracle triage permits a cited intentional divergence with our side asserted. It does not require changing the renderer to imitate every conflicting oracle.

Inspected the per-case map, retained raw comparison and extracted pinned ECMA text. Clauses17.10.1/5/6 support section-start parity, same-type inheritance and first-section blank ownership. Clause17.4.57 places floating tables outside normal flow, with17.4.56table-only overlap settlement. Clause17.6.23defines main-story vertical alignment relative to text margins. Generic frame text-default prose is not misrepresented as proof of the observed omitted-margin furniture context. That context and tbRl numeric behavior retain their separately qualified native authorities and approved narrow boundaries. No LibreOffice import/layout cause is claimed as established.

All11LibreOffice discrepancies remain failed numeric comparisons, with original inputs, thresholds and measured outputs preserved. Their bounded HLD compatibility limit and cited existing Word-target disposition do not waive an explicitly promised matching consumer, unexplained mismatch, complete font/furniture/raster coverage or operation/composition evidence. The revised plan expressly preserves those limits. The per-case map's older optional literal36recommendation is superseded by its explicit v2 correction and the source-matched assertion actually applied. No failed raw result is erased by that correction.

This review accepts the consistency of the clarification, not final disposition of every acceptance row. BODY-only results cannot qualify actual table settlement/paint, frame/FLOW positions, note content/geometry, complete fonts, saved field operations or union composition. Those remain separate pending criteria.

## Evidence applicability and packaging overlap

The geometry03 collector records98successful CLI artifact rows. Actual BODY comparison03 reports Word42pass/0fail and pinned LO31pass/11fail at0.5pt with exact physical pages. Collection source/font/input/binary guards match before and after. Comparing the collector's source inventory to current files finds only engine.rs changed, and its runtime prefix is unchanged. Integration assertions, documentation and measurement constants are outside that runtime. Accordingly these are applicable executions from producer1431cb40with the recorded original binary and input/font configurations. They are not new executions at the dirty review snapshot, and they do not establish full native acceptance.

Scoped26 records93lower and10all-feature nativeAPI tests passing, plus strict three-crate all-target/all-feature clippy, fmt, prose, adapters and whitespace. Its actual inputs and test assertions were inspected. No absent full-suite/check/CLI execution is invented for this scoped sequence.

The two affected archive payloads were actually packaged with no verification in preparation27. Independently authenticated both retained archive hashes and compressed sizes, normalized member sizes under the existing VCS normalization algorithm and member counts. They match the new constants: rdocx4,139,012compressed/54,306,876normalized member bytes/37members, rdocx-layout407,983/2,235,061/15. Both are below10MiB. README copies match those rows. Validation28 checks current README payloads and four affected measurement policies. Its normalized rdocx compression differs by2bytes within the existing tolerance, with exact normalized member size/count preserved. The20other archive rows and historical performance claims remain unchanged.

These are affected factual preparation measurements, not verified publication artifacts. F-X275's final clean verified22-package dry run, archive compilation and exact-main release approvals remain required. No version, dependency pin, family order or public runtime signature changes in this diff.

## Not found and remaining boundary

Correctness, contract, panics, OOXML, tests and structure:0 additional findings in new material and affected consumers. No new production abstraction, XML operation, module, source file, trait, generic or public field is introduced.

All remaining native/full/hosted/installed/risk gates stay pending, including complete font/furniture geometry, deterministic PDF/PNG/text and reviewed49hash/7pixel deltas, F-X272/F-X273/F-X274 composition, installed Python3.9/3.12runtime/typing/stubs, WASM, genuine verified archives, README qualification and final metadata. Actual full06 failures, scoped25failure and native UI permission/recovery failures remain attributed to their original observations. No acceptance waiver, baseline update, publication or final completion is granted.

End review and return sole tracked-writer control to root. Later material source changes require their affected evidence and review. No redundant unchanged-scope confirmation pass is requested.
