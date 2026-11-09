# F-X278 DeepSeek implementation review, pass 2 of 3

Date: 2026-11-02
Reviewer: DeepSeek V4.1 Flash (DGX)
Mode: independent implementation review. Report-only lease. No source, test, HLD, plan, progress, handoff, metadata, TEMP, Git, configuration, build, harness, oracle, network or delegation action was taken. This file is the only write.

## Verdict

Not accepted. Four source or documentation defects remain, nine coverage or hygiene smells remain and nine nitpicks remain. The repair1 handoff is treated as a claim and not as proof. The seven A278 criteria are reviewed individually below. A278-7 fails on D1. A278-1 through A278-6 pass with the caveats recorded. The final union work remains pending and is not called passed.

## Scope reviewed

Complete implementation diff, not only repair1:

- Base `5820539cd277b417ae060e15010eb73bd1d00b01`, branch `work/f-x278-deepseek`, intentionally dirty and uncommitted.
- 21 modified files, 2633 insertions, 302 deletions.
- Rust: `crates/oxml-layout/src/font.rs`, `crates/rdocx-layout/src/engine.rs`, `paginator.rs`, `table.rs`, `crates/rdocx-oxml/src/paragraph_properties.rs`, `crates/rdocx/src/document.rs`, `epub.rs`, `lib.rs`, `odt.rs`, `svg.rs`, `crates/rdocx-py/src/document.rs`, `crates/rdocx-cli/src/commands.rs`, `main.rs`.
- Tests: `crates/rdocx/tests/integration_test.rs`, `crates/rdocx/tests/regression_test.rs`, `crates/rdocx-cli/tests/integration.rs`, `crates/rdocx-py/tests/test_core.py`.
- Bindings surface: `crates/rdocx-py/python/rdocx/_rdocx.pyi`.
- HLD: `docs/hld/08-rendering-spec.md`, `10-bindings-spec.md`, `15-build-and-toolchain.md`.
- Evidence: approved plan, repair1 progress checkpoint, `AUTHOR_HANDOFF.md`, frozen peer1 report, repair1 logs under `/private/tmp/rdocx-deepseek-harness/f278-repair1/logs/`, author transcript excerpts.

## Provenance authentication

- Frozen pin manifest `/private/tmp/rdocx-cross-review/F278_REPAIR1_FINAL_FROZEN_PINS.json`: 112 of 112 file hashes match.
- Approved plan SHA-256 `cacd4844683b63f7c9fb2fbcde658d7b3b03fcdc36eb747768ac46fe9af02bae` matches.
- Worktree HEAD `5820539cd277b417ae060e15010eb73bd1d00b01`, branch `work/f-x278-deepseek`, 21 modified files. Dirty state is expected and is not a finding.
- F2 RED: `CHILD_STATUS=101`, `EXIT 101`, log `20261102-f2-image-export-red.log` records `assert_success` failure with `Font 'Calibri' substituted: selected 'Carlito' from the bundled tier`. GREEN `CHILD_STATUS=0`.
- F3 race: `20261102-cli-feature-on-postfmt.log` records the feature-off refusal. Sequential `postfmt2` passed child 0. Feature-off `postfmt2` passed child 0. This is a test-harness executable race and is not promoted to a product defect.
- F4: extension rebuilt at 09:20 with `cargo rustc -p rdocx-py --features extension-module ... -C link-arg=-undefined -C link-arg=dynamic_lookup`, copied to `pyext/rdocx/_rdocx.so`. `run_f4.py` exited 0 with `F4_PYTHON_TEST_PASS`. The earlier run failed at `test_core.py:954` with `AssertionError`.
- Historical CLI RED provenance: the initial full CLI RED log was overwritten and only an exact tool excerpt is retained. That provenance is kept separate from the new F2 RED log above. The overwritten log is not claimed as present evidence.

## Findings

Severity scale: defect blocks acceptance, smell requires a decision or evidence before acceptance, nitpick is non-blocking.

### Defects

#### D1. Python `render_pages` selects pages from a different layout than the one it renders when caller fonts change pagination

Path: `crates/rdocx-py/src/document.rs:13875-13885` computes `selected` from `layout_deterministic_with_options`, and `:13897-13903` renders the caller layout with that list.

Trigger: call `render_pages(pages=None, fonts=<nonempty bundled caller fonts>)` where the caller fonts change pagination, for example by making the document longer or shorter than the deterministic layout.

Expected: the selected page range must be derived from the same layout that is rendered, and the diagnostic twin `render_pages_with_diagnostics` must agree with it.

Observed: `selected` is built from the deterministic layout, then `layout_with_fonts_and_bundled_fallback_and_options` produces a different layout and `oxml_pdf::render_pages` is called with the stale list. Caller fonts that add pages drop them. Caller fonts that remove pages produce out-of-range indices. The diagnostic twin computes its range from the caller layout, so C9 sibling agreement fails. The pre-repair code computed `selected` from the caller layout, so this is a regression introduced in this diff.

Severity: defect. This is the one A278-7 acceptance blocker in this pass.

#### D2. Public Rust documentation still describes the old normal-fallback policy

Path: `crates/rdocx/src/document.rs:46308`.

Trigger: read the doc comment for `to_pdf_with_diagnostics`.

Expected: the comment names the bundled plus system policy helper that the function now uses.

Observed: the comment still says `Encode one normal-fallback bundle and return its owned diagnostics.` The body uses the shared policy helper at `crates/rdocx/src/document.rs:24934` and no longer matches the sentence.

Severity: defect. Public API documentation contradicts the implementation.

#### D3. Python stub docstring still describes normal fallback as the default

Path: `crates/rdocx-py/python/rdocx/_rdocx.pyi:1171`.

Trigger: read the `to_pdf` docstring.

Expected: the docstring states that bundled is the default and that caller fonts are ordered on top.

Observed: it still says `Render PDF with normal fallback plus ordered caller fonts.` The implemented default is the bundled tier, as the HLD and the plan C7 contract state.

Severity: defect. Published typing surface contradicts the implementation.

#### D4. HLD08 TOC style matcher wording overstates the accepted names

Path: `docs/hld/08-rendering-spec.md:2465`.

Trigger: read the TOC ink section against the matcher.

Expected: the HLD describes the matcher as requiring the identifier or name to trim to `toc` plus a numeric level 1 to 9, with the id form `toc` plus the level.

Observed: the HLD says a style `whose identifier or name begins with \`toc\` followed by a level from 1 to 9`. The implementation does not accept arbitrary suffixes. A style such as `tocCustom` is not matched by the implemented rule. The HLD therefore documents a wider contract than the code implements.

Severity: defect. This is the contract the next implementer will read.

### Smells

#### S1. `word_substitution_diagnostics` silently caps at 64 entries

Path: `crates/oxml-layout/src/font.rs:1925`.

Trigger: a document that produces more than 64 distinct substitution messages.

Expected: the cap is either in the approved plan, documented in the HLD, or tested with the truncation behaviour stated.

Observed: the cap is a bare `>= 64` early return with no comment, no plan reference and no test. Diagnostics after the sixty-fourth are dropped without a signal.

Severity: smell. Behaviour not covered by the contract.

#### S2. Repair1 checksum change lacks executed evidence

Path: `crates/rdocx/tests/regression_test.rs:31333-31341`.

Trigger: review the `dense_form` assertions.

Expected: a repair1 log shows the changed test executed and passing, and the comment explains the geometry delta.

Observed: the checksum, non-white pixel and foreground pixel expectations were updated. No repair1 log shows this test executed. The nearby comment still describes the old side-padding change.

Severity: smell. Incomplete local evidence, not a demonstrated source defect.

#### S3. Python diagnostic tuples are not content-asserted

Path: `crates/rdocx-py/tests/test_core.py:948-949`.

Trigger: compare `render_pages` bytes with `render_pages_with_diagnostics(...)[0]`.

Expected: at least one test asserts the returned diagnostic tuple content for a substitution route, so a route that silently drops diagnostics fails.

Observed: the tests assert bytes equality only. A regression that returns an empty diagnostic vector for the caller route would pass.

Severity: smell. Coverage gap.

#### S4. Python `render_pages` system-omission sibling is not tested

Path: `crates/rdocx-py/tests/test_core.py:969-987`.

Trigger: `render_pages(dpi=24.0, system_fonts=True)` with a document whose system route omits pages.

Expected: the system route omission is asserted, matching the deterministic route at `:960-961`.

Observed: the caller plus system route is tested, and the deterministic omission is tested, but the no-caller system omission is not.

Severity: smell. Coverage gap.

#### S5. CLI image branch is not exercised with `--system-fonts` or caller fonts

Path: `crates/rdocx-cli/src/commands.rs:908` and `crates/rdocx-cli/tests/integration.rs`.

Trigger: run the feature-on and feature-off CLI tests that were added in repair1.

Expected: the image branch, which contains its own layout selection at `commands.rs:908-925`, is covered with `--system-fonts` and with caller fonts.

Observed: the repair1 feature-on and feature-off tests exercise `--to pdf`, render and layout. The image branch layout selection is not covered. The F3 refusal path and the caller-only image path are therefore untested on the image branch.

Severity: smell. Coverage gap on a branch with a distinct policy decision.

#### S6. A278-4 test lacks the `cantSplit` control named in plan C4

Path: `crates/rdocx-layout/src/paginator.rs:1829`.

Trigger: delete the `row.cant_split` term from the bail condition.

Expected: the A278-4 test fails because the control row with `cantSplit` still moves whole.

Observed: the test lacks the control, so removing the `cant_split` bail would still pass. The `cant_split` semantics are not discriminated.

Severity: smell. The test does not pin the contract it claims to pin.

#### S7. Caller-over-embedded priority is not locally exercised

Path: caller font tests in `crates/rdocx/tests/integration_test.rs` and `crates/rdocx-cli/tests/integration.rs`.

Trigger: use a caller family that also exists in the bundled tier.

Expected: plan C8 priority, caller face over an embedded face of the same family, is asserted.

Observed: tests use a family absent from the bundled tier. The priority rule is implemented but not locally discriminated.

Severity: smell. Coverage gap.

#### S8. Python `render_page_to_svg(0)` bundled-omission route is not tested

Path: `crates/rdocx-py/tests/test_core.py:954-961`.

Trigger: call `render_page_to_svg(0)` on a document whose bundled route omits the first page.

Expected: the bundled route omission is asserted, matching the caller route at `:960-961`.

Observed: only caller and system routes are covered for the bundled omission. The default bundled route is not.

Severity: smell. Coverage gap.

#### S10. The `tbRl` test enters the arm but the asserted positions come from the normal layout path

Path: `crates/rdocx-layout/src/engine.rs:4803` and `:4806` merge into a local `effective` that is used only for the line policy at `:4808-4859`. The positions asserted by `issue433_tb_rl_paragraph_merge_clears_the_absent_sibling` at `:26930-26963` come from the normal layout merges at `:11352` and `:11357`.

Trigger: remove the two `merge_direct_ppr` calls inside the `tbRl` arm.

Expected: the test fails, proving the arm's merge is load-bearing.

Observed: the arm still sets `tb_rl_exact_baseline`, and the asserted indent positions are produced by the shared normal path. Removing the arm's merge calls would still pass. F5 is only partially fixed.

Severity: smell. Partial fix, and the test does not discriminate the repaired behaviour.

### Nitpicks

- N1. `docs/hld/08-rendering-spec.md:1162` lists row barriers without nested-table, while the next sentence includes nested-table. Make the list complete.
- N2. `docs/hld/08-rendering-spec.md:903` describes automatic-colour precedence in a way that reads as applying only to explicit `auto`. State the inheritance rule explicitly.
- N3. The A278-3 fixture uses style id `Grid`, not the `TableGrid` id named in the plan. It still exercises the contract but the identity differs from the plan text.
- N4. The generated TOC test asserts zero blue lines rather than zero underline lines. The HLD contract is colour plus underline omission, so assert both.
- N5. The same-element both-attributes hanging case is not asserted at layout level. The merge helper test covers it, but the engine path does not.
- N6. `render_pages` performs a redundant layout when `pages` is `None`. Compute the range from the layout that is rendered once.
- N7. `layout_with_font_policy` reports the system-font error before revision-style validation. Error precedence is not the plan's stated order.
- N8. The test named `font_modes_share...` does not assert what its name says. Rename or assert the shared mode.
- N9. The `dense_form` comment near `crates/rdocx/tests/regression_test.rs:31333` is stale after the checksum update.

## Peer1 finding dispositions

Peer1 report is the frozen `/private/tmp/rdocx-cross-review/f278-peer1-final-frozen/REPORT.md`. Each disposition cites source or evidence.

- F1, stale Rust docs for `layout()` and `layout_page()`: fixed. `crates/rdocx/src/document.rs:45691` and `:46420` now name the bundled plus system policy.
- F2, CLI image-export tests require empty stderr: fixed. New RED `CHILD_STATUS=101` with the Calibri to Carlito message, GREEN `CHILD_STATUS=0`, log `20261102-f2-image-export-red.log`.
- F3, CLI caller plus system route untested with feature on: fixed for convert PDF, render and layout. Caveat S5 for the image branch.
- F4, Python caller plus system and SVG routes untested: fixed for the named routes. Caveats S3 and S8 for diagnostic content and bundled SVG omission.
- F5, `tbRl` test never enters the engine arm: still open and only partially fixed. The test now enters the arm, but S10 shows the asserted positions do not depend on the arm's merges.
- F6, TOC stored-segment filter and generated route unexercised: fixed.
- F7, Wingdings 2 `w:sym` consumer unexercised: fixed.
- F8, row-split test lacks contract discriminators: fixed. Caveat S6 for the missing `cantSplit` control.
- F9, indent test used relative equality and 720 twips: fixed to the absolute contract.
- F10, style-only table borders asserted `>= 3`: fixed to exactly 3 and 3.
- F11, nested-table diagnostics allowed any substitution fact: fixed.
- F12, automatic-colour comment claimed a LibreOffice provenance: fixed.
- F13, TOC style matcher trimmed only leading whitespace: fixed. Caveat D4 for the HLD wording.
- F14, HLD08 said only `w:cantSplit` moves the whole row: fixed. Caveat N1 for the barrier list.
- F15, HLD08 automatic-colour wording covered only explicit `auto`: fixed. Caveat N2 for precedence wording.
- F16, HLD08 Wingdings wording named only `U+F0A1`: fixed.

No peer1 finding is recorded as a justified rebuttal. None was withdrawn.

## A278 outcomes

- A278-1, TOC entry ink: pass, with nit N4.
- A278-2, first-line indent: pass, with smells S6 and S7 and nit N5.
- A278-3, table borders: pass, with nit N3.
- A278-4, row split: pass, with smell S6.
- A278-5, automatic text colour: pass, with nit N2.
- A278-6, Wingdings 2: pass.
- A278-7, shared font policy: fail on D1. D2, D3 and D4 are additional documentation defects. S1 through S5 and S8 remain coverage or hygiene smells. The Rust control and implicit publishers are otherwise consistent with the plan.

## Zero-finding categories

Explicit zero findings were recorded for these categories after inspecting the full diff:

- Schema child order: no violation found.
- Unmodelled XML preservation: no regression found.
- `oxml-*` dependency direction: no violation found.
- Feature-flag gating: no ungated system-font access found.
- Revision-view parsing: no new mismatch found.
- WASM surface: no new API divergence found in the reviewed diff.
- Prose rules in tracked Markdown: no em dash, en dash or prose semicolon found in the reviewed files.

## Remaining gates, not passed

These were not run in this pass and are not called passed:

- Bulk suites: preserved baselines are integration 462 passed, 10 failed, 8 ignored, and library 391 passed, 17 failed, 1 ignored. Not rerun, not green.
- Hash harness: not run. The 14 keys were not re-recorded. `report:page1.png` is unchanged at the donor pin. `dense_form` is not a hash key.
- `hash14`, `hash49`, `pixels7`.
- Native Word.
- LibreOffice 26.2.5.2 at `cd7284b4cbbfeb507e630c1aac019f4157393acb`.
- `pdftoppm` 26.01.0.
- Installed Python 3.9 and 3.12 typing and stubs.
- WASM build.
- Wheels.
- Archives.
- Hosted CI.
- `cargo publish --dry-run`.
- Size assertion.
- Clippy.
- Full workspace tests.
- Integrated union.

Earlier-phase local evidence remains applicable only where repair1 did not change the source path. Repair1 local runs: `fmt` child 0, engine `issue433` 7 passed, rdocx integration `issue433` 1 passed, nested 1 passed, CLI feature-on 1 passed, CLI feature-off 1 passed, Python F4 passed, prose 0 violations, `py_compile` child 0, `cargo check` six crates all-targets child 0. Earlier applicable evidence: oxml-layout Wingdings 1 passed, rdocx-oxml 2 passed, rdocx consumers 4 passed, rdocx feature-off 1 passed, cache-nine 9 passed.

## Counts

- Defects: 4. D1, D2, D3, D4.
- Smells: 9. S1, S2, S3, S4, S5, S6, S7, S8, S10.
- Nitpicks: 9. N1 through N9.

## Attribution

Root supplied the earlier cache-fixture and empty-directory corrections and the initial finding directions. Peer1 supplied the F1 to F16 list. This pass independently authenticated the frozen pins, the plan hash, the RED and GREEN provenance and the diff, and it records its own dispositions. This pass does not claim acceptance on Root's behalf and does not reset any count.

No remediation was performed. Control returns to the orchestration command.
