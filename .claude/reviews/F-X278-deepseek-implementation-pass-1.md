# F-X278 implementation review, pass 1 (DeepSeek)

- Base commit: `5820539cd277b417ae060e15010eb73bd1d00b01`
- Branch: `work/f-x278-deepseek`
- Plan: `.claude/plans/F-X278-design.md`, SHA-256 `cacd4844683b63f7c9fb2fbcde658d7b3b03fcdc36eb747768ac46fe9af02bae`
- Author handoff: `/private/tmp/rdocx-deepseek-harness/f278-implementation/AUTHOR_HANDOFF.md`, SHA-256 `fe6fa1a9582b311faa7a1396495e15c96b4fc08cb006f3e32facf8abbfea72c1`, 12468 bytes
- Frozen pins: 49 of 49 matched the current worktree
- Diff scope: 21 modified files, 2238 insertions, 296 deletions
- Worktree state: intentionally dirty and uncommitted. This review is not story acceptance and does not approve integration.

## Verdict

No product defect was confirmed in any of the seven A-IDs. The bundled-default policy, the explicit system opt-in, the feature-off refusal, caller priority, empty-input omission, the shared `Arc` cache, the TOC ink rule, the first-line and hanging merge, table-border precedence, row splitting, automatic colour and the Wingdings 2 map all match the approved contract in the code paths I read.

The durable problems are test-evidence gaps, one likely stale CLI test family, and documentation drift. They are listed below with citations. Zero-finding categories are stated explicitly.

## Per-A-ID coverage

| A-ID | Code path reviewed | Test reviewed | Result |
| --- | --- | --- | --- |
| A278-1 TOC ink | `crates/rdocx-layout/src/engine.rs:10988` `is_toc_paragraph_style`, `:11005` `is_hyperlink_character_style`, stored-segment filter `:12395` | `crates/rdocx-layout/src/engine.rs:26594` | Logic matches. Stored-segment and generated-TOC routes are not exercised (F6). |
| A278-2 indent merge | `crates/rdocx-layout/src/engine.rs:14185` `merge_direct_ppr`, `crates/rdocx-oxml/src/paragraph_properties.rs` merge change, `tbRl` arms `engine.rs:4803` and `:4806` | `crates/rdocx-layout/src/engine.rs:26768`, `:26795` tbRl test | Merge rule matches. `tbRl` test bypasses the engine arm (F5). The absolute 72.0 pt contract is not asserted (F9). |
| A278-3 table borders | `crates/rdocx-layout/src/table.rs:2165-2172` promotion removed, `resolve_base_table_properties`, paint fallback | `crates/rdocx-layout/src/engine.rs:26530` | Direct borders override style borders and all-`none` yields zero. Style-only exact 3/3 is not asserted (F10). |
| A278-4 row split | `crates/rdocx-layout/src/paginator.rs:1844` `v_align` bail removed | `crates/rdocx-layout/src/engine.rs:26720` | Splitting now ignores `w:vAlign`. The contract discriminators are missing from the test (F8). |
| A278-5 automatic colour | `crates/rdocx-layout/src/engine.rs:16607` `automatic_text_color`, `resolve_run_color` and named shading calls | `crates/rdocx-layout/src/engine.rs:26628` | White on dark, black on light, nested and body controls match. Provenance comment overstates LibreOffice (F12). HLD wording is narrower than the code (F15). |
| A278-6 Wingdings 2 | `crates/oxml-layout/src/font.rs:1988` `resolve_word_symbol_text`, `unknown_word_symbol_diagnostics` | `crates/rdocx-layout/src/engine.rs:26688` | Numbering marker maps `U+F0A1`, keeps `U+F0A7` and emits one diagnostic. `w:sym` consumer is not exercised (F7). HLD wording is narrower than the low-byte mapping (F16). |
| A278-7 font policy | `crates/rdocx/src/document.rs` policy helper and `RenderOptions`, `crates/rdocx-cli/src/commands.rs:868-935`, `crates/rdocx-py/src/document.rs` route table, `rdocx-wasm` `toPdf` | `crates/rdocx-cli/tests/integration.rs:7958` and `:8078`, `crates/rdocx-py/tests/test_core.py:877` | Policy logic matches. Public docs are stale (F1). One CLI test family likely fails (F2). Caller plus system and SVG routes are untested (F3, F4). |

Confirmed policy facts read directly from the tree:

- `RenderOptions::default().use_system_fonts` is false at `crates/rdocx/src/document.rs:64443`.
- Feature-off explicit system request returns `system fonts are not available in this build (the system-fonts feature is disabled)` before the normal engine, including for a nonempty caller list (`crates/rdocx-cli/src/commands.rs:874-880`).
- Empty caller input is omission and uses the bundled tier. Nonempty caller input keeps caller priority (`commands.rs:868-890`, `crates/rdocx-py/src/document.rs` route table).
- Caller results are wrapped in `Arc::new`. Accepted omission returns the shared bundled cache through `Arc::clone` (`crates/rdocx/src/document.rs:45680` region).
- `inspect` stays on the deterministic path and rejects `--system-fonts` through clap.
- `rdocx-wasm` `toPdf` follows `Document::to_pdf` into the deterministic bundled cache and has no font flag.

## Findings

### F1. Stale public Rust documentation for `layout()` and `layout_page()`

Severity: documentation defect, user-facing.

`crates/rdocx/src/document.rs:45691` documents `layout()` as returning "the cached normal-font layout". `crates/rdocx/src/document.rs:46418` documents `layout_page()` as returning a page "from the cached normal-font layout". Both methods now delegate to `layout_with_options(RenderOptions::default())`, and the derived default selects the bundled tier (`document.rs:64443`). A reader following the doc comment would expect system fonts by default, which is the opposite of the approved policy. Expected text is the bundled tier with the explicit `use_system_fonts` opt-in.

### F2. CLI image-export tests still require empty stderr under the new substitution diagnostics

Severity: test-suite defect, likely red.

The feature adds bundled substitution facts to layout diagnostics and prints them on the image and PDF routes at `crates/rdocx-cli/src/commands.rs:890-892` and `:981-983`. `crates/rdocx-cli/tests/integration.rs:3181`, `:3209` and `:3229` call `assert_success`, which requires `output.stderr.is_empty()` (`integration.rs:58-69`). The fixture is built with `fixture_document` (`integration.rs:138-148`) from `Document::new()`, whose document defaults request Calibri (`crates/rdocx-oxml/src/styles.rs:1780-1784`). The new policy test proves that a plain paragraph in that same fixture emits `Font 'Calibri' substituted: selected 'Carlito' from the bundled tier` (`integration.rs:8109-8113`). Therefore the three image-export assertions would fail on the current tree. The handoff says "existing empty-stderr assertions were replaced with explicit diagnostic-shape assertions" (`AUTHOR_HANDOFF.md:64`), but the diff replaces only the two Helvetica tests at `integration.rs:2812` and `:2840`. Expected behavior is either the diagnostic-shape helper or an exact expected fact set on these image tests.

### F3. CLI caller plus system route is untested with the feature on

Severity: coverage gap.

Plan rider C7b requires `convert`, `render` and `layout` with one caller font and the flag both off and on (`.claude/plans/F-X278-design.md:430`). `cli_font_policy_routes_bundled_omission_and_the_system_opt_in_per_command` covers caller priority with the flag off (`crates/rdocx-cli/tests/integration.rs:8243-8260`) and system selection with no caller list (`:8165-8225`). It never exercises the `font_refs` nonempty and `system_fonts` true branch at `crates/rdocx-cli/src/commands.rs:874-880`. Expected is one caller font with `--system-fonts` on each of the three commands.

### F4. Python caller plus system and SVG routes are untested

Severity: coverage gap.

`test_issue433_python_font_policy_keeps_one_tier_per_folder` covers omission, a caller-only route with `system_fonts=False`, and a system route with no caller list (`crates/rdocx-py/tests/test_core.py:877-955`). It does not cover `fonts=caller` with `system_fonts=True`, and it does not call any SVG sibling. The plan names `render_page_to_svg_with_fonts_and_options` in the Python route table and requires the diagnostic twins to share the public route (`.claude/plans/F-X278-design.md:302-311`, `:432`). Expected are the combined caller plus system call and the SVG sibling on both tiers.

### F5. `tbRl` test calls `merge_direct_ppr` directly and never enters the engine arm

Severity: coverage gap.

`issue433_tb_rl_paragraph_merge_clears_the_absent_sibling` builds `CT_PPr` values and calls `merge_direct_ppr` at `crates/rdocx-layout/src/engine.rs:26813` and `:26832`. The `tbRl` branch that the plan names is the `text_direction == Some("tbRl")` arm at `engine.rs:4802-4806`. The test never sets a section direction or runs `deterministic_layout`, so it cannot catch a regression in that arm. Expected is a paragraph in a `tbRl` section whose inherited hanging and later firstLine exercise `engine.rs:4803` and `:4806`.

### F6. TOC stored-segment filter and generated-TOC route are unexercised

Severity: coverage gap.

`issue433_toc_entries_drop_the_hyperlink_character_style` builds a hyperlink run with a `Hyperlink` run style in a TOC paragraph (`crates/rdocx-layout/src/engine.rs:26594-26626`). It exercises the run-style site, not the stored-segment filter at `engine.rs:12395-12399`, and it does not build a TOC through the `rdocx` API. Expected is a stored `w:rPr` with the Hyperlink character style inside a TOC paragraph, plus one generated TOC entry.

### F7. Wingdings 2 `w:sym` consumer is unexercised

Severity: coverage gap.

`issue433_wingdings_2_bullets_map_to_unicode_or_report_the_code` covers numbering markers only (`crates/rdocx-layout/src/engine.rs:26688-26718`). The `RunContent::Symbol` path at `engine.rs:12640-12664` calls `resolve_word_symbol_text` for `w:sym` runs and is not reached by the test. Expected is one `w:sym` run with `Wingdings 2` `F0A1` and one with `F0A7`, with the same substitute and diagnostic contract.

### F8. Row-split test lacks the contract discriminators

Severity: coverage gap.

Plan A278-4 expects the row to start on page 1 and stay within one line of `[23, 17]` (`.claude/plans/F-X278-design.md:414`). The test asserts only that there are two pages with at least one `Row` line on each, and that the three `vAlign` values equal the plain case (`crates/rdocx-layout/src/engine.rs:26753-26763`). A regression that pushes the row to page 2, or splits it at a different line, would still pass. Expected are the page-1 start and the split-point assertion.

### F9. Indent test uses relative equality and 720 twips rather than the absolute contract

Severity: coverage gap.

Plan A278-2 expects four issue cases at 72.0 pt with 0.05 pt tolerance (`.claude/plans/F-X278-design.md:392`). `issue433_first_line_indent_is_one_property_across_levels` compares each `origin.x` to `starts[0].1` within `1e-6` (`crates/rdocx-layout/src/engine.rs:26786-26790`) and builds the styles with 720-twip values (`:26773`). A uniform shift of all five paragraphs would pass. Expected is the absolute 72.0 pt position with the stated tolerance.

### F10. Style-only table borders assert `>= 3` instead of exactly 3/3

Severity: coverage gap.

Plan A278-3 accepts exactly 3 horizontal and 3 vertical ruled lines for style-only TableGrid (`.claude/plans/F-X278-design.md:88`, `:413`). `issue433_direct_table_borders_override_the_style_grid` asserts `lines[0].0 >= 3 && lines[0].1 >= 3` (`crates/rdocx-layout/src/engine.rs:26585`). A regression that duplicates an edge and paints 4 or more would pass. The zero case and the partial ordering are asserted correctly. Expected is the exact pair `(3, 3)`.

### F11. Nested-table diagnostics allow any substitution fact

Severity: coverage gap, contract-justified but weak.

`independent_nested_tables_measure_to_one_final_height` changed from requiring empty diagnostics to accepting any message containing `substituted:` (`crates/rdocx/tests/integration_test.rs:9071-9078`). The contract says substitution facts are intended output, so the relaxation is justified, but it does not pin the requested family, the selected family, the tier or the count. A wrong-tier substitution would pass. Expected is the exact expected fact set for the fixture.

### F12. Automatic-colour comment claims a LibreOffice provenance the plan disclaims

Severity: documentation defect.

`crates/rdocx-layout/src/engine.rs:16605-16607` says "The threshold is LibreOffice's, which turns white up to a luminance of 156 in 255, weighted 76, 151 and 29 in 256." The approved plan says the coefficients and floor are the donor's chosen policy and "They are not a pinned LibreOffice source file" (`.claude/plans/F-X278-design.md:100`). Expected is wording that names the donor policy without asserting a pinned LibreOffice source.

### F13. TOC style matcher trims only leading whitespace

Severity: minor robustness.

`is_toc_paragraph_style` parses `name[3..].trim_start()` at `crates/rdocx-layout/src/engine.rs:10989-10993`. A style name such as `toc 1 ` with a trailing space fails the numeric parse, although Word tolerates surrounding whitespace. Expected is `trim()` on the suffix, or the same normalisation used by the existing style-name lookup.

### F14. HLD08 says only `w:cantSplit` moves the whole row

Severity: documentation defect, internal contradiction.

`docs/hld/08-rendering-spec.md:1162` and `:1175` say "Only `w:cantSplit` moves the whole row". The same section says merged, rotated, anchored, nested-table and exact-height rows keep whole-row placement (`docs/hld/08-rendering-spec.md:1164-1167`). The plan keeps `cant_split`, header rows, merges, clipping, rotation, anchored content, keep-next and page-break-before as whole-row barriers (`.claude/plans/F-X278-design.md:94`). Expected is wording that names the full barrier set and reserves "only" for the `vAlign` case.

### F15. HLD08 automatic-colour wording covers only explicit `auto`

Severity: documentation defect.

`docs/hld/08-rendering-spec.md:905` says a run in `w:color/@w:val="auto"` takes a contrast colour. The implementation also applies automatic colour when no `w:color` is present (`crates/rdocx-layout/src/engine.rs:16600-16604` falls back to `automatic`). Expected is wording that covers both an explicit `auto` and an absent colour.

### F16. HLD08 Wingdings wording names only `U+F0A1`

Severity: documentation defect.

`docs/hld/08-rendering-spec.md:1565` says qualified numbering bullets in `Wingdings 2` map `U+F0A1` to the black square. The implementation also maps the low byte `0xA1` in the same bounded table (`crates/oxml-layout/src/font.rs`, `resolve_word_symbol_text`). Expected is wording that names both accepted source forms without widening the mapping claim.

## Zero-finding categories

- **Product defects in the seven A-IDs.** Zero confirmed. Each A-ID matched its contract on the paths I read.
- **Panics, aborts and error-path regressions.** Zero confirmed. The feature-off refusal returns a typed error before layout, and the Python mapping goes through `rdocx_to_pyerr`.
- **Raw XML, namespace and schema-order preservation.** Zero confirmed. The changed model code preserves unmodelled XML and the diff does not reorder `w:ind`, `w:tblBorders` or symbol children.
- **Structural rules.** Zero confirmed. No new trait, generic parameter, crate, module or test binary was added. The change reuses existing methods and adds three documented `#[doc(hidden)]` `FontManager` methods.
- **PowerPoint parity.** Zero confirmed. The PowerPoint call sites are unchanged at `crates/rpptx-render/src/text.rs:515`, `:912`, `:3729`, `:3752`, `:3710` and `crates/rpptx-render/src/lib.rs:367`. PowerPoint symbol and diagnostic output is untouched.
- **HLD impact completeness.** Zero confirmed. All three named files changed: `docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, `docs/hld/15-build-and-toolchain.md`.

## Provenance authentication

- The two CLI green reruns are authenticated by Cargo output. `logs/20261102-cli-on.log` shows 3 passed, 0 failed. `logs/20261102-cli-off-attempt2.log` and `logs/20261102-cli-off-attempt3-emptyfix.log` each show 1 passed, 0 failed.
- The cache log `logs/20261102-cache-nine.log` shows 9 passed, 0 failed.
- The original CLI RED run is not preserved. Only the transcript excerpt in `AUTHOR_HANDOFF.md:102` records the two empty-stderr failures. That is historical RED evidence, not current green evidence.
- The cli-on deciding wrapper ended in `tail` without `exit "$rc"` per the preserved transcript, so child-status propagation for that run is not authenticated by the log even though `CHILD_STATUS=0` was recorded. The attempt-3 wrapper did propagate `exit "$rc"`.
- The bulk suites were not rerun. The handoff records a preserved integration baseline of 462 passed, 10 failed, 8 ignored and a library baseline of 391 passed, 17 failed, 1 ignored (`AUTHOR_HANDOFF.md:108-110`). Those are not green results and do not cover F2.
- No helper-only assertion was accepted as all-consumer proof. F5 through F11 exist because the named tests do not reach the consumer paths the plan names.

## Pending gates

- Rerun the CLI integration and library bulk suites on the current tree, including the image-export tests named in F2.
- Re-record the 14 hash keys. No hash baseline was moved in this worktree.
- Native Word, LibreOffice `26.2.5.2` at `cd7284b4cbbfeb507e630c1aac019f4157393acb` and pdftoppm `26.01.0` final union observations.
- Installed Python 3.9 and 3.12 runtime, typing and stub checks.
- Wheels, archives, WASM build and hosted CI.
- `cargo publish --dry-run` and the size assertion.
- Final union and oracle render checks for TOC ink, automatic colour, Wingdings square and row split.

## Disposition

Pass 1 of at most 3. The implementation is product-correct on the paths I read. The blocking work before acceptance is F2, which is likely a red test family, plus the coverage gaps F3 through F11. The documentation findings F1, F12, F14, F15 and F16 are small and should be corrected in the same feature. F13 is minor. This review changes no code and does not accept the story.
