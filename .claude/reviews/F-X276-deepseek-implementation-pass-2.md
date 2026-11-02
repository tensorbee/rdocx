# F-X276, deepseek implementation, pass 2

**Reviewed**: working tree of `work/f-x276-deepseek` against Base
`9a3b2960ccda7c3208ecb9b753c7019b64c086f5`, 12 changed files plus the untracked
pass 1 review. HEAD is the Base, no commit. Approved plan
`.claude/plans/F-X276-design.md` SHA-256
`7ac52101673baac1ec65e485478488d1cf7a636a3e237fc83b6abde84e350382`.

**Verdict**: 1 defect, 5 smells, 0 nitpicks. Not a clean whole-story gate. A1,
A2, A3, A4 and A5 each have local coverage, but A2 and A5 retain open contract
or evidence gaps, and the Markdown repair evidence is partly stale.

## Source and evidence authentication

All 15 changed or untracked file SHA-256 values in
`/private/tmp/rdocx-cross-review/F276_REPAIR1_FROZEN_PINS.json` match the
reviewed worktree. The plan, the pass 1 review
(`02f0143f9635bfb1c51e37fe836d1c75522a49a5c9d69baf84b8cb2d16b262fc`), the
progress file (`e646c8c1bbf9789f36e188d014c8841a5c0f5fe0110cb5ee11ebf0061647d08c`)
and the handoff (`4b6421df601c6b95ec10da3c185733a5b76432455c0643da617feb3d4daa1fe6`,
14262 bytes) match. The frozen log copies under
`/private/tmp/rdocx-cross-review/f276-author-final-frozen/logs/` match
`LOG_PATH_MAP.json` (`e2f5dd26a859715f72c46d10a4b0a5a122a3ea7c3ae9d4aaf899d0d7f4cb11e1`).
The recovered abort log
`/private/tmp/rdocx-cross-review/F276_ISSUE431_CLI_ABORT_RECOVERED.log` has SHA-256
`ddb5074ab6bbead3fd684da5b095c09bc74340bb3adcab8f65d623842cdf58cd`. The later
green log `/private/tmp/rdocx-cross-review/F276_ISSUE431_CLI_GREEN_20261102T0244.log`
has SHA-256 `b45b5d2e3e4b605a9fb908b2b21b51e42857ea544668466d39b6841a15f6ab61`.
The Root lead file `F276_ROOT_REVIEW_LEADS.json` has SHA-256
`6395617d7c9059dd8e05f729dfd21a08be8811470b01a69141860125ab7ec11b`.

The handoff is an author claim, not acceptance. The source pins authenticate the
reviewed bytes. They do not authenticate every claim in the handoff.

## Actual inspected scope

Focused source and test regions were read, not whole giant files:

- `crates/rpptx/src/lib.rs` lines 300 to 340, 1100 to 1160, 3935 to 4334,
  27181 to 27250, 27600 to 28430, 36290 to 36945.
- `crates/rpptx-render/src/lib.rs` lines 611 to 1185, tests 2739 to 2825.
- `crates/rpptx/src/html.rs` lines 180 to 280, 350 to 720, 1160 to 1300, test
  3554 to 3571.
- `crates/rpptx-cli/src/commands.rs` lines 666 to 1900.
- `crates/rpptx-cli/tests/integration.rs` diff, especially 428, 2993 to 3279.
- `crates/rpptx/tests/integration.rs` diff, especially 94, 170, 339, 377, 484,
  888 and the new issue431 tests.
- `crates/rpptx-py/src/slide.rs` lines 438 to 480, 1532 to 1545.
- `crates/rpptx-py/src/presentation.rs` line 197.
- `crates/rpptx-py/python/rpptx/_rpptx.pyi` line 634.
- `crates/rpptx-py/tests/test_documented_examples.py` diff and 11155 to 11180,
  12515 to 12550.
- `docs/hld/08-rendering-spec.md` line 3095 and
  `docs/hld/10-bindings-spec.md` lines 902, 1228, 4080 and 4680.
- The plan, handoff, pass 1 review, progress, frozen pins, Root leads, log map,
  every frozen repair log, the recovered abort log and the green log.
- Author tool events were extracted in memory from
  `/private/tmp/rdocx-deepseek-harness/f276-implementation/author-repair1.stdout.jsonl`
  for command, child exit and semantic result only. Reasoning content was not
  read.

Not inspected: whole giant source files, unrelated crates, the complete Python
suite, the complete CLI suite, the complete rpptx suite, or the full JSONL
history.

## Defects

### P2-D1, A2 width-failure published-trial clause is unmet
`.claude/plans/F-X276-design.md:204`, `crates/rpptx/src/lib.rs:4195`,
`crates/rpptx/src/lib.rs:4214`, `crates/rpptx/src/lib.rs:36669`

The plan says a width failure keeps `overflow` true and "A published trial has
`width_fits` false." The implementation's width-failure fixture uses no-wrap
text in a table cell whose stored width does not fit
(`crates/rpptx/src/lib.rs:36646`). `fit_report` produces `overflow` true, which
is correct, but it also produces `needed_scale=None`, `quantized_trial=None` and
`persisted_trial=None`. The committed test asserts that absence at
`crates/rpptx/src/lib.rs:36669` to `36673`.

The source reason is that `select_autofit_pair` accepts only candidates where
`q.fits && s.fits` (`crates/rpptx/src/lib.rs:27181` to `27203`) and `fits`
includes `width_fits` (`crates/rpptx-render/src/text.rs:680` to `681`). A width
failure therefore has no admissible pair.

The author's rebuttal is technically correct about the current pair admission,
but it narrows an approved A2 clause without a Root plan amendment. Expected per
the approved plan is a published trial with `width_fits` false. Actual is no
trial. The observable overflow result is correct, so severity is medium. This is
a contract and source mismatch for Root adjudication, not a repair performed by
this review.

## Smells

### P2-S1, S4 absent-owner `IndexError` clause remains unproven
`.claude/plans/F-X276-design.md:228`, `crates/rpptx-py/src/slide.rs:438`,
`crates/rpptx-py/tests/test_documented_examples.py:12515`

A5 says "Other methods still raise `IndexError` for an absent owner." The deleted
pass 1 assertion expected `held.slide_id` after `remove` to raise `IndexError`.
The actual behavior is causal stale-handle behavior, `StaleElementError`, pinned
at `crates/rpptx-py/tests/test_documented_examples.py:11170` to `11172`. The
rebuttal correctly separates the removed-handle case from an absent-owner case.
`PySlide::validate` still has the defensive `IndexError` branch at
`crates/rpptx-py/src/slide.rs:438` to `455`, and `SlideCollection.index` maps it
to `ValueError` at `crates/rpptx-py/src/slide.rs:1532`.

What remains open is the A5 assertion itself. No committed test exercises an
absent-owner path for another method, and the rebuttal does not demonstrate a
public path to `current_path` returning `None` after `validate_path` succeeds.
Expected is either a test for that path or a plan amendment. Actual is no test
and an asserted-unreachable defensive branch. Severity is low. Do not conflate
causal stale-handle behavior with absent-owner `IndexError`.

### P2-S2, paint-delta control is self-referential and the named existing fixture was not run
`.claude/plans/F-X276-design.md:61`, `.claude/plans/F-X276-design.md:181`,
`crates/rpptx-render/src/lib.rs:2790`, `crates/rpptx-render/src/lib.rs:2824`

The plan says paint equality means the extracted deltas stay the painter's
deltas, and an existing painted-table fixture must stay unchanged. The repair
adds `painted_table_fills_use_the_grown_row_offsets`
(`crates/rpptx-render/src/lib.rs:2790`), which compares painter fill offsets and
heights against `grown_row_heights`. That is a meaningful self-consistency check
for offsets and heights. It is not a before-and-after check against the old
`lower_table` output, and it is not the existing
`banded_merged_table_renders_correct_fills_without_duplicated_borders` fixture at
`crates/rpptx-render/src/lib.rs:2824`. The repair logs show only the new paint
test (`f276-repair1-render-paint-attempt1.log:8`), not the existing fixture.

The source diff at `crates/rpptx-render/src/lib.rs:1136` to `1185` moves the
growth loop faithfully and `lower_table` still computes row offsets from the
returned heights. No pixel-moving defect was found by inspection. The named
existing-fixture control remains unperformed. Severity is low.

### P2-S3, span test does not assert the exact `needed - height` delta
`.claude/plans/F-X276-design.md:181`, `crates/rpptx-render/src/lib.rs:2739`,
`crates/rpptx-render/src/lib.rs:1177`

The plan says a span grows its last row by `needed - height` only. The code does
`required_heights[row_index + row_span - 1] += needed - height` at
`crates/rpptx-render/src/lib.rs:1177`. The new test
`grown_row_heights_grow_the_last_span_row_and_short_neighbors_add_nothing` at
`crates/rpptx-render/src/lib.rs:2739` asserts only that `heights[1] > 6.0` at
`crates/rpptx-render/src/lib.rs:2771`. A wrong implementation that adds
`needed` rather than `needed - height` would still pass that assertion. Expected
is an exact delta assertion or a fixture where the two values differ. Actual is
a qualitative greater-than check. The source is correct, so severity is low.

### P2-S4, the handoff's universal wrapper-exit claim is false for the initial batch
`/private/tmp/rdocx-deepseek-harness/f276-implementation/AUTHOR_HANDOFF.md:140`,
`/private/tmp/rdocx-cross-review/F276_ISSUE431_CLI_ABORT_RECOVERED.log:8`,
`/private/tmp/rdocx-deepseek-harness/f276-implementation/author-repair1.stdout.jsonl:162`

The handoff says each command captured Cargo's child status immediately and the
wrapper returned that status. That is not true chronologically for the initial
repair1 Issue431 CLI batch. The tool event at JSONL line 162 printed
`cli issue431 exit=101` while the wrapper metadata exit was 0 because the
command ended with a `grep` and did not `exit "$rc"`. The recovered abort log
lines 8 to 14 show the test-thread stack overflow and SIGABRT. The author later
read the complete log and reran with the already-established
`RUST_MIN_STACK=8388608` at JSONL line 172. That rerun printed child exit 0 and
the green log line 8 records 2 passed. No new Root stack-setting cue was
supplied.

The abort is execution failure, not semantic behavior RED. The green rerun is
authenticated separately. The handoff's universal wording is overstated, so
severity is low. It does not invalidate the green rerun.

### P2-S5, Markdown and integration GREEN evidence predates the repair
`/private/tmp/rdocx-deepseek-harness/f276-implementation/AUTHOR_HANDOFF.md:206`,
`crates/rpptx/src/html.rs:1193`, `crates/rpptx/tests/integration.rs:484`

The handoff retains `rpptx integration issue431` 4 passed and `fx218_` 21 passed
as earlier evidence. Those runs predate the repair cycle that changed
`crates/rpptx/src/html.rs` and `crates/rpptx/src/lib.rs`. The repair logs contain
the new `markdown_long_plain_line_keeps_its_characters_and_syntax_boundaries`
test (`f276-repair1-html-long.log:11`) and the new lib issue431 tests
(`f276-repair1-lib-issue431-attempt5.log:14`), but they do not contain a
post-repair `rpptx --test integration issue431` run or a post-repair `fx218_`
run. The `fx218_` corpus covers cross-run spaces, notes, alt, label, breaks and
punctuation. The long-line test does not.

Source inspection of the repair shows the line-end scan is lazy and the
space/escape logic is unchanged, so no weakening was found. The evidence for
those behaviours is still stale. Expected is a post-repair `fx218_` and
integration rerun. Actual is a pre-repair GREEN. Severity is low to medium.

## Pass 1 disposition

| Pass 1 finding | Disposition | Evidence |
|---|---|---|
| D1, Python `ValidationIssue.message` used `Debug` | Fixed | `crates/rpptx-py/src/presentation.rs:197` uses `issue.to_string()`. The Python test pins the Display sentence at `crates/rpptx-py/tests/test_documented_examples.py:3583`. `f276-repair1-py-tests.log:2` records 2 passed. |
| D2, CLI validate test asserted the variant name | Fixed | `crates/rpptx-cli/src/commands.rs:669` prints `{issue}`. The test at `crates/rpptx-cli/tests/integration.rs:431` asserts `InvalidXmlChild` is absent and the human message is present. `f276-repair1-cli-validate.log:12` records 1 passed. |
| D3, A2 scale assertion was weak | Fixed | `crates/rpptx/src/lib.rs:36456` asserts `Some(40000)`. The off-slide branch asserts `is_some()` at `crates/rpptx/src/lib.rs:36454`. `f276-repair1-lib-issue431-attempt5.log:14` records 6 passed. |
| D4, A2 geometry and span controls absent | Fixed, with P2-S3 | The geometry controls are at `crates/rpptx/src/lib.rs:36470` to `36720`. The span and paint tests are at `crates/rpptx-render/src/lib.rs:2739` and `2790`. The exact span delta is not asserted. |
| S1, empty frame untested | Fixed | `crates/rpptx/src/lib.rs:36678` asserts status empty, overflow false and no cannot-assess cause. The lib issue431 log records 6 passed. |
| S2, notes export and CLI exit 2 untested | Fixed | `crates/rpptx/src/lib.rs:36941` asserts `to_notes_pdf_deterministic().is_err()`. `crates/rpptx-cli/tests/integration.rs:3229` asserts exit 2. The CLI fit log records 5 passed. |
| S3, CLI picture and contextless routes untested | Fixed | `crates/rpptx-cli/tests/integration.rs:3111` covers the picture `Diagnostic:` line. `crates/rpptx-cli/tests/integration.rs:3187` covers the contextless SmartArt cause. The CLI fit log records 5 passed. |
| S4, absent-owner `IndexError` assertion deleted | Rebutted for the removed handle, A5 clause still open | The removed handle is causal stale behavior, `StaleElementError`, at `crates/rpptx-py/tests/test_documented_examples.py:11170`. The A5 absent-owner clause has no test. See P2-S1. |
| S5, paint delta unauthenticated | Partially repaired | The new paint test is at `crates/rpptx-render/src/lib.rs:2790`. The existing painted-table fixture at `crates/rpptx-render/src/lib.rs:2824` was not run. See P2-S2. |

## A-ID coverage and pending checks

| A-ID | Local coverage inspected | Disposition | Exact pending checks |
|---|---|---|---|
| A1 | Display sentences at `crates/rpptx/src/lib.rs:1102`. Duplicate leaf behavior, picture diagnostic, empty frame, notes reasons, contextful and contextless causes, Python message and validate sentence are covered by `crates/rpptx/src/lib.rs:36296` to `36945`, `crates/rpptx-cli/tests/integration.rs:2993` to `3279`, and `crates/rpptx-py/tests/test_documented_examples.py:3563`. Lib issue431 6 passed, CLI fit 5 passed, Python focused 2 passed. | Covered locally, no blocking source defect found. The duplicate test asserts both frames exist but does not assert both text frames' `measured_font_scale`. | Post-repair `rpptx --test integration issue431` rerun. Full integrated union. |
| A2 | On-slide `Some(40000)` and off-slide overflow at `crates/rpptx/src/lib.rs:36421`. Geometry controls at `crates/rpptx/src/lib.rs:36470`. Width failure at `crates/rpptx/src/lib.rs:36646`. Span and paint tests at `crates/rpptx-render/src/lib.rs:2739` and `2790`. Lib issue431 6 passed, render span and paint 1 passed each. | Partial. P2-D1 is an open A2 contract mismatch. P2-S2 and P2-S3 are evidence gaps. | Existing painted-table fixture run. Exact span delta assertion. Root adjudication of the published-trial clause. Hash 49 and pixels 7. |
| A3 | CLI line, exit rule and JSON reason at `crates/rpptx-cli/src/commands.rs:1811` to `1900`. CLI tests at `crates/rpptx-cli/tests/integration.rs:2998`, `3111`, `3187` and `3229`. CLI fit 5 passed. | Covered locally, no blocking defect found. | Full integrated union. |
| A4 | Markdown escape state at `crates/rpptx/src/html.rs:1185` to `1281`. Long-line test at `crates/rpptx/src/html.rs:3554`. A4 integration tests at `crates/rpptx/tests/integration.rs:484` and the new issue431 Markdown tests. `f276-repair1-html-long.log:11` records 1 passed. | Partial evidence. P2-S5 is the stale `fx218_` and integration GREEN. Source inspection found no weakening of paragraph assembly, cross-run spaces, hard breaks, notes, punctuation or alt. | Post-repair `fx218_` 21 and `rpptx integration issue431` rerun. Full export benchmark. HTML unchanged check. |
| A5 | `SlideCollection.index` at `crates/rpptx-py/src/slide.rs:1532`, stub at `crates/rpptx-py/python/rpptx/_rpptx.pyi:634`, Python test at `crates/rpptx-py/tests/test_documented_examples.py:12515`. Python focused 2 passed. | Partial. Live index, two wrappers, move, foreign, twin, reopen, empty, non-slide and removed handle are covered. P2-S1 is the open absent-owner clause. | Plan amendment or an absent-owner test. Installed Python 3.9 and 3.12 wheels, typing and stubs. |

## Evidence authentication, separated

**Build and artifact identity**: The repair logs name
`/private/tmp/rdocx-s96-build-f276` and the same CLI integration binary
`integration-7770667fa5bf89cb` for the abort and green runs. The source pins
authenticate the reviewed files. The log pins authenticate the log copies.

**Compiler and child exits**: Later repair commands printed child exit lines and
the author tool metadata exit 0. The initial Issue431 CLI batch at JSONL line 162
printed child exit 101 while the wrapper metadata exit was 0, because the wrapper
did not return the child status. The green rerun at JSONL line 172 printed child
exit 0 and metadata exit 0, with 2 passed at green log line 8. The abort log
lines 8 to 14 record SIGABRT. A tail or grep pipeline is not treated as Cargo
exit 0.

**Semantic results**: The semantic results are read from the `test result` lines
in the frozen logs, separately from exit status. Lib issue431 6 passed
(`f276-repair1-lib-issue431-attempt5.log:14`), CLI fit 5 passed
(`f276-repair1-cli-fit-attempt2.log:13`), render span 1 passed
(`f276-repair1-render-span-attempt1.log:13`), render paint 1 passed
(`f276-repair1-render-paint-attempt1.log:8`), Python focused 2 passed
(`f276-repair1-py-tests.log:2`), CLI validate 1 passed
(`f276-repair1-cli-validate.log:12`), and the Markdown long-line test 1 passed
(`f276-repair1-html-long.log:11`). The abort is execution failure, not semantic
behavior RED. The green rerun does not establish every D and S finding, and this
review does not infer that it does.

**Stale evidence**: The `rpptx integration issue431` 4 passed and `fx218_` 21
passed runs named at handoff line 206 predate the repair. They are not current
GREEN for the changed `html.rs` and `lib.rs`. The repair logs do not contain a
post-repair run of those carriers. See P2-S5.

## Markdown performance lead

The Root lead is at `/private/tmp/rdocx-cross-review/F276_ROOT_REVIEW_LEADS.json`.
It is Root discovery, not peer 1 independent discovery. The pre-repair helper SHA
in the lead is `11f481a98cd0b74c1186b549f40511d4f98f0210995a5e0f0712779dca1814bb`.
The current `crates/rpptx/src/html.rs` SHA is
`0f9903554a86b0cc50dfd09c99819c81f494c76ebf68c1c9e5d2151e5a1f94bf`.

Current `semantic_md_chars` at `crates/rpptx/src/html.rs:1202` calls
`md_line_end` at `crates/rpptx/src/html.rs:1193` only in the `-` and `=` setext
branches. The `-` and `=` scans run only when the character is at the line
content start, so the line-end scan is once per line rather than once per
ordinary character. The source is bounded by inspection. The long ordinary line
test at `crates/rpptx/src/html.rs:3554` passed. The Root probe was isolated
helper timing on the pre-repair helper, not a full export benchmark. No
wall-clock CI assertion was offered. The full export benchmark remains final
debt. No weakening of paragraph assembly, cross-run spaces, hard breaks, notes,
punctuation or image-alt behaviour was found in the source. The A4 integration
evidence for those behaviours is stale, as recorded in P2-S5.

## Unperformed rider debt

These are not claimed as local passes and are not used to excuse the local gaps
above:

- Full integrated workspace union.
- Hash harness, expected hash 49 and pixels 7.
- Native, PDF and PNG rendering, and LibreOffice.
- Installed Python 3.9 and 3.12 wheels, typing, stubs and archive.
- CI.
- Full Markdown export benchmark.
- The existing painted-table fixture run.
- Post-repair `fx218_` and `rpptx integration issue431` reruns.

## Not found

No source defect was found in the approved Display sentences, the
`authoring_table_height` extraction, the duplicate-leaf publication shape, the
notes reason strings, the two unsupported publishers, the width-qualified page
test for `overflow`, the CLI line, exit and JSON `reason`, the Markdown escape
state, or the `SlideCollection.index` live path. Those aspects produced no
additional defect beyond the evidence and contract gaps recorded above. The
review is not a clean whole-story gate.
