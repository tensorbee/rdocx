# F-X276, deepseek implementation, pass 3

**Reviewed**: working tree of `work/f-x276-deepseek` against Base
`9a3b2960ccda7c3208ecb9b753c7019b64c086f5`. HEAD is the Base, no commit. 13
changed files plus the two untracked earlier implementation reviews. Approved
plan `.claude/plans/F-X276-design.md` SHA-256
`cd69e353b511bf9eadeb57ffe157b1cb27b660d0870a6054fb7986e1c7f38847`.

**Verdict**: 0 defects, 1 smell, 0 nitpicks. No blocking source defect was
found. One contract ambiguity remains around an empty notes placeholder. The
five A-IDs have local coverage. The final union riders are not claimed.

## Source and evidence authentication

All 99 entries in `/private/tmp/rdocx-cross-review/F276_REPAIR2_FINAL_FROZEN_PINS.json`
match the current worker bytes and the named evidence. That includes the 13
changed source, test, HLD and plan files, the progress file
(`5fa186ff528d1341bc8b17435faa47990c5165d09a1b9d48337b7ff2ee9545eb`), the two
earlier implementation reviews (`02f0143f9635bfb1c51e37fe836d1c75522a49a5c9d69baf84b8cb2d16b262fc`
and `94f85a587d04dc6de36ecba613a810354d4acf1433d355afa156084e43910eda`), the
current handoff (`b976b717ef6cb52fd3627c091bd2e64e7faeb2644924e842fe320abda2ae9ccf`),
the adjudication report (`6e8fce22bf08250faf9f006d383f8f5c07c681e19e62e15681587d1dc72acae3`),
the frozen logs, the Root lead file and the local Python extension
(`b92c0389c29b6f825504ef5374e82bef50448499e6849068b84c4fd22a2411a5`).

Source identity is separate from artifact identity. The reviewed source is the
current worktree. The Python artifact is the authenticated extension under
`/private/tmp/rdocx-s96-py-f276/pkg`. The extension was built at repair1 and
only `cfg(test)`, pytest, documentation and plan changes followed it. The
repair1 frozen pins and the peer2 frozen source both equal the current
production files for `presentation.rs`, `slide.rs`, `html.rs` and the production
regions of `rpptx/src/lib.rs` and `rpptx-render/src/lib.rs`. The diffs against
`/private/tmp/rdocx-cross-review/f276-peer2-final-frozen/worker-relative` are
confined to test modules. Reuse is therefore justified by unchanged production
inputs, not by the handoff statement alone.

Child exit, wrapper exit and semantic result are recorded separately.

- p2-a1 measured leaves: child 0, wrapper 0, `1 passed, 211 filtered out`.
- p2-d1 admitted pair: child 0, wrapper 0, `1 passed, 211 filtered out`.
- p2-d1 width fixture: child 0, wrapper 0, `1 passed, 211 filtered out`.
- p2-s1 Python index: child 0, wrapper 0, `1 passed, 1033 deselected`.
- p2-s2 named paint fixture: child 0, wrapper 0, `1 passed, 134 filtered out`.
- p2-s3 span: child 0, wrapper 0, `1 passed, 134 filtered out`.
- p2-s5 fx218: child 0, wrapper 0, `21 passed, 415 filtered out`.
- p2-s5 issue431 integration: child 0, wrapper 0, `4 passed, 432 filtered out`.
- p2-check-rpptx: child 0, wrapper 0.
- p2-clippy-rpptx: child 0, wrapper 0.
- p2-fmt-final attempt 1: child 1, wrapper 1, one rustfmt reflow.
- p2-fmt-final attempt 2: child 0, wrapper 0.
- prose_check after the HLD edit: 0 violations.

The counts above are per execution. They are not summed into a claim of unique
coverage. The earlier repair1 runs remain separately authenticated: lib
issue431 `6 passed, 206 filtered out`, CLI fit `5 passed, 60 filtered out`, CLI
validate `1 passed, 61 filtered out`, Python focused `2 passed, 1032 deselected`,
and the Markdown long-line test `1 passed`. The CLI fit and validate production
and test files are unchanged since repair1, so those runs stay applicable. The
lib issue431 run predates the final measured-leaves assertion, which has its own
current `1 passed` run.

The initial CLI abort is an execution failure, not semantic RED. The recovered
log records a test-thread stack overflow and SIGABRT. The tool event printed
child exit 101 while the wrapper metadata exit was 0 because the command ended
in a grep without `exit "$rc"`. The later rerun with `RUST_MIN_STACK=8388608`
printed child exit 0 and wrapper exit 0, and the green log records `2 passed`.
No historical wrapper is claimed to have propagated every child status.

## Inspected scope

Coherent bounded reads of the changed functions and their consumers:

- `crates/rpptx/src/lib.rs`: 307 to 330, 1097 to 1160, 3947 to 4312, 27603 to
  28124, 28296 to 28390, 36292 to 36945.
- `crates/rpptx-render/src/lib.rs`: 608 to 620, 1131 to 1210, tests 2735 to
  2850.
- `crates/rpptx/src/html.rs`: 205 to 265, 369 to 500, 500 to 605, 688 to 700,
  1176 to 1305, tests 3550 to 3575.
- `crates/rpptx-cli/src/commands.rs`: 666 to 672, 1811 to 1900.
- `crates/rpptx-cli/tests/integration.rs`: 429 to 432, 2994 to 3279.
- `crates/rpptx/tests/integration.rs`: 95, 171 to 178, 340 to 350, 378 to 390,
  485 to 700, 1086 to 1092.
- `crates/rpptx-py/src/presentation.rs`: 194 to 200.
- `crates/rpptx-py/src/slide.rs`: 438 to 465, 1531 to 1542.
- `crates/rpptx-py/python/rpptx/_rpptx.pyi`: 631 to 635.
- `crates/rpptx-py/tests/test_documented_examples.py`: 3579 to 3585, 12510 to
  12559.
- `docs/hld/08-rendering-spec.md`: 3092 to 3110.
- `docs/hld/10-bindings-spec.md`: 899 to 912, 1231 to 1238, 4083 to 4095, 4680
  to 4695.
- The plan, progress, current handoff, both earlier implementation reviews, the
  canonical adjudication report, the frozen pins, the named logs, the log path
  map, the Root lead file and the relevant tool metadata.

Not inspected: whole giant source files, unrelated crates, the complete Python,
CLI or rpptx suites, and the full JSONL history.

## Prior finding dispositions

### Pass 1

| Finding | Disposition | Evidence |
|---|---|---|
| D1, Python message used `Debug` | Fixed | `crates/rpptx-py/src/presentation.rs:197` uses `issue.to_string()`. The Python test pins the Display sentence at `crates/rpptx-py/tests/test_documented_examples.py:3583`. |
| D2, CLI validate test asserted the variant | Fixed | `crates/rpptx-cli/src/commands.rs:669` prints `{issue}`. `crates/rpptx-cli/tests/integration.rs:432` asserts `InvalidXmlChild` is absent and the human message is present. |
| D3, A2 scale assertion weak | Fixed | `crates/rpptx/src/lib.rs:36459` asserts `Some(40000)` for the on-slide case. The off-slide branch asserts a scale is present. |
| D4, A2 geometry and span controls absent | Fixed | Geometry controls at `crates/rpptx/src/lib.rs:36470` to `36720`, span at `crates/rpptx-render/src/lib.rs:2739`, paint at `:2790`. |
| S1, empty frame untested | Fixed | `crates/rpptx/src/lib.rs:36678` asserts status `empty`, overflow false and no cannot-assess cause. |
| S2, notes export and CLI exit 2 untested | Fixed | `crates/rpptx/src/lib.rs:36941` asserts `to_notes_pdf_deterministic().is_err()`. `crates/rpptx-cli/tests/integration.rs:3229` asserts exit 2. |
| S3, CLI picture and contextless routes untested | Fixed | `crates/rpptx-cli/tests/integration.rs:3111` covers the picture `Diagnostic:` line. `:3187` covers the contextless SmartArt cause. |
| S4, absent-owner `IndexError` assertion deleted | Resolved by adjudication and current tests | `crates/rpptx-py/src/slide.rs:438` to `455` keeps the defensive branch. `crates/rpptx-py/tests/test_documented_examples.py:12545` to `12558` asserts removed-handle `StaleElementError`, numeric `IndexError` and `index` `ValueError`. The branch is not claimed publicly reachable. |
| S5, paint delta unauthenticated | Fixed | The named existing fixture `banded_merged_table_renders_correct_fills_without_duplicated_borders` passed in `p2-s2-existing-painted-table-attempt1.log`. The self-consistency test at `crates/rpptx-render/src/lib.rs:2790` is not used as the before and after proof. |

### Pass 2

| Finding | Disposition | Evidence |
|---|---|---|
| P2-D1, A2 width-failure published-trial clause | Resolved by the approved plan revision and implemented | `crates/rpptx/src/lib.rs:36681` to `36692` asserts both needed values and both trials `None` for the no-pair width fixture. The admitted-pair control at `:36459` keeps `Some(40000)` and asserts both trials present with `width_fits` true. |
| P2-S1, A5 absent-owner `IndexError` clause | Resolved by the approved plan revision and implemented | The current plan at `.claude/plans/F-X276-design.md:228` states the clarified accessor behavior. The Python test asserts causal `StaleElementError`, numeric `IndexError` and `index` `ValueError`. |
| P2-S2, paint fixture not run | Fixed | `p2-s2-existing-painted-table-attempt1.log` records the named existing fixture at `1 passed`. |
| P2-S3, span delta not exact | Fixed | `crates/rpptx-render/src/lib.rs:2765` to `2789` derives `needed` independently and asserts the last span row equals `needed - 6.0`. It also asserts the first row and the short neighbor stay `6.0`. |
| P2-S4, historical wrapper-exit claim | Fixed | The handoff and progress withdraw the universal claim. The abort log and the green log are separated. |
| P2-S5, stale Markdown and integration GREEN | Fixed | `p2-s5-fx218-attempt1.log` records `21 passed`. `p2-s5-issue431-attempt1.log` records `4 passed`, both on the post-repair source. |

## New finding

### S1, an empty notes placeholder gains the notes reason while status stays `empty`
`.claude/plans/F-X276-design.md:79`, `crates/rpptx/src/lib.rs:28090` to `28094`,
`crates/rpptx/src/lib.rs:4138` to `4150`

The plan says an empty owner keeps status `empty`, overflow false and does not
gain the cause. The notes fallback rule at `.claude/plans/F-X276-design.md:77`
says the inventory pushes the reason onto `update.diagnostics` once. The code
follows the second sentence without checking `owner.empty` at
`crates/rpptx/src/lib.rs:28090` to `28094`. The empty branch at `:4138` to
`:4150` then clones that update, so the frame is `empty` but its
`update.diagnostics` still contains the notes reason.

Concrete trigger: a notes slide with an empty placeholder whose idx matches no
notes-master placeholder and whose type has two notes-master candidates, or an
empty placeholder with no idx and two type candidates. Expected per plan line
79 is status `empty`, overflow false and no cause. Observed is status `empty`,
overflow false and a notes reason in `update.diagnostics`, visible through
Python `frame.update.diagnostics` (`crates/rpptx-py/src/layout.rs:251` to `255`
per the plan). The CLI line does not print it because the status is `empty`.

Severity low. This is a contract ambiguity rather than a production regression
in the named A1 fixtures. The notes fallback sentence and the empty-owner
sentence conflict. Root should clarify whether the empty-owner rule suppresses
the notes reason. No named A1 fixture exercises an empty notes placeholder with
an unresolved key, so the current tests do not catch it.

## A-ID outcomes

| A-ID | Local outcome | Evidence inspected | Remaining final gates |
|---|---|---|---|
| A1 | Covered locally, no blocking source defect. The one new smell is the empty notes placeholder reason above. Display and CLI/Python parity, physical duplicate leaves, read-only reporting versus mutating refusal, empty and unsupported frames, contextful and contextless causes, notes placeholder cardinality and collision behavior are all exercised. | `crates/rpptx/src/lib.rs:36292` to `36945`, `crates/rpptx/tests/integration.rs:487` to `700`, `crates/rpptx-cli/tests/integration.rs:2994` to `3279`, `crates/rpptx-py/tests/test_documented_examples.py:3579` to `3585`. Lib issue431 `6 passed`, measured-leaves `1 passed`, CLI fit `5 passed`, Python focused `2 passed`. | Full integrated union. |
| A2 | Covered locally, no blocking source defect. Stored width, grown geometry, page bounds, rotations and flips, shared paint deltas and trial fields are exercised. The no-pair width fixture asserts both needed values and both trials `None`. The admitted-pair control keeps the original height-only `Some(40000)` and asserts both trials present with `width_fits` true. | `crates/rpptx/src/lib.rs:36459` to `36720`, `crates/rpptx-render/src/lib.rs:1131` to `1210`, tests `2735` to `2850`. Admitted-pair `1 passed`, width `1 passed`, span `1 passed`, paint `1 passed`, named existing paint fixture `1 passed`. | Hash 49 and pixels 7. Full integrated union. |
| A3 | Covered locally, no blocking source defect. Plain CLI lines, context, exit codes and JSON reasons are exercised, including the picture `Diagnostic:` line, the contextless SmartArt route, the notes collision exit 2 and the non-diagnostic line count. | `crates/rpptx-cli/src/commands.rs:1811` to `1900`, `crates/rpptx-cli/tests/integration.rs:2994` to `3279`. CLI fit `5 passed`, CLI validate `1 passed`. | Full integrated union. |
| A4 | Covered locally, no blocking source defect. Assembled paragraph spacing, cross-run punctuation, hard breaks, notes, image alt text and the bounded line scan are exercised. The line-end scan is lazy through `md_line_end` and runs only for the `-` and `=` setext checks. | `crates/rpptx/src/html.rs:1176` to `1305`, tests `3550` to `3575`, `crates/rpptx/tests/integration.rs:95` to `700`, `1086` to `1092`. Post-repair `fx218_` `21 passed`, `issue431` integration `4 passed`, long-line `1 passed`. | HTML unchanged check. Full Markdown export benchmark. Full integrated union. |
| A5 | Covered locally, no blocking source defect. Live index identity, two wrappers, `move`, foreign, twin and reopened objects, empty collections, stale handles and invalid types are exercised. A removed handle through `slide_id` is `StaleElementError`, `slides[999]` is `IndexError` and `index` rejected inputs are `ValueError`. | `crates/rpptx-py/src/slide.rs:438` to `465` and `1531` to `1542`, `crates/rpptx-py/python/rpptx/_rpptx.pyi:634`, `crates/rpptx-py/tests/test_documented_examples.py:12515` to `12558`. Index test `1 passed`. | Installed Python 3.9 and 3.12 wheels, typing, stubs and archive. |

## Preserved XML, schema order and structural rules

No production change adds, removes or reorders serialized XML. `fit_report`,
`validate`, `index` and Markdown export do not mutate package bytes. The
duplicate deck, the notes deck and the Markdown fixtures assert unchanged bytes
where the plan requires it. `materialize_notes_bounds` mutates a measurement
clone, not the stored notes slide. The duplicate and cause diagnostics are
read-only report data. The new `AutofitTable` is private with three fields and
no builder. The new `MdNewline` is a private two-variant enum. No new trait,
generic, crate, module, file, feature flag or test binary was added. The two
clippy allowances at `crates/rpptx/src/lib.rs:27690` and `:27931` carry a
one-line justification. The HLD impact is exactly the approved sections in HLD
08 and HLD 10, including the clarified accessor error behavior. README
measurements, release scripts, `AS_BUILT.md`, `SPRINT_TRACKER.md` and the
baseline are untouched.

## Root supplied leads

The Markdown performance lead is Root's, recorded in
`/private/tmp/rdocx-cross-review/F276_ROOT_REVIEW_LEADS.json` at SHA-256
`6395617d7c9059dd8e05f729dfd21a08be8811470b01a69141860125ab7ec11b`. The
pre-repair helper SHA in that lead is
`11f481a98cd0b74c1186b549f40511d4f98f0210995a5e0f0712779dca1814bb`. The
current helper at `crates/rpptx/src/html.rs:1202` is lazy for the line-end scan
and the current file SHA is
`0f9903554a86b0cc50dfd09c99819c81f494c76ebf68c1c9e5d2151e5a1f94bf`. This is
Root discovery, not a blind peer discovery. The provenance and frozen pins are
also Root supplied. The full export benchmark remains final work.

## Remaining gates

These are not claimed as local passes and are not used to excuse the one smell
above.

- Full integrated workspace union.
- Hash harness, expected unchanged, hash 49 and pixels 7.
- Native, PDF and PNG rendering, LibreOffice.
- Installed Python 3.9 and 3.12 wheels, typing, stubs and archive.
- Hosted CI.
- Full Markdown export benchmark.
- Final union clippy, check and formatting over the integrated result. The
  p2-clippy and p2-check logs predate the final test-only A1 measured-leaves
  assertion reflow. The final A1 run compiled that test and the final
  `cargo fmt --all --check` passed, but the final union gate remains owed.
- Root final review after this pass.

## Not found

- No source defect in the approved `Display` sentences.
- No source defect in the `authoring_table_height` extraction or the shared
  `grown_row_heights` use by `lower_table`.
- No source defect in the duplicate-leaf publication shape, the first-leaf
  exclusion, the picture-only diagnostic or the mutation refusal.
- No source defect in the notes reason strings, the single-match copy, the
  zero-candidate and several-candidate reasons, or the master-key collision
  checked error.
- No source defect in the two unsupported publishers or the one-copy rule.
- No source defect in the width-qualified page test, the rotation and flip
  bounds, the trial fields or the original `Some(40000)` control.
- No source defect in the CLI line, context, exit rule or JSON reason.
- No source defect in the Markdown escape state, cross-run spaces, hard breaks,
  notes, alt text or the bounded line scan.
- No source defect in the `SlideCollection.index` live path or the stub.
- No source defect in the HLD impact sections.
- No source defect in preserved XML, schema order or the structural rules.

This is a report-only pass. The only write is this review file. No source,
test, HLD, plan, progress, canonical metadata, TEMP file, Git state, build,
test, hash harness, native app, oracle, network or credential was touched.
