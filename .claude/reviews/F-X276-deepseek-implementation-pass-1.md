# F-X276, deepseek implementation, pass 1

**Reviewed**: working tree of `work/f-x276-deepseek` against Base
`9a3b2960ccda7c3208ecb9b753c7019b64c086f5`, 11 files, 1338 insertions and
119 deletions. Source snapshot authenticated against
`/private/tmp/rdocx-cross-review/F276_AUTHOR_FROZEN_PINS.json`, all 11 file
SHA-256 values and the author handoff SHA-256 match. Approved plan
`.claude/plans/F-X276-design.md` SHA-256
`7ac52101673baac1ec65e485478488d1cf7a636a3e237fc83b6abde84e350382`.

**Verdict**: 4 defects, 5 smells, 0 nitpicks. Blocking, no clean acceptance.

**Scope note**: the author was controller-stopped after writing source, the
progress checkpoint and `AUTHOR_HANDOFF.md`. The frozen snapshot and the
11 changed files were identical before and after exit. The author aggregate
claim of five RED-then-GREEN outcomes is only partly supported by the logs
below.

## Defects

### D1, Python `ValidationIssue.message` still uses Rust `Debug`, not `Display`
`crates/rpptx-py/src/presentation.rs:197`

`PyValidationIssue::from` still builds `message` with `format!("{issue:?}")`,
and `crates/rpptx-py/src/presentation.rs` is not in the changed set. A broken
deck therefore returns
`DuplicateShapeId { slide: 0, id: 2 }` instead of the approved
`slide 1: duplicate shape ID 2, shared by more than one shape`. A1 requires
"Python `message` equals the CLI line", HLD 10 now states the message equals
the `Display` sentence, and the plan explicitly rejects "Python `message` left
as `Debug`" (`.claude/plans/F-X276-design.md:248`). The existing Python test
still pins the wrong text at
`crates/rpptx-py/tests/test_documented_examples.py:3582`, so it would also need
updating when this is fixed. Severity: high, public binding contract.

### D2, existing CLI `validate` test contradicts the new `Display` line
`crates/rpptx-cli/tests/integration.rs:432`

`validate_reports_duplicate_paragraph_properties_while_text_remains_readable`
asserts `diagnostic.contains("InvalidXmlChild")`. After
`crates/rpptx-cli/src/commands.rs:672` changed to `eprintln!("{issue}")`, the
stderr line is `{part} {path}: {message}` and no longer contains the variant
name. That existing test now fails, and the author did not run the CLI
integration binary outside the two new `issue431` filters, so the regression
was not observed. The plan requires the CLI to print the `Display` sentence
(A1, `.claude/plans/F-X276-design.md:200` and `:302`). Severity: high, existing
test breakage hidden by filtered execution.

### D3, A2 local assertion weakens the named `needed_scale` value
`crates/rpptx/src/lib.rs:36453`

The new A2 test asserts `frame.needed_scale.unwrap() < 100000`. A2 names
`needed_scale` `Some(40000)` and says "If this HEAD measures another scale,
stop and report it. Do not retune it"
(`.claude/plans/F-X276-design.md:181`). The runtime RED log does show
`needed_scale: Some(40000)` for the on-slide case, but the GREEN test would
also pass a retuned 50000 or 99999. The approved local discriminator is not
proved by the committed test. Severity: medium, weakened local evidence.

### D4, A2 named geometry and span controls are absent
`crates/rpptx/src/lib.rs:36421`

The only A2 test runs a 2 by 2 table at two vertical positions and checks
`overflow` for cell `(1, 1)`. It has no width failure, no right edge 580
control, no right edge 760 control, no rotated `max_x` or `min_x` fixture, no
on-page rotation, no `flip_v`, no row span and no short neighbor
(`.claude/plans/F-X276-design.md:202` through `:209`, test matrix C4 at
`:259`). A wrong implementation that compares `max_x` against page height, or
returns false for every non-zero rotation, or checks only the bottom, still
passes this test. The plan rejects exactly those implementations
(`.claude/plans/F-X276-design.md:236` through `:239`). Severity: high, the
named local discriminator is missing.

## Smells

### S1, A1 empty-frame observation is not tested
`crates/rpptx/src/lib.rs:36297`

The duplicate test has two non-empty text boxes and a picture. Nothing asserts
that an empty text frame stays `status=empty`, `overflow=false` and gains no
cannot-assess cause, which A1 names explicitly
(`.claude/plans/F-X276-design.md:191`). A regression that marks empty frames
`unsupported` would pass the current tests.

### S2, A1 matching-key deck does not check notes export or CLI exit 2
`crates/rpptx/src/lib.rs:36603`

The notes test asserts `collision.fit_report(&[]).is_err()` but never checks
that notes export returns the same collection error, nor that the CLI exits 2
(`.claude/plans/F-X276-design.md:199` and `:162`). A regression in the export
path would pass.

### S3, A3 CLI picture `Diagnostic:` line and the contextless route are untested
`crates/rpptx-cli/tests/integration.rs:2998`

`issue431_fit_text_output_names_each_frame_once` uses a clean table deck with
no duplicate id and no picture, and
`issue431_fit_reports_an_unsupported_frame_with_cannot_assess` uses a notes
placeholder reason. The plan's CLI test row names a "picture `Diagnostic:`
line" and "one cause on each route"
(`.claude/plans/F-X276-design.md:282`), and A1 requires the SmartArt CLI line
to contain `opaque or unresolved text owner` once
(`.claude/plans/F-X276-design.md:193`). Neither is asserted.

### S4, A5 `IndexError`-on-other-methods assertion was deleted
`crates/rpptx-py/tests/test_documented_examples.py:12550`

The first Python run failed because `held.slide_id` after `remove` raises
`StaleElementError`, not `IndexError`. The author removed the assertion rather
than covering the absent-owner path. A5 still states "Other methods still raise
`IndexError` for an absent owner"
(`.claude/plans/F-X276-design.md:228`), and `PySlide::validate` is the contract
point (`crates/rpptx-py/src/slide.rs:438`). No committed test now proves that
path.

### S5, A2 paint-delta local check was not authenticated
`crates/rpptx/src/lib.rs:36421`

The plan makes "the paint-delta check" part of A2 local evidence and says an
existing painted-table fixture must stay unchanged
(`.claude/plans/F-X276-design.md:181` and `:61`). The authenticated GREEN runs
are `issue431` and `fx218_` only. `fx218_` is semantic export, not table paint.
The `grown_row_heights` extraction looks faithful, but the named local paint
delta was not observed.

## Inspected behavior

- `Display for ValidationIssue` matches the approved sentences, including
  one-based slide, master and shape numbering
  (`crates/rpptx/src/lib.rs:1102`).
- `authoring_table_height` has the exact approved signature and sums the
  grown rows (`crates/rpptx-render/src/lib.rs:611`). `grown_row_heights` is a
  faithful extraction of the former growth loop, and `lower_table` still uses
  the same cumulative offsets (`crates/rpptx-render/src/lib.rs:1136`).
- `autofit_inventory` returns the approved tuple and takes `read_only`
  (`crates/rpptx/src/lib.rs:27934`). Mutation callers pass false, `fit_report`
  passes true. Duplicate ids produce one tuple per later physical leaf, the
  first leaf is not labelled, and the picture leaf is skipped as an owner
  (`crates/rpptx/src/lib.rs:27720`).
- Both unsupported publishers set `status=unsupported` and add the cause once
  when absent (`crates/rpptx/src/lib.rs:4151` and `:4230`). Empty owners keep
  `empty`, `overflow=false` and no cause.
- The page-growth exception is gated on `!fits && width_fits`, uses
  `authoring_table_height` and `shape_transform`, and compares the transformed
  bounding box against `(0, 0, page)` with the 0.01 epsilon
  (`crates/rpptx/src/lib.rs:4188` and `:27605`). Width failure falls back to
  `!measured.fits`, so it stays overflow.
- The CLI line, exit rule and JSON `reason` match A3
  (`crates/rpptx-cli/src/commands.rs:1849` and `:1813`).
- `SlideCollection.index` returns the live path position and maps extraction,
  foreign-presentation and `validate` failures to
  `ValueError("slide is not in this collection")`
  (`crates/rpptx-py/src/slide.rs:1532`). The stub is present
  (`crates/rpptx-py/python/rpptx/_rpptx.pyi:632`).
- The Markdown escaper assembles one paragraph across runs, keeps one inner
  space, escapes the approved always set and the conditional list and setext
  positions, resets state after a hard break, and applies the notes prefix
  outside the escaped text (`crates/rpptx/src/html.rs:1180`).
- HLD 08 and HLD 10 contain the approved replacement sentences and no fragment
  starting with "and" (`docs/hld/08-rendering-spec.md:3096`,
  `docs/hld/10-bindings-spec.md:903`).

## Authenticated executed evidence

- The earliest Cargo call was piped through `tail` without `pipefail`, so its
  pipeline exit is not Cargo's exit. Its visible output is an API compile RED,
  `E0599 ValidationIssue does not implement Display`, before the Display
  surface existed. It cannot establish runtime RED by itself.
- Later focused runs do contain runtime RED after Display was implemented:
  A2 reported `left: true, right: false` with `needed_scale: Some(40000)` for
  the on-slide case, A1 duplicate returned
  `Err "autofit /ppt/slides/slide1.xml: duplicate shape ID 2"`, A4 printed the
  old `&#32;` and mid-sentence escapes, and A3 printed `Array [Number(...)]`
  debug paths and failed the unsupported and unique-line assertions. The first
  table run stack-overflowed, and the 8 MiB rerun reached the assertions.
- Authenticated GREEN, read from the visible `test result` lines rather than
  pipeline exit: `rpptx --lib issue431` 3 passed,
  `rpptx --test integration issue431` 4 passed, `fx218_` 21 passed,
  `rpptx-cli issue431` 2 passed, Python `issue431` 1 passed against a freshly
  built extension, `cargo fmt --all --check` pass, scoped clippy pass after the
  two allowances, `prose_check` 0 violations, and `cargo check` for rpptx,
  rpptx-cli and rpptx-py pass.
- The two clippy allowances are declared with one-line justification at
  `crates/rpptx/src/lib.rs:27689` and `:27932`.

## Unverified claims and remaining checks

- A1 Python `message` equality is not merely unverified, it is wrong (D1).
- A2 exact `needed_scale` is not asserted in GREEN (D3), and the geometry,
  span and paint controls were not run (D4, S5).
- A1 empty-frame behavior, notes-export collision error, CLI exit 2 and the
  Python validate message were not run (S1, S2, D1).
- A3 picture `Diagnostic:` and contextless SmartArt CLI lines were not run
  (S3).
- A5 absent-owner `IndexError` was not run and the assertion was removed (S4).
- Existing unfiltered CLI integration and Python suites were not run by the
  author, which is why D2 and the Python Debug pin at
  `crates/rpptx-py/tests/test_documented_examples.py:3582` were not seen.
- Final union riders remain pending and are not claimed here: full workspace
  union, hash harness hash 49 and pixels 7, PDF/PNG and native viewers,
  installed Python 3.9 and 3.12 wheels, typing and stub checks, archive and
  CI. These are not used to excuse the local gaps above.

## Per-A276 coverage

| A-ID | Coverage | Status |
|---|---|---|
| A1 | Rust duplicate, notes reason and mutation refusal, SmartArt and table-cell causes, Rust `Display` | Partial, blocked by D1, S1, S2 |
| A2 | On-slide and off-slide overflow, row-growth extraction | Partial, blocked by D3, D4, S5 |
| A3 | Plain line count, uniqueness, table cell lines, exit rule, JSON `reason`, one notes unsupported route | Partial, blocked by S3 |
| A4 | Cross-run spaces, punctuation, breaks, notes, alt, label and existing `fx218_` updates | Substantially covered, no blocking defect found |
| A5 | Live index, two wrappers, move, foreign, twin, reopen, empty, non-slide, removed | Partial, blocked by S4 |

## Not found

- No defect found in the `Display` sentence text, the `authoring_table_height`
  extraction, the duplicate-leaf publication shape, the notes reason strings,
  the two unsupported publishers, the width-qualified page test, the CLI line
  and exit rule, or the Markdown escape state. Those aspects produced no
  additional defect beyond the evidence gaps recorded above.
