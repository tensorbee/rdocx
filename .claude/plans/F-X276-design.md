# F-X276, Report rpptx fit, plain Markdown, and slides.index

**Status**: approved

**Sprint**: S96

**Size**: M

**Depends on**: none

**Test gate**: regression

Root adopts this final-reviewed V4 preparation as the canonical draft on 2026-11-01. The source is GROK_PROPOSAL_V4.md at SHA-256 `90f3a88bb74ef63822198a405f720687945778e538feef1e9ce6f91c05a6edc9`, accepted as an adoption basis by the Codex affected final recheck fx276-codex-final-preparation-pass2.md at SHA-256 `268e633f9d0ba6baa035273d924bd7cded8ac21ec408e996af268606a764bfab`. Root independently verified four accepted candidate hashes and the 3854 other baseline pins unchanged before this adoption. The canonical branch is `sprint/s96` at `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd` with the preserved dirty source. The source preparation author has stopped, and the separate metadata intake adopter has returned its lease. F-X276 is already registered as pending and unclaimed in HLD 14, `BACKLOG.md`, `CURRENT_SPRINT.md` and `SPRINT_PLAN.md`, so this adoption does not rewrite those four registration files. This adoption is not approval and is not a source implementation review cycle or reset. Semantic closure, any material product questions and the global batch approval are still required. A clean Base is required before any claim and before a shared approved design commit.

## Operative rolling approval, 2026-11-01

Root approves this contract for the user's explicit rolling parallel start after the recorded final design reviews and canonical fidelity checks. Earlier preparation and batch-blocking wording is historical. The immutable product source Base is reviewed commit `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`, excluding the preserved unfinished F-X272 working changes. No product dependency exists. Implementation must recheck exact source locations on its worker Base and preserve every acceptance criterion, signature, HLD impact and pending final rider. Q-S96-07 is answered as bundled fonts by default, system fallback opt-in and caller fonts highest. This is contract approval, not implementation acceptance or a provider checkpoint. Three design peers remain used. Implementation has a separate cap of three independent passes. The current A2 and A5 contracts include the factual clarifications in `.claude/reviews/F-X276-peer2-contract-adjudication-codex-final-1.md`. Pair selection and existing accessor error behavior are unchanged. The active worker must acknowledge this contract revision before continuing affected work. At that historical checkpoint implementation peer2 was used, with one pass available. All three implementation peers have since completed. Codex final pass2 is clean after separate remediation, while final sprint acceptance remains pending.


## Problem

On main `24af95d3`, and on the pinned `rpptx` paths, one unassessable frame aborts `rpptx fit` and `fit_report` for the whole deck. A table cell taller than its stored row is `overflow=true` with `needed_scale=40000` when render grows the row and the table stays on the slide. Issue 431 records that table ending at 4.0 in on a 7.5 in slide. Plain `rpptx fit` prints Rust debug paths, omits the shape and cell, and repeats lines. Markdown writes every space as `&#32;` and escapes mid-sentence punctuation. `prs.slides.index` is missing.

The duplicate check returns `Err` for every physical leaf before the text projection (`crates/rpptx/src/lib.rs:27563-27568`). A picture has an empty `cells` vector and is skipped at `:27591`, so it never becomes an `AutofitOwner`. SmartArt is kept as an owner at `:27581-27584`. A missing projection stores its cause on `owner.unsupported` at `:27690-27694` and leaves `update.diagnostics` empty at `:27740`. `fit_report` publishes a contextful unsupported frame at `:4060-4074` and a contextless frame at `:4121-4134`. Neither copy copies `owner.unsupported` into `update.diagnostics`. CLI `fit` maps `Ok(false)` to exit 0, `Ok(true)` to exit 1, and `Err` to exit 2 (`crates/rpptx-cli/src/main.rs:557-569`). It currently turns status `unsupported` into `Err` (`commands.rs:1814-1838`). HLD 08 says "Unsupported mandatory owners refuse rather than vanish, and no permitted fitting trial retains a null needed scale" at `docs/hld/08-rendering-spec.md:3095-3096`.

Stored fit overflow is `!measured.fits` (`lib.rs:4108`). `StackedText::fits` is `width_fits && height <= content.height + 0.01` (`crates/rpptx-render/src/text.rs:680-681`). `width_fits` is cleared when ink exceeds the line (`text.rs:1065`).

## Spec reference

- `docs/hld/08-rendering-spec.md`, "Tables" and "Read-only fit and count consumers".
- `docs/hld/10-bindings-spec.md`, the `SlideCollection` paragraph, "`Presentation.validate()`", "PowerPoint fit", and "Presentation semantic text export".

Section names are the contract. Line numbers are a working-tree snapshot and will drift.

## Approach

No new crate, module, file, trait, generic, feature flag, or test binary. No new public page field. Carriers are the existing `crates/rpptx/src/lib.rs` tests, `crates/rpptx/tests/integration.rs`, `crates/rpptx-cli/tests/integration.rs`, `crates/rpptx-py/tests/test_documented_examples.py`, and `crates/rpptx-py/python/rpptx/_rpptx.pyi`. This story must not edit `crates/rdocx-py/tests/test_core.py`.

Existing signatures stay:

- `Presentation::fit_report(&self, font_files: &[(&str, &[u8])]) -> Result<FitReport>`
- `Presentation::validate(&self) -> Vec<ValidationIssue>`
- `Presentation::to_markdown(&self) -> Result<PresentationTextExport>`
- `Presentation::to_markdown_with_options(&self, options: &PresentationTextExportOptions) -> Result<PresentationTextExport>`
- `refresh_autofit` and `fit_text_at` stay mutating.

Additive public surface, minor 0.15.0:

- `impl std::fmt::Display for ValidationIssue`. `Debug` stays derived.
- This exact helper in `rpptx-render`. `rpptx` calls the render crate. There is no trait.

```rust
pub fn authoring_table_height(
    table: &ResolvedTable,
    fonts: &mut FontManager,
    page_number: usize,
    directions: &[Vec<oxml_layout::TextDirection>],
) -> Result<f64, RenderInputError>
```

The value is the sum of the grown row heights. `lower_table` row offsets equal those cumulative sums. The last row of a span grows by `needed - height` when `needed > height + 0.01` (`crates/rpptx-render/src/lib.rs:1162-1164`). A short cell adds nothing. Paint equality means those deltas stay the painter's deltas. An existing painted-table fixture must stay unchanged. If the extraction moves a pixel, fix the extraction. Do not move the hash baseline.

Python, additive, with no default and no Rust, CLI, or WASM twin:

```python
def index(self, slide: Slide) -> int
```

`SlideLayouts.index` is unchanged (`crates/rpptx-py/src/slide.rs:294-304`).

Private `read_only` flag on the existing `autofit_inventory` and `materialize_notes_measurement_bounds`. `fit_report` passes true. `refresh_autofit` and `fit_text_at` pass false.

Private returns change, with no new struct. `autofit_inventory` returns `Result<(Vec<AutofitOwner>, Vec<(String, Vec<usize>, oxml_layout::Diagnostic)>)>`. The tuple is part, child path, and one `Diagnostic { message }`. `refresh_autofit`, `fit_text_at`, and the debug comparison at `lib.rs:4203` discard the vector. `fit_report` keeps it.

`append_autofit_context` already visits every physical leaf. When `read_only` is false, a repeated id still returns `autofit {part}: duplicate shape ID {id}` for a picture, a connector, or a text shape. When `read_only` is true, a repeated id pushes one tuple and does not return `Err`. The leaf is still skipped when `cells` is empty, so a picture still creates no owner. `fit_report` appends each diagnostic to `FitReport.diagnostics` once per leaf, after the font facts, not once per consuming context. If an owner has the same part and child path, it appends that diagnostic to `update.diagnostics` once, after the font-diagnostic assignment at `:4033`. Each frame cloned from that owner then carries the sentence once. The context loop does not push it again. A picture appears only in `FitReport.diagnostics`. The first leaf is not labeled a duplicate.

`materialize_notes_measurement_bounds` returns `Result<Vec<(Vec<usize>, String)>>`. The path is the same child path `owned_items` uses. After `append_autofit_context`, the inventory finds the notes owner with that part and path and sets `unsupported` to the reason. For a non-empty owner only, it pushes `Diagnostic { message: reason }` onto `update.diagnostics` once. An empty notes placeholder follows the general empty-owner rule below and receives no unsupported cause.

Unsupported cause publication. This is one rule for both publishers. Immediately before the contextful push at `:4060` and immediately before the contextless push at `:4121`, if the owner is non-empty and `unsupported` is `Some(cause)`, set status to `unsupported` and, when `update.diagnostics` does not already contain that cause, push one `Diagnostic { message: cause }`. Do not push it again. Do not append that cause to `FitReport.diagnostics`. An empty owner keeps today's status `empty`, overflow false, and does not gain this cause. Notes reasons already stored on `update.diagnostics` stay one copy. Duplicate-leaf diagnostics stay on their own channel and are not rewritten by this rule.

Page-growth exception. It applies only to a measured table cell whose stored measurement has `width_fits == true` and `fits == false`. Overflow is then false only when the grown transformed rectangle lies on the page. When `width_fits == false`, overflow stays true. The exception does not run for a non-table text frame, an empty frame, or an unsupported frame. `needed_scale`, `needed_reduction`, and both trials stay the stored-row hypothetical. Candidate-pair selection is unchanged. If no pair fits every context in both measurements, both needed values and both trials are `None`. If a pair is admitted, both trials are present and `width_fits` is true. Do not synthesize a non-fitting trial.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Whole-deck fit report | Rust `fit_report`, Python `fit_report`, CLI `fit` and `--json` | Exit 1 when any frame has `overflow` or status `unsupported`. Status `empty` with `overflow` false does not force exit 1. A non-empty unsupported cause is on the frame diagnostics once | Read-only. No save. Reopen the same bytes |
| Grown-table overflow | Those three, plus paint through `authoring_table_height` | The exception runs only for stored height failure with `width_fits` true. Width failure stays overflow | Stored-box measure, grown page test, PNG and PDF unchanged |
| Duplicate id | `fit_report` versus `refresh_autofit` and `fit_text_at` | One report diagnostic per later physical leaf. Mutation refuses every later physical leaf | Report, then a refused mutation, bytes unchanged |
| Notes idx | `fit_report` only | Do not change `PlaceholderKey::matches`. Do not change `collect_export_placeholder_shapes` or notes export | Measurement clone only. Saved idx and absent transform stay |
| Validate sentence | CLI `validate`, Python `ValidationIssue.message` | The message equals the CLI line. `kind` stays snake_case | Same deck, both consumers |
| Markdown spaces | Rust `to_markdown` and `to_markdown_with_options`, Python `to_markdown`, `export_markdown`, `save_markdown`, CLI `convert --to md` | `inline` for shape runs. Notes and alt use `semantic_md_preserved`. The non-decimal label uses `semantic_md_text` | One run, split runs, alt, label, break, grave accent, notes marker, and the `-5` boundaries. HTML and WASM untouched |
| `slides.index` | Python `SlideCollection.index` and the stub | Live path on that presentation | Fresh lookup, held handle, two wrappers, `move`, foreign, reopened, empty, removed, and a non-slide |

Unreduced, and not dropped: Issue 431 leaves PowerPoint-saved layout, master, and chart refusals out of this reduction. Owners that already become `unsupported` are reported with a cause and exit 1. This story does not add chart fitting, general off-page detection, or a general layout-table repair. There is no native PowerPoint claim.

## Behavior contracts and evidence

| Surface | Exact behavior and boundary | Primary source | Discriminator |
|---|---|---|---|
| `authoring_table_height` | Signature above. Epsilon `0.01` points | `pulls/pr-438.diff:326-335`. Growth at `rpptx-render` `lib.rs:1162-1164` | A2. Unexecuted |
| Page test | The exception in Approach. Page origin is `(0, 0)`. Slides use `slide_size` EMU `/ 12700`, default `(720.0, 540.0)` at `lib.rs:27765-27767`. Notes use `export_page_size` (`:26175-26179`). Local rect is `(0, 0, column-width sum, grown-height sum)` through `shape_transform` (`rpptx-render` `lib.rs:1013-1018`, `:2120-2148`). On page means `transform_rect_bbox` (`oxml-layout` `transform.rs:68-96`) has `min_x >= -0.01`, `min_y >= -0.01`, `max_x <= width + 0.01`, and `max_y <= height + 0.01` | `width_fits` at `text.rs:1065`. `fits` at `text.rs:680-681`. Current overflow at `lib.rs:4108` | A2, including the on-page width failure. Unexecuted |
| Unsupported cause | Non-empty contextful and contextless frames each carry `unsupported` once on `update.diagnostics`. Python reads that through `PyAutofitUpdate.diagnostics` (`layout.rs:251-255`). CLI `cannot assess` prints it once. `FitReport.diagnostics` does not gain a second copy | Contextful publish `:4060`. Contextless publish `:4121`. SmartArt owner `:27581` | A1 SmartArt and table-cell fixtures. Unexecuted |
| Duplicate id | Message `{part} {plain path}: shape ID {id} is shared with an earlier shape`. Plain path is zero-based ordinals joined by `/`, such as `0` or `1/0`. One report diagnostic per later leaf | `Err` at `lib.rs:27563-27568`. Picture skip at `:27591`. Report assembly at `:4159-4161` | A1 both text frames and the picture. Unexecuted |
| Notes fallback | Only when `read_only` is true and collection returned `Ok`. Exactly one key match copies that transform. Zero key matches and exactly one equivalent type copies that master. Otherwise the transform stays unset and the returned vec carries the reason. `old_bounds` is `(off.x, off.y, ext.cx, ext.cy)` when copied (`lib.rs:27700-27737`) and `(0, 0, 0, 0)` when unset | `PlaceholderKey::matches` at `placeholder.rs:289-296`. Collection `Err` at `lib.rs:26378-26382` | A1 strings below. Unexecuted |
| CLI line | One line per frame. Omitted pieces are absent, not bracketed: `{owner_part} {plain path} {name:?} cell {row},{column} in {context}: status={status}, overflow={true\|false}, needed_scale={int\|none}[, cannot assess: {reason}]`. `in {context}` appears when context differs from the owner part. Name uses `Debug` quoting | `commands.rs:1868-1877` | A3. Unexecuted |
| Exit and JSON | Exit 0 only when every frame has a status other than `unsupported` and `overflow` false. Exit 1 when any `overflow` is true or any status is `unsupported`. Exit 2 stays a checked error. JSON `reason` is `cannot-assess` for status `unsupported`, even when `overflow` is false. Otherwise it stays `unfit-at-floor` when `overflow` is true and `needed_scale` is absent, and null otherwise | Empty overflow false at `lib.rs:4046-4057`. Non-empty unsupported overflow true at `:4065` and `:4127` | A1. Unexecuted |
| `Display` | Sentences below. Spoken slide, master, and shape numbers are 1-based. Fields stay zero-based. Python `message` is the same sentence | HLD 10 requires that equality. The frozen snapshot pins `DuplicateShapeId { slide: 0, id: {first} }` | A1. Unexecuted |
| Markdown | Always-escape set and conditions below. A space is literal only when both neighbors in the assembled paragraph are non-whitespace. Otherwise `&#32;`. Tab and CR stay `&#9;` and `&#13;`. Alt newlines stay `&#10;` (`html.rs:657-664`) | Current escaper at `html.rs:1147-1160`. Notes prefix at `:254-259`. Break insert at `:512-518` | A4. Unexecuted |
| `index` | Live zero-based position. Identity is the same `Py<PyPresentation>` plus the live path from `current_path` (`presentation.rs:3424-3438`). Every miss uses `slide is not in this collection`. A non-slide, including `slide`, is `ValueError`, so `extract` is not the only gate. `PySlide::validate` keeps its defensive absent-owner `IndexError` branch (`slide.rs:438-450`). Valid public handles whose owner is removed instead fail scope validation with `StaleElementError` before that branch. `index` converts its rejected inputs, including a stale generation, to `ValueError` | python-pptx 1.0.2 `Slides.index` at `src/pptx/slide.py:285-293`. Equality is element identity in `src/pptx/shared.py:24-33`. Annotated tag `799d6b153eb69cf237191cf0ccd8568591dd9599`, commit `278b47b1dedd5b46ee84c286e77cdfb0bf4594be` | A5. Unexecuted |

Always escape these characters and no others:

- U+005C REVERSE SOLIDUS
- U+0060 GRAVE ACCENT
- U+002A ASTERISK
- U+005F LOW LINE
- U+005B LEFT SQUARE BRACKET
- U+005D RIGHT SQUARE BRACKET
- U+003C LESS-THAN SIGN
- U+003E GREATER-THAN SIGN
- U+0023 NUMBER SIGN
- U+007C VERTICAL LINE
- U+007E TILDE
- U+0026 AMPERSAND

U+0020 SPACE is not in that set. Conditional escapes:

- U+002D HYPHEN-MINUS and U+002B PLUS SIGN at the content column only when the next character is whitespace or the line ends.
- U+003D EQUALS SIGN, or a line of U+002D, only when that whole line is a setext underline.
- U+002E FULL STOP and U+0029 RIGHT PARENTHESIS only when they end an ordered-list number on that same line.

`Display` sentences from `pulls/pr-438.diff:499-550`:

- `DuplicateShapeId`: `slide {slide+1}: duplicate shape ID {id}, shared by more than one shape`
- `SlideIdOutOfRange`: `slide {slide+1}: slide ID {id} is outside 256 to 2147483647`
- `DuplicateSlideId`: `two slides share slide ID {id}`
- `MissingContentTypeOverride`: `{part} has no content type override`
- `DanglingRelationship`: `{part}: relationship {r_id} to {target} is referenced without being defined, or defined without being referenced`
- `UnreachableRelationshipTarget`: `{part}: relationship target {target} is not in the package`
- `EmptyTextBody`: `slide {slide+1}: shape {shape+1} has a text body without a paragraph`
- `DuplicatePlaceholderIdx`: `slide {slide+1}: two placeholders share placeholder idx {idx}`
- `OrphanMedia`: `{part} is not the target of any relationship`
- `CustomShowReference`: `a custom show references relationship rId{slide_id}, which names no slide`
- `MissingLayoutRel`: `slide {slide+1} has no slide layout relationship`
- `MissingThemeRel`: `slide master {master+1} has no theme relationship`
- `InvalidXmlChild`: `{part} {path}: {message}`

Markdown position. Track the digits-only line prefix incrementally rather than rescanning it for each period or closing parenthesis. Each input character participates in at most one ordered-prefix digit check. A deterministic work counter in the existing long-line regression distinguishes a long digit prefix followed by repeated punctuation, without a wall-clock threshold. State crosses runs of one paragraph. `TextRun::Break` inserts `  \n` and resets the line column and the digit count. Heading text after `# ` is not a list start. A list body after `- ` or `1. ` escapes a nested marker. Notes are written as `> ` plus `semantic_md_preserved` of each line (`html.rs:254-259`). A marker at the start of that line is escaped. The generated `> ` is not part of the escaped text. A plain paragraph starts at column 0. HTML stays on `semantic_escape`. WASM has no semantic export.

Notes reasons, exact:

- Idx present, zero type candidates: `notes placeholder idx {idx} matches no notes-master placeholder, and no notes-master placeholder has its type`
- Idx absent, zero type candidates: `notes placeholder has no idx, and no notes-master placeholder has its type`
- Idx present, several type candidates: `notes placeholder idx {idx} matches no notes-master placeholder, and {count} notes-master placeholders have its type`
- Idx absent, several type candidates: `notes placeholder has no idx, and {count} notes-master placeholders have its type`

The absent-idx several case is the existing two-key-match path. Masters with different idx values do not match each other, so collection succeeds. A shape with no idx matches both by type (`placeholder.rs:292-295`). On the read-only path this returns the vec and does not return `ambiguous notes placeholder geometry`. On the mutation path that `Err` stays.

A foreign idx that equals a different-type master's idx keeps today's match. The type fallback does not run. Do not add a type-agreement check.

Collection stays `notes slide has ambiguous placeholder {key:?}` for fit and for notes export when two master keys match each other. CLI exit is 2.

Package bytes, revisions, raw XML, and schema order stay unchanged for `fit_report`, `validate`, `index`, and Markdown export.

## Dependency readiness

There is no product dependency, and no F-X272 provider is invented. F-X272 does not supply a fit, index, or Markdown rule. Q-S96-07 is answered as bundled default with an explicit system opt-in and caller priority. It is not an F-X276 product choice, and this story has no blocking product question of its own.

Historical evidence only: peer pass 2 described an F-X272 handoff as frozen and named Word field regions of HLD 10 at working-tree regions 2915 and 4880. That sentence is not current writer state, and Codex final 1's live-author observation is historical qualification, not a reexecuted current test. The source preparation author has stopped, and the separate metadata intake adopter has returned its lease. The preparation turn re-hashed the files in the pin table and did not re-count the tree. Those pins, CI facts and dates remain historical.

Registration is complete. F-X276 is pending and unclaimed in both `CURRENT_SPRINT.md` and `BACKLOG.md`, and it is registered in HLD 14 and `SPRINT_PLAN.md`. This adoption does not write those four files again.

The clean-base gate is separate. Claim runs only from `sprint/s96` when the tree is clean except the approved F-X276 plan. The canonical branch is `sprint/s96` at `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd` with the preserved dirty source. This adoption does not create `work/f-x276-grok`.

## Acceptance criteria

| ID | Setup | Observation | Consumers | Local test | Final gate |
|---|---|---|---|---|---|
| A1 | Issue case 1. Text id 2, then picture id 2. Two text boxes with id 2. One empty text frame. A contextful table cell in shape autofit. A contextless SmartArt frame. Notes `sldNum` idx 10 with one type match and non-empty text. Idx 10 against masters idx 1 and idx 2. No idx against those two masters. Idx 10 with zero type candidates. No idx with zero type candidates. Two master keys that match each other | See the A1 observations below | Rust, Python, CLI | Existing `rpptx` and CLI carriers, `test_documented_examples.py` | Full union. No native PowerPoint. Unexecuted |
| A2 | Issue case 2. Short neighbor, span, width failure, and the geometry below. Slide size is 720 by 540 points | On-slide issue cell: `overflow` false and `needed_scale` `Some(40000)` with deterministic bundled fonts and empty caller fonts. If this HEAD measures another scale, stop and report it. Do not retune it and do not move the baseline. Off-slide cell: `overflow` true. Width failure and the geometry booleans below hold. A span grows its last row by `needed - height` only | Rust, Python, CLI, paint helper | `rpptx` integration and the paint-delta check | Hash 49 and pixels 7 unchanged. Unexecuted |
| A3 | Issue case 3 | No exact duplicate non-diagnostic lines. No `Array [` or `Number(`. Four slide-3 lines name `Table 3` and a cell. Lines that do not start with `Diagnostic:` equal `frames.len()` | CLI text and JSON | `rpptx-cli` integration | Full union. Unexecuted |
| A4 | Issue case 4 plus the fixtures below | Lines equal `# Q3 revenue, by region`, `- North grew 12% (EUR 1.2m) - see table.`, and `> Thank the team.`. The repeated-space line contains `&#32;` and its inner single spaces stay spaces. HTML alt stays `alt="alt [x] &lt; &amp;\nnext"` | Rust, Python, CLI `--to md` | `rpptx` integration | HTML unchanged. Unexecuted |
| A5 | Issue case 5 plus `move`, two wrappers, a fresh lookup, a twin, a reopen, an empty collection, `slide`, and a removed or stale handle | See the A5 observations below | Python and stub | `test_documented_examples.py` | Installed 3.9 and 3.12, typing, stubs, wheels. Not a local worker gate. Unexecuted |

A1 observations:

- `fit_report` is `Ok` on the duplicate, SmartArt, table-cell, and notes decks. Read-only calls leave the package bytes equal.
- The first text leaf with id 2 has a frame, is measured, and its diagnostics do not contain `shape ID 2 is shared`. The second text leaf has its own frame, is measured, and its diagnostics contain its sentence once. `FitReport.diagnostics` contains that sentence once. Dropping either frame fails.
- `FitReport.diagnostics` contains `{part} {picture path}: shape ID 2 is shared with an earlier shape` once. The picture has no frame. `refresh_autofit` returns `Err` containing `duplicate shape ID 2`.
- The empty text frame stays status `empty`, overflow false, and does not gain a cannot-assess cause.
- The shape-autofit table cell has a context other than `uninstantiated`, status `unsupported`, and `update.diagnostics` equal to one message, `table-cell shape autofit has no independently owned extent`. The CLI line contains that cause once. `FitReport.diagnostics` does not contain it. Python `frame.update.diagnostics` equals that one message.
- The SmartArt frame has context `uninstantiated`, status `unsupported`, overflow true, and `update.diagnostics` equal to one message, `opaque or unresolved text owner`. The CLI line contains that cause once. `FitReport.diagnostics` does not contain it.
- The single-match notes frame `old_bounds` equals that master's EMU tuple. After `fit_report` and after a mutation call, the stored transform is still absent and the stored idx is still 10.
- The idx-10 several deck has status `unsupported`, `old_bounds` `(0, 0, 0, 0)`, and reason `notes placeholder idx 10 matches no notes-master placeholder, and 2 notes-master placeholders have its type`, once.
- The absent-idx several deck has status `unsupported` and reason `notes placeholder has no idx, and 2 notes-master placeholders have its type`, once. Mutation on that deck still returns `ambiguous notes placeholder geometry`.
- The idx-10 zero-candidate deck has status `unsupported` and reason `notes placeholder idx 10 matches no notes-master placeholder, and no notes-master placeholder has its type`, once.
- The absent-idx zero-candidate deck has status `unsupported` and reason `notes placeholder has no idx, and no notes-master placeholder has its type`, once.
- The matching-key deck returns the collection `Err`. Notes export of that master returns the same `Err`. CLI exit is 2.
- The validate sentence and Python `message` match.

A2 geometry, identity group unless noted. The page test runs only for the guarded height failure.

- Width failure: no-wrap text in a table cell whose stored and grown rectangle lies inside 720 by 540. Stored `width_fits` is false. Expected `overflow` true. For the W-times80 fixture with no admitted pair, `needed_scale`, `needed_reduction`, `quantized_trial` and `persisted_trial` are `None`. Published trials for an admitted pair remain the stored-row hypothetical and have `width_fits` true. The page-growth exception must not clear a stored width failure. Do not infer that every stored width failure has no fitting pair.
- Width discriminator: origin `(500, 100)`, width 80, rotation 0, no flip, `width_fits` true, grown bottom at most `540.01`. Right edge is 580. Expected `overflow` false. Comparing `max_x` with page height 540 reports true and fails this fixture. Keep the off-page control at `(680, 100)`, width 80, right edge 760, expected `overflow` true, and the on-page control at x 100, right edge 180, expected `overflow` false.
- Rotated `max_x`: bounds `(660, 100, 40, 100)`, rotation 90, no flip, `width_fits` true. `Transform::rotate_about(90, 20, 50)` then translation by `(660, 100)` maps local `(x, y)` to `(730 - y, 130 + x)`. `max_x` is 730. Expected `overflow` true while grown bottom stays at most `540.01`.
- Rotated `min_x`: bounds `(0, 200, 20, 80)`, rotation 90, no flip, `width_fits` true, one row whose stored height is 80. The same rotation maps local `(x, y)` to `(50 - y, 230 + x)`. Grown `H` must satisfy `80 < H <= 340`. Expected `overflow` true. If a font build misses the range, change the text until `H` is inside it.
- On-page rotation: bounds `(300, 100, 100, 40)`, rotation 90, `width_fits` true, grown `H` with `40.01 < H <= 150`. Page point is `(370 - y, 70 + x)`. Expected `overflow` false.
- `flip_v`: bounds `(100, 0, 80, 40)`, `flip_v`, rotation 0, `width_fits` true, grown `H` with `40.01 < H <= 540`. Page `Y = 40 - local_y`. Expected `overflow` true.

A4 fixtures:

- `Q3` plus ` revenue` is one literal space. `Hello ` plus `world` is one literal space. `a` plus `  b` keeps both spaces as `&#32;`. Leading and trailing space runs stay `&#32;`.
- `see ` plus `- table` does not escape the hyphen. A plain paragraph or list body that starts with `- ` does escape it.
- `1` plus `. item` at column 0 escapes the dot. The same split after `see ` does not. An asterisk and a low line stay escaped in every run.
- A run whose text is a grave accent, then `code`, then a grave accent, escapes both grave accents. It does not remain a raw fence.
- Alt contains `alt \[x\] \< \&&#10;next` and does not contain `alt&#32;`. The label fixture contains `4. IV) item 0\n5. V) item 1`.
- `Hello` plus a break plus `- next` contains `Hello  \n\- next`. `1` plus a break plus `. item` does not escape the dot. The same text without a break, at column 0, does.
- Notes text `- next` contains `> \- next`.
- A plain paragraph `-5` contains `-5` and does not contain a backslash before `5`. `+5` contains `+5` and does not contain a backslash before `5`. `= x` contains `= x` and does not contain a backslash before `=`.

A5 observations:

- `index(slides[2]) == 2`.
- Two wrappers fetched as `slides[2]` both return 2.
- After `move(2, 0)`, the held wrapper returns 0, and a fresh `slides[0]` fetched after the move returns 0. Both resolve through the same presentation and the same live slide path.
- A twin, a reopened deck, a foreign slide, an empty collection asked for a foreign slide, the string `slide`, and a removed or stale handle raise `ValueError` with `slide is not in this collection`.
- Other accessors keep their existing errors. A removed handle used through `slide_id` raises `StaleElementError`, and an invalid numeric collection index raises `IndexError`. The defensive absent-owner `IndexError` branch is unchanged, but no public absence path is promised or fabricated.

Wrong implementations rejected:

- Aborting the deck, or an empty cannot-assess reason.
- Copying the unsupported cause only at `:4060`, so SmartArt at `:4121` has an empty cause.
- Copying the cause twice, or also appending it to `FitReport.diagnostics`.
- Turning an empty frame into status `unsupported`.
- Treating stored-row overflow as slide overflow, checking only the bottom, or treating every rotation as overflow.
- Clearing overflow for an on-page cell whose `width_fits` is false.
- Using page height as page width, so right edge 580 overflows.
- Checking stored x plus vertical edges, so the rotated `max_x` 730 or `min_x` fixture stays on the page.
- Recording a duplicate diagnostic only on text owners, dropping the first text frame, or emitting one duplicate diagnostic per context.
- Letting `refresh_autofit` continue after a picture with a repeated id.
- Leaving a zero or several type candidate at status `unchanged`.
- Copying notes geometry on a mutation path, changing stored idx 10, or guessing among several candidates.
- Changing `PlaceholderKey::matches`, or turning the collection `Err` into a report.
- Changing package bytes on `fit_report`, `validate`, `index`, or Markdown export.
- Collapsing shared frames. Restarting escape state per run. Carrying the column across a hard break. Escaping `-5`, `+5`, `= x`, or a literal space. Leaving a grave accent or a notes-line hyphen unescaped.
- `index` by slide id, by the construction index, or `TypeError` for a string. Two wrappers or the post-move fresh lookup disagreeing.
- Python `message` left as `Debug`.
- A README measurement edit or a hash-baseline edit.
- Exit 1 implemented from overflow alone.

## Test matrix

| Case | Modes or state | Expected difference | Wrong implementation rejected |
|---|---|---|---|
| C1 | Read-only fit and mutation, first text leaf, second text leaf, picture | Both text frames exist. Only the later leaf is named. The picture has no frame. Mutation returns the current `Err`. Bytes agree | A missing first frame, or a text-owner-only diagnostic |
| C2 | Contextful shape-autofit cell and contextless SmartArt, plus one empty text frame | Each non-empty cause appears once on the frame and in Python and the CLI line. The empty frame stays empty | A repair only at `:4060`, or a second copy in `FitReport.diagnostics` |
| C3 | Single type match, both several reasons, both zero-candidate reasons, matching master keys | The exact reasons, stored idx 10, absent transform, and equal package bytes. Export still errors for matching keys | A dropped zero-candidate deck or a changed idx |
| C4 | Eligible height failure on the page, on-page width failure, right edge 580, right edge 760, rotated `max_x`, rotated `min_x`, and `flip_v` | Height failure can be `overflow` false. Width failure stays true. The other booleans match A2 | An unconditional page sentence, or height used as width |
| C5 | One run, two runs, alt, label, break, grave accent, notes hyphen, `-5`, `+5`, and `= x` | The A4 strings | A space in the always-escape set, or a notes hyphen left plain |
| C6 | Two wrappers, held handle, fresh `slides[0]`, twin, reopen, empty, non-slide, removed or stale | A5 | A cached construction index |
| C7 | CLI, JSON, and Python on one unsupported frame with overflow false | Exit 1 and `reason` `cannot-assess` | Exit 0 from an overflow-only return |

Word validation, Word Markdown, WASM semantic export, chart fitting, and `SlideLayouts.index` stay outside the issue. HTML is covered by the unchanged alt assertion.

## Rejected alternatives

- Cherry-pick of `65305305`, the README measurements. Re-check apply on a clean post-integration HEAD. Do not treat the old apply log as that result.
- Donor `grown_table_fits_page`, which returns false for every non-zero rotation and then tests only the bottom (`pulls/pr-438.diff:659-685`).
- Donor per-run reset of `line_digits`, and donor escaping of every leading hyphen, plus sign, or equals sign (`pulls/pr-438.diff:409-422`).
- Donor `TypeError` for `slide`, and Python `message` left as `Debug`.
- A new struct, module, or file for the diagnostic tuple, the notes reason vec, or the cause copy.
- Changing `PlaceholderKey::matches` or `collect_export_placeholder_shapes`.
- Revising the approved F-X275 plan so its inventory names F-X276.
- Waiving the restored A1 or A5 assertions as optional test debt.

## Test plan

| Timing | Category | Test | Asserts |
|---|---|---|---|
| local scoped | regression | Existing `rpptx` tests for A1 through A4 | The observations above. Issue 431 is the RED description. None of these tests have been run |
| local scoped | regression | Existing `rpptx-cli` integration for A3, the exit rule, and both unsupported routes | Non-diagnostic count, picture `Diagnostic:` line, one cause on each route, exit 1 when status is `unsupported` |
| local scoped | regression | `test_documented_examples.py` for A5 and the validate message | Wrapper and fresh-lookup agreement, `ValueError` text, Python message equal to the CLI line |
| final sprint | regression | `/verify` on the integrated sprint, including the hash harness | Hash unchanged. Pixels 7 unchanged |
| final sprint | risk riders | WASM check, Python 3.9 and 3.12 wheels, typing, stubs, archive, and the hosted debt already owed by the union | Local tests are not that gate. LibreOffice is a final-union viewer rider |

The test gate is `regression`. A live python-pptx process is not required. No assertion in this plan is optional.

## HLD impact

`/complete-feature` updates exactly these sections.

`docs/hld/08-rendering-spec.md`, "Read-only fit and count consumers". Replace the whole sentence at lines 3095-3096, from "Unsupported mandatory owners refuse rather than vanish" through "no permitted fitting trial retains a null needed scale." Do not leave a fragment that starts with "and". The replacement is:

An unsupported mandatory owner is named with its cause and does not vanish or abort the deck. The cause is published once on every non-empty unsupported frame, whether that frame has a measurement context or is contextless. The CLI exits 1. A table cell whose stored measurement fits the width and not the height overflows only when the grown table's transformed rectangle leaves the page. A cell whose stored measurement does not fit the width stays overflow. The test uses the shared row growth and `shape_transform`, including rotation and flip. `needed_scale` stays the stored-row hypothetical. A repeated shape id is measured and named once for every later physical leaf. `refresh_autofit` and `fit_text_at` still refuse. A notes placeholder whose idx matches no notes-master placeholder is measured from the single notes-master placeholder of its type, on the read-only fit path only. Zero or several type candidates after a successful collection are reported, not guessed. A master-key collision that collection already rejects stays a checked error. No permitted fitting trial retains a null needed scale.

Do not rewrite "Tables" unless the shared helper forces one sentence. Do not copy the donor wording that the table ends below its page, or that an unsupported owner counts as an overflow.

`docs/hld/10-bindings-spec.md`:

- `SlideCollection`. `index` returns the live zero-based position and raises `ValueError` with `slide is not in this collection` for another presentation, a structural twin, a reopened deck, a non-slide, a removed or stale handle, and a foreign slide against an empty collection. Two wrappers and a fresh lookup after `move` agree. Other accessor errors stay unchanged, including causal `StaleElementError` for a removed handle and `IndexError` for a numeric index outside the collection. The defensive absent-owner branch is not presented as a reachable public case.
- "`Presentation.validate()`". The printed line is the `Display` sentence, and Python `message` equals it.
- "PowerPoint fit", current text at lines 4087-4089. Keep "No fitting permitted trial reports null needed scale and unfit-at-floor." Replace "Unsupported mandatory owners refuse. Fit does not save, exiting0 without overflow,1 for overflow and2 for checked error." Exit 0 when every frame is assessed and none overflow. Exit 1 when any frame overflows or has status `unsupported`. Exit 2 for a checked error. A non-empty unsupported frame carries its cause once. JSON `reason` is `cannot-assess` when status is `unsupported`. The page-growth exception is the same width-qualified rule as HLD 08. Fit does not save.
- "Presentation semantic text export", section starting at line 4656. Keep the existing method sentences. Add this paragraph: Markdown keeps one space between two non-whitespace characters as a space. U+0020 is not always escaped. A leading, trailing, or repeated space, and a space beside whitespace, is written `&#32;`. Tab and CR stay `&#9;` and `&#13;`. The always-escaped characters are U+005C, U+0060, U+002A, U+005F, U+005B, U+005D, U+003C, U+003E, U+0023, U+007C, U+007E, and U+0026. U+002D and U+002B are escaped only where they would start a list marker. U+003D is escaped only where the line would be a setext underline. U+002E and U+0029 are escaped only when they end an ordered-list number on that same line. The decision uses the assembled paragraph, including image alt text and a notes line after the generated `> ` prefix. An inserted hard break resets the line column and the digit count.

Do not edit HLD 03, HLD 04, or the Word field regions of HLD 10.

## Risk routing

- Public API of a published crate. The 0.15.0 additions and behavior changes are listed below. No 0.14.x patch break. Archive measurement stays with F-X275.
- WASM or PyO3. `index` is PyO3. No new WASM export. Workspace tests keep `--exclude rdocx-py --exclude rpptx-py`.
- Layout and text. Deterministic bundled fonts. Do not re-record a baseline.
- Compatibility and collection order. `index` follows python-pptx 1.0.2 element identity.
- New projection. Cover absence, `move`, two wrappers, and a foreign object.
- Cached live path. A held slide, a second wrapper, and a fresh `slides[i]` agree after `move`.
- Contributed PR. PR 438 is open, unmerged, and not acceptance. Every issue row is an A-ID.
- Several consumers. The map above is the check.

Not matched: a parser or serializer grammar change, a new trait or file, theme colour, unit truncation, a feature flag, a performance budget, and a raw-XML input API.

## Hash harness

Expected unchanged. Saved OPC bytes, PDF, and PNG stay the same. The integrated union remains hash 49 and pixels 7. This story does not move that baseline.

## Implementation checklist

- [x] Add `Display` and `authoring_table_height` without a new file.
- [x] Return the duplicate tuple from `autofit_inventory` and append it once per leaf in `fit_report`.
- [x] Return the notes reason vec and copy it once onto `unsupported` and `update.diagnostics`.
- [x] Publish a non-empty unsupported cause once at both `:4060` and `:4121`.
- [x] Apply the page-growth exception only when `width_fits` is true and `fits` is false.
- [x] Change the CLI line, exit rule, and JSON `reason`.
- [x] Add `SlideCollection.index` and the stub.
- [x] Add the A1 through A5 fixtures to the existing carriers.
- [x] Update only the HLD sentences listed above.
- [x] Leave README measurements, release scripts, `AS_BUILT.md`, and `SPRINT_TRACKER.md` untouched.

## Final sprint acceptance checklist

- [ ] Final integrated workspace and CLI union.
- [ ] Hash49 and pixels7, with expected unchanged F-X276 output.
- [ ] Named native, PDF, PNG and LibreOffice riders.
- [ ] Installed Python3.9 and3.12, typing, stubs and archive qualification.
- [ ] Hosted CI and full public Markdown export benchmark.

## Final Markdown measurement interpretation

The full public Markdown export benchmark in the preceding checklist uses the
Root-approved seven-route recipe of2026-11-02. This records the existing
measurement decision, not a new product choice, performance budget or executed
gate. Original benchmark provenance remains unlocated. Do not claim that this
new technical recipe recovered the original benchmark. Plan status stays
approved, final acceptance stays pending and prior review caps do not reset.

Measure Rust default export, options export and save, Python string export,
owned export and save, and CLI conversion. Use21independently generated,
byte-frozen inputs: plain1/8/32slides, digit-prefix lengths1024/4096/16384
with period/parenthesis and1/64runs, mixed1/8/32slides and chart1/8/32slides.
Chart roles are title Quarter, categories Q1/Q2, series Zeta region and values
1150/1190.5. The mixed cases retain notes, hidden slides, images, tables,
whitespace, line breaks and split punctuation. Default configuration reaches
all seven routes. Mixed cases also exercise admissible notes, hidden selection
and detached-image options. Preserve actual defaults and detached string-export
refusal as separate untimed controls. This yields213route/configuration rows,
2130paired samples and84separate untimed chart-control observations.

Each SDK model is opened fresh for five exports, with the first cold after open
and four subsequent calls on that model. Preflight uses different instances.
Each CLI sample includes a fresh process, open, export and publication. Alternate
baseline/current order. Retain all raw samples and min/median/max. No warmup is
discarded. Hashing, result inspection and independent Markdown structure checks
are outside the measured call. Save includes actual publication. Run without
concurrent compilation or other measurement work.

Use actual release artifacts from baseline9a3b2960ccda7c3208ecb9b753c7019b64c086f5
and the final integrated source, with pinned compiler, features, lock, profile,
interpreter, host, fixtures and imported/binary identities. Authenticate real
successful build receipts. Validate semantic/default/refusal behavior through
a pinned independent CommonMark XML producer with tables and strikethrough.
Retain changed escaping and added chart work as different output classes.
Plain output is equivalent only after actual equality. Missing cases, failed
semantics, source/artifact drift or incomplete results leave the gate pending.
Report timing and growth without inventing an absolute cap or speed ratio.
Word, HTML and WASM throughput remain excluded. All existing native, installed,
hash/pixel, hosted and other final riders remain independently mandatory.

The immutable preparation manifest SHA256 is
`d243506bf42ec0e8ff83dee5b5f7cadd0fea28af8444c030604dc3a89e5af44d`.
The Root recipe-decision SHA256 is
`63c7fd140a2033419cd345d7efc3749be73168b796824b357af29a5082f27b19`.
Their local preparation lives at
/private/tmp/rdocx-cross-review/S96_MARKDOWN_BENCHMARK_PREPARATION.
Root static inspection found no actionable preparation-source finding, but no
Rust consumer build, fixture production or benchmark execution was performed.
Tool-only validator qualification is recorded below. Final source/artifact
refresh and actual product runtime results remain required. This clarification
changes no HLD product intent.

The qualified independent parser is cmark-gfm0.29.0.gfm.13. Its executable
SHA256 is
`f05a86b6961401e2f732f91b7864a43fe6d61598c67b96a17e5d4eda51eb94b2`.
Use its recorded absolute executable with `--to xml --extension table
--extension strikethrough <input>`. The eight actual tool controls pass. This
is parser-tool evidence only, with no product fixture or benchmark execution.
Authenticate its retained qualification/build/source receipts and binary before
final use. Do not rebuild or rerun unchanged qualification merely for metadata.

Final execution must import the accepted table-cell collector successor from
`table-cell-collector-v2`, utility SHA256
`a4e0a492a1f13b08b383bebc7cebc44fdbcf58bbc157a4d959385bcefb54109b`,
alongside the unchanged prepared companion scripts. Its independent final2
review is clean. Sixteen actual tiny Markdown controls reproduce the old
15-cell false acceptance and prove the successor counts real table cells only.
The existing16-cell threshold, selection/default behavior and image checks stay
unchanged. Bind the actual imported utility, validator, companions and execution
context digests in the final artifact envelope. Keep the immutable original
manifest/context untouched. This tool repair is not product semantic acceptance,
timing, a new performance limit or a reset of the feature review history.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision | Status |
|---|---|---|---|
| Q-FX276-1 | A5 | `index` maps a removed or stale handle to `ValueError`. Other accessors keep existing errors, including causal `StaleElementError` for removed handles and numeric-index `IndexError`. The defensive absent-owner branch remains unchanged, not a claimed public runtime case | Answered. python-pptx 1.0.2 and `slide.rs:438-450` |
| Q-FX276-2 | A1 | `PlaceholderKey::matches` stays index-first. Type fallback runs only when no key matches | Answered. `placeholder.rs:289-296` |
| Q-FX276-3 | A4 | The spacing rule applies to alt text and the non-decimal label | Answered. `html.rs:473-478` and `:657-664` |
| Q-FX276-4 | A1 | The donor `Display` sentences are the public text | Answered. `pulls/pr-438.diff:499-550` |
| Q-FX276-5 | A1 | The absent-idx several-candidate sentence is `notes placeholder has no idx, and {count} notes-master placeholders have its type` | Answered. `placeholder.rs:292-295` |

No blocking row remains. Q-S96-07 is answered globally: bundled default, explicit system fallback and caller fonts highest. It adds no F-X276 product choice.

## Compatibility for 0.17.0 / 0.15.0

Listed behavior changes are the width-qualified grown-table overflow rule, fit exit 1 for status `unsupported`, one published cause on every non-empty unsupported frame, the `Display` sentences, Markdown spacing, and `index` raising `ValueError` for its rejected inputs while other accessors keep their existing errors. The defensive absent-owner `IndexError` branch is unchanged and a removed accessor remains causally stale. Width failure and unsupported overflow stay true. Listed additions are `SlideCollection.index`, `Display`, and `authoring_table_height`. No existing argument, default, or `kind` string changes.

F-X275's approved plan at `.claude/plans/F-X275-design.md:172` lists issues 423, 424, 425, 426, 427, 428, 429, and 430. It does not name F-X276. F-X276 does not revise that plan, and it does not edit `CHANGELOG.md` or the release scripts. There is no required order between the two stories. When `/release-notes` is later run for a reviewed sprint that contains this story, the family notes name Issue 431, `SlideCollection.index`, the `Display` sentences, the fit exit rule, and the Markdown spacing rule. A draft built only from the F-X275 inventory is regenerated at that point. Publication stays a separate exact-main approval.

## Tracker text for the integrator

```markdown
### F-X276, Report rpptx fit, plain Markdown, and slides.index

**Sprint**: S96
**Size**: M
**Depends on**: none
**Test gate**: regression
**Spec reference**: `08-rendering-spec.md`, "Tables" and "Read-only fit and count consumers". `10-bindings-spec.md`, SlideCollection, validate message, PowerPoint fit, and Presentation semantic text export.

Finish every row of Issue 431. Fit and fit_report cover the whole deck, name each unassessable frame and why, measure a duplicate shape id without rewriting it, and resolve a notes placeholder idx by the single master placeholder of its type during fit measurement only. A grown table row that stays on the slide is not an overflow when its stored width fits. A width failure stays an overflow. A table that leaves the slide still is. The page test uses the painted growth and shape_transform, including rotation and flip. Plain fit text is one unique line per frame, with shape, cell, and plain indexes. Markdown keeps a single inner space across runs and escapes punctuation only where the assembled line needs it. slides.index matches python-pptx 1.0.2 element identity: live zero-based position, ValueError for another presentation, a structural twin, a non-slide, and a removed slide. Python ValidationIssue.message stays equal to the CLI line. No README measurement edit, no hash move, no Word or package-validation scope.
```

`BACKLOG.md` and `CURRENT_SPRINT.md` already carry the pending row with owner `-` until claim. `SPRINT_PLAN.md` records the issue, PR 438, the A-IDs, and that PR 438 is open, unmerged, and not acceptance. The design file is `.claude/plans/F-X276-design.md`, adopted from the `# F-X276` heading down with the review metadata above dropped. This adoption does not change F-X271 through F-X275.

## Claim prerequisites

The claim prerequisites are not met yet.

1. The affected Codex final recheck has completed and accepted V4 as an adoption basis. Peer preparation passes are complete. There is no peer 4. A clean preparation audit is not the later implementation microscope.
2. Registration is complete in HLD 14, the backlog row, the current-sprint row and the sprint-plan section. This adoption writes only `.claude/plans/F-X276-design.md` from the `# F-X276` heading down and does not write those four registration files again.
3. The plan stays `draft` after adoption. It becomes `approved` only after semantic closure, any material product questions and the global batch approval. Spec reference, test plan, risk routing, and HLD impact stay non-empty.
4. The canonical worktree is `sprint/s96` and clean except the approved plan. The branch is at `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd` with the preserved dirty source, so a clean Base is still required.
5. F-X276 is `pending` in both `CURRENT_SPRINT.md` and `BACKLOG.md`.
6. Claim runs from that clean worktree, owner `grok`, branch `work/f-x276-grok`, worktree path given and absent.
7. `work/f-x276-grok` does not exist locally or on the remote.
8. The clean-base gate in item 4 is met. This story must not edit `crates/rdocx-py/tests/test_core.py`. F-X272 is not a product prerequisite.
9. The worker's first command is `/start-feature F-X276 --claimed`.
10. Implementation needs a new lease after that claim. This adoption does not start it.

## Evidence pins

The preparation turn re-hashed the files below. The ten hashes carried since V1 still match. `text.rs` and `layout.rs` are new pins for the lines this recheck cites. The rest of the dirty tree was not hashed and is not frozen.

| Fact | Value |
|---|---|
| V3 reviewed by Codex final 1 | SHA-256 `ef910063cca0f3a99a7d515a468e0bc352ebd8ca5230b6a898e911c378151742` |
| Codex final 1 report | SHA-256 `fefde607cc165e8d298e5d05606895055df2de9ba0e807ffa48dff3507a13d2b` |
| V2 proposal | SHA-256 `417f3bc3c0587a3bf42c2c14a41323f8500a17fd8693f4eabf8d00d7f81149f3` |
| V1 proposal | SHA-256 `91e5d4fa3af00981781aabb0beaa2f438e3fdd65afdf3ccb6cafa5e6ec6ff643` |
| Issue 431 body | 21641 bytes, SHA-256 `992ea123f7818d87ed3ef5595365085de6171952eb620782782d2a2596be90bb` |
| PR 438 | Head `653053058a9a9c6647732179219f40ed2725caba`, base `24af95d3d94c25e95f61a75e7e21f18fa95b6363`, body SHA-256 `5746b1d907c961ed9998fc95b92d08b0fa0ffdaf71369f95b62cb9557cddef69` |
| API diff | 52049 bytes, SHA-256 `5fc3ea1513c798808f804c4822a12c824edf31a966eee2c46d201a47eaeecd91` |
| Historical branch and HEAD | `sprint/s96`, `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`, authenticated by the earlier passes. The preparation turn did not re-read `HEAD` |
| CI | Run `41212837717`. README doctests belong to the contributor patch. LibreOffice 404 and the rdocx `KeyError` match the main baseline. Green checks are not acceptance |

| File | SHA-256 |
|---|---|
| `crates/rpptx/src/lib.rs` | `23ee69f59227f73f617645c2bb2f67fbaefbb346339ba4855f17a0e323a9d2cd` |
| `crates/rpptx/src/html.rs` | `8391459f8125ef219bc1c001b02832572f5747a85667e18e07c990bb9408d01b` |
| `crates/rpptx-render/src/lib.rs` | `c5e2b92a323bd156f7307a7dd18df65dcbe6286132390ffe4fe673b8b5bba12c` |
| `crates/rpptx-render/src/text.rs` | `19557f9eb0c421675a8595dd96532cbe51d000961953283a082b58b53aa9f9f6` |
| `crates/rpptx-cli/src/commands.rs` | `93cbe5c0b1d402a733709a51840539324ecf0fe5f62b931eabc8f543eaeff8bb` |
| `crates/rpptx-cli/src/main.rs` | `3c595ae6dd5713cf7b0753a72f43921b11585394f120669fd0d6f2d9fab9317a` |
| `crates/rpptx-py/src/slide.rs` | `2d4d8432fd3cd66a35ce1eff9452da374facf9a9620c56aec3a7a89ecb183d4e` |
| `crates/rpptx-py/src/presentation.rs` | `8ea7008175106161a8f9cd5d8883eb6deeeb10408fb21dd139c42f123e6eba35` |
| `crates/rpptx-py/src/layout.rs` | `04c550124281ead1020882493b73ac3e3c02eb88d07fddfa201c5ace0675691c` |
| `crates/rpptx-py/python/rpptx/_rpptx.pyi` | `75486defda8bdf5060b465267b4df313df1360404026211609a2b70bfe70ba65` |
| `crates/rpptx-oxml/src/placeholder.rs` | `3eb8a2ad0622875d8ec5bd2ebe3e769718352ed9b23cf526aca67927c31f1521` |
| `docs/hld/08-rendering-spec.md` | `656064fcef1e130bf507b2463b40921548054e39b89e7bc3254cfd81b1058bcf` |
| `docs/hld/10-bindings-spec.md` | `c09cf197a66315737897aa597361ff2d32dfa1a109d8db67930924e985410ade` |

`git apply` was not re-run. Full Actions logs were not re-downloaded. No test, build, claim, or Git mutation ran in the preparation turn.
## Current measurement preparation successor, 2026-11-07

The original preparation manifest and recipe-decision bytes remain unavailable.
Their historical hashes above are retained. The versioned successor manifest
`b7384e9967a1030878d2b50f6477dcbc105b094df41d630b6bf3a6b45636e3c9`
is a new reconstruction of the same approved seven-route measurement recipe,
not recovered original evidence. Actual baseline9a3b and currentH8 Rust worker
and CLI release builds pass. The baseline Python wheel is built and authenticated.
The current PPT Python artifact retains its actual4cf producer and the bounded
unchanged production-source applicability assessment.

Actual untimed preflight found two measurement-tool premise errors. Literal NaN
is an existing chart-reader refusal, rather than a successful dense projection
containing a missing numeric cell. The corrected control removes point0 while
retaining ptCount2 and point1 at1190.5, matching the existing sparse-cache test.
All21 performance fixtures and the other five controls remain byte-identical.
The original literal-NaN input and failed preflight remain preserved separately.

The title predicate also confused hidden W2-TITLE-1 with visible W2-TITLE-10.
The versioned utility now compares complete numbered title identities. Its
SHA256 is `a44dae8057eb497ed3da2fa26ed421a0e76941ffa8b05846ee1a2b41eaf67047`.
The accepted table-cell counting rule and its16-cell threshold remain unchanged.
Actual old-helper RED, corrected valid-input GREEN and leaked-hidden/missing-visible
negative controls pass. The qualified cmark producer remains unchanged.

The actual successor handoff is
/private/tmp/rdocx-s96-prepared-integration-readiness/markdown-H8-artifact-successor-20261107/UNTIMED_ALL_READY_HANDOFF.json,
SHA256 `327e87110bd5fb400f04418f1da1028f844f95e61eccffa3bae6c983b1dfd682`.
It records84 actual chart controls and426 actual route/profile preflights over
213 rows, retaining original product exports and revalidating affected captured
W2 results through the corrected independent helper. Failed attempts remain
failed. No2130 timing samples have run and quiet admission remains unset.
The cold/four-warm sampling rule, workloads, publication boundaries, chart
semantics and uncapped timing/growth report are unchanged. Actual final-source
applicability, quiet observations and the complete benchmark remain due.
