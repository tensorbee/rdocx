# 12, Testing strategy

## Test taxonomy

Six categories. Every story's design plan picks the applicable ones and names
exactly one as its test gate.

| Category | Purpose | Where |
|---|---|---|
| `unit` | Pure logic, no I/O | `crates/<crate>/src/*.rs` under `#[cfg(test)]` |
| `integration` | Multi-crate behaviour through the public API | `crates/<crate>/tests/` |
| `regression` | Locks down one named past failure | `crates/rdocx/tests/regression_test.rs` and the rpptx equivalent |
| `round-trip` | Parse, serialise, reparse, compare | corpus-driven |
| `golden` | Byte or pixel comparison against a recorded baseline | the hash harness |
| `differential` | Compared against an external oracle | LibreOffice for renders, python-docx and python-pptx for the bindings |

The existing repository convention is preserved: **no binary fixture files.**
Fixtures are constructed in code, including hand-assembled PNG and JPEG headers
with precomputed CRCs. It keeps the `.crate` payload small and the diffs
readable. The corpus in the next section is the one deliberate exception, and it
lives outside the published crates.

Regression tests are named as sentences describing the failure they prevent, so
a reintroduction is obvious from the test name alone rather than from a diff.
The existing file is the model: `zero_column_tables_do_not_panic`,
`saving_is_reproducible`.

## The hash harness

The single highest-value mechanism in the plan is
`scripts/hash_harness.py --check`. It deletes the expected generated outputs,
runs `generate_all_samples`, and records the flushed `word/document.xml`,
`word/styles.xml`, and `word/numbering.xml` state plus the page-one PNG for each
of the seven samples. PNGs are rendered at 150 dpi through the deterministic
font path.

The sorted `scripts/hash_baseline.json` manifest has 28 entries. Each entry is
either a SHA-256 digest or JSON `null` when an optional XML part is absent.
Check mode reads the manifest without modifying it and reports added, removed,
and changed entries. Baseline writes require `--update --reason <text>`, and an
empty reason is rejected. Generated PNGs remain ignored under `samples/`.

It exists because the extraction changes unit conversion and text-shaping input
types, and both alter output **without failing to compile**. Structural
round-trip tests cannot see that class of defect.

Rules:

- Every PR in M1 through M6 gates on it.
- Baseline updates require a non-empty review reason.
- An intentional behavioural change lands as its own labelled commit with its
  expected delta stated in the message and reviewed.
- An unexplained delta blocks the merge.

## The golden-PNG gate

Specifically for the PDF coordinate-system change in M5. Render the `samples/`
corpus before and after and compare **pixels**. Do not compare PDF bytes: the
operator stream legitimately changes when the per-element Y flip becomes one
global CTM, so a byte comparison would be all noise.

## The deck corpus

Roughly 50 real `.pptx` files, stored outside the published crates and fetched
by a script. Deliberately spanning producers, because non-Microsoft writers are
where the parser assumptions break:

- PowerPoint 2016 and Microsoft 365
- Google Slides export
- Keynote export
- LibreOffice Impress
- A multi-master corporate template
- Decks containing SmartArt, charts, embedded video, and ink

Three gates run against it:

1. **Raw round-trip** (M8 entry): open and save with everything treated as
   opaque parts, assert byte-identical output. This proves the OPC layer and the
   corpus harness before any XML modelling exists.
2. **Modelled round-trip** (M8 exit): parse, serialise, reparse, compare
   structurally, and compare the saved package part by part.
3. **Opens without repair** (M8 and M11): every saved deck opened manually in
   PowerPoint once per milestone. Not automatable, and not skippable.

## The render fidelity gate

Roughly 50 decks rendered to PNG at 150 dpi and compared with
`libreoffice --convert-to png` using a perceptual diff.

**Target: at least 0.95 SSIM on at least 80 percent of slides, and 100 percent
of slides rendering without a panic or a dropped shape.**

The second half of that is the hard gate. The first is a trend line.

LibreOffice is the oracle only because PowerPoint is not scriptable on CI
runners, and LibreOffice has its own rendering bugs. **SSIM regressions are
therefore review-required, not automatic failures.** Spot-check against real
PowerPoint output once per milestone.

Stand this harness up in M10 alongside the first text rendering, not afterwards.

## New tests the extracted crates need

These crates have never seen a non-docx package, so the existing tests do not
cover the cases that matter now.

**`oxml-opc`**
- `with_main_part("ppt/presentation.xml", ...)` then `main_document_part()`
  resolves, and the package round-trips.
- A pptx-shaped package: package rels to `presentation.xml`, slide rels to
  `slide1.xml`. Assert
  `resolve_rel_target("/ppt/slides/slide1.xml", "../slideLayouts/slideLayout1.xml")`
  resolves correctly. The `..` traversal is currently exercised only by docx
  headers.
- Every `rel_types` constant is unique and well-formed. Cheap, and the only
  thing that catches a copy-paste typo among the new constants.
- Zip-slip: a part named `../../etc/passwd` and an absolute-path entry are
  normalised or rejected. The code handles it, nothing tests it, and the crate
  is about to become a public shared component.

**`oxml-core`**
- New unit round-trips: `Centipoints::from_pt(18.0).0 == 1800`,
  `Angle::from_degrees(90.0).0 == 5_400_000`,
  `Percent1000::from_percent(75.0).to_fraction() == 0.75`.
- Existing `Length`, `Twips` and `Emu` constructors have positive and negative
  truncation-pinning tests that move with the units into `oxml-core`.
- `xml_text` becomes public API, so add CDATA, mixed content, nested elements,
  unknown entities, and the `GeneralRef` split case.
- `AppProperties` parses a Word `app.xml` **and** a PowerPoint one, leaving the
  other format's fields `None`, and omits them on write.

**`oxml-media`**
- Sniffing every format from magic bytes, and **sniff beats extension**: a
  `.png` that is really a JPEG resolves to JPEG.
- DPI from PNG `pHYs` with unit 1 and unit 0, and from JPEG JFIF density units
  1 and 2, including a file with EXIF before the SOF.
- **A truncation loop per format**: `for n in 0..data.len()`, assert no panic.
  Cheap, and it catches every slice-index bug in one shot.
- The counter fix, named as a sentence:
  `next_image_name_uses_the_highest_existing_index_not_the_part_count`.

**`oxml-layout`**
- `Transform` composition order matches the PDF `cm` operator.
- `walk()` flattens nested groups and accumulates the transform correctly.
- `FontManager` with no fonts returns an error rather than panicking, and
  `--no-default-features` is in the CI matrix so the bundled-fonts-off path is
  actually exercised.

**`oxml-pdf`**
- `Path` with fill only, stroke only, and both, produces `f`, `S` and `B`, and
  `q`/`Q` counts balance. One assertion that catches the classic unbalanced
  graphics-state bug.
- **`Group` containing `Text` finds the font.** The regression test for the
  recursion hazard.
- `Group` containing `Image` registers the XObject.
- `Group` containing `LinkAnnotation` emits it with a transformed rectangle.
- Raster: a rotated rectangle at 72 dpi has a filled interior pixel and an empty
  corner. Deterministic, no golden files.

## Binding tests

The parity suites are worth more than any number of Rust-side assertions,
because the whole value proposition is compatibility:

- Write a document with `rdocx`, open it with `python-docx`, assert text, styles
  and tables survive. Then the reverse.
- The same for `rpptx` and `python-pptx`.

Both libraries are free CI dev dependencies.

Plus `mypy --strict` over a typing smoke file and `stubtest` against the stubs,
so hand-written stubs cannot drift.

## What CI runs

| Job | Command |
|---|---|
| test | `cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py` |
| clippy | `--workspace --all-targets --all-features -- -D warnings` |
| fmt | `cargo fmt --all -- --check` |
| doc | `cargo doc --workspace --no-deps` with `RUSTDOCFLAGS=-D warnings` |
| msrv | the pinned MSRV toolchain |
| supply-chain | `cargo-deny check` |
| no-default-features | `cargo test -p oxml-layout --no-default-features` |
| wasm | `cargo check --target wasm32-unknown-unknown` plus `wasm-pack test --node` |
| python | `maturin develop && pytest`, plus `mypy` and `stubtest` |
| packaging | `cargo publish --dry-run` plus a `.crate` size assertion |
| prose | the voice rules over tracked Markdown |

The `--exclude` on the binding crates is required, not cosmetic:
`pyo3/extension-module` tells the linker that Python symbols come from the host
interpreter, which is false for a test binary, and on Linux this is an
unresolved-symbol link failure that is easy to misdiagnose.

## Gaps being closed

Stated plainly, because they are why two shipped defects went unnoticed:

- **`rdocx-cli` has zero tests** despite being a published binary.
- **`rdocx-wasm` has zero tests and no CI job**, which is exactly why its
  part-dropping save path survived a release.
- **PDF and PNG output is only checked for non-emptiness**, so layout
  regressions are invisible. The hash harness closes this.
