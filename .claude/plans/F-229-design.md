# F-229, OfficeMath layout and PDF rendering

**Status**: completed
**Sprint**: S65
**Size**: M
**Depends on**: F-228

## Problem

F-228 will make OfficeMath typed and authorable, but the Word layout engine
currently converts ordinary paragraph runs into line items at
`crates/rdocx-layout/src/engine.rs:4602` and has no equation projection. The
shared positioned output already supports nested groups at
`crates/oxml-layout/src/output.rs:300`, so equation geometry does not need a
math-specific PDF writer. The existing inline group carrier is top-aligned,
however, and cannot report the ascent and descent required for an equation
baseline.

F-229 must measure and position the F-228 expression tree through the existing
font, line, page-frame, PDF, and raster boundaries. Fractions, scripts,
radicals, matrices, limits, n-ary operators, delimiters, and accents must affect
line height and baseline correctly. Unsupported or unrenderable content must
remain in the document and produce stable layout diagnostics rather than being
silently discarded.

## Spec reference

- `docs/hld/03-architecture.md`, "The dependency rule", "Why these seams", and "What stays put".
- `docs/hld/08-rendering-spec.md`, "The seam that makes this cheap", "Extending PositionedElement", "The PDF backend", "The rasteriser", and "The renderer's input".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability" and "Packaging".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", "The golden-PNG gate", "The Word render fidelity gate", and "What CI runs".
- `docs/hld/14-development-backlog.md`, "F-229, OfficeMath layout and PDF rendering".
- `docs/hld/15-build-and-toolchain.md`, the deterministic font rules.

## Approach

Add one approved private `crates/rdocx-layout/src/math.rs` module. It consumes
the F-228 `MathExpression` tree directly and returns one measured equation value
containing width, ascent, descent, and a backend-neutral `GroupElement`. It uses
the existing `FontManager` for every glyph and the existing positioned text,
line, path, and group values for output. No PDF or raster backend learns about
OfficeMath.

Extend the existing shared inline group carrier with an optional baseline:

```rust
InlineItem::Group {
    group: GroupElement,
    width: f64,
    height: f64,
    baseline: Option<f64>,
}
```

Carry the same value through `LineItem::Group`. `None` preserves the current
top-aligned chart and drawing behavior. `Some(ascent)` contributes the measured
ascent and descent to line metrics and positions the group at
`line_baseline_y - ascent`. This is one general baseline-aware group path rather
than a math-specific backend variant.

Measure and lower the supported expressions recursively:

- Runs shape through the effective math run font and size, defaulting to the
  bundled Caladea family through deterministic font resolution.
- Fractions center numerator and denominator around a rule with measured gaps.
- Scripts scale child size by a fixed reviewed factor and preserve base,
  superscript, subscript, and pre-script baselines.
- Radicals draw the root glyph and stretch its overbar over the measured base.
- Matrices compute per-column widths and per-row ascent and descent, then place
  cells on shared row baselines.
- Limits and n-ary operators use above or below placement only when the model
  requests it. Inline placement otherwise follows script metrics.
- Delimiters choose a glyph at the base size and apply bounded vertical scaling
  to cover the argument. Separators use the same measured height.
- Accents center the requested glyph or line over the base with a measured gap.

Integrate equations into the paragraph source-order projection added by F-228.
Inline equations participate in ordinary line breaking as indivisible groups.
Display equations occupy their own centered paragraph line and respect the
typed display justification. Unsupported characters use the normal missing
glyph diagnostics. Unsupported raw expression siblings remain preserved and
produce one stable source-path diagnostic when they affect visible content.

The golden uses a source-built one-page DOCX containing each supported construct
and Caladea-family math runs. Rust renders with bundled deterministic fonts.
Microsoft Word 16.104 build 16.104.25121423 exports the same source to PDF. The
test records exact text tokens, page size, and Word glyph bounding boxes. The
harness derives each expression's ink width and vertical bounds directly from
the Word and Rust PDFs, with fixed raster windows only for the delimiter and
accent that Poppler coalesces into one Word token. It keeps no second set of
normalized Rust truth. Aggregate and per-expression geometry use the 1.0 point
tolerance. A deterministic complete-page raster comparison uses 64 by 64 pixel
block luminance at 150 DPI with a 0.99 SSIM floor. Mutations move a baseline,
delimiter, operator, and an actual rendered OfficeMath group by more than the
declared 1.0 point tolerance.

Add the approved source-only `scripts/officemath_oracle_harness.py` and
`scripts/officemath_oracle_manifest.json`. The harness builds the DOCX from
source, verifies the exact Word and Poppler identities, and revalidates the text
manifest against caller-supplied Word PDF output. Binary oracle output remains
untracked.

Carry the F-228 document-wide defaults through the concrete optional
`LayoutInput.math_properties: Option<MathProperties>` field. The `rdocx`
document facade populates it from relationship-resolved settings. Existing
layout input literals explicitly use `None`, which preserves behavior when no
settings part or math properties are present. This is a pre-1.0 source-breaking
addition for callers that construct `LayoutInput` with a struct literal. It
adds no wrapper, trait, binding, WebAssembly entry point, or command-line
surface.

## Rejected alternatives

- Rendering equations directly in `oxml-pdf` would add a Word grammar
  dependency below the format-neutral backend.
- Flattening an equation into ordinary text would lose stacked structure,
  baselines, stretch, and matrix geometry.
- Treating a group as top-aligned would overstate ascent, erase descent, and
  move surrounding text.
- Adding a math-specific positioned element would increase every backend's case
  count when existing group, text, line, and path primitives are sufficient.
- Adding or recording a system-font baseline would make the golden depend on
  the host and is forbidden.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `baseline_aware_inline_groups_contribute_exact_ascent_and_descent` | The shared line item preserves old `None` behavior and positions a measured baseline group correctly. |
| unit | `supported_math_expressions_measure_and_lower_to_bounded_groups` | Every F-228 expression returns finite width, ascent, descent, and ordered children. |
| regression | `equation_descent_does_not_move_the_surrounding_text_baseline` | A deep denominator and subscript increase line descent without top-aligning adjacent text. |
| regression | `unsupported_math_content_is_diagnostic_and_remains_preserved` | Visible unsupported content reports a stable source path while save output retains its raw XML. |
| integration | `inline_and_display_equations_share_the_document_page_frame_and_pdf_backends` | Inline and display equations paginate with ordinary text and produce one normal layout result and PDF. |
| golden | `officemath_baselines_and_glyph_geometry_match_the_pinned_word_pdf_oracle` | Page size, tokens, baselines, and per-expression bounds match the pinned Word export within 1.0 point. |
| regression | `officemath_golden_rejects_baseline_delimiter_and_operator_mutations` | Deliberate geometry and pixel perturbations fail the declared predicate. |

The **test gate** is the backlog's golden gate:
`officemath_baselines_and_glyph_geometry_match_the_pinned_word_pdf_oracle`.
The test binds the exact Word build and source-built input, records the declared
tolerance, and uses deterministic bundled fonts on the Rust side. No binary
fixture enters the repository.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/08-rendering-spec.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`

## Risk routing

- Layout, pagination, line breaking, and text shaping. Use bundled deterministic
  fonts for every Rust baseline, keep system font discovery out of the golden,
  and make every baseline update deliberate and mutation-sensitive.
- Public API of published crates. The shared `InlineItem` and `LineItem`
  extension and `LayoutInput.math_properties` addition are pre-1.0 source API
  changes. Update affected struct literals and facade projection tests. Run
  rustdoc with warnings denied, `cargo publish --dry-run` for `oxml-layout` and
  `rdocx-layout`, and assert both package archives remain below 10 MiB.
- New module or file. Obtain explicit approval for
  `crates/rdocx-layout/src/math.rs` and the source-only oracle harness and
  manifest. Add no new crate, test binary, dependency, trait, generic
  parameter, builder, wrapper, or feature flag.
- External oracle. Bind the source-built DOCX and exported PDF to Microsoft Word
  16.104 build 16.104.25121423, state the 1.0 point geometry tolerance and the
  pixel metric, and prove the comparison rejects the intended perturbations.

## Hash harness

Expected unchanged. Existing samples contain no equations. Any sample PDF or
PNG delta blocks the story and does not authorize a baseline update.

## Implementation checklist

- [x] Add the approved private equation layout module.
- [x] Add optional baseline metrics to the shared inline and line group path.
- [x] Measure and lower every F-228 expression through shared primitives.
- [x] Integrate inline and display equations with paragraph line breaking and pagination.
- [x] Project optional document-wide math settings through `LayoutInput` and the `rdocx` document facade.
- [x] Emit stable diagnostics for visible unsupported or unrenderable content.
- [x] Add focused metrics, regression, integration, and mutation-sensitive golden tests.
- [x] Record the pinned Word PDF oracle evidence and declared tolerance.
- [x] Add the approved source-only oracle harness and text manifest.
- [x] Run deterministic font, archive, rustdoc, focused, and full verification gates.
- [x] Update exactly the listed HLD files.

## Open questions

Resolved for S65. The new layout module, shared optional baseline metrics,
bundled Caladea fallback, source-only oracle harness and manifest, Word 16.104
build 16.104.25121423, Poppler 26.01.0, 150 DPI, 1.0 point geometry tolerance,
and 1.01 point negative perturbation are approved. Tagged-PDF math semantics
remain outside F-229. The optional concrete `LayoutInput.math_properties`
field, the `rdocx` document projection, affected literals and tests, pre-1.0
source impact, HLD 10 coverage, rustdoc gate, and package dry runs are approved
as a contract clarification required to carry F-228 global settings into
F-229 layout.
