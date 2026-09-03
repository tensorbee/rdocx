# Current Sprint, S65

**Milestone**: M22 Word depth.

**Goal**: author, convert, lay out, and render modern Word equations. The typed
OfficeMath model preserves unsupported siblings and remains independent of
legacy Equation Editor objects, while supported expressions share the existing
font, layout, PDF, and facade boundaries.

## Spec references

- `docs/hld/02-scope-and-non-goals.md`, for the modern OOXML boundary and the
  permanent exclusion of legacy document-format compatibility.
- `docs/hld/03-architecture.md`, for WordprocessingML grammar ownership, the
  single document facade, dependency direction, and shared layout projection.
- `docs/hld/04-opc-and-packaging.md`, for namespace-aware parsing,
  schema-ordered serialization, and verbatim preservation of unmodelled XML.
- `docs/hld/08-rendering-spec.md`, for the shared font, page-frame, PDF, and
  raster boundaries used to measure equation baselines and glyph geometry.
- `docs/hld/10-bindings-spec.md`, for additive native Word facade rules and the
  requirement that Python, WASM, and CLI exposure remain explicit.
- `docs/hld/12-testing-strategy.md`, for round-trip, golden, differential,
  deterministic-font, and pinned external-oracle evidence.
- `docs/hld/14-development-backlog.md`, for the F-228 through F-230 contracts,
  dependency order, acceptance gates, and the M22 completion boundary.

## The wave

| F-ID | Title | Size | Status | Owner |
|------|-------|------|--------|-------|
| F-228 | OfficeMath model and authoring | L | done | - |
| F-229 | OfficeMath layout and PDF rendering | M | in-progress | codex |
| F-230 | MathML and LaTeX conversion | M | in-progress | codex |

## Sequencing note

Rows are listed in dependency order, not F-ID order.

F-228 establishes the typed, schema-ordered OfficeMath model and must complete
before either consumer begins. F-229 and F-230 both depend only on F-228, so
layout and rendering can proceed independently from conversion once that model
is reviewed and integrated. Their combined sprint review proves that authoring,
conversion, layout, and rendering use one normalized equation structure without
weakening unsupported-sibling preservation.

## Definition of done for this sprint

- The source-built equation corpus parses, mutates, saves, and reopens typed
  runs, fractions, scripts, radicals, matrices, limits, n-ary operators,
  delimiters, accents, and equation properties in schema order.
- Supported OfficeMath content remains editable while unsupported sibling XML
  is preserved verbatim, and legacy Equation Editor objects remain outside the
  typed model.
- Supported equations use the shared deterministic font and page-frame path for
  baseline, stretch, delimiter, and operator sizing.
- Equation baselines and glyph geometry match the pinned Word PDF oracle within
  an explicit calibrated tolerance and a mutation-sensitive golden gate.
- Supported source-built expressions preserve their normalized expression tree
  through both MathML and LaTeX import and export.
- Conversion loss is reported through stable diagnostics rather than silently
  substituting unsupported constructs.
- Full verification passes with every deterministic hash explained, every
  package archive below 10 MiB, and the bounded sprint review clean.
