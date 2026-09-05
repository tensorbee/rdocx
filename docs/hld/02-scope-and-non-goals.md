# 02, Scope and non-goals

This document decides whether something is in v1. `03-architecture.md` decides
which crate owns it.

## The shape of v1

One release, containing:

1. The `oxml-*` infrastructure extracted from rdocx, with rdocx migrated onto it
   and released as 0.3.0.
2. `rpptx` at feature parity with `python-pptx`, including charts.
3. PDF and PNG rendering of slides.
4. Rust crates, CLIs, WASM modules and Python wheels for both rdocx and rpptx.

There are no partial-feature interim releases. This was a deliberate choice and
its cost is recorded in `00-vision.md`.

## In scope for rpptx v1

### Presentation and slides

| Capability | Notes |
|---|---|
| `Presentation::new / open / from_bytes / save / to_bytes` | `new()` uses a bundled template |
| Slide collection, iteration, indexing, lookup by id | |
| `add_slide(layout)` | Synthesises placeholders, does not deep-copy |
| `remove_slide`, `move_slide`, `duplicate_slide` | Beyond python-pptx |
| Slide size get and set | |
| Slide masters and layouts, layout lookup by name | Read |
| Core, app and custom properties | Shared with rdocx via `oxml-core` |
| Notes slides | Read and write |
| Notes-master and handout-master header and footer settings | Native Rust read and write |
| Ordered presentation sections and slide membership | Native Rust read and write |
| Modern comments, authors and threaded replies | Native Rust read and write. Legacy comments remain opaque |
| Slide background, follow-master-background | |
| Hidden slides | Skipped when rendering, preserved on save |
| Modern package classes | Native Rust reads, preserves, inspects, and output-selects PPTX, PPTM, POTX, POTM, PPSX, and PPSM. Binary `.ppt` remains out of scope |
| OLE, ActiveX, and VBA executable payloads | Native Rust relationship-owned inventory, byte-exact extraction and replacement, and ownership-aware removal. Payloads remain opaque and are never executed. OLE renders from its stored preview image |
| OpenDocument Presentation interchange | Native Rust bounded read and deterministic write for slides, ordinary rectangles and text boxes, tables, embedded images, and speaker notes. Other safe content is reported through stable diagnostics |
| HTML slide content import | Native Rust bounded import of HTML5 documents and fragments into editable slides, explicitly positioned shapes, formatted text, tables, caller-supplied images, and links. Browser layout, scripting, fetching, and unsupported CSS remain diagnostic |
| PDF page content import | Native Rust bounded import into either one preserved full-slide graphic per page or an editable subset of text, raster images, paths, and URI links. Unsupported operators and font substitutions remain diagnostic |

### Shapes

| Capability | Notes |
|---|---|
| `add_textbox`, `add_picture`, `add_table`, `add_shape`, `add_connector`, `add_group_shape` | |
| Shape id, name, type, rotation | |
| Position and size, with placeholder inheritance | `Option`-returning plus an `effective_bounds` accessor |
| Fill, line, shadow | Fill and line full, shadow read-only |
| Adjustment values, `a:avLst` | |
| Click actions and hyperlinks | |
| Placeholders by index and by type | |
| Picture crop and intrinsic size | Via `oxml-media` |
| Image deduplication by content hash on insert | |
| SmartArt inspection, bounded node-text editing, and native rendering for six pinned layouts | Data, layout, style, colour, cached drawing, and relationship ownership are typed. The exact pinned list, hierarchy, cycle, relationship, matrix, and pyramid resources lower through the shared DrawingML engines. Unsupported algorithms and unmodelled XML remain preserved |

### Text

| Capability | Notes |
|---|---|
| Text frame, paragraphs, runs, line breaks | |
| Alignment, level, line spacing, space before and after | |
| Font: bold, italic, underline, strike, size, name, colour, caps, language | |
| Bullets: character, auto-number, none, size percent, colour | python-pptx has no bullet API. This is beyond parity |
| Margins, vertical anchor, word wrap, auto-size | |
| Nine-level list style inheritance | |

### Tables

Rows, columns, cells, cell text and text frames, cell fill and margins,
`merge` and `split`, merge-origin and span queries, and the banding flags.

### Charts

`add_chart` with bar, line, pie, scatter, area, doughnut and radar plots.
Series, categories, axes, gridlines, legend, data labels and number formats.
Each chart writes its own part, its relationship, and an embedded workbook.

### Rendering

Preset and custom geometry, solid, gradient, pattern and picture fills, lines
with dash, cap, join and arrowheads, rotation, flips and nested groups, the full
inheritance chain, shape text with anchoring, insets, wrap, bullets and stored
autofit, tables, connectors, hyperlinks, slide-number fields and backgrounds.

### Distribution

`rpptx` and `rdocx` as crates, `rpptx-cli` and `rdocx-cli`, `rpptx-wasm` and a
rewritten `rdocx-wasm`, and `rdocx-py` and `rpptx-py` wheels on PyPI.

## Explicitly not in v1

Each of these is **preserved verbatim on round-trip**. Nothing in this list
causes data loss, only reduced fidelity when rendering.

| Area | v1 behaviour |
|---|---|
| Animations, transitions, `p:timing` | Preserved, irrelevant to static rendering |
| Unsupported SmartArt algorithms and unmodelled `dgm:` content | Preserved. Rendering uses the drawing fallback part, else its cached picture, else its bounding box. The six exact pinned native layouts are handled before this fallback |
| OLE objects, ActiveX | Preserved, rendered as the stored preview image |
| Video and audio | Preserved, rendered as the poster frame |
| 3-D, `a:scene3d` and `a:sp3d` | Preserved, rendered flat |
| Blur on shadows, glow, reflection, soft edges | Shadow renders as a hard offset silhouette. The rest are dropped |
| WordArt text warp, `a:prstTxWarp` | Rendered as plain unwarped text |
| EMF and WMF images | Outline placeholder. Writing an EMF interpreter is out of scope |
| `eaVert` upright stacked CJK | Falls back to rotated vertical text |
| `mongolianVert` upright stacking | Falls back to rotated vertical-270 text |
| `wordArtVert` and `wordArtVertRtl` glyph stacking | Fall back to rotated vertical and vertical-270 text respectively |
| Gradient stop alpha | Stop colour composited, alpha dropped |
| Justified text inside shapes | Treated as left-aligned |
| Table cell text autofit | Not attempted |
| Legacy comments, ink, `p:contentPart` | Preserved |

Every one of these records a diagnostic, surfaced by `rpptx inspect --json` and
by the render API, so a user can tell approximation from fidelity.

## Non-goals, permanently

**`oxml-sml` is not a spreadsheet library.** It writes one worksheet with the
cells a chart needs. It is not a foundation for an `rxlsx` and should not grow
into one without a separate decision.

**Drop-in `python-docx` and `python-pptx` compatibility is not promised.** Those
libraries' real-world surface is inseparable from lxml, and a large fraction of
production code reaches through `._p`, `._r` and `qn()`. Source compatibility
is bounded to the completed public binding surface. The rdocx gate pins the
seventeen executable python-docx 1.2.0 documentation examples that fit the S33
API to stable tagged sources. Sixteen change only their import namespace. The
Quickstart held-row example re-fetches the row through the public document path
before its second cell assignment because the first structural text replacement
intentionally stales every pre-write handle under strict global revision.
Touching a private lxml-shaped attribute raises a clear error naming the
equivalent, rather than failing five frames away.

**Not a PowerPoint clone.** The renderer targets business decks built from
stock or corporate templates. Decks that lean on 3-D, heavy effects or WordArt
will render legibly but not faithfully, and will say so.

## Beyond v1

v1 shipped. This section records what changed after it, and it is the only place
a v1 non-goal may be superseded. A non-goal not named here still stands.

The shape of the roadmap is in `14-development-backlog.md`, M14 through M22. The
principle behind it: v1 proved the model and the renderer can live in one
codebase, which is the thing no other library in Python or Rust has. Everything
after v1 leans on that rather than away from it.

Modern Transitional OfficeMath is part of the post-v1 Word authoring surface.
Native Rust callers can inspect, mutate, and author inline and display
equations through the normalized Word model. Legacy Equation Editor, OLE, and
pre-OOXML equation payloads remain opaque under the permanent legacy-format
boundary.

The native Rust Word facade inventories relationship-owned OLE objects,
ActiveX controls, and VBA projects without decoding or executing their
payloads. Callers can extract and replace exact bytes or remove one validated
owner while shared targets and unrelated producer content survive. Package and
VBA signature evidence is either retained as explicitly invalidated evidence or
removed through an explicit mutation policy. Python, WASM, CLI, binary `.doc`,
payload decoding, and execution remain outside this surface.

Native Rust callers can also import and export the supported normalized
equation subset as Presentation MathML or LaTeX. Lossy format and OfficeMath
properties remain visible through ordered diagnostics. Python, WASM, CLI,
legacy equation formats, and a second conversion model remain outside this
surface.

Modern OOXML legacy form fields and glossary entries are part of the post-v1
native Word surface. Native Rust callers can inventory supported form fields
across internal Word stories, update their typed values, and replace existing
AutoText and building-block entries. Binary `.doc` input, field execution,
implicit entry expansion, new glossary authoring, and additional binding
surfaces remain outside this scope.

### Superseded

| v1 position | Superseded by | Why it changed |
|---|---|---|
| Charts are a PowerPoint capability | **M15** | `oxml-chart` now owns the format-neutral engine. `rpptx-chart` remains a deprecated compatibility shim |
| Animations, transitions, and `p:timing` are preserved but never executed | **M21** | The static renderer and corpus now provide the geometry, timing-independent frame state, and output backends needed to add bounded timeline execution without making it a prerequisite for ordinary slide rendering |
| Video and audio are preservation-only poster content | **M21** | The native package model edits embedded or linked media, and the additive media-aware timeline path returns poster or labelled fallback output with synchronized playback state while static rendering remains poster-only |

These entries are decisions, not corrections. The v1 positions were right when
they were written.

Bounded timeline execution is additive to ordinary presentation rendering.
Native callers select a slide-local elapsed time and click count, then receive
one deterministic page frame, the evaluated frame state, and ordered
diagnostics. Supported entrance, exit, emphasis, motion, transition, and
explicit-name morph cases execute without making timing a prerequisite for
static rendering. Unsupported or malformed timing stays visible through
diagnostics and does not acquire guessed behavior.

Audio and video editing is a package operation, not playback or decoding.
Native callers can inspect and atomically mutate media sources, poster images,
relationships, and bounded playback settings. Unknown safe payloads remain
extractable and diagnostic. Static rendering continues to use the poster.
The additive deterministic media timeline facade returns the ordinary timeline
frame with ordered audio and video playback states. A valid poster uses the
existing image path. An unresolved poster can become a deterministic labelled
fallback or a closed error according to explicit caller policy. Media payloads
remain outside renderer image input, and no codec is decoded.

Native callers can export explicit slide segments as deterministic animated
GIF or Motion JPEG AVI. Each segment declares its duration, fixed click count,
and optional outgoing transition slide. Sampling uses bounded integer
millisecond timestamps at a declared frame rate, reuses one prepared package,
resolver, and media context, and renders one opaque frame at a time. GIF loop
metadata and cumulative centisecond timing are explicit. AVI quality, frame
rate, dimensions, duration, chunks, and index are deterministic. Fixed frame,
pixel, and byte caps fail closed, and no system codec or subprocess is used.

### Conditional expansion

`oxml-sml` remains chart-workbook support rather than a spreadsheet library.
M19 may supersede that position only if F-184 finds a material gap still exists
in the Rust ecosystem at S70. A basic reader, writer, or formula evaluator is
not enough. The required gap is one loss-aware lifecycle covering advanced
editing, calculation, local pivot refresh, selected Power Query execution,
Office Scripts-compatible automation, and rendering. If a credible maintained
crate provides that boundary by then, M19 is archived rather than implemented.

### Still non-goals, and still permanent

- **Not a PowerPoint clone, and not a Word clone.** The renderer targets
  business documents. Decks and documents that lean on 3-D, heavy effects or
  WordArt render legibly and say so.
- **Drop-in `python-docx` and `python-pptx` compatibility is not promised.**
  Unchanged and for the unchanged reason: their real surface is inseparable from
  lxml.
- **EMF and WMF interpretation.** Still an outline placeholder. M18 adds
  formats, and this is not one of them.
- **Legacy binary Office formats.** Binary `.doc`, `.xls`, and `.ppt`, Word 2003
  XML, and equivalent pre-OOXML authoring surfaces are not scheduled. They do
  not share the OOXML package, model, preservation, or rendering foundations,
  and adding them would create separate legacy engines rather than deepen the
  current product.
- **Universal Excel service compatibility.** If M19 proceeds, it executes
  worksheet and table-backed pivots, a declared Power Query M and connector
  subset, and an explicitly versioned Office Scripts-compatible API. It
  preserves and reports unsupported OLAP and Power Pivot execution, proprietary
  and tenant-bound connectors, custom functions, Python cells, VBA, XLM, and
  Microsoft-hosted storage or automation services without claiming to run them.

### The WASM packages are deliberately unpublished

`@tensorbee/rdocx-wasm` and `@tensorbee/rpptx-wasm` are built, optimised,
packed as bundler tarballs and install-tested on every pull request. **They are
not published to npm, and npm publication is not authorised.**

This is enforced rather than intended. The WASM CI job is asserted to contain
none of `npm publish`, `wasm-pack publish`, `npm login`, `npm adduser`,
`npm token`, `NODE_AUTH_TOKEN`, `NPM_TOKEN`, `--registry`, `id-token:`,
`git tag` or `gh release`, so a step that could publish cannot be added without
failing the release preflight.

Both crates are `publish = false` for crates.io and inherit their Rust family's
version. That inheritance is harmless while nothing ships, and it is the only
thing that would need revisiting on the day npm publication is authorised.
F-X030 was filed against that inheritance and archived once this position was
confirmed, and its entry records what the work would be if the position changes.

### Deliberately not scheduled

Named so a reader knows they were considered rather than missed.

- **Legacy binary `.doc`, `.xls` and `.ppt`.** Each is a compound-file format
  with no relation to OOXML, and each would require a separate legacy engine.
  They remain excluded after the modern Office depth milestones complete.
- **A collaborative editing server.** Out of the shape of a library.

## The measurable bar

For a business deck built from a stock Office template, with title and content
slides, bullets, tables, images, theme colours and a gradient title bar, a
150 dpi PNG should be indistinguishable from PowerPoint's own export at a
glance: text baselines within about one point, shape edges exact, colours exact.

CI compares the pinned 50-deck corpus with LibreOffice's render and records at
least 0.95 SSIM on at least 80 percent of slides as a trend reference. The hard
automatic gate requires every slide to render without panic, missing output,
dimension mismatch, or a dropped bounded shape. LibreOffice is the CI oracle
only because PowerPoint is not scriptable on runners, so SSIM regressions are
review-required rather than automatic failures. A pinned native PowerPoint
representative review is the hard manual fidelity gate.
