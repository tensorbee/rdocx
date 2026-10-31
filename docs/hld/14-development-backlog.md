# 14, Development backlog

Solo-developer build plan. Ordered by dependency, biased toward small,
incrementally-testable slices so something verifiable lands every few days.

## How to read this

- **Milestones (M1, M2, ...)**, each ends with a concrete, testable gate. Pause
  at any milestone boundary and the workspace is coherent.
- **Stories (F-001, F-002, ...)**, each sized for a solo dev. `S = 1d`,
  `M = 2-3d`, `L = 4-5d`, `XL = split me`.
- **Depends on**, hard dependencies. If unstated, the story can start as soon as
  its milestone begins.
- **Test gate**, the smallest test that proves the story works. Every story has
  exactly one. Nothing merges without it.
- `F-X###` marks cross-cutting work belonging to no milestone.

## Velocity assumption

### v1, as planned and as delivered

The v1 backlog was **162 stories**, roughly **408 developer-days**, forecast at
17 to 18 months solo. It delivered as **182 stories** across 43 sprints, the
extra 20 being cross-cutting `F-X###` work that no milestone predicted: an
external contribution, four releases, two dependency events and the defect
follow-ups each of those filed.

The forecast held at the story level and failed at the sprint level. Sprints
came in far under their estimates whenever a story arrived with its cause
already written up by the sprint that filed it, and the escalation record
carries the variance for each.

### Post-v1, M14 through M24

M14 through M22 delivered 72 non-spreadsheet stories. M23 and M24 add 71 Word
stories, roughly 331 developer-days, before the conditional 21-story M19
spreadsheet programme. Four S70 cross-cutting stories add roughly 12 days for
the confirmed Issue 67 closure and the three independently measured Issue 69
performance corrections.

F-X137 through F-X171 form a contribution and issue repair programme across
S76 through S85. The 2 October intake had 31 open PRs and 8 open issues.
Issue 264 joined on 2 October after that intake and is included in S84.
F-X172 through F-X174 address Issue 266's unified family release in S86.
The new tolerance, baseline, revision and accepted-view work continues after
the original S82 acceptance wave. These are cross-cutting stories, separate
from the 47 planned M24 feature stories. Their shared files and full issue
contracts require integrated gates at each sprint boundary.

M23 closes the five-document from-scratch business-document boundary. M24 then
classifies and closes the broader modern DOCX authoring surface before M19 may
begin. The spreadsheet programme remains a business decision and proceeds only
if F-184 confirms a material gap in the Rust ecosystem at S103.

The stopping and compression choices are:

- **Stop after M23.** S73 can generate the five private reference documents
  from `Document::new()` through public modeled APIs, with no base template,
  raw OOXML, or LibreOffice field-update pass.
- **Stop after M24.** S101 provides the complete modern DOCX authoring boundary.
  Every in-scope feature is authorable, readable, mutable, round-trip safe,
  rendered where applicable, and classified across the public bindings.
- **Archive M19 at its decision gate.** F-184 may still find that the advanced
  spreadsheet lifecycle no longer represents a material ecosystem gap.
- **Parallelise dependency-independent stories.** Each larger sprint contains
  explicit waves. The sprint boundary remains one integrated verification and
  review result.

---

## Milestone 1, Preparation and safety net (about 2 weeks)

**Goal**: rdocx behaves identically to today, but every future change is
measurable. Nothing has moved yet.

**Why first**: the extraction changes unit conversion and text-shaping inputs,
both of which alter output silently. Without a byte-level baseline, every later
step is unverifiable.

**End-of-milestone gate**: `cargo test --workspace` green, the hash harness
records a baseline that reproduces on a second machine, and `v0.4.1` is tagged.

### F-001, Deterministic font mode (M)
Add `FontManager::new_deterministic()` using bundled fonts only, bypassing
`load_system_fonts()` at `crates/rdocx-layout/src/font.rs:93`.
**Test gate**: rendering the same document twice with system fonts installed and
absent produces identical PNG bytes.

### F-002, rust-toolchain.toml (S)
Pin 1.97.1 with `rustfmt`, `clippy` and the `wasm32-unknown-unknown` target.
**Depends on**: none.
**Test gate**: `rustup show` reports the pinned channel in a clean clone, and
the MSRV job still pins 1.93 separately.

### F-003, Output-stability hash harness (L)
Digest each sample's `document.xml`, `styles.xml`, `numbering.xml` and page-one
PNG at 150 dpi in deterministic font mode. Store the baseline. Provide a
`--update` mode that requires an explicit reason string.
**Depends on**: F-001.
**Test gate**: the harness passes on an unmodified tree and fails when a
whitespace change is injected into a writer.

### F-004, Caladea licence and the false OFL claim (S)
Add `LICENSE-Caladea` plus NOTICE. Correct `bundled_fonts.rs:12`, which claims
all bundled fonts are SIL OFL when Caladea is Apache-2.0.
**Test gate**: a test asserts a licence file exists for every distinct font
family in `fonts/`.

### F-005, Fix the image counter (S)
`crates/rdocx/src/document.rs:135-138` counts matching parts instead of parsing
the maximum suffix.
**Test gate**: `next_image_name_uses_the_highest_existing_index_not_the_part_count`,
asserting `image1` + `image5` yields `image6`, and `image1,2,4` yields `image5`.

### F-006, Fix the JPEG standalone-marker walk (S)
`crates/rdocx-pdf/src/image.rs:51` treats every marker as length-bearing.
**Test gate**: a JPEG with an `RST` marker before the `SOF` still reports correct
dimensions, and a truncation loop over the file panics nowhere.

### F-007, Resolve core properties through the relationship (S)
Replace the hardcoded `/docProps/core.xml` lookup with
`rel_types::CORE_PROPERTIES`.
**Test gate**: a package storing core properties at a non-standard path round-trips
with its metadata intact.

### F-008, Non-consuming setter twins (M)
Add `set_*` siblings for every consuming builder in `paragraph.rs`, `run.rs`,
`table.rs`, with the builders delegating.
**Test gate**:
`doc.paragraph_mut(0).unwrap().add_run("text").set_bold(true)` compiles and has
the same effect as the builder form.

### F-009, Cache the layout result (M)
Separate `Mutex<Option<Arc<LayoutResult>>>` caches for normal and deterministic
font modes on `Document`, invalidated before public mutation and mutable access,
plus a cloned `layout_page` entry point. Caller-supplied font layouts remain
uncached.
**Test gate**: rendering all pages of a 20-page document performs exactly one
layout, asserted with a counter.

### F-010, Reserve crate names (S)
Publish `0.0.0` placeholders for every `oxml-*` and `rpptx*` name.
**Test gate**: `cargo info` resolves each name.

### F-011, Pin unit truncation behaviour (S)
Tests locking the current `as i64` truncation in every `Length`, `Twips` and
`Emu` constructor, before anyone changes it to rounding.
**Test gate**: the pinning tests, which must fail if truncation becomes rounding.

### F-012, Tag v0.4.1 (S)
A known-good published state immediately before the churn.
**Depends on**: F-003 through F-011.
**Test gate**: the release tag builds and publishes from a clean clone.

---

## Milestone 2, Shared infrastructure extraction (about 2 weeks)

**Goal**: `oxml-core` and `oxml-opc` exist as isolated staged crates, and no
released rdocx dependency or behaviour changes.

**End-of-milestone gate**: hash harness unchanged, and `OpcPackage` opens a real
`.pptx` in a test.

### F-013, Create oxml-core (M)
Copy `units.rs`, `raw_xml.rs`, `xml_text.rs`, the generic half of
`namespace.rs`, `core_properties.rs`, `error.rs`, plus
`crates/rdocx/src/length.rs`. Leave released rdocx consumers unchanged. Make
`xml_text` public. Consolidate the duplicate `local_name` and `get_attr`
helpers in the staged crate.
**Test gate**: the moved tests pass unchanged in their new crate.

### F-014, New unit types (M)
`Centipoints`, `Angle` in 60000ths of a degree, `Percent1000`, `Length::mm`.
**Depends on**: F-013.
**Test gate**: round-trip assertions including `Angle::from_degrees(90.0).0 == 5_400_000`.

### F-015, rdocx-oxml becomes a facade (S)
`rdocx-oxml` re-exports the shared modules, error surface and namespace helpers
from published `oxml-core` 0.1.2. Existing public paths and all internal call
sites remain source compatible. `Cargo.lock` records the one-way dependency.
**Depends on**: F-013, F-X005.
**Test gate**: the crate-local diff changes only `lib.rs`, `namespace.rs` and
`Cargo.toml` plus five deletions. The workspace tests and hash harness pass.

### F-016, Length re-export (S)
Delete `crates/rdocx/src/length.rs`, re-export from `oxml-core`.
**Depends on**: F-013, F-X005.
**Test gate**: workspace compiles with no call-site changes.

### F-017, App and custom properties (M)
`AppProperties` as a union struct with `Option` fields, plus `CustomProperties`.
Neither exists today.
**Depends on**: F-013.
**Test gate**: a Word `app.xml` and a PowerPoint `app.xml` each parse, leave the
other format's fields `None`, and round-trip without emitting them.

### F-018, Create oxml-opc (M)
Copy `rdocx-opc` into an isolated staged crate. Replace `new_docx` with
`with_main_part` and `ContentTypes::minimal` without changing `rdocx-opc`.
**Test gate**: the 11 moved tests pass, with the two docx-specific ones rebuilt
on a local fixture helper.

### F-019, PresentationML relationship and content types (S)
Add the package-namespace, extended and custom property, and PresentationML
constants, plus a `content_types` constants module.
**Depends on**: F-018.
**Test gate**: a table test asserting every constant is unique and well-formed.

### F-020, oxml-opc reads a pptx (M)
A pptx-shaped package fixture built in code: package rels to `presentation.xml`,
slide rels to `slide1.xml`, a layout one directory up.
**Depends on**: F-019.
**Test gate**: `main_document_part()` resolves `/ppt/presentation.xml`, and
`resolve_rel_target("/ppt/slides/slide1.xml", "../slideLayouts/slideLayout1.xml")`
resolves correctly.

### F-021, Zip-slip hardening tests (S)
Part names escaping the package root, and absolute-path entries.
**Depends on**: F-018.
**Test gate**: `../../etc/passwd` is clamped to the root, and an absolute entry is
normalised.

### F-022, rdocx-opc deprecation shim (S)
`pub use oxml_opc::*` with a deprecation note, description updated, consumers
flipped to `oxml_opc` directly.
**Depends on**: F-018, F-X005.
**Test gate**: workspace compiles, and `rdocx::Error::Opc` wraps the new type.

---

## Milestone 3, Media (about 1 week)

**Goal**: one isolated staged crate owns everything about an image byte string.

**End-of-milestone gate**: the staged crate passes its tests and the hash
harness remains unchanged. F-027 later proves sniffed content types with a
focused package regression.

### F-023, oxml-media format sniffing (M)
`ImageFormat::sniff`, `from_extension`, `extension`, `content_type`, `resolve`.
**Test gate**: every supported format sniffs from magic bytes, and a `.png` that
is really a JPEG resolves to JPEG.

### F-024, Image probing and DPI (L)
`probe() -> ImageInfo` for PNG, JPEG, GIF, BMP and WebP, including `pHYs` units
0 and 1, JFIF density units 1 and 2, EXIF before the SOF, and progressive JPEG.
**Depends on**: F-023.
**Test gate**: dimension and DPI assertions per format, plus a truncation loop
`for n in 0..data.len()` that panics nowhere.

### F-025, MediaNamer (S)
`scan` parses the maximum existing suffix rather than counting.
**Test gate**: the naming assertions from F-005, now in the shared crate.

### F-026, native_size with explicit DPI (S)
`native_size(default_dpi) -> Option<NativeSize>` returns dependency-free EMU
dimensions. Declared finite positive DPI wins per axis, otherwise the explicit
caller default applies. Conversion truncates toward zero, and an invalid
effective DPI returns `None`. Callers use 72 for python-docx parity against
Word's 96.
**Depends on**: F-024.
**Test gate**: a 96 dpi PNG probed at `default_dpi = 72` yields the expected EMU.

### F-027, rdocx adopts oxml-media (M)
`rdocx::Document` uses `MediaNamer` for scanned collision-free allocation and
shared byte-first format resolution for package metadata, HTML, and layout
inputs. The facade has no local image numbering, extension, or MIME helper.
**Depends on**: F-023, F-025, F-X005.
**Test gate**: a mislabelled image is stored with its sniffed extension and
content type, naming remains collision-safe, and the hash harness is unchanged.

### F-028, add_picture_auto (S)
`Document::add_picture_auto` probes and sizes image bytes at a 72 DPI caller
default before mutation, converts the shared EMU dimensions with `Length::emu`,
and delegates successful insertion to the existing `add_picture` path. This is
an additive API, so the explicit-size signature and its existing callers stay
unchanged. Unavailable dimensions return a typed error carrying the filename
without adding a part, relationship, drawing, or paragraph.
**Depends on**: F-026, F-027.
**Test gate**: a picture added with no explicit size has exact 72 DPI EMU
dimensions before and after round-trip, while unavailable dimensions fail
atomically.

---

## Milestone 4, Layout primitives (about 2 weeks)

**Goal**: the format-neutral layout types live in `oxml-layout` and can express
a rotated, clipped, gradient-filled shape.

**End-of-milestone gate**: hash harness unchanged. This is the milestone where
that matters most.

### F-029, Create oxml-layout (M)
Copy `output.rs`, `font.rs`, `bundled_fonts.rs` and `fonts/`, `error.rs` into an
isolated staged crate. Move `FontFile` within that staged implementation and
leave `rdocx-layout` unchanged.
**Test gate**: the copied tests pass in `oxml-layout`, and the existing
`Document::load_fonts_from_dir` remains unchanged.

### F-030, Decouple line.rs (L)
In staged `oxml-layout`, replace the four docx imports with `TabStop`, `Align`,
`TabAlign`, `Underline` and `LineSpacing`. Add `wrap: bool`. The rdocx-side
converter waits for the deferred consumer cutover.
**Depends on**: F-029.
**Test gate**: `line.rs`'s 11 tests rewritten on the new types pass, and the hash
harness is unchanged.

### F-031, Transform (M)
The 2x3 affine, `rotate_about`, `then`, `apply`, `is_identity`,
`transform_rect_bbox`.
**Depends on**: F-029.
**Test gate**: composition order matches the PDF `cm` operator, verified against
a hand-computed matrix.

### F-032, Path and PathCommand (M)
Four command variants, fill rule, `bounds()` documented as conservative
control-point bounds, plus `rect`, `round_rect` and `ellipse` constructors.
**Depends on**: F-029.
**Test gate**: an ellipse path's bounds contain the ellipse and lie within its
control hull.

### F-033, Paint and Stroke (M)
Solid, linear, radial and tile paints. Stroke width, cap, join and dash.
**Depends on**: F-032, F-036.
**Test gate**: a single-stop gradient degrades to solid at construction time.

### F-034, Path and Group arms (M)
Add both `PositionedElement` variants, `PageFrame::background`,
`LayoutResult::diagnostics`, and `#[non_exhaustive]` on `PositionedElement`,
`Effect`, `PageFrame` and `LayoutResult`, with constructors on the two structs.
**Depends on**: F-031, F-033.
**Test gate**: the staged `oxml-layout` construction sites compile, and the hash
harness is unchanged.

### F-035, The walk helper (S)
`walk(elements, &mut f)` flattening groups and accumulating the transform.
**Depends on**: F-034.
**Test gate**: a three-deep nested group yields every leaf exactly once with the
correct accumulated transform.

### F-036, MediaId (S)
Content-addressed media handles replacing `embed_id` as the renderer's key.
**Depends on**: F-029.
**Test gate**: the same image bytes inserted twice produce one `MediaId`.

---

## Milestone 5, PDF backend (about 2 weeks)

**Goal**: staged `oxml-pdf` renders rotated, clipped, gradient-filled paths and
nested groups. Released rdocx keeps its dependency graph and publication state,
with only the F-039 global CTM source change mirrored into `rdocx-pdf` before
the F-046 cutover.

**End-of-milestone gate**: golden-PNG diffs of the whole sample corpus show zero
pixel changes.

### F-037, Create oxml-pdf (S)
Copy `rdocx-pdf` into an isolated staged `oxml-pdf`, rewire the copy to
`oxml-layout` and `oxml-media`, and delete duplicated header parsers from the
copy. Leave the `rdocx-pdf` dependency cutover and publication until F-046.
F-039 is the only approved mirrored source change before that cutover.
**Depends on**: F-029, F-024.
**Test gate**: the eight moved tests pass.

### F-038, Golden-PNG harness (M)
Render the sample corpus to PNG and compare pixels. Distinct from the hash
harness, and specifically for F-039.
**Depends on**: F-037, F-001.
**Test gate**: passes on an unmodified tree, fails on an injected one-pixel offset.

### F-039, Global CTM flip (L)
Replace the per-element Y flip with one `q 1 0 0 -1 0 H cm`. Text `Tm` becomes
`[1 0 0 -1 x y]`, images `cm [w 0 0 -h x y+h]`.
**Depends on**: F-038.
**Test gate**: the old manifest differs only at the four declared Poppler
26.01.0 antialias pixels in `invoice` and `quote`, then all seven buffers match
the reviewed manifest exactly.

### F-040, Group rendering (M)
`q`, `cm`, optional clip via `W n`, optional `/ExtGState` for opacity, recurse,
`Q`. Effects and raster group support remain owned by later renderer work.
**Depends on**: F-039.
**Test gate**: `q`/`Q` counts balance in the content stream for a three-deep
nesting.

### F-041, Path rendering (M)
`m`, `l`, `c`, `h` then `f`, `f*`, `S`, `B` or `B*`. Stroke state via `w`, `J`,
`j`, `M`, `d`. This story renders solid paint components. Gradient shading
dictionaries remain owned by F-043.
**Depends on**: F-039.
**Test gate**: fill-only emits `f`, stroke-only `S`, both `B`.

### F-042, Rewrite the three collection passes on walk (M)
Font subsetting, XObject registration and link annotations use `walk`.
Depth-first leaf ordinals align resources with recursive emission, and link
rectangles apply the accumulated group transform.
**Depends on**: F-035, F-040.
**Test gate**: three tests, one per pass, each with the target nested inside a
group. This is the R3 regression gate.

### F-043, Gradient shading dictionaries (L)
Type 2 axial and type 3 radial, with a type 3 stitching function over type 2
exponentials, deterministic occurrence names, page-local pattern resources,
and an accumulated `/Matrix` so gradients rotate with their shape. Fill and
stroke pattern operators preserve the supported solid half of mixed paint.
**Depends on**: F-041.
**Test gate**: a rotated linear gradient renders with its axis rotated, asserted
on sampled raster pixels at 72 dpi with Poppler 26.01.0.

### F-044, ExtGState alpha (S)
One document-wide state per distinct normalized alpha, with page-local resource
references. Differing fill and stroke alpha paint the path in two operations.
**Depends on**: F-039.
**Test gate**: a 50 percent alpha fill over white rasterises to the midpoint colour.

### F-045, Rasteriser: groups, paths, gradients, dashes, background (L)
The raster backend recursively composes group transforms, intersects clip
masks, composites group opacity, translates path geometry and paint to
tiny-skia, honours line and path dashes, and paints supported page backgrounds.
**Depends on**: F-040, F-041, F-043.
**Test gate**: a rotated rectangle at 72 dpi has a filled interior pixel and an
empty corner, and a dashed line has gaps.

---

## Milestone 6, shared publication and rdocx cutover (after PowerPoint development)

**Goal**: after PowerPoint development is complete, the shared crates are
published through an approved release plan and rdocx moves onto them.

**End-of-milestone gate**: `cargo publish --dry-run` passes for every crate and
the `.crate` sizes are under the limit.

### F-046, rdocx layout and PDF cutover (M)
Move `rdocx-layout` onto the published `oxml-layout` types through its retained
flow-model facade, add the `rdocx-pdf` deprecation shim over published
`oxml-pdf`, and install the rdocx-side conversion boundary deferred from F-030.
**Depends on**: F-030, F-037, F-047 through F-050, F-X005.
**Test gate**: the workspace compiles, `rdocx::Error::Layout` wraps the new
type, and the hash harness is unchanged.

### F-047, Packaging include and size gate (M)
`include` on `oxml-layout`, drop `--no-verify`, assert `.crate` size in CI.
**Depends on**: F-037.
**Test gate**: `cargo package --list` contains every TTF and the licence files,
and the archive is under 10 MiB.

### F-048, Automate split-family release preparation (M)
Add `cargo-release` preparation for the stable and incubating tag namespaces.
**Test gate**: a dry-run bump of the workspace version updates
`[workspace.package]` and every `[workspace.dependencies]` pin, and touches no
README prose.

### F-049, Extend publish.yml to the extracted workspace (M)
Publish the expanded dependency graph and support both release tag namespaces.
**Depends on**: F-048.
**Test gate**: a dry-run publish of the full workspace succeeds in dependency
order.

### F-050, CI matrix additions (S)
`--no-default-features` for `oxml-layout`, the wasm check job, the prose gate.
**Test gate**: all new jobs pass on a clean tree.

### F-051, CHANGELOG and migration notes (S)
Document the crate moves, the deprecations, and the eventual breaking cutover.
**Depends on**: F-015, F-016, F-022, F-027, F-028, F-046, F-X005.
**Test gate**: every renamed crate is named in the CHANGELOG with its replacement.

---

## Milestone 7, DrawingML (about 4 weeks)

**Goal**: `oxml-drawing` models enough of the `a:` namespace to describe any
shape a business deck contains.

**End-of-milestone gate**: every `a:txBody` and `a:spPr` in the deck corpus
parses, serialises and reparses to a structurally equal value. F-067 executes
this carried gate at S16 entry after it creates the external corpus harness.

### F-052, Create oxml-drawing and namespace constants (S)
**Test gate**: crate compiles, namespace URIs match the spec.

### F-053, OrderedRawChildren (M)
The schema child-order helper that keeps unmodelled siblings in their slots.
**Test gate**: an element with a modelled child between two unmodelled ones
round-trips with all three in the original order.

### F-054, Colour choices (M)
`a:srgbClr`, `a:schemeClr`, `a:sysClr`, `a:prstClr`.
**Test gate**: each form parses and round-trips.

### F-055, The colour transform stack (L)
All transforms, applied in document order, with RGB-to-HSL conversion and
linear-gamma tint and shade per ECMA-376 20.1.2.3.
**Depends on**: F-054, F-014.
**Test gate**: a table of 40 (theme colour, transform) pairs sampled from real
PowerPoint renders resolves to exact RGB.

### F-056, Colour map resolution (M)
`p:clrMap` and `p:clrMapOvr` applied before the theme lookup.
**Depends on**: F-055.
**Test gate**: a dark master inverting `bg1` and `tx1` resolves correctly.

### F-057, a:xfrm (M)
Offset, extent, child offset and extent, rotation, flips.
**Test gate**: a nested group transform composes to the hand-computed matrix.

### F-058, Guide evaluator (L)
The full `GuideOp` set, the seeded environment, adjust values, and `a:arcTo`
flattened to cubics.
**Depends on**: F-014.
**Test gate**: a hand-written `custGeom` with guides produces the expected path
coordinates.

### F-059, a:custGeom (M)
Path lists, adjust value lists, guide lists, the text rectangle.
**Depends on**: F-058.
**Test gate**: a corpus `custGeom` shape round-trips and evaluates to a closed path.

### F-060, Fills (L)
`a:noFill`, `a:solidFill`, `a:gradFill` with linear and path variants,
`a:pattFill`, `a:blipFill` with stretch, tile and `a:srcRect`.
**Depends on**: F-054.
**Test gate**: each fill form round-trips, and a gradient's stops are ordered.

### F-061, Lines (M)
`a:ln` with width, dash presets, cap, join, head and tail ends.
**Depends on**: F-054.
**Test gate**: every `ST_PresetLineDashVal` maps to a dash array.

### F-062, Effects (S)
`a:effectLst` with outer shadow modelled, everything else preserved.
**Test gate**: a shape with a glow round-trips with the glow intact as raw XML.

### F-063, Shape properties and style references (M)
`a:spPr`, and `a:lnRef` / `a:fillRef` / `a:effectRef` / `a:fontRef` including the
`idx > 1000` background-fill rule.
**Depends on**: F-060, F-061.
**Test gate**: `fillRef@idx = 1001` resolves to background fill style 1.

### F-064, DrawingML text model (XL, split at implementation)
`a:txBody`, `a:bodyPr`, `a:lstStyle` with nine levels, `a:p`, `a:pPr`, `a:r`,
`a:rPr`, `a:t`, `a:fld`, `a:br`, and the bullet family.
**Depends on**: F-053.
**Test gate**: every `a:txBody` in the corpus round-trips structurally, and
`a:t` whitespace survives via `xml:space`.

F-064 is split into the four implementation stories below. The parent closes
only after every child closes.

### F-064a, Text body properties and shell (M)
`a:txBody` ownership and `a:bodyPr` insets, anchoring, wrapping, vertical
direction, and autofit forms.
**Depends on**: F-053.
**Test gate**: every `a:bodyPr` autofit form round-trips in schema order with
unmodelled children preserved.

### F-064b, Text paragraphs and runs (L)
`a:p`, `a:pPr`, `a:r`, `a:rPr`, `a:t`, `a:fld`, and `a:br`, including the
DrawingML centipoint and percentage conventions.
**Depends on**: F-064a.
**Test gate**: leading and trailing `a:t` whitespace survives a structural
round-trip through `xml:space="preserve"`.

### F-064c, Text bullets (S)
`a:buChar`, `a:buAutoNum`, `a:buNone`, `a:buFont`, `a:buSzPct`, `a:buSzPts`,
and `a:buClr` on paragraph properties.
**Depends on**: F-064b.
**Test gate**: every modelled bullet form round-trips with colour, font, and
size children in schema order.

### F-064d, Nine-level list styles (M)
`a:lstStyle` with nine level-specific paragraph property slots, completing the
modelled `a:txBody` hierarchy.
**Depends on**: F-064b, F-064c.
**Test gate**: a schema-valid `a:txBody` fixture using all nine list levels
serialises, reparses, and remains structurally equal.

### F-065, Theme read and write (L)
`CT_OfficeStyleSheet` including `a:fmtScheme`, plus `office_default()`.
**Depends on**: F-060, F-061.
**Test gate**: the Office theme generated by PowerPoint 16.104 build
16.104.25121423 round-trips structurally, and `office_default()` produces a
theme that the same pinned build opens without repair.

### F-066, The rdocx Theme adapter (S)
`impl From<&CT_OfficeStyleSheet> for rdocx_oxml::theme::Theme`, leaving the Word
tint and shade path untouched.
**Depends on**: F-065.
**Test gate**: the hash harness is unchanged.

---

## Milestone 8, PresentationML (about 4 weeks)

**Goal**: open any deck in the corpus, model what will be rendered, preserve the
rest verbatim, and save it byte-comparably.

**End-of-milestone gate**: all 50 corpus decks round-trip, and every one opens in
PowerPoint without a repair prompt.

### F-067, Create rpptx-oxml and the corpus harness (M)
Crate skeleton, corpus fetch script, and a raw open-and-save test treating every
part as opaque. Once the corpus is present, execute the carried M7 DrawingML
structural gate before beginning M8 model work.
**Test gate**: the carried M7 DrawingML gate passes, and all 50 decks round-trip
byte-identically with no XML modelling.

### F-068, presentation.xml (M)
`CT_Presentation`, `p:sldSz`, `p:notesSz`, `p:sldIdLst`, `p:sldMasterIdLst`,
`p:defaultTextStyle`.
**Test gate**: every corpus deck's presentation part round-trips.

### F-069, Slide, layout and master parts (L)
`CT_Slide`, `CT_SlideLayout`, `CT_SlideMaster`, `p:cSld`, `p:clrMap`,
`p:clrMapOvr`, `p:txStyles`.
**Depends on**: F-064.
**Test gate**: every corpus slide, layout and master round-trips structurally.

### F-070, The shape tree (L)
`p:spTree`, `p:nvGrpSpPr`, `p:grpSpPr`, and the six-variant child union.
**Depends on**: F-063.
**Test gate**: a deck with nested groups round-trips with tree shape preserved.

### F-071, Placeholders (M)
`p:ph`, `PhType`, `PlaceholderKey` and its matching rule.
**Depends on**: F-070.
**Test gate**: matching by idx, by type, absent type defaulting to body, and both
equivalence classes.

### F-072, Pictures (M)
`p:pic`, `p:blipFill`, `a:srcRect` crop.
**Depends on**: F-060.
**Test gate**: a cropped picture round-trips with its crop rectangle.

### F-073, Graphic frames (M)
`p:graphicFrame` and the `a:graphicData` uri dispatch for tables, charts,
SmartArt and OLE.
**Depends on**: F-070.
**Test gate**: each payload kind is recognised and its unmodelled forms preserved.

### F-074, DrawingML tables (L)
`a:tbl`, `a:tblPr`, `a:tblGrid`, `a:tr`, `a:tc`, merges and banding flags.
**Depends on**: F-064.
**Test gate**: a table with merged cells round-trips with merge origins intact.

### F-075, Connectors (S)
`p:cxnSp` with start and end connections.
**Test gate**: a corpus connector round-trips.

### F-076, mc:AlternateContent (M)
Preserved verbatim, with the fallback branch selected for rendering.
**Depends on**: F-070.
**Test gate**: a deck with `AlternateContent` round-trips byte-identically in that
subtree.

### F-077, Notes slides and notes master (M)
**Depends on**: F-069.
**Test gate**: notes text extracts, and a deck with notes round-trips.

### F-078, relmap rewrite_rel_ids (M)
**Depends on**: F-067.
**Test gate**: a preserved blob containing `r:embed`, `r:link` and `r:dm` has all
three rewritten, and everything else is byte-identical.

### F-079, The rpptx read facade (L)
`Presentation::open`, `from_bytes`, `to_bytes`, `slides`, `text`, plus the
`*Ref` handle types and shape iteration.
**Depends on**: F-069, F-070.
**Test gate**: a `dump_deck` example printing every slide's shapes and text
matches python-pptx's output on the corpus.

### F-080, Modelled round-trip gate (M)
Parse, serialise, reparse, compare structurally, plus part-by-part byte
comparison of the saved package.
**Depends on**: F-079.
**Test gate**: all 50 decks pass, and each opens in PowerPoint without repair.

---

## Milestone 9, Inheritance resolver (about 2 weeks)

**Goal**: a `ResolvedSlide` in which every inherited and theme-derived value is
already concrete.

**End-of-milestone gate**: the contract is frozen and published to the render
track.

### F-081, ResolveCtx skeleton and placeholder chain (M)
**Depends on**: F-071.
**Test gate**: a slide placeholder resolves to its layout and master counterparts.

### F-082, Effective transform and body properties (M)
**Depends on**: F-081, F-057.
**Test gate**: a slide placeholder with no `a:xfrm` inherits the layout's position.

### F-083, The seven-step list style merge (L)
**Depends on**: F-081, F-064.
**Test gate**: a run inheriting from `p:defaultTextStyle` through five
intermediate levels resolves to the expected size and typeface.

### F-084, Format scheme reference resolution (M)
Including `phClr` substitution and the `idx > 1000` rule.
**Depends on**: F-063, F-065.
**Test gate**: a shape with `p:style` resolves to the theme's fill with its own
colour substituted.

### F-085, Typeface resolution (S)
`+mn-lt`, `+mj-lt`, `+mn-ea`, `+mn-cs` and per-script overrides.
**Depends on**: F-065.
**Test gate**: `+mn-lt` resolves to the theme's minor Latin typeface.

### F-086, Draw order and the flattener (L)
Background resolution, the master and layout non-placeholder passes,
`showMasterSp`, the placeholder suppression rules, and latent placeholder
handling.
**Depends on**: F-081.
**Test gate**: a rendered slide contains no "Click to edit Master title style",
and a master logo appears exactly once.

### F-087, ResolvedSlide contract (M)
The full type set, frozen and documented.
**Depends on**: F-082 through F-086.
**Test gate**: a corpus slide resolves with no unresolved theme references
remaining anywhere in the output.

### F-088, Visual differential tests (M)
Decks whose correct appearance can be eyeballed, plus the 40-pair colour table.
**Depends on**: F-087.
**Test gate**: the colour table resolves exactly, and the differential decks are
reviewed once manually.

---

## Milestone 10, Renderer (about 4 weeks)

**Goal**: a deck renders to PDF and PNG at the quality bar in
`02-scope-and-non-goals.md`.

**End-of-milestone gate**: the pinned 50-deck SSIM harness renders every slide
without panic, missing output, dimension mismatch, or a dropped bounded shape,
retains the 0.95 SSIM on 80 percent trend result, and has an accepted native
PowerPoint representative review.

### F-089, Resolve the preset geometry licensing question (S)
Settle Q1 from `13-risks-and-open-questions.md` before writing the generator.
**Test gate**: a written decision recorded in the HLD with its licence basis.

### F-090, Preset table generator (L)
`tools/gen-presets/` emitting a checked-in generated file.
**Depends on**: F-089, F-058.
**Test gate**: the generated table covers every preset name in the corpus, and
the file regenerates byte-identically.

### F-091, Preset evaluation and fallback (M)
**Depends on**: F-090.
**Test gate**: an unknown preset emits its bounding box, keeps its text, and
records a diagnostic.

### F-092, rpptx-render skeleton and RenderInput (M)
`RelScopes`, `SlideBundle`, media resolution to `MediaId`.
**Depends on**: F-087, F-036.
**Test gate**: a slide, layout and master each using `rId2` for different targets
all resolve correctly.

### F-093, Shape geometry, fills and lines (L)
**Depends on**: F-091, F-092.
**Test gate**: a slide of solid, gradient and outlined shapes rasterises with
correct colours at sampled pixels.

### F-094, Rotation, flips and groups (M)
**Depends on**: F-093, F-031.
**Test gate**: a rotated shape's corners land at hand-computed coordinates.

### F-095, Arrowheads (S)
Lowered into filled paths.
**Depends on**: F-093.
**Test gate**: a line with a triangular tail end emits an extra filled path.

### F-096, Pictures with crop and tile (M)
**Depends on**: F-092, F-072.
**Test gate**: a cropped picture renders only its crop region.

### F-097, Backgrounds (S)
**Depends on**: F-086.
**Test gate**: a slide inheriting a master gradient background renders it.

### F-098, Shape text layout (XL, split at implementation)
`bodyPr`, insets, anchoring, wrap, the content box from the preset text
rectangle.
**Depends on**: F-083, F-030.
**Test gate**: text anchored bottom-centre in an inset box lands at the computed
baseline.

F-098 is implemented through the four stories below. F-098a owns content-box
geometry, F-098b owns shaped inline resolution, F-098c owns line stacking, and
F-098d owns horizontal and vertical anchoring. The parent is complete only when
all four child gates pass together in deterministic font mode.

### F-098a, Text content box (M)
Use the preset or custom geometry text rectangle, falling back to the shape
bounds, then apply the resolved body insets without producing negative extents.
**Depends on**: F-083, F-030.
**Test gate**: a preset text rectangle minus four unequal insets produces the
hand-computed content box.

### F-098b, Paragraph inline resolution (L)
Resolve concrete run style into shaped inline items and preserve explicit line
breaks without introducing a second text model.
**Depends on**: F-098a.
**Test gate**: resolved text runs emit glyph items with the expected font size,
colour, style, and explicit break boundaries.

### F-098c, Line stacking (M)
Break paragraphs against the content width, apply paragraph indents and spacing,
and stack their lines in shape-local coordinates.
**Depends on**: F-098b.
**Test gate**: wrapped paragraphs stack at hand-computed baselines while
`wrap="none"` breaks only at explicit line breaks.

### F-098d, Text anchoring (S)
Lower stacked line items to glyph runs, apply horizontal paragraph alignment,
and place the complete block through the resolved vertical anchor.
**Depends on**: F-098c.
**Test gate**: text anchored bottom-centre in an inset box lands at the computed
baseline.

### F-099, Bullets (M)
Character, auto-number with the eight common schemes, none, size, colour, and
the Wingdings codepoint table.
**Depends on**: F-098d.
**Test gate**: a Wingdings `F0B7` bullet renders as a visible bullet glyph, not
a missing-glyph box.

### F-100, Autofit (M)
Stored `normAutofit` applied verbatim, `spAutoFit` trusted, `noAutofit`
overflowing without clipping, and the 2.5 percent ladder for the bare case.
**Depends on**: F-098d.
**Test gate**: a stored `fontScale` of 62500 renders at exactly 62.5 percent.

### F-101, Vertical text (S)
Transposed layout wrapped in a rotated group, with `eaVert` degrading.
**Depends on**: F-098d.
**Test gate**: vertical text renders rotated and records a diagnostic for `eaVert`.

### F-102, Table rendering (L)
**Depends on**: F-074, F-098.
**Test gate**: a banded table with merged cells renders with correct fills and no
duplicated borders.

### F-103, Hyperlinks, fields and diagnostics (M)
Link annotations, slide-number fields reusing the existing field machinery, and
the diagnostic surface.
**Depends on**: F-092.
**Test gate**: a slide-number field renders the correct number and a hyperlink
emits an annotation.

### F-104, SSIM fidelity harness (L)
Corpus renders compared with LibreOffice.
**Depends on**: F-102.
**Test gate**: all pinned corpus slides render without panic, missing output,
dimension mismatch, or a dropped bounded shape. The harness records 0.95 SSIM
on 80 percent as a trend, and the native PowerPoint representative review is
accepted.

---

## Milestone 11, Write API (about 3 weeks)

**Goal**: build and edit decks, and produce files PowerPoint accepts.

**End-of-milestone gate**: a generated 10-slide deck opens clean in PowerPoint,
Keynote, Google Slides and LibreOffice.

### F-105, Bundled default.pptx (M)
The 16:9 template with one master, eleven layouts, a full theme, and zero slides.
**Depends on**: F-065.
**Test gate**: `Presentation::new()` produces a deck PowerPoint opens without
repair.

### F-106, ShapeIdAllocator and MediaStore (M)
Tree-wide id scanning, and content-hash media deduplication.
**Depends on**: F-070, F-036.
**Test gate**: ids are unique across nested groups and `AlternateContent`, and the
same image inserted twice creates one part.

### F-107, add_slide (L)
The nine-step synthesise recipe.
**Depends on**: F-105, F-106.
**Test gate**: a deck with three added slides opens without repair, and every
`p:sldId/@id` is at least 256 and unique.

### F-108, validate() (M)
Every `ValidationIssue` variant, run under `debug_assertions` before save.
**Depends on**: F-107.
**Test gate**: one deliberately corrupted deck per variant is detected, and the
whole corpus validates clean.

### F-109, Shape mutation facade (L)
Position, size, rotation, name, fill, line, adjust values.
**Depends on**: F-079.
**Test gate**: every setter round-trips through save and reload.

### F-110, add_textbox, add_shape, add_connector, add_group_shape (M)
**Depends on**: F-109.
**Test gate**: each produces a shape PowerPoint opens without repair.

### F-111, add_picture (M)
Owning-facade picture insertion uses 72-DPI native sizing, truncating one-axis
aspect inference, package-wide media deduplication, and slide-scoped image
relationships. Every fallible operation completes before package or shape-tree
state is committed.
**Depends on**: F-106, F-026.
**Test gate**: a picture added with no explicit size uses its native dimensions.

### F-112, Text frame mutation (L)
Text frame, paragraphs, runs, font properties, bullets.
**Depends on**: F-109.
**Test gate**: setting text on a placeholder round-trips and renders.

### F-113, Table facade (L)
`add_table`, cells, merge and split, banding, column widths.
**Depends on**: F-074, F-109.
**Test gate**: merge then split restores the original grid.

### F-114, remove_slide, move_slide, duplicate_slide (M)
Including deep copy through `rewrite_rel_ids` and media transfer.
**Depends on**: F-078, F-107.
**Test gate**: a duplicated slide's images resolve to the new slide's own
relationships.

### F-115, Slide and presentation properties (S)
Slide size, background, hidden flag, core properties, `save_as_show`.
**Depends on**: F-017.
**Test gate**: each property round-trips.

### F-116, Cross-viewer acceptance (M)
**Depends on**: F-107 through F-115.
**Test gate**: a generated 10-slide deck exercising every feature opens clean in
all four viewers.

---

## Milestone 12, Charts (about 7 weeks)

**Goal**: create and render charts.

**End-of-milestone gate**: a chart created by rpptx opens in PowerPoint, its
data is editable, and it renders.

### F-117, oxml-sml workbook writer (L)
One worksheet, numeric and string cells, shared strings, defined ranges.
**Test gate**: the produced `.xlsx` opens in Excel and LibreOffice Calc.

### F-118, ChartML core types (L)
`CT_ChartSpace`, `CT_Chart`, `CT_PlotArea`, `CT_Title`, `CT_Legend`.
**Depends on**: F-063.
**Test gate**: a corpus chart part round-trips.

### F-119, Series and data references (L)
`c:ser`, `c:cat`, `c:val`, string and numeric references, and the caches.
**Depends on**: F-118.
**Test gate**: a chart written with a cache and a formula reference has both
consistent with one source of data.

### F-120, Axes (L)
`c:catAx`, `c:valAx`, `c:dateAx`, `c:serAx`, scaling, gridlines, tick marks,
label position, and paired `crossAx` ids.
**Depends on**: F-118.
**Test gate**: axis id pairing is consistent, and a corpus chart's axes round-trip.

### F-121, Bar and line plots (M)
**Depends on**: F-119, F-120.
**Test gate**: each round-trips and renders.

### F-122, Pie, doughnut, area, scatter and radar plots (L)
**Depends on**: F-121.
**Test gate**: each round-trips and renders.

### F-123, Data labels and number formats (M)
**Depends on**: F-119.
**Test gate**: a percentage-formatted label renders with the correct text.

### F-124, add_chart (L)
Writes the chart part, the workbook, both relationship sets, both content-type
overrides, and the graphic frame.
**Depends on**: F-117, F-121.
**Test gate**: a created chart opens in PowerPoint and "Edit Data" shows the
source values.

### F-125, Chart rendering: geometry (L)
Bars, lines, wedges, areas and markers as paths.
**Depends on**: F-121, F-093.
**Test gate**: a bar chart rasterises with bars at computed positions.

### F-126, Chart rendering: axes, gridlines and labels (L)
Nice-number tick selection, axis lines, tick labels, legend.
**Depends on**: F-125, F-098.
**Test gate**: a chart with a 0 to 100 value axis produces the expected tick set.

### F-127, Chart colour resolution (M)
Series colours from `c:spPr` or the theme accent cycle.
**Depends on**: F-125, F-055.
**Test gate**: an unstyled four-series chart uses accent1 through accent4.

### F-128, Preserved chart fallback (S)
Cached image if present, else a labelled placeholder with a diagnostic.
**Depends on**: F-125.
**Test gate**: a 3-D chart renders its cached image and records a diagnostic.

---

## Milestone 13, Bindings and tooling (about 4 weeks)

**Goal**: both libraries ship as crates, CLIs, WASM modules and Python wheels.

**End-of-milestone gate**: wheels install and pass the parity suites on every
target platform.

### F-129, oxml-py-support (M)
Word `ContentPath` and `PathSeg` values, the revision counter, the Rust
`StaleElementError`, and canonical `Length` conversion helpers. Presentation
path variants wait for F-136.
**Test gate**: a stale path raises the named error with both revisions in the
message.

### F-130, rdocx-py core (L)
`PyDocument`, lazy collections, paragraph and run handles.
**Depends on**: F-129, F-008.
**Test gate**: `doc.paragraphs[3]` held across `remove_content(1)` raises
`StaleElementError` rather than reading the wrong paragraph.

### F-131, rdocx-py formatting and tables (L)
Path-only font and paragraph-format sub-handles expose the bounded S33
formatting inventory with tri-state clearing. Lazy table, row, cell and nested
paragraph handles cover table style, alignment and width, plus cell text,
width and vertical alignment. Public facade accessors re-resolve every path.
**Depends on**: F-130, F-132.
**Test gate**: `r.font.bold` returns `None` when unset, not `False`.

### F-132, Python enums, units and exceptions (M)
The bounded `IntEnum` shims for paragraph alignment, table alignment, cell
vertical alignment and underline, pure-Python `Length` and `RGBColor` values,
the package exception hierarchy, and concrete mapping from Rust binding errors.
The types are top-level exports and retain the `rdocx.shared`,
`rdocx.enum.text` and `rdocx.enum.table` compatibility paths.
**Depends on**: F-129, F-130.
**Test gate**: `WD_ALIGN_PARAGRAPH.CENTER == 1` and `Inches(1) == 914400`.

### F-133, rdocx-py rendering with allow_threads (S)
**Depends on**: F-130.
**Test gate**: four concurrent `to_pdf` calls from a thread pool complete faster
than serial execution.

### F-134, Type stubs and py.typed (M)
Both mixed packages ship hand-written native-extension stubs and `py.typed`
markers. Strict installed-wheel smoke programs cover concrete handles,
collections, overloads, iterators, path-like inputs, byte outputs, and optional
values without duplicating inline-typed pure-Python modules. Bounded enums and
Length returns retain their semantic types, and factory-only native handles
remain non-constructible at type-check time.
**Depends on**: F-131, F-136.
**Test gate**: exact `mypy==2.3.0 --strict` and `stubtest` both pass against
freshly installed cp39-abi3 wheels.

### F-135, python-docx parity suite (M)
**Depends on**: F-131.
Pin and assert python-docx 1.2.0. Execute an explicit manifest of all executable
documentation examples inside the completed S33 surface from stable v1.2.0
tagged sources. Sixteen examples change only the import namespace. The exact
Quickstart held-row example uses the minimal public row re-fetch required by
strict global revision before its second cell assignment. Author the approved
structure with both writers, read both outputs with both libraries, and compare
normalized public records rather than package bytes. Preserve relative float
line spacing separately from absolute lengths and compare explicit table style
after save and reopen.
**Test gate**: `documented_s33_examples_run_with_declared_transformations`
passes for the exact seventeen-entry manifest, and the two-way normalized
differential agrees.

### F-136, rpptx-py (L)
An unpublished abi3-py39 mixed-layout binding over `Presentation`, using lazy
path-only slide, shape, text and table handles. The bounded surface includes
pure-Python presentation units and required shape enum values.
**Depends on**: F-129, F-116.
**Test gate**: the seven python-pptx 1.0.2 Getting Started workflows run with
the package namespace changed and minimal public re-fetches after structural
writes. Both readers agree on each writer, and normalized structures from the
two writers agree directly with that exact oracle version.

### F-137, wheels.yml (M)
Build `rdocx` and `rpptx` with maturin as abi3-py39 wheels for
manylinux_2_28 x86_64 and aarch64, musllinux_1_2 x86_64, macOS x86_64 and
arm64, and Windows x86_64. Build one source distribution per package. Every
compatible wheel is installed and tested in a fresh environment. A separate
job selects the exact six wheels and one source distribution for either
`py-rdocx-v*` or `py-rpptx-v*` and receives PyPI OIDC authority only for those
tag namespaces. Manual dispatch builds both projects and never publishes.
**Depends on**: F-134, F-136.
**Test gate**: the local exact-product contract and its negative mutations
pass, both native wheels and source distributions build, and both native wheels
install and pass their compatible package, typing, and stub gates. The first
reviewed hosted dispatch supplies cross-platform execution evidence.

### F-138, PR-time Python job (S)
`maturin develop && pytest`.
**Depends on**: F-137.
**Test gate**: the job fails when a binding test fails.

### F-139, Rewrite rdocx-wasm (L)
Wrap `rdocx::Document` and keep the existing JavaScript method names. The
default-on `system-fonts` feature is forwarded through `rdocx-layout` and
`rdocx`, while `rdocx-wasm` disables it and retains unconditional bundled font
data. An inline Node regression exercises the same package-preserving contract
as the native gate.
**Depends on**: F-029.
**Test gate**: a document with images, headers and numbering round-trips through
`fromBytes` and `toDocxBytes` with every part intact. This is the R-class
regression gate.

### F-140, wasm CI job (S)
**Depends on**: F-139, F-142.
**Test gate**: locked `cargo check --target wasm32-unknown-unknown` and
`wasm-pack test --node` run for both WASM packages on PRs.

### F-141, to_pdf in the browser (M)
**Depends on**: F-139, F-001.
**Test gate**: a wasm-pack node test produces a non-empty PDF with embedded fonts.

### F-142, rpptx-wasm (M)
Wrapping the real facade, in two feature profiles.
**Depends on**: F-116.
**Test gate**: the default profile is under 1 MB gzipped and round-trips a deck.

### F-143, oxml-cli-support (S)
Range parsing, output-path defaulting, the versioned JSON envelope.
**Test gate**: `2,4-6` parses to the expected set, and the envelope carries
`"schema": 1`.

### F-144, rpptx-cli (L)
`inspect`, `text`, `convert`, `diff`, `replace`, `validate`, `render`.
**Depends on**: F-143, F-116, F-104.
**Test gate**: `validate` exits non-zero on a corrupted deck and zero across the
corpus.

### F-145, rpptx-cli thumbnail and outline (M)
**Depends on**: F-144.
**Test gate**: `thumbnail` produces a proportional 320-pixel-wide PNG of slide
one, and `outline` prints each title once followed by the recursive paragraph
tree with stable level indentation.

### F-146, npm publication (S)
`@tensorbee/rdocx-wasm` and `@tensorbee/rpptx-wasm` build as release bundler
packages under exact checksum-pinned wasm-opt 125. Pull-request CI packs and
installs both tarballs locally without registry credentials or publication
authority.
**Depends on**: F-140, F-142.
**Test gate**: `npm pack` produces an installable tarball for each, and both
installed packages retain their exact metadata, WASM, JavaScript glue,
TypeScript declaration, and import.

---

## Milestone 14, Word collaboration layer (about 4 weeks)

**Goal**: the parts of a document that exist because more than one person
touched it. All four are preserved verbatim today and none is addressable.

Commercial libraries treat this as the dividing line. Aspose.Words, Spire.Doc
and GemBox all sell revision and comment APIs, and `python-docx` has offered
neither in a decade of requests. Nothing in the Rust ecosystem has any of it.

**End-of-milestone gate**: a document carrying tracked changes, comments,
content controls and bookmarks round-trips byte-identically in the parts this
milestone does not model, and every one of the four is readable and writable
through the public API.

### F-147, Comment model and part (M)
`word/comments.xml`, `CT_Comment` and `CT_Comments`, plus the
`w:commentRangeStart`, `w:commentRangeEnd` and `w:commentReference` anchors in
the body. Today the part survives because `OpcPackage` writes every part it
holds, which means a comment is never lost and never reachable.
**Depends on**: none.
**Test gate**: round-trip. A document with three comments, one spanning two
paragraphs, reloads with every anchor in the same place and saves byte-identical.

### F-148, Comment API (M)
`Document::comments`, `add_comment` over a run range, `reply_to`, `resolve` and
`remove`. Replies use `w:commentsExtended` and the paragraph-id linkage, which
is what Word itself reads.
**Depends on**: F-147.
**Test gate**: regression. A comment added over a range, replied to and resolved
opens in Word with the thread intact.

### F-149, Revision model (L)
`w:ins`, `w:del`, `w:delText`, `w:moveFrom`, `w:moveTo`, and the property-change
elements `w:rPrChange`, `w:pPrChange`, `w:tblPrChange` and `w:sectPrChange`.
These are captured as raw XML today, listed in the modelled-children exclusions
in `numbering.rs` and `text.rs`.
**Depends on**: none.
**Test gate**: round-trip. Every revision element survives a load and save
unchanged, and each is reported with its author, timestamp and kind.

### F-150, Accept and reject revisions (L)
`accept_all`, `reject_all`, and the same two scoped to an author, a date range
or a single revision id. Rejecting an insertion removes content, rejecting a
deletion restores it, and a property change reverts to the recorded prior value.
**Depends on**: F-149.
**Test gate**: regression. Accepting every revision produces the document Word
produces from the same input, compared as normalised body XML.

### F-151, Revision display in the renderer (M)
Rendering shows insertions underlined, deletions struck through, and a change
bar in the margin, or renders the accepted view. The choice is a render option
and the default is the accepted view, because that is the document a reader
means when they ask for a PDF.
**Depends on**: F-149.
**Test gate**: golden. Both views of one document render, and the accepted view
is pixel-identical to the same document with revisions accepted and removed.

### F-152, Content control model (L)
`w:sdt`, its `w:sdtPr` properties and `w:sdtContent`, at block, row, cell,
paragraph and run level. `table.rs` already unwraps these to find rows and
cells, so the traversal exists and the model does not.
**Depends on**: none.
**Test gate**: round-trip. Controls at all five nesting levels survive, and each
is reported with its tag, alias, id and type.

### F-153, Content control binding (M)
Read and write a control's value by tag or alias, and bind a control set to a
key-value map in one call. Includes the `w:dataBinding` XPath into a custom XML
part, which is how document-assembly products drive Word.
**Depends on**: F-152.
**Test gate**: regression. A control set bound to a map produces the expected
text, and a bound custom XML part updates both the part and the display text.

### F-154, Bookmarks and cross-references (M)
`w:bookmarkStart` and `w:bookmarkEnd`, a bookmark collection, insertion over a
range, and `REF` and `PAGEREF` targets resolved against them.
**Depends on**: none.
**Test gate**: regression. A bookmark inserted over a range is listed, its text
is readable, and a cross-reference to it resolves to the right page after
pagination.

### F-155, Document protection (M)
`w:documentProtection` in settings: read-only, comments-only, tracked-changes-
forced and forms-only, with the hash and salt Word writes. Reading the setting
matters more than enforcing it, because a consumer needs to know the author's
intent.
**Depends on**: none.
**Test gate**: regression. Each protection mode round-trips with its hash
intact, and the mode is reported through the public API.

---

## Milestone 15, Charts beyond PowerPoint (about 2 weeks)

**Goal**: one chart engine, two document families. `oxml-chart` owns the
format-neutral model and renderer. `rpptx-chart` remains an exact deprecated
re-export for existing consumers.

`python-docx` has no chart API at all. The standard workaround is rendering a
chart to PNG and pasting it, which loses every bit of editability. Apache POI
and docx4j both have native Word charts, and so does every commercial library.

**End-of-milestone gate**: a Word document gains a native chart that opens
editable in Word, and renders identically to the same chart in a deck.

### F-156, Extract oxml-chart (L)
Move `rpptx-chart` to `oxml-chart` with no behaviour change. A pure rename and
re-export, with the deprecation shim pattern F-015 and F-022 already
established.
**Depends on**: none.
**Test gate**: regression. The hash harness is byte-identical across the move,
and every existing chart test passes against the new path. This is a file move,
so folding any behaviour change into it is forbidden.

### F-157, Word chart part and embedded workbook (M)
The chart part, its relationship from `document.xml`, and the embedded
`.xlsx` workbook Word requires. `oxml-sml` already writes exactly the one
worksheet a chart needs, which is the whole reason it exists.
**Depends on**: F-156.
**Test gate**: round-trip. A document with a chart part saves with the part, its
relationship, its content type and its embedded workbook, and Word opens it
without repair.

### F-158, Document::add_chart (M)
The Word-side authoring API, matching the shape of `rpptx`'s `add_chart` so a
reader who knows one knows the other.
**Depends on**: F-157.
**Test gate**: regression. A bar, line and pie chart added to a document carry
the series, categories and number formats they were given.

### F-159, Chart rendering in the Word paginator (M)
An anchored or inline chart lays out and renders through the same path as an
image, delegating to the chart renderer for its content.
**Depends on**: F-158.
**Test gate**: golden. A chart in a Word document renders pixel-identical to the
same chart on a slide at the same size.

---

## Milestone 16, Document automation (about 5 weeks)

**Goal**: generate documents from data rather than editing them by hand. This
is the largest commercial category. Aspose sells a LINQ reporting engine,
docxtpl is one of the most-used Python packages in the space, and every
document-assembly product is built on fields, content controls and merges.

`rdocx` already has `replace_text`, `replace_regex`, `replace_all` and
`replace_many_in_chart_xml`, which covers substitution and nothing structural.

**End-of-milestone gate**: a template with loops, conditionals and a repeating
table row produces a correct document from a JSON data model, and every field in
it evaluates to the value Word computes.

### F-160, Field instruction parser (L)
`w:fldSimple` and the `w:fldChar` plus `w:instrText` run sequence, parsed into a
field name, arguments and switches. `text.rs` already extracts `w:instr` for the
simple form.
**Depends on**: none.
**Test gate**: unit. Every field form in the corpus parses, including nested
fields and instructions split across runs, which is how Word actually writes
them.

### F-161, Field evaluation engine (L)
`IF`, `REF`, `PAGEREF`, `SEQ`, `DOCPROPERTY`, `DOCVARIABLE`, `STYLEREF`,
`INCLUDETEXT`, `DATE`, `TIME`, `FILENAME`, `AUTHOR` and `MERGEFIELD`, plus the
formatting switches. `PAGE` and `NUMPAGES` already evaluate during pagination.
**Depends on**: F-160, F-154.
**Test gate**: regression. Each supported field evaluates to the value Word
computes for the same document, checked against a pinned expected set.

### F-162, Field update policy (M)
Update on demand, update on save, and leave alone, with the dirty flag Word
sets. A field whose result is cached must not be silently recomputed, because a
document may legitimately carry a stale result on purpose.
**Depends on**: F-161.
**Test gate**: regression. Each policy produces the expected result cache, and
an unsupported field keeps its cached result rather than blanking.

### F-163, Template syntax (L)
A tag syntax over the existing placeholder machinery, resolving inside runs that
Word has split mid-tag, which is the failure every naive implementation hits.
**Depends on**: none.
**Test gate**: unit. A tag split across five runs with different formatting
resolves, and the surrounding formatting is preserved.

### F-164, Loops and conditionals (L)
Block-level repetition and inclusion over a data model, at paragraph, row and
section granularity.
**Depends on**: F-163.
**Test gate**: regression. A template with a nested loop and a conditional
produces the expected document from a fixture data model.

### F-165, Repeating table rows and lists (M)
The two structures that need their own handling: a row that repeats keeps its
formatting and its merged cells, and a repeated list item keeps its numbering
continuous.
**Depends on**: F-164.
**Test gate**: regression. A three-row template over ten records produces thirty
rows with the banding and numbering intact.

### F-166, Mail merge (M)
A record set driving `MERGEFIELD`, with one document per record or one document
with a section per record.
**Depends on**: F-161, F-164.
**Test gate**: regression. A merge over a fixture record set produces the
expected documents, and an absent field renders empty rather than failing.

### F-167, Document comparison (L)
Compare two documents and express the difference as tracked revisions, scoped to
body text, tables and list structure. Formatting-only differences are recorded
as a diagnostic rather than a revision, which keeps this one story instead of
three.
**Depends on**: F-149.
**Test gate**: regression. Comparing a document with its edited copy produces
revisions that, when accepted, reproduce the edited copy exactly.

### F-168, Watermarks (S)
Text and image watermarks through the header `w:pict` shape Word uses, readable
and writable, and rendered.
**Depends on**: none.
**Test gate**: golden. A watermark renders behind body text on every page.

---

## Milestone 17, Security and compliance (about 3 weeks)

**Goal**: files an enterprise or a public body can accept. Encryption and
signatures are table stakes in commercial libraries and absent from every open
source Office library in Python and Rust. Apache POI is the only open
implementation of OOXML agile encryption worth reading.

Tagged PDF is a legal requirement for public-sector documents in the EU and the
United States, and a LibreOffice-based pipeline cannot produce it well. The PDF
backend here is ours, so it can.

**End-of-milestone gate**: an encrypted document opens with its password, a
signed document verifies, and a rendered PDF passes a PDF/UA structure check.

### F-169, Agile encryption, read (L)
ECMA-376 Part 4 agile encryption: the `EncryptionInfo` stream, key derivation,
and AES decryption of the package. This is the difference between opening a
protected file and telling the user to go and find Word.
**Depends on**: none.
**Test gate**: regression. A password-protected document produced by Word opens
with the right password and fails cleanly with the wrong one.

### F-170, Agile encryption, write (M)
Save with a password, using the same parameters Word writes, so the result opens
in Word rather than only in this library.
**Depends on**: F-169.
**Test gate**: round-trip. A document encrypted here decrypts here, and the
parameters match a Word-encrypted reference byte for byte where the spec fixes
them.

### F-171, Digital signature verification (L)
Read `_xmlsignatures`, verify the signature over the declared part set, and
report which parts a signature actually covers, since a signature over a subset
is the usual attack.
**Depends on**: none.
**Test gate**: regression. A validly signed document verifies, and a document
modified after signing fails with the changed part named.

### F-172, Digital signature creation (M)
Sign a package with a supplied key and certificate.
**Depends on**: F-171.
**Test gate**: round-trip. A document signed here verifies here and in Word.

### F-173, Tagged PDF structure tree (L)
Emit `/StructTreeRoot`, marked content, heading levels, list structure, table
headers and alternate text from the document's own semantics, which the layout
engine already knows because `audit_accessibility` reads them.
**Depends on**: none.
**Test gate**: regression. A rendered PDF carries a structure tree whose heading
and list nesting matches the source document.

### F-174, PDF/A conformance (M)
PDF/A-2b and PDF/A-3b output: embedded fonts already, plus the output intent,
metadata and the prohibited-feature checks.
**Depends on**: F-173.
**Test gate**: regression. A rendered PDF passes a conformance check for the
declared level.

### F-175, Redaction (M)
Remove text and its traces rather than drawing a black box over it, covering the
body, comments, revisions, metadata and the embedded workbook of any chart.
**Depends on**: F-147, F-149.
**Test gate**: regression. Redacted text is absent from every part of the saved
package, checked by scanning the raw zip rather than the model.

---

## Milestone 18, Format breadth (about 5 weeks)

**Goal**: read and write the formats users actually have, rather than the one
format we prefer. Aspose.Words converts between roughly twenty. The gap that
costs real users is inbound: a library that cannot read RTF or HTML cannot be
put in front of a corpus nobody curated.

Rendering is already format-neutral below the facade, so every writer here is a
new front end onto a layout engine that exists.

**End-of-milestone gate**: each format round-trips at its declared fidelity
level, and every lossy conversion records a diagnostic naming what it dropped.

### F-176, RTF reader (L)
The native `rdocx` facade reads the Word-written RTF subset through
`Document::from_rtf_bytes` and `Document::open_rtf`. The reader owns bounded
control-word scanning, destination and group state, Unicode fallback handling,
font and colour tables, list tables and overrides, code-page decoding, table
rows, and PNG or JPEG picture projection. It converts text, run and paragraph
formatting, tables, lists, and images into the normal `Document` tree. Safe
lossy skips return stable diagnostics naming the dropped destination or
formatting control, while malformed RTF fails closed through `Error::Rtf`.
**Depends on**: none.
**Test gate**: differential. An RTF file converted to docx here matches the same
file opened and saved as docx by the pinned oracle, compared structurally.

### F-177, RTF writer (M)
The native `rdocx` facade writes the F-176 RTF fidelity boundary through
`Document::to_rtf_bytes` and `Document::save_rtf`. The writer allocates font,
colour, list, and image references deterministically, emits header tables
before body content, resets formatting at paragraph, run, cell, and row
boundaries, and writes non-ASCII text as signed UTF-16 RTF Unicode controls.
It preserves supported text, run and paragraph formatting, tables, multilevel
lists, and PNG or JPEG pictures with truncating goal dimensions. Unsupported
or lossy public properties and retained raw XML produce one stable
location-aware diagnostic while supported siblings continue. Output growth,
picture hex expansion, and diagnostics are bounded. Path saves serialize
first, stage a same-directory temporary file, sync it, and publish with the
shared portable replacement helper.
**Depends on**: F-176.
**Test gate**: round-trip. A document written to RTF and read back preserves
text, formatting, tables, lists and images.

### F-178, HTML import (L)
The native Word facade accepts bounded UTF-8 HTML5 documents and fragments from
strings or paths. Browser-grade tree repair projects source-ordered paragraphs,
runs, headings, block quotes, preformatted text, hard breaks, nested lists, and
spanned tables directly into the existing Word model. Inline declarations and
embedded type, class, id, descendant, and child selectors apply the supported
font, decoration, colour, alignment, spacing, and indentation subset by
specificity and source order. Stable path-aware diagnostics report parser
repairs, unsupported CSS, external resources, and dropped visible constructs.
Input, retained text, DOM, projection, table, and diagnostic ceilings fail
closed, and every candidate saves and reopens before publication. No resource
is fetched and no Python, WASM, or CLI API is added.
**Depends on**: none.
**Test gate**: regression. A fixture set of HTML documents produces the expected
paragraph, run, table and list structure, with unsupported CSS recorded as a
diagnostic.

### F-179, ODT reader (L)
The native Word facade reads bounded OpenDocument Text ZIP packages and
projects supported text, formatting, lists, tables, and images into a fresh
editable document. Archive names, encryption, compression, and expansion are
validated before namespace-aware XML parsing. Default, named, parent, and
automatic styles resolve into effective Word formatting. Stable source-path
diagnostics identify safe lossy skips, and fatal failures expose no partial
document. The ODT boundary is a private two-way facade conversion rather than
an OPC package or a retained second document model. Python, WASM, and CLI
surfaces do not gain ODT entry points.
**Depends on**: none.
**Test gate**: differential. A source-built ODT converted here matches the exact
pinned LibreOffice conversion by normalized body structure, formatting, lists,
tables, and image content without comparing package serialization details.

### F-180, ODT writer (L)
The native Word facade writes the F-179 fidelity boundary through
`Document::to_odt_bytes` and `Document::save_odt`. The private writer walks the
owned document tree without mutating it, materializes effective paragraph and
run formatting, emits nested lists and valid table spans, and copies supported
inline image bytes at their truncating EMU dimensions. Automatic styles,
media paths, manifest entries, ZIP metadata, and package order are
deterministic and bounded. Unsupported Word content returns stable path-aware
diagnostics while supported siblings continue. Path saves serialize first,
stage and sync a sibling file, and publish through the portable replacement
primitive. Python, WASM, and CLI surfaces do not gain ODT export methods.
**Depends on**: F-179.
**Test gate**: round-trip. Text, formatting, tables, lists and images survive.

### F-181, EPUB export (M)
The native Word facade exports bounded deterministic EPUB 3 bytes and atomic
path saves with stable lossy-conversion diagnostics. Outline roots split the
source-ordered spine. Pre-heading content becomes front matter, nested headings
remain nested navigation entries, and a document without headings produces one
item. The private writer packages semantic XHTML, shared CSS, metadata, and
relationship-resolved core PNG, JPEG, and GIF images through the existing `zip`
dependency. Media eligibility requires byte sniffing and structural validation,
never a filename fallback. SVG and malformed media are diagnosed and omitted.
Standard ordered marker formats remain semantic list styles. Stable diagnostics
cover unsupported marker details, table-cell list flattening, paragraph style
effects, deep headings, revision flattening, and dropped document metadata.
Numbered headings remain semantic headings inside their list items. Supported
image descriptions become alternative text, while other simplified drawing and
text-spacing properties are diagnosed. Page breaks are emitted as conforming
flow content, column breaks are diagnosed as simplified, and supported absolute
links pass an RFC 3986 syntax check before emission.
Heading labels exclude dropped content-control trees. Final section properties,
style-derived deep headings, non-basic underline styles, patterned or invalid
shading, table-cell shading, document backgrounds, visible default paragraph
style and document-default effects, and both preserved deleted-text losses have
stable diagnostics. Revision, change, and raw-only defaults are inert. Indexed
PNG palettes must fit the IHDR bit-depth capacity, and HTTP user information
accepts only RFC 3986 user-information characters.
Python, WASM, and CLI surfaces remain unchanged.
**Depends on**: none.
**Test gate**: regression. A source-built generated EPUB passes exact
EPUBCheck 5.3.0 and its spine matches the document outline.

### F-182, SVG page export (M)
A rendered page as SVG, from the same `PageFrame` the PDF and PNG backends
consume. Text stays text, so the output is searchable and scalable.
**Depends on**: none.
**Test gate**: golden. An SVG page rasterises to the same pixels as the PNG
backend at the same dpi, within the recorded tolerance.

### F-183, Image export options (S)
Multi-page TIFF, JPEG quality, transparent PNG backgrounds, and a page range on
every image entry point.
**Depends on**: none.
**Test gate**: regression. Each option produces the declared output and a page
range selects exactly the requested pages.

---

## Milestone 19, Advanced spreadsheets (about 17 weeks)

**Goal**: `rxlsx`, a loss-aware, headless spreadsheet lifecycle engine rather
than another cell reader or report writer.

**This milestone may supersede a recorded permanent non-goal only after its
go or no-go gate.**
`docs/hld/02-scope-and-non-goals.md` states that `oxml-sml` is not a spreadsheet
library and must not grow into one without a separate decision. F-184 is that
decision, and nothing else in this milestone may start before an affirmative
decision lands.

OPC, DrawingML, the chart engine, the layout engine and the PDF backend all
exist and are format-neutral, which lowers the cost of a third family. That is
not sufficient reason to build one. F-184 must reassess the Rust ecosystem when
S103 begins. M19 proceeds only if no credible maintained crate provides the
combined lifecycle required here: open an existing advanced workbook, preserve
what is not executed, edit typed features, recalculate formulas and local
pivots, refresh a declared Power Query subset, automate it through an Office
Scripts-compatible surface, save it, and render it without Excel.

Support is always classified as preserve, model and edit, or execute and
refresh. Unsupported execution never destroys the stored workbook state or
silently substitutes a result. Power Pivot and OLAP models, proprietary cloud
connectors, VBA, XLM, custom functions, Python cells, and Microsoft-hosted
Office Scripts storage remain preservation and diagnostic boundaries in this
milestone.

**End-of-milestone gate**: a representative advanced workbook round-trips
without losing unsupported parts, formulas and worksheet-backed pivots
recalculate to the pinned Excel values, selected Power Query M transformations
refresh through allowed connectors, an Office Scripts-compatible automation
fixture edits the workbook in a sandbox, and the resulting sheets and charts
render to PDF.

### F-184, Advanced spreadsheet go or no-go (S)
The go or no-go decision record. Reassess the maintained Rust spreadsheet
ecosystem at S103, state whether the combined lifecycle gap still exists, and
archive M19 if it does not. If it does, amend `02-scope-and-non-goals.md`, define
the boundary between `oxml-sml` as chart support and `rxlsx` as a library, and
publish the preserve, model, and execute classification for every advanced
feature in this milestone. Compare the planned boundary with `calamine`,
`rust_xlsxwriter`, `umya-spreadsheet`, `xls`, and any credible successor without
claiming that simple read or write support is a differentiator.
Assess the concrete demand in [Issue 281](https://github.com/tensorbee/rdocx/issues/281)
from `hadim`: template filling, small preservation-safe edits and a consistent
CLI/Python workflow with rendering. Compare integration, installation,
maintenance and release costs, including reuse behind a consistent facade.
Evaluate the utility and limits of a reader/editor stage before calculation
and pivots, retaining formula and pivot state with explicit stale-result
diagnostics when edits cannot be recalculated. The existing S112 distribution
boundary remains until a separate reviewed roadmap decision changes it.

The reporter's [follow-up](https://github.com/tensorbee/rdocx/issues/281#issuecomment-6063901190)
defines a Google Sheets review workflow: read and change cell text, edit a
shared-formula column and preserve reviewer comments. Ordinary spreadsheet
reading is already covered in that workflow. The differentiators are one
CLI/Python shape and installation, Rust performance and safety, loss-aware
editing and later rendering without Excel. The requested first cut preserves
all unmodeled parts, including threaded comments and `commentsmeta*`, data
validation, conditional formatting and defined names. It reads values and
types, shared strings, formats and styles, merged ranges, formulas and cached
results. Edits include values, formulas and styles, rows and columns, sheets,
widths and heights, frozen panes and auto-filters, with atomic output. Shared
formula structure must survive. The decision must choose explicit stale-cache
invalidation and recalculate-on-load, or a separately bounded evaluator for
arithmetic, references, SUM, AVERAGE, MIN, MAX and IF. Chart authoring is not a
requested first-cut need. Rendering is a desired second step. These are inputs
to F-184, not S90 implementation or an approved earlier release boundary.

The [maintainer commitment](https://github.com/tensorbee/rdocx/issues/281#issuecomment-6064533994)
confirms preservation tests and the review workflow first, then reading and
styled edits with explicit stale-cache handling. A bounded evaluator and
rendering follow, with chart authoring deferred. Pivot refresh, Power Query
and scripting remain longer-term goals. F-184 must reconcile the scheduled
stories and any proposed reader/editor distribution boundary with that order
before approving implementation. S90 records this direction without claiming
that the conditional spreadsheet programme or an earlier release is approved.
**Depends on**: none.
**Test gate**: regression. The scope document and capability matrix state one
non-contradictory boundary, and every scheduled spreadsheet story maps to a
declared preserve, model, or execute outcome.

### F-204, Spreadsheet corpus and compatibility matrix (M)
A pinned, licensed corpus of ordinary and advanced xlsx workbooks covering
formulas, tables, charts, conditional formats, validation, pivots, slicers,
Power Query metadata, external connections, script associations, and preserved
unsupported extensions. Record Excel and LibreOffice identity and expected
cached results separately.
**Depends on**: F-184.
**Test gate**: regression. The fetcher verifies every checksum and licence,
refuses an unpinned workbook, and reports the declared capability class for
every advanced part in the corpus.

### F-185, Workbook and worksheet model (L)
Workbook, sheets, rows, columns, cells, cell types, merged ranges and defined
names. The ownership model keeps unsupported package parts attached to their
relationships so an edit does not turn into a lossy rewrite.
**Depends on**: F-184.
**Test gate**: round-trip. Every element survives a load and save unchanged.

### F-186, Shared strings, styles and number formats (L)
The three tables that make xlsx compact and make naive implementations wrong:
the shared string table, `styles.xml` with its indexed formats, and the built-in
plus custom number format codes.
**Depends on**: F-185.
**Test gate**: round-trip. A workbook with every built-in format and twenty
custom ones preserves each cell's displayed value.

### F-189, Formula parser (L)
The A1 and R1C1 grammars, operators, ranges, cross-sheet and cross-workbook
references, and the shared-formula compression Excel writes.
**Depends on**: F-185.
**Test gate**: unit. Every formula in the corpus parses and re-serialises
identically.

### F-205, Excel tables and structured references (L)
Typed worksheet tables, totals rows, calculated columns, table styles,
autofilters, sorting, and structured formula references. Table growth and
column mutation update dependent ranges without rewriting unrelated worksheet
content.
**Depends on**: F-186, F-189.
**Test gate**: differential. Table edits, filters, totals, and structured
references save to the same effective values and ranges as the pinned Excel
oracle.

### F-206, Advanced worksheet objects (L)
Comments, hyperlinks, rich text cells, images, drawings, row and column groups,
hidden state, freeze panes, page breaks, sparklines, and modern image cells.
External content follows an explicit offline-by-default policy with allowed
schemes, limits, and diagnostics.
**Depends on**: F-186.
**Test gate**: round-trip. Every supported object remains typed and editable,
unsupported siblings remain byte-preserved, and external content is never
fetched without an explicit policy.

### F-187, Reader (L)
Streaming read of the sheet XML, because a spreadsheet is the one Office format
that is routinely too large to hold in memory as a tree.
**Depends on**: F-186, F-206.
**Test gate**: regression. A 100 MB fixture reads within a bounded memory
ceiling, asserted rather than assumed.

### F-188, Writer (L)
Streaming write, with the same ceiling.
**Depends on**: F-187, F-205.
**Test gate**: round-trip. A generated workbook opens in Excel without repair.

### F-190, Calculation engine (L)
Dependency graph, evaluation order, cycle detection, and the function set that
covers the overwhelming majority of real sheets: maths, statistics, text,
logical, date, lookup, dynamic arrays, spill ranges, and structured references.
**Depends on**: F-189.
**Test gate**: differential. Recalculated values match the values Excel stored
in a pinned corpus, cell for cell, with unsupported functions listed rather than
silently wrong.

### F-191, Charts in spreadsheets (M)
The chart part on a worksheet, reusing `oxml-chart` for the third time.
**Depends on**: F-156, F-185.
**Test gate**: round-trip. A chart on a sheet saves, reopens and renders.

### F-192, Conditional formatting and data validation (M)
Both are widely used and both are commonly dropped by libraries that claim
round-trip fidelity.
**Depends on**: F-186.
**Test gate**: round-trip. Every rule type survives with its ranges and
priorities.

### F-193, Pivot cache and table model (L)
Typed pivot definitions, cache definitions, cache records, row, column, data,
and filter fields, grouping, calculated fields, layouts, and worksheet or table
sources. External and OLAP sources remain attached and preserved when they
cannot be executed locally.
**Depends on**: F-188, F-190.
**Test gate**: round-trip. A workbook with worksheet, external, and OLAP pivots
preserves every source and cache, exposes the supported local model, and never
claims that an unavailable source refreshed.

### F-207, Pivot recalculation engine (L)
Refresh worksheet and table-backed pivot caches, aggregate supported fields,
apply filters and grouping, and regenerate the transient output cells after
source edits. Unsupported aggregation or source kinds retain their last cached
result with a diagnostic.
**Depends on**: F-193.
**Test gate**: differential. Mutating each source fixture and refreshing its
pivot produces the same fields, aggregates, filters, cache records, and visible
cells as the pinned Excel oracle.

### F-208, Slicers, pivot charts, and Data Model boundary (L)
Model and edit slicer caches, slicers, timelines, and pivot-chart relationships
over supported local pivots. Preserve and inspect Power Pivot and Data Model
parts, relationships, and measures without promising VertiPaq or DAX execution.
**Depends on**: F-191, F-207.
**Test gate**: differential. Slicer selections and pivot charts follow a local
pivot refresh, while Data Model parts remain relationship-complete and
byte-preserved after unrelated edits.

### F-209, Power Query package and M language (L)
Preserve and model workbook queries, connections, load destinations, refresh
metadata, and M source. Parse and evaluate the bounded M language core needed
for tables, records, lists, functions, `let` expressions, joins, grouping,
filtering, projection, and type conversion.
**Depends on**: F-188, F-190.
**Test gate**: differential. Corpus M programs parse and reserialize without
semantic drift, and pure transformations produce the pinned Power Query tables
or an explicit unsupported-function diagnostic.

### F-210, Power Query execution and connectors (L)
Execute an allowlisted first connector set for workbook tables, CSV, JSON, and
HTTP. Enforce credential isolation, privacy levels, source-combination rules,
timeouts, byte and row limits, deterministic caching, and offline operation.
Query folding is limited to connectors whose contract is explicitly tested.
**Depends on**: F-209.
**Test gate**: differential. Source-built refresh scenarios match the pinned
Power Query outputs, unsafe source combinations fail closed, and the same
fixture is deterministic when the network is disabled and cached input is
provided.

### F-211, Office Scripts artifacts and ExcelScript surface (L)
Model external `.osts` source and workbook associations without pretending the
script lives inside xlsx. Provide an explicitly versioned compatibility surface
for workbook, worksheet, range, table, chart, pivot, and query operations.
Microsoft OneDrive, SharePoint, Power Automate, and tenant identity remain
external services rather than hidden runtime dependencies.
**Depends on**: F-191, F-192, F-207, F-209.
**Test gate**: regression. Typed automation examples compile against the
declared compatibility surface, associations survive round-trip, and missing
external scripts produce diagnostics without modifying the workbook.

### F-212, Sandboxed Office Scripts runtime (L)
Execute the supported TypeScript and JavaScript subset against the same Rust
workbook model with CPU, memory, call-count, and output limits. Network access
is denied by default and uses the same explicit policy boundary as Power Query
when enabled.
**Depends on**: F-210, F-211.
**Test gate**: differential. Representative Office Scripts that edit ranges,
tables, charts, pivots, and query results match Excel's resulting workbook
state, while infinite loops, excessive allocation, unavailable APIs, and
unapproved external calls fail without partial mutation.

### F-194, Sheet rendering (L)
Page setup, print areas, repeating rows and columns, scaling, and the grid
itself, through the existing layout and PDF backends. Rendered output includes
supported conditional formats, drawings, charts, refreshed pivots, and print
objects.
**Depends on**: F-191, F-192, F-206, F-208.
**Test gate**: golden. A rendered sheet matches the pinned oracle render within
the recorded SSIM threshold.

### F-195, rxlsx distribution (L)
The facade, `rxlsx-cli`, `rxlsx-wasm` and the Python wheel, following the shape
M13 established for the other two families.
**Depends on**: F-188, F-194, F-212.
**Test gate**: regression. The parity suite passes on every target platform,
and each target reports the same unsupported feature and execution-policy
diagnostics.

---

## Milestone 20, Fidelity at scale (about 3 weeks)

**Goal**: prove the Word renderer against documents nobody here wrote.

PowerPoint fidelity is measured against 50 fetched decks with an SSIM harness.
Word fidelity rests on seven samples this project generates itself, so it can
catch a regression against its own output and can never catch a disagreement
with how Word actually renders. That asymmetry is the largest untested surface
in the workspace.

**End-of-milestone gate**: the Word corpus renders at the declared SSIM
threshold, and text shaping is correct for the scripts the corpus contains.

### F-196, Word corpus (M)
A pinned, fetched document corpus with the same provenance and licence
discipline as the deck corpus, covering business letters, reports, forms, legal
documents with revisions, and multi-script text.
**Depends on**: none.
**Test gate**: regression. The fetcher verifies every checksum and refuses a
corpus that does not match.

### F-197, Word SSIM harness (L)
The analogue of `pptx_ssim_harness.py`, comparing rendered pages against the
pinned oracle with the same trend-reference and hard-gate split.
**Depends on**: F-196.
**Test gate**: regression. The harness reports per-page SSIM, and a deliberate
layout change moves it.

### F-198, Hyphenation (L)
Liang hyphenation with language-specific patterns, which changes line breaking
and therefore every subsequent line. Word hyphenates and this renderer does not,
so any hyphenated document currently differs from the first hyphenated line
onward.
**Depends on**: F-197, F-X059, F-X066.
**Test gate**: golden. A hyphenated document matches the oracle's line breaks
within the recorded tolerance, and the harness delta is declared.

### F-199, Complex script shaping (L)
Arabic joining and shaping, Indic reordering and clusters, Thai breaking, and
CJK line-breaking rules. The shaper handles these and the line breaker does not
know their rules.
**Depends on**: F-196, F-X059.
**Test gate**: golden. Multi-script corpus pages match the oracle within the
recorded threshold.

### F-200, Vertical and bidirectional text (M)
Right-to-left paragraph direction, mixed-direction runs, and the vertical text
directions the deck renderer currently approximates.
**Depends on**: F-199, F-X059.
**Test gate**: golden. A bidirectional document renders with the correct visual
order.

### F-201, Large document performance (L)
A bounded memory ceiling and a stated throughput floor for a thousand-page
document, with the paginator and the renderer both measured.
**Depends on**: none.
**Test gate**: regression. A thousand-page fixture paginates and renders within
the asserted ceiling and floor.

### F-202, Incremental layout (L)
Re-lay out only what a mutation invalidated, rather than the whole document. The
layout cache added in F-009 is all or nothing, which is what makes an editing
session quadratic.
**Depends on**: F-201.
**Test gate**: regression. Editing one paragraph of a thousand-page document
re-lays out a bounded number of pages, asserted by counting layout invocations.

### F-203, Reader compatibility corrections (M)
Namespace-aware table-cell property recognition and schema-slot preservation for
numbering-level raw XML. Foreign same-local-name elements remain opaque,
byte-identical XML, and raw content before `w:suff` remains before that typed
element after a round trip.
**Depends on**: none.
**Test gate**: regression. Foreign `tcW` XML remains unmodelled and
byte-identical, and an `isLgl` raw child stays before `suff` after parse and
write.

### External Word benchmark, S102

Issue 377 reports that `neurotic_docx_bench` scores `rdocx convert` against
pinned Microsoft Word PDF exports. Its 2026-10-14 cohort ranks rdocx 0.16.0
eighth of ten at an ITT median of 45.90. These stories reproduce that score
locally, then close the largest measured gaps. The benchmark corpus is never
committed, because ODC-By covers the collection and not the individual files.

### F-313, External Word benchmark reproduction loop (S)
A pinned local run of the benchmark's own scorer over its 400-document
conversion set and Word PDFs, producing a per-document score, page-count and
failure table for one rdocx build. No Word installation is required.
**Depends on**: none.
**Test gate**: regression. The released 0.16.0 binary reproduces the published
ITT median within the scorer's recorded noise, and the per-document table is
written outside the repository.

### F-314, Benchmark conversion failures and startup latency (M)
Fix the three documents that fail to convert in the published cohort, and
measure where process startup time goes, including system font discovery,
before reducing it.
**Depends on**: F-313.
**Test gate**: regression. All three documents convert, and the measured
startup cost has a recorded before and after on the same machine.

### F-315, Word tracked-change markup print view (L)
An opt-in revision view that paints tracked changes the way Word's Save as PDF
does with all markup shown: per-author insertion and deletion colours,
underline and strikethrough, change bars and moved text. The default
accepted view and every hash baseline stay unchanged.
**Depends on**: F-313.
**Test gate**: golden. Benchmark documents with tracked changes and no comments
match their Word PDFs within the recorded tolerance, and default output is
byte identical to before this story.

### F-316, Word comment and revision balloon pane (L)
The markup pane Word draws when comments or formatting revisions are present:
page content scaled into the reduced area, comment and formatting balloons on
the right and their connector lines. Builds on the F-315 view.
**Depends on**: F-315.
**Test gate**: golden. Benchmark documents with comments match their Word PDFs
within the recorded tolerance, including the balloon geometry.

### F-317, Benchmark pagination and layout fidelity wave (L)
Group the lowest-scoring clean documents into root-cause clusters such as page
breaks, table splits, spacing and font substitution. Fix the largest clusters.
Every fix must also hold on the F-197 Word SSIM harness, so the work does not
overfit one Word build and one scorer.
**Depends on**: F-313.
**Test gate**: regression. The page-count mismatch rate and the clean-document
median both improve on the F-313 baseline, with no F-197 regression and every
hash delta declared.

### F-318, Benchmark release, adapter update and README badge (M)
Prepare the stable family release that carries F-314 through F-317. Propose
the benchmark adapter change that passes the markup view, request a rerun, and
add the README badge chosen from the Issue 377 options. Publication itself
stays with `/release` and its separate final approval.
**Depends on**: F-314, F-315, F-316, F-317.
**Test gate**: regression. The prepared release passes the full gate, and the
F-313 loop on the release candidate records the expected ranking before the
rerun request.

---

## Milestone 21, Presentation depth (about 12 weeks)

**Goal**: take the existing PowerPoint family beyond static business slides
while preserving a bounded, testable rendering contract.

The milestone covers modern PresentationML capabilities that already share the
OPC, DrawingML, chart, layout, media, security, and rendering foundations. It
does not add the legacy binary `.ppt` format. Executable VBA, ActiveX controls,
and arbitrary embedded objects remain inventory and preservation surfaces.

**End-of-milestone gate**: one representative modern deck round-trips its
comments, sections, authentic pinned-resource SmartArt, media, animation
timeline, signatures, and package variant without repair. Its static frames,
animated export, notes, and handouts
match the pinned PowerPoint 16.104 oracle at their declared fidelity boundaries.
An embedded manifest pins the no-repair signed canonical source, its exact
observed active file name, and its four directly bound outputs. The release
oracle consumes only the captured bundle from a configured directory. An
ignored macOS reference-only writer owns access to external authentic SmartArt
resources. The portable source-built signature proof has byte-identical
non-signature parts and relationships. One shared assertion applies the full
package, collaboration, section, media, playback, timing, signature, slide, and
SmartArt semantic contract to both exact source bytes and their save/reopen
result. Authentic mode rejects unsupported SmartArt fallback. All three static
pages and the three aligned movie samples require exact normalized token
cardinality and order, all non-media ink within 6 pixels at 150 DPI, and at
least 0.45 SSIM per complete ink region after masking only the page-one
audio-poster rectangle. The third static page must prove the complete SmartArt
graph and relationships plus visible three-node SmartArt text and ink. Notes
and handout page sizes are recorded absolutely. Notes require exact per-page
tokens and bounded monochrome-band cardinality. Corresponding semantic notes
components compare by normalized size within 0.06 and ink occupancy within
0.35. Placement is not equated across different notes masters and page sizes.
Handout text and thumbnail geometry compare in normalized coordinates within
0.05 of one page dimension.

### F-213, Animation and transition timing model (L)
Typed timing nodes, sequences, parallel groups, triggers, entrance and exit
effects, motion paths, transitions, and morph metadata. Unsupported timing
extensions remain relationship-complete raw XML.
**Test gate**: round-trip. The corpus timeline parses into the declared model,
serializes in schema order, and preserves every unsupported sibling.

### F-214, Timeline evaluation and transition rendering (L)
Evaluate supported timing trees into deterministic frame states and render
entrance, exit, emphasis, motion-path, ordinary transition, and bounded morph
effects without changing static slide rendering.
**Depends on**: F-213.
**Test gate**: differential. Pinned timestamps match the PowerPoint frame oracle
within the declared geometric and pixel tolerances.

### F-215, Audio and video package model (L)
Read, write, add, replace, extract, and remove linked or embedded audio and
video with poster frames, trim ranges, volume, looping, and playback triggers.
Unsupported codecs remain packaged and diagnosable.
**Test gate**: round-trip. Media bytes, relationships, playback settings, and
unsupported metadata survive save and reopen without duplication.

### F-216, Media poster and playback rendering (M)
Render poster frames and deterministic media placeholders in static output,
then expose synchronized media events to animated exporters. The library does
not decode a codec unless a named bounded backend supports it.
**Depends on**: F-214, F-215.
**Test gate**: golden. Static poster output and timestamped playback state match
the source-built oracle fixtures.

### F-217, Presentation collaboration and navigation model (L)
Typed comments and replies, slide sections, slide numbers, dates, footers,
notes headers, and handout settings with ordered mutation and preservation.
**Test gate**: round-trip. Every collaboration and navigation object survives
reordering, mutation, save, and reopen with its relationships intact.

### F-218, Embedded object and macro inventory (L)
Safe inventory, extraction, replacement, and removal for OLE objects, ActiveX
controls, and VBA projects. Executable content is never run, and signatures
are invalidated or preserved according to an explicit mutation policy.
**Test gate**: regression. Inventory reports exact hashes and relationships,
safe removal leaves a valid deck, and ordinary edits do not alter retained
payload bytes.

### F-219, SmartArt typed model (L)
Model diagram data, layout, style, colour, text, and relationship ownership for
the bounded SmartArt corpus while preserving unsupported algorithms.
**Test gate**: round-trip. Supported nodes remain editable and unsupported
diagram parts remain byte-preserved after unrelated mutations.

### F-220, SmartArt layout and rendering (L)
Resolve the six pinned authentic list, hierarchy, cycle, relationship, matrix,
and pyramid programs through bounded private instruction evaluation and the
shared DrawingML paint and text engines. The exact three-node `cycle1` resource
uses a private PowerPoint 16.104 compatibility profile that rejects any
identity, resource SHA-256, instruction, or node-count variation.
**Depends on**: F-219.
**Test gate**: differential. The common-source PowerPoint corpus retains exact
ownership, diagnostics, dimensions, and provenance, stays within 1 point for
shape bounds and 3 points for ordered text ink metrics, and reaches at least
0.90 symmetric text-masked non-text SSIM for every family.

### F-221, Presentation encryption and signatures (M)
Expose password-based read and write plus signature inspection, verification,
creation, and invalidation policy through the Presentation facade by reusing
the shared package security implementation.
**Depends on**: F-169, F-170, F-171, F-172.
**Test gate**: integration. Pinned PowerPoint opens encrypted output, signature
verification matches the trusted certificate fixtures, and mutation never
leaves a signature falsely reported as valid.

### F-222, ODP read and write (L)
Import and export slides, ordinary rectangles and text boxes, tables, embedded
images, slide names, and speaker notes through a declared OpenDocument fidelity
boundary. Charts, transitions, media, animation, SmartArt, and unsupported
appearance semantics produce stable diagnostics.
**Depends on**: F-214, F-215, F-217, F-220.
**Test gate**: differential. Source-built ODP and PPTX conversions match the
pinned LibreOffice structural and render records in both directions.

### F-223, Modern presentation package variants (M)
The native facade maps the six exact PPTX, PPTM, POTX, POTM, PPSX, and PPSM
main-part content types to `PresentationPackageClass`. Ordinary saves retain
the opened class. Output-specific conversion changes only the staged main
override, preserves opaque executable payloads and relationships, and records
retained package signatures as invalidated. Binary `.ppt` remains out of
scope.
**Depends on**: F-218.
**Test gate**: round-trip. PPTM, POTX, POTM, PPSX, and PPSM fixtures reopen in
their original package class with preserved executable payloads.

### F-224, HTML slide content import (L)
The native `rpptx` facade projects a bounded HTML5 and CSS subset into fresh,
editable slide shapes, formatted text, tables, caller-supplied images, and
links. Explicit absolute geometry, supported cascade semantics, stable DOM-path
diagnostics, and closed resource limits define the conversion boundary. The
candidate is serialized, reopened, and validated before publication. Browser
layout, scripts, external fetching, transforms, and unsupported CSS remain out
of scope and diagnostic.
**Depends on**: F-110, F-112.
**Test gate**: differential. Source-built HTML matches the browser reference at
the declared structure, text, one-pixel geometry, and 0.95 full-image luminance
SSIM boundary after save and reopen with Google Chrome 152.0.7977.65.

### F-225, PDF page content import (L)
Import PDF pages as either preserved page graphics or a bounded editable subset
of text, raster images, nonzero paths, and URI links. Strict bounded parsing,
CropBox and rotation normalization, equal effective page sizes, deterministic
font resolution, and transactional save, reopen, and validation define the
conversion boundary. Font substitution and unsupported PDF operators remain
explicit ordered diagnostics. Editable dash arrays require strictly positive
members and phase zero or an exactly representable dash boundary. Zero members,
interior phases, and positive members that convert to a zero DrawingML stop
diagnose and omit affected strokes until valid dash state or graphics-state
restore. JavaScript, encryption, malformed graphs, and declared resource-limit
failures are rejected.
**Depends on**: F-109, F-110, F-111.
**Test gate**: differential. Pinned PDF pages preserve page geometry and match
the source render at the declared Poppler 26.01.0, 150 DPI, exact-dimension,
and 0.995 raw full-image luminance SSIM boundary. The editable subset retains
text and link mappings. Geometry, one-pixel imported geometry, text, link, and
fill mutations prove the final predicate remains sensitive.

### F-226, Notes and handout export (M)
Render relationship-resolved speaker notes and all six audience handout grids
to deterministic PDF and PNG. Notes pages use `notesSz`, master-first overlay,
vector slide thumbnails, exact placeholder ownership, and typed slide numbers,
dates, headers, and footers. Handouts preserve the handout master below
aspect-fitted, clipped, bordered, and numbered thumbnails. The three-up layout
adds ruled writing space.
**Depends on**: F-217.
**Test gate**: source-built deterministic regression. Noncanonical relationship
targets, cross-scope id collisions, absent notes, placeholder ambiguity, all six
layouts, exact vector geometry, PDF text and page count, PNG dimensions and
pixels, 1.01-point sensitivity, and source-byte preservation pass. The 49-entry
render hash manifest remains unchanged.

### F-227, Animated GIF and video export (L)
Sample deterministic timeline states into animated GIF and a bounded video
backend with explicit frame rate, duration, resolution, transition, and media
fallback policy.
**Depends on**: F-214, F-216.
**Test gate**: golden. Frame hashes, timestamps, loop behavior, and output
dimensions match the reviewed manifest on two machines.

---

## Milestone 22, Word depth (about 9 weeks)

**Goal**: complete the modern WordprocessingML features already identified as
valuable without opening a legacy document-format programme.

The milestone deepens modern DOCX, DOCM, DOTX, and DOTM workflows. Binary
`.doc`, Word 2003 XML, and other pre-OOXML formats remain permanent non-goals.
Macro projects and embedded executable content are preserved and inspectable,
never executed.

**End-of-milestone gate**: a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads.

### F-228, OfficeMath model and authoring (L)
`rdocx-oxml` owns one Transitional OfficeMath tree for inline and display
equations, runs, fractions, scripts, radicals, matrices, limits, n-ary
operators, delimiters, accents, and document-wide math defaults. The native
`rdocx` facade exposes source-ordered borrowing, indexed mutation, bounded
authoring, and relationship-resolved settings access. Prefix aliases are read
by expanded name, fixed `m:` output is schema ordered, unsafe namespace
collisions fail closed, and unsupported or legacy content remains raw.
**Test gate**: round-trip.
`officemath_corpus_parses_mutates_saves_and_reopens_without_losing_supported_or_raw_siblings`
covers every supported expression plus opaque root, property, and argument
siblings through mutation and reopen.

### F-229, OfficeMath layout and PDF rendering (M)
Lay out supported equations through the shared font and page-frame boundary
with baseline, stretch, delimiter, and operator sizing. The Word engine lowers
the typed tree to shared groups, text, lines, and paths, carries optional global
math defaults through `LayoutInput`, and keeps existing top-aligned group
behavior when no baseline is present. Deterministic rendering uses bundled
Caladea and a source-built Word PDF oracle. Tagged-PDF math semantics remain
outside this story.
**Depends on**: F-228.
**Test gate**: golden. Equation baselines and glyph geometry match the pinned
Word 16.104 PDF oracle within 1.0 point, and the complete 150 DPI page meets the
declared luminance SSIM floor.

### F-230, MathML and LaTeX conversion (M)
The native `rdocx` facade imports and exports the supported normalized
OfficeMath subset through bounded Presentation MathML and LaTeX converters.
Four free functions return the existing `MathArgument` tree or canonical text
with stable ordered loss diagnostics. MathML uses expanded W3C names, accepts
`mfenced` on input, and emits explicit fences. LaTeX uses a local bounded
recursive-descent parser. Python, WASM, and CLI surfaces remain unchanged.
**Depends on**: F-228.
**Test gate**: differential.
`mathml_and_latex_conversion_matches_pinned_pandoc_texmath_trees` checks
source-built equations structurally in both directions against exact Pandoc
3.10 and records its intentional wrapper divergences.

### F-231, Extended field evaluation (L)
Evaluate TOC, TC, formula, mail-merge control, and barcode fields while
retaining unavailable field instructions and cached results.
**Depends on**: F-161, F-162.
**Test gate**: differential. Supported field results match the pinned Word
values, and unsupported instructions remain intact with diagnostics.

### F-232, Dynamic table of contents rebuild (L)
Rebuild an existing TOC from headings, custom styles, outline levels, TC
entries, bookmarks, and page numbers without replacing its unrelated field
formatting.
**Depends on**: F-154, F-231.
**Test gate**: differential. Heading, style, and TC mutations produce the same
entries, links, levels, and page numbers as the pinned Word update.

### F-233, Advanced mail merge (L)
Add merge regions, nested records, multiple named data sources, images,
document fragments, and caller-provided formatting hooks to the existing merge
engine.
**Depends on**: F-166.
**Test gate**: regression. Nested source-built records generate the expected
ordered paragraphs, lists, tables, images, and formatting without stale fields.

### F-234, Full-story document comparison (L)
Extend comparison through headers, footers, comments, fields, text boxes,
footnotes, endnotes, and formatting while preserving story order and source
mappings.
**Depends on**: F-167.
**Test gate**: differential. The pinned document pairs produce the same
insertions, deletions, moves, and story placement as Word at the declared
boundary.

### F-235, Comparison granularity and ignore policy (M)
Add character and word granularity plus explicit ignore rules for formatting,
whitespace, fields, comments, and selected stories.
**Depends on**: F-234.
**Test gate**: regression. Each policy changes only the declared comparison
records and remains deterministic under repeated runs.

### F-236, Embedded object and macro inventory (L)
Inventory, extract, replace, and remove embedded objects, VBA projects, and
their signatures without executing payloads or weakening package preservation.
**Depends on**: F-171, F-172.
**Test gate**: regression. Inventory hashes and relationship paths remain
stable, safe removal leaves a valid document, and unrelated edits preserve
payload bytes.

### F-237, Forms, glossary, and building blocks (L)
Typed inventory and bounded mutation for legacy form fields stored inside
modern OOXML, glossary entries, AutoText, and building blocks. This does not
add a binary `.doc` reader.
**Test gate**: round-trip. Supported entries remain editable and every
unsupported subtree survives unrelated document edits.

### F-238, Flat OPC and modern Word package variants (M)
Read and write Flat OPC plus DOCM, DOTX, and DOTM while preserving package
identity, macros, templates, relationships, and content types. Word 2003 XML
and binary `.doc` remain out of scope.
**Depends on**: F-236, F-X077, F-X079.
**Test gate**: round-trip. Each modern package class reopens without repair and
retains its executable payload and template semantics.

### F-239, MHTML import and export (M)
The native Word facade converts its supported document surface to and from one
bounded `multipart/related` MHTML entity. Import resolves only unique contained
Content-ID and Content-Location resources, never fetches external content, and
publishes after DOCX save and reopen. Export is deterministic, deduplicates
equal image bytes, reparses before publication, writes paths atomically, and
reports stable loss diagnostics.
**Depends on**: F-178.
**Test gate**: differential. Source-built MHTML and DOCX conversions preserve
body order, formatting, tables, lists, images, links, and declared loss records
against Microsoft Word 16.104 build 16.104.25121423.

---

## Milestone 23, From-scratch business documents (about 22 weeks)

**Goal**: generate the five private proposal and order-form references from
`Document::new()` through the public `rdocx` facade. The generation path uses no
base DOCX, caller-supplied raw XML, direct `rdocx-oxml` mutation, or LibreOffice
field-update pass.

The first sprint is an evidence-gathering and planning boundary. Its audit owns
the final capability matrix and the remaining M23 and M24 work. Client
documents, extracted parts, and rendered pages stay in a private ignored
directory. Only sanitized source-built fixtures and non-identifying requirements
may be tracked.

**End-of-milestone gate**: all five private references are generated from a
blank public facade, reopen without repair, match the required package
semantics, and meet the reviewed deterministic visual thresholds. Repeated
generation produces identical DOCX bytes, and no generated document reports an
unexplained preservation-only fallback.

### F-240, Modern DOCX completeness audit and private corpus matrix (L)
Audit every modern DOCX capability across create, read, mutate, remove,
save-reopen, story placement, layout, rendering, determinism, bindings, and
diagnostics. The closed matrix contains 85 stable rows and maps the five private
documents only as anonymous, non-identifying capability families. The audit
found no duplicate scope, missing owner, dangling dependency, dependency cycle,
or scheduling conflict in F-243 through F-310. Their boundaries, sizes, and
dependencies remain authoritative. The current plan places them in S71 through
S75, S86, S89 through S91, S95.2 through S95.4, and S99 through S101.
**Capability matrix owner**: `docs/hld/02-scope-and-non-goals.md`, "Modern DOCX
capability matrix".
**Test gate**: regression. Every in-scope matrix row has evidence, an owner
story, an explicit preservation boundary, or a permanent non-goal, and every
roadmap duplicate or dangling dependency is rejected.

### F-241, Public authoring conformance harness (L)
Build the reusable source-built and private-corpus harness for public-facade
generation, package comparison, save-reopen checks, deterministic rendering,
and unsupported-content diagnostics. Private client inputs and outputs remain
ignored, and tracked fixtures contain no identifying text, media, or extracted
package bytes.
**Depends on**: F-240.
**Test gate**: differential. A sanitized representative fixture proves every
gate locally, while missing private inputs report a clear skip rather than
weakening required private-corpus mode.

### F-242, Root README product and capability overview (M)
Simplify the root README into a current product entry point with major feature
categories, an evidence-based comparison with common alternatives, the
authoring and preservation distinction, and a concise roadmap. Detailed status
continues to live in the canonical backlog and sprint plan, and every public
feature story updates the summary when its classification changes.
**Depends on**: F-240.
**Test gate**: regression. README examples compile, links resolve, version
requirements match manifests, and every capability claim maps to the approved
matrix.

### F-243, Word-compatible fresh package profiles (L)
`Document::new()` creates a deterministic Word-compatible DOCX package.
`Document::new_with_profile` selects minimal or Word-compatible DOCX, DOCM,
DOTX, and DOTM graphs that own their declared parts, content types,
relationships, and package metadata without copying a template. Macro-capable
profiles do not invent VBA, and minimal output remains explicit.
**Depends on**: F-240.
**Test gate**: round-trip. Each profile saves, reopens with the same identity,
and passes strict package validation and the pinned no-repair check.

### F-244, Corpus settings and document properties (L)
Author the settings, core properties, application properties, custom
properties, document variables, compatibility facts, and defaults required by
the private corpus. The bounded settings set is compatibility settings, default
tab stop, character spacing control, document variables, and theme font
languages. Values are typed, deterministic, and removable through the public
facade. Core, application, and custom properties use relationship-resolved
package parts with selective custom-property and whole-part removal.
**Depends on**: F-243, F-249.
**Test gate**: round-trip. Every authored value survives save and reopen, and
removal deletes only its owned package content.

### F-245, Corpus themes, font tables, and embedded fonts (L)
The native facade creates and selects the shared DrawingML theme, language
defaults, font-table records, font relationships, and licensed embedded-font
parts required by the private documents. Font embedding accepts caller bytes
only after explicit authorization, preserves the exact license identity, and
uses the caller's OOXML font key for deterministic obfuscation and layout.
**Depends on**: F-243, F-249.
**Test gate**: differential. Public-authored theme and font resolution matches
the pinned Word references in deterministic layout without system-font input.

### F-246, Corpus style authoring (L)
Create, update, remove, and resolve the paragraph, character, and table styles
used by the private corpus, including defaults, inheritance, linked and next
styles, and conditional table regions. Style references are validated before a
transaction publishes.
The native facade uses fallible create, update, default-selection, removal, and
validation operations over one staged style graph. Parser and serializer
support includes the style UI and locking flags in schema order while retaining
unmodelled XML.
**Depends on**: F-243, F-245.
**Test gate**: differential. The source-built style graph resolves to the same
effective formatting and visible output as the sanitized Word oracle.

### F-247, Complete numbering level and instance model (L)
Model and author full numbering levels and instances, including level text,
suffix, alignment, indentation, marker properties, legal numbering, restart
controls, level overrides, and start overrides. Imported paragraph-style links
are typed, inspectable, and preserved. Their two-sided mutation belongs to
F-248. The public format type accepts the complete standard set, including
`none`, without requiring raw XML. Definition and instance CRUD validates the
complete candidate graph before publishing.
**Depends on**: F-243, F-249.
**Test gate**: round-trip. Every typed level and override survives save and
reopen with schema-correct order and reports no unmodeled properties when
created solely through the public API.

### F-248, Style-linked numbering, counters, TOC, and REF (L)
The native facade provides atomic link and unlink operations that write the
numbering-level style link and the style's numbering properties together.
Result-local counters are keyed by concrete instance and level and cover
continuation, overrides, restarts, table cells, sections, suppression through
`numId` zero, numbered TOC entries, and numbering-aware REF switches. Distinct
concrete instances start independently, including when they share an abstract
definition. This is an intentional divergence from the captured Word 16.112.3
shared-definition behavior.
**Depends on**: F-246, F-247.
**Test gate**: differential. Three-level numbered headings inside and outside
tables match Word 16.112.3 in body text, TOC entries, cross-references, and
restart behavior. Twenty-seven exact Word records and three normalized TOC
records are the semantic authority. The two-page 150 DPI Word PDF oracle
requires at most one pixel of raster dimension difference, at least 0.95 paired
ink coverage, at most 0.08 normalized ink bounding-edge delta, at most 0.27
total variation across a 32-region ink distribution, and at most 0.04 row or
column projection distance on every page. A synthetic five percent shift must
exceed the projection threshold. The recorded minimum coverage is 0.983131.
The recorded maximum edge delta is 0.062144, the maximum distribution delta is
0.257661, and the maximum projection distance is 0.038314. The recorded minimum
synthetic-shift distance is 0.049960.

### F-249, Deterministic package identifier allocation (M)
The Word facade owns deterministic allocation for relationships, bookmarks,
comments, drawings, numbering definitions, numbering instances, parts, and
content types. Category scopes remain independent, relationships are scoped by
source part, and allocation follows final recursive document order. Package
open rejects duplicate normalized part names, content-type identities,
relationship identifiers, and typed or preserved XML definitions. Mutations
and serialization publish only a complete staged candidate.
**Depends on**: F-243.
**Test gate**: regression. Equivalent construction orders produce the declared
stable identifiers, repeated saves are byte-identical, and collision or
overflow failures are atomic.

### F-250, Ordered mutable section facade (L)
The native facade exposes every paragraph-level and schema-final section owner
in document order through concrete immutable and mutable handles. Handles carry
their ordinal and final-owner identity, inspect orientation, and retain complete
property access. Mutable handles normalize page dimensions for orientation
changes and configure first-page behavior. Total lookup, insertion, and removal
stage the complete document and package without a second section tree. Removal
preserves effective same-variant header and footer inheritance, body order, and
unmodelled XML, and prunes only unreachable facade-owned related stories.
**Depends on**: F-249.
**Test gate**: round-trip. A portrait, landscape, portrait document retains
three ordered sections and proves each independent or inherited story variant,
relationship type, target, and content after mutation, save, and reopen.

### F-251, Complete section and page geometry (L)
The ordered section handles read and author M23 page size, orientation, margins,
gutter, equal-width columns, page-number start, header and footer distance,
title-page state, and break type through checked atomic setters. Legacy
final-section document setters remain unchecked and infallible. The OXML model
authors only `w:pgNumType/@w:start`, preserves number format, chapter style,
chapter separator, and unsupported children for M24, and replays raw children
at schema and repeated-reference boundaries. Distinguishable references follow
their value. Indistinguishable equal duplicates use a deterministic source
ordinal without changing the public `Vec<HdrFtrRef>` fields. Pagination keeps
physical and displayed page identities separate, substitutes PAGE from the
displayed value, and continues final-section numbering onto endnote pages.
**Depends on**: F-250.
**Test gate**: differential.
`mixed_orientation_sections_match_word_geometry_and_page_numbers` matches
Microsoft Word 16.112.3 build 16.112.26083020 for three exact page geometries,
physical identities, and displayed PAGE values.

### F-252, Rich per-section headers and footers (L)
The native facade creates, links, unlinks, inherits, replaces, and removes
default, first, and even header and footer stories per section. Effective
lookup follows same-type inheritance and reports the source section. Removal
authors an explicit empty story, while inheritance removes the direct
reference. Unlink and replacement clone the complete part-local relationship
set, rebase internal targets, and freshen drawing identities. Each story accepts
paragraphs, tables, fields, links, images, drawings, and nested supported
content through the common public container operations. First-page creation
enables `titlePg`, and even-page selection has an explicit document setting.
**Depends on**: F-250, F-253.
**Test gate**: differential.
`section_header_footer_variants_match_word_width_and_inheritance` matches
Microsoft Word 16.112.4 build 16.112.26090911 across nine pages at three exact
widths and proves default, first, and even selection through direct, inherited,
and replaced stories.

### F-253, Container-neutral story editing (L)
The native Word facade provides one public content-location and mutation model
for the body, table cells, headers, footers, ordinary notes, comments, and text
boxes. Deterministic owner and item traversal exposes paragraphs, tables,
controls, fields, drawings, and preserved nodes over existing typed and package
sources without introducing a second document tree. Checked text mutation
resolves an operation-scoped location on a staged package and leaves the
document unchanged on every location or serialization failure.
**Depends on**: F-240.
**Test gate**: integration. One generic mutation visits and edits the same
supported content shape in every story with identical error behavior.

### F-254, Generic insert, move, clone, and remove operations (L)
The native Word facade performs transactional insertion, removal, cloning, and
same-owner movement of direct story content through owned `ContentFragment`
values. Existing destinations use canonical actual-item locations that resolve
to direct owner children. `ContentLocation::end` represents the boundary after
final direct content, including empty and self-closing owners, and stays before
body section properties. Clones allocate fresh document identities.
Relationship-bearing fragments require the unchanged owner scope. Invalid,
stale, cross-owner, ambiguous, or structurally incomplete operations leave the
document unchanged.
**Depends on**: F-253, F-249.
**Test gate**: regression. Interleaved operations across body and cell content
preserve exact order, references, and untouched raw XML.

### F-255, Part-scoped assets, links, and relationships (M)
The native Word facade resolves image, hyperlink, chart, and ordinary related
content through an explicit OPC owner rather than a conventional main-document
path. Public story-scoped picture and hyperlink insertion, relationship lookup,
and internal relationship validation use checked `StoryId` owners. Cells and
text boxes inherit the containing part. Header, footer, note, and comment
stories use their resolved related parts. Staged serialization reconciles live
relationship occurrences, preserves same-owner shared references, assigns each
authored picture occurrence a global drawing identity, and publishes only a
validated reopened candidate.
**Depends on**: F-253, F-249.
**Test gate**: round-trip. Equal content in body, header, footer, note, and text
box stories resolves only through its correct owner relationships.

### F-256, Transactional cross-document fragment import (L)
The native facade imports a nonempty half-open main-body selection while
remapping selected styles, numbering, bookmarks, comments, media, drawings,
charts, embedded workbooks, fields, and recursive internal relationships.
Caller-selected equivalent reuse or renaming is deterministic for styles,
numbering, and related parts. Exact retained body and comment XML remains
package-authoritative, and unsupported, external, dangling, malformed,
split-range, or exhausted dependencies abort without changing the destination.
**Depends on**: F-246 through F-255.
**Test gate**: regression. A dependency-rich fragment imports twice without
collisions and reopens with every reference resolved.

### F-257, Complete M23 table authoring (L)
The public native table facade authors auto, fixed, and percentage width modes,
alignment, indentation, layout, shading, aggregate and individual borders,
cell margins, complete active grids, and style-look flags. Checked mutation
validates units, colors, grid omissions, spans, coverage, and overflow before
publication. Grid replacement synchronizes table, column, and covering-cell
widths. Invisible borders remain explicit modeled values rather than absence
inferred by convention, and unrelated producer XML remains exact.
**Depends on**: F-253.
**Test gate**: differential. `m23_layout_and_data_tables_match_word` proves the
sanitized layout and data-table XML semantics, column geometry, reviewed
pagination record, and deterministic rendering. Companion regressions prove
complete typed reopen and failure atomicity.

### F-258, Complete M23 row and cell authoring (L)
The native facade authors checked exact and minimum row heights, explicit
repeating-header and split toggles, row alignment and conditional regions,
grid omissions, horizontal and vertical merges, per-cell borders and margins,
width, shading, vertical alignment, six text directions, conditional regions,
wrapping, and checked nested tables. Topology changes reconcile only untouched
empty cells and publish only a complete validated table. Aliased modeled
properties remain readable, foreign same-local children remain raw, and
unrelated property and border-extension slots remain exact.
**Depends on**: F-257, F-X100.
**Test gate**: differential. `m23_nested_rows_and_cells_match_word` proves the
sanitized nested-table property matrix, canonical schema order, typed reopen,
merge topology, and deterministic rendering. Companion regressions prove
failure atomicity and exact producer XML preservation.

### F-259, Container measurement and equal-height layout (M)
Measure supported paragraphs and tables at a caller-supplied width using the
same deterministic fonts and layout rules as whole-document pagination. The
owned `ContentMeasurement` reports fractional point height and ordered layout
diagnostics. `Document::measure_content` accepts an existing checked
`ContentLocation`, positive `Length` width, and `RenderOptions`. It resolves one
paragraph or table, builds the ordinary layout input, and runs the production
paragraph or recursive table path through a fresh deterministic engine without
pagination, mutation, or cache publication. Callers round upward at the twip
boundary before applying an equal minimum row height.
**Depends on**: F-257, F-258.
**Test gate**: regression. Two independently measured nested tables align to
the same final row height and match whole-document layout. Companion coverage
proves wrapping, margins, borders, spans, ordered diagnostics, checked failure,
byte purity, and deterministic-cache purity.

### F-260, Ordered run content authoring (L)
Author tabs, line, page, and column breaks, drawings, fields, symbols, and text
inside one run while preserving mixed-content order. Existing formatting
setters operate on the run without replacing non-text children. Fields are
normalized to schema-valid paragraph children on write, with the direct run
properties copied to the cached result and each surrounding physical run.
**Depends on**: F-253.
**Test gate**: round-trip. A run containing every supported child reopens in
the same order and renders each break and tab at the expected position.

### F-261, Rich HTML fragments in arbitrary containers (L)
`Document::insert_html_fragment` inserts bounded inline and embedded CSS,
nested lists, tables, data-URI images, links, and exact-key explicit image
resources at a checked body, cell, header, or footer `ContentLocation`. The
owned result returns a refreshed story identity, direct range, and ordered
diagnostics. Destination-scoped numbering, media, relationships, and drawing
identifiers publish atomically after reopen. No external resource is fetched.
**Depends on**: F-253 through F-260.
**Test gate**: differential. The supported fragment subset produces equivalent
Word content and rendering in body, cell, header, and footer containers.

### F-262, Corpus drawings, text boxes, and watermarks (L)
`PictureOptions` authors inline and floating images with crop, size, relative
position, wrapping, distances, z-order, and behind-text control. Checked story
text boxes carry rotation and all three supported text directions in a WPS
primary branch with a self-contained VML fallback. Section-aware text
watermarks target one header variant without changing caller-owned parity
settings. Every mutation publishes only after save and reopen. Raw header XML
is not used by the generator.
**Depends on**: F-252, F-255, F-260.
**Test gate**: differential. Every private-corpus drawing and watermark matches
the reviewed Word geometry and compatibility structure. Companion exact XML,
relationship, option-matrix, header-selection, and rollback tests cover the
portable gate when Word GUI automation is unavailable.

### F-263, Layout-backed fields and M23 corpus gate (L)
Materialize PAGE, NUMPAGES, PAGEREF, and supported TOC caches from deterministic
layout, including numbered headings inside table cells. The native and Python
operations publish atomically and return an owned report with separate field
counts and ordered diagnostics. Build all five private documents through pure
Rust public `Document::new()` programs, rerun each generator for exact byte
determinism, and enforce package, semantic, visual, repair, and no-runtime-
fallback acceptance.
**Depends on**: F-241 through F-262.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/93>.
**Test gate**: differential. Required private-corpus mode passes all five
references without LibreOffice post-processing, raw XML, or source-template
access at generation time.

---

## Milestone 24, Modern DOCX authoring completeness (about 44 weeks)

**Goal**: close the modern, non-executable DOCX authoring surface after the
private corpus proves the architecture. Every in-scope capability is created,
read, mutated, removed, saved and reopened through public APIs, placed in every
valid story, rendered where the product claims rendering, and classified across
native, Python, WASM, and CLI surfaces.

Binary `.doc`, Word 2003 XML, VBA execution, ActiveX execution, OLE application
execution, hosted add-in execution, Microsoft cloud services, and undocumented
Word layout parity remain permanent non-goals. Their modern-package payloads
may be preserved, inventoried, attached, extracted, replaced, removed, and
diagnosed without execution.

**End-of-milestone gate**: the approved capability matrix has no unexplained
partial row. A broad source-built document and the pinned Word corpus pass
strict and transitional validation, public authoring, mutation, save-reopen,
deterministic layout and rendering, accessibility, package preservation, and
binding-parity checks without Word repair.

**Tracked human action**: Word GUI capture is not available on the development
machine, so the "without Word repair" confirmation is performed by hand at this
gate and recorded as performed or not performed. No automated story test
depends on it, and no gate skips in its absence.

### F-264, Complete paragraph property authoring (L)
Expose the full supported paragraph-property model through public setters and
readers, including logical indentation, automatic spacing, borders, shading,
tabs, pagination, frames, outline, direction, and paragraph-mark properties.
The paragraph grammar owns `w:divId` and its paragraph accessor, while F-270
owns the matching `CT_WebSettings` projection. It also owns ordered attribute
retention on `CT_BorderEdge`, which F-269 consumes for page borders.
Positioned frame placement is layout work this story does not build, so the
`DOCX-030` layout and render columns stay partial.
**Depends on**: F-253.
**Test gate**: round-trip.
`every_public_paragraph_property_reopens_and_preserves_unrelated_xml` proves
every public-authored paragraph property reopens as modeled content and
preserves unrelated producer XML.

### F-265, Complete run property and inline authoring (L)
Expose full run fonts, theme references, colors, complex-script formatting,
shading, effects, language, symbols, special characters, and ordered inline
content. Explicit-font replacement has a documented theme-clearing policy.
The story also owns the `w:rFonts`, `w:color`, and `w:shd` theme-attribute
sweep, and it gives `rdocx_oxml::theme::apply_tint_shade` its first production
caller without changing its arithmetic. Stroke, relief, character-border,
kerning, and fitted-text rendering is layout work this story does not build, so
the `DOCX-032` layout and render columns stay partial.
**Depends on**: F-260.
**Test gate**: differential. Effective run formatting and inline ordering match
the pinned Word reference across save, reopen, and render.

### F-266, International and vertical typography (L)
Author and render bidirectional, East Asian, complex-script, vertical, ruby,
phonetic, emphasis-mark, character-grid, and locale-sensitive text behavior.
**Depends on**: F-264, F-265.
**Test gate**: golden. Mixed Arabic, Hebrew, Korean, Japanese, and Latin pages
match the pinned deterministic geometry and reading order.

F-266 is split into the three implementation stories below. The parent closes
only after every child closes. The split was taken in the S74 consolidated
design round, because the six work groups separate cleanly and the bundled
font decision belongs to the first child alone.

### F-266a, Script identity and font slot resolution (L)
Hangul and Kana script identity, `w:rFonts` script-slot font resolution, the
East Asian and complex-script theme references, and the bundled deterministic
Hebrew, Korean, and Japanese subset faces.
**Depends on**: F-264, F-265.
**Test gate**: golden.
`mixed_script_page_matches_the_pinned_geometry_and_reading_order` pins the
mixed Arabic, Hebrew, Korean, Japanese, and Latin page geometry and reading
order in deterministic font mode.

Two boundaries are recorded rather than closed here. Shaping does not cross a
`w:r` boundary, so one Arabic word split across two runs loses its joining
forms, because Word multilingual reassembly requires every shaped span to stay
inside the inline item it came from. Lifting that is a redesign of the
reassembly contract and belongs to its own story. Separately, a paragraph on
the rich shaping path cannot enter the paragraph block cache, which now
includes Korean alongside Arabic, Hebrew, and CJK.

### F-266b, Ruby and emphasis marks (L)
`w:ruby` typed paragraph content with its base and phonetic lines, and `w:em`
emphasis marks projected into layout.
**Depends on**: F-266a.
**Test gate**: golden.
`ruby_and_emphasis_page_matches_the_pinned_geometry_and_reading_order` pins the
ruby and emphasis page and asserts the F-266a digest is unmoved.

`w:ruby` models `w:rubyPr`, `w:rt` and `w:rubyBase` in
`crates/rdocx-oxml/src/ruby.rs`, reusing `CT_R` for both lines rather than
introducing a second run grammar. The base runs live in the paragraph's own run
list behind a recorded span, so text extraction, search and redaction see the
base and never the phonetic line. `w:em` was already modeled by F-265, so this
story adds its render projection alone.

Two limits are recorded rather than closed here. An annotated span is one
unbreakable inline item, so a ruby never breaks inside its base and an
emphasis-marked run breaks only at the whitespace it was split on. A marked run
also leaves the rich shaping path, which costs nothing for the scripts `w:em`
applies to and would cost a slice of every glyph cluster to avoid.

### F-266c, Character grid and vertical text (L)
`w:eastAsianLayout`, `w:docGrid`, the East Asian paragraph toggles, and the
`w:textDirection` render projection for vertical cell and section text.
**Depends on**: F-266a, F-269.
**Test gate**: golden.
`grid_and_vertical_page_matches_the_pinned_geometry_and_reading_order` pins
the vertical and character-grid page and asserts the earlier digests are
unmoved.

**Delivered**: `CT_SectPr` gains a typed `w:docGrid` at its own `xsd:sequence`
slot, and `w:printerSettings` moves to the slot after it so the producer
children that shared the old raw slot still round-trip byte for byte. The
section grid reaches paragraph and table layout as a threaded value and a cache
key, so a gridded and an ungridded section can never share a cached block.
`lines`, `linesAndChars` and `snapToChars` put line advance on `w:linePitch`,
the last two add `w:charSpace` to every character advance, and `default` stays
off the new arithmetic entirely. `w:eastAsianLayout`, modeled by F-265, gains
its render projection: `w:combine` compresses the run into one base-character
advance inside the `w:combineBrackets` pair, `w:vert` rotates it 90 degrees
within the line, and `w:vertCompress` narrows the rotated run to one advance.
`w:tcPr/w:textDirection` and `w:sectPr/w:textDirection` both lower onto a
same-centre transposed box wrapped in a rotated `Group`, and a rotated cell
contributes the transposed box's measure to its row height. The seven East
Asian paragraph toggles F-264 left raw-preserved are typed on `CT_PPr` with a
public authoring surface, and `w:snapToGrid` gates the grid. The six
line-breaking policy toggles carry no break-opportunity projection, which is
recorded on the DOCX-033 row.

### F-267, Complete table style and conditional formatting authoring (L)
Create and mutate table styles, conditional regions, band sizes, table look,
row and cell conditional selectors, and their paragraph, run, table, and cell
property layers.
**Depends on**: F-246, F-257, F-258.
**Test gate**: differential. Every conditional region resolves and renders like
the pinned Word-authored table.
**Delivered**: `CT_TblStylePr` gains its `w:rPr` and `w:trPr` layers and a
closed `TableStyleRegion` whose declaration order is Word's priority order,
`CT_Style` gains its base `w:trPr` and `w:tcPr`, `CT_TblPr` gains
`w:tblStyleRowBandSize` and `w:tblStyleColBandSize`, and `CT_PPr` gains the
`w:cnfStyle` F-264 handed over raw-preserved. Resolution now flattens the
`basedOn` chain per region before region precedence applies, gives the
horizontal band the higher priority, counts bands by the resolved band size
after the header row or first column, and threads a table-style run layer into
`resolve_run_properties`. The facade gains the typed region parameter, the run
and row layers, per-region removal, a typed read side, `clear_look`, checked
band sizes, a `set_look` that writes the legacy bitmask beside the booleans,
and a paragraph conditional selector. The conditional `w:trPr` is modeled and
round-tripped, and its row geometry moves to F-268a with the DOCX-034 layout
and render columns, which the F-268 parent owns in the capability matrix.
**Tracked human action**: Word GUI capture is not available on the development
machine, so confirming that Word reopens an authored thirteen-region table
without offering to repair it is performed by hand at the Milestone 24
end-of-milestone gate above and recorded as performed or not performed. No
automated story test depends on it, and no gate skips in its absence.

### F-268, Floating and advanced table layout (L)
Author floating table positioning, overlap, bidirectional visual order, complete
width modes, autofit, captions, descriptions, and advanced row-grid behavior.
**Depends on**: F-257 through F-259.
**Test gate**: golden. Fixed, autofit, nested, and floating tables match the
reviewed Word page geometry and pagination.

F-268 is split into the two implementation stories below. The parent closes
only after both children close. The split was taken in the S74 consolidated
design round, to separate the authoring and geometry work from the paginator
float, which is the part that touches the shared obstacle machinery.

### F-268a, Advanced table authoring and geometry (L)
The `w:tblpPr`, `w:tblOverlap`, `w:bidiVisual`, `w:tblCellSpacing`,
`w:tblCaption`, and `w:tblDescription` grammar, the row `w:wBefore` and
`w:wAfter` offsets, the public authoring surface, autofit column widths, and
bidirectional visual column order. It also owns applying a conditional region's
`w:trPr` at layout, which F-267 modeled and round-tripped without applying, and
with it the remaining DOCX-034 layout and render columns.
**Depends on**: F-267.
**Test gate**: golden.
`fixed_autofit_and_nested_table_geometry_matches_reviewed_word_pages` pins the
page count, per-row origins, and column widths for a fixed-grid, an autofit,
and a nested table in deterministic font mode.
**Delivered**: `CT_TblPr` gains `w:tblpPr` as a typed `CT_TblPPr`, plus
`w:tblOverlap`, `w:bidiVisual`, `w:tblCellSpacing`, `w:tblCaption`, and
`w:tblDescription`, and `CT_TrPr` gains `w:wBefore`, `w:wAfter`, a row
`w:tblCellSpacing`, and `w:hidden`, all at their existing schema slots. The
three new enums are `ST_TblAnchor`, `ST_YAlign`, and `ST_TblOverlap`, and
`w:tblpXSpec` reuses `AnchorAlignH`. The facade gains ten checked setters and
their readers, and `has_unmodeled_properties` narrows by exactly those names.
Layout gains `TableBlock::bidi_visual`, `TableRow::offset_left`, and
`autofit_column_widths`, which engages only for an autofit or absent layout
mode with an auto or absent width. `w:gridBefore`, `w:gridAfter`, `w:wBefore`,
`w:wAfter`, and `w:tblCellSpacing` now change geometry, and a conditional
region's `w:trPr` resolves into row height, header repetition, and row grid
offsets, which closes the DOCX-034 layout and render columns. Floating
placement stays with F-268b, so `w:tblpPr` round-trips and a floating table
still renders in the flow.

### F-268b, Floating table placement and wrap (M)
Floating table lowering, `place_floating_table`, the wrap and `ResolvedWraps`
integration, and float against float resolution within one page.
**Depends on**: F-268a.
**Test gate**: golden.
`floating_tables_match_reviewed_word_page_geometry_and_pagination` pins the
float origins, page count, and the wrapped line boxes beside a margin-anchored,
a page-anchored, and a text-anchored float.
**Delivered**: `TableBlock` gains `floating`, a boxed `FloatingTable` lowered
from the `CT_TblPPr` F-268a modeled onto the drawing anchor frames, so
`resolve_anchor_h` and `resolve_anchor_v` take it unchanged. The paginator's
table arm branches on it, and `Pager::place_floating_table` resolves the rect,
renders every row at that origin, records the body fragments, pushes one square
`PlacedWrap` carrying the four from-text distances, and never advances
`cursor_y`. `has_paragraph_relative_wrap` and `lookahead_wraps` now see a
floating table, so the text above a float is pushed aside and a `text`-anchored
float settles across the existing two passes. `document_has_wrapping_drawing`
sees one too, which is what routes a float document onto the two-pass path
rather than the single-pass restart path. A float that does not fit moves whole
to the next page, never splits, and never repeats a header row, and
`w:tblOverlap` resolves float against float within one page. The reciprocal
half, a non-floating table narrowing beside a float, and `w:cantSplit` row
splitting stay named follow-ups. No sample floats, so all 49 hash entries and
the seven-entry golden pixel manifest are unchanged.

### F-269, Complete section page semantics (L)
Add page borders, line numbering, variable-width columns, separators, vertical
page alignment, mirrored margins, book-fold settings, paper source, and section
footnote and endnote configuration.
**Depends on**: F-250, F-251.
**Test gate**: differential. Every supported section property survives
round-trip and changes only its declared layout behavior.
**Delivered**: `CT_PageBorders`, `CT_LineNumber`, `CT_PaperSource` and the
shared `CT_NoteProperties` join `CT_SectPr` beside typed `w:vAlign` and
`w:textDirection`, so `w:footnotePr`, `w:endnotePr`, `w:paperSrc`,
`w:pgBorders`, `w:lnNumType`, `w:vAlign` and `w:textDirection` leave
`extra_xml` while the retained children keep their slots through a new
sub-slot. The central defect is closed: `Section::set_columns` wrote `w:cols`
that `sect_pr_to_geometry` never read, and column tracks now reach pagination
with a separator rule, while the single-column path bypasses the track
arithmetic so all 49 hash entries stay unchanged. Page borders, margin line
numbering excluded from the PDF reading order, vertical page alignment and
mirrored margins all render. `ST_VerticalJc` gains `Both` and
`#[non_exhaustive]`, which is a breaking change to `rdocx-oxml`. DOCX-036
has checked note policies and placement through F-274.

### F-269a, Word GUI confirmation for section page semantics (S)
Record the Word-authored oracle for columns, page borders, line numbering,
vertical alignment and mirrored margins. F-269 could not: Word GUI automation
is not available on the machine that produced it, so the capture path landed as
the `#[ignore]` test `capture_f269_word_section_evidence` instead.
**Depends on**: F-269.
**Test gate**: differential. The capture asserts the installed Word build
before it records anything, and the recorded set becomes a pinned oracle.

### F-269b, True vertical distribution for section `w:vAlign="both"` (S)
`both` preserves its source value, lays out as `top` and emits one diagnostic
per section. Distributing the body band's paragraphs across the unused vertical
measure is the remaining work.
**Depends on**: F-269.
**Test gate**: integration. A `both` section distributes its blocks and emits
no diagnostic, while `top`, `center` and `bottom` are unchanged.

### F-269c, Column balancing and per-track line breaking (M)
Two gaps left open deliberately by F-269. Word balances column heights at a
continuous section break, and this workspace fills tracks left to right without
balancing. Word also breaks each track to its own measure, and this workspace
breaks a whole section to the first track's measure, which only differs when
`w:equalWidth` is `0` with tracks of different widths.
**Depends on**: F-269.
**Test gate**: differential. A balanced continuous break and an unequal track
list both match the pinned LibreOffice render.

### F-270, Complete settings and web settings authoring (L)
Model and author the remaining modern document settings, compatibility options,
proof state, update policy, default tabs, theme language, mail-merge settings,
and web settings with typed removal and diagnostics.
**Depends on**: F-244.
**Test gate**: round-trip. A public-authored settings package reports no
unmodeled supported children and preserves unknown extensions byte for byte.
**Delivered**: `SUPPORTED_SETTINGS` closes the thirty-one top-level names over
one `SETTINGS_ORDER` table, `CompatibilityOption` covers the complete closed
`CT_Compat` on-off set, `MailMerge` authors thirteen members at their own schema
positions, `SettingsDiagnostic` separates duplicated from malformed occurrences,
and `crates/rdocx-oxml/src/web_settings.rs` owns `w:webSettings` with a
read-only `div_ids` projection. `w:defaultTabStop` now reaches
`oxml-layout::LineBreakParams::default_tab_interval_pt`. DOCX-007 and DOCX-037
are `complete`.

### F-271, Uniform rich header and footer editing (L)
Complete all valid header and footer content, fields, controls, annotations,
drawings, tables, links, and inherited variant operations through the common
story API.
**Depends on**: F-252 through F-255.
**Test gate**: integration. The same rich subtree, including authored comment
anchors, can be authored in every header and footer variant and reopens with
correct part-scoped relationships.

### F-272, Rich footnote authoring (L)
The native facade creates normal footnotes with body references, resolves their
story IDs, reorders their exact part elements, and removes them with all body
references in one staged mutation. Common story operations edit rich
paragraphs, tables, fields, links, drawings, comments, and content controls.
Internal note IDs stay stable while visible labels follow body reference order.
**Depends on**: F-253 through F-255.
**Test gate**: differential. Rich notes and their references match Word in
numbering, page placement, continuation, and round-trip structure.

### F-273, Rich endnote authoring (L)
Create, edit, reorder, and remove endnotes with the same content and relationship
surface as footnotes while retaining an independent identifier namespace.
The native facade stages each mutation and reopens the package before
publication. Common story editing preserves exact untouched XML, separator
records, and endnote-owned picture and hyperlink relationships. Body references
across sections number independently from footnotes, and the endnote stream
flows from the section or document boundary under the selected placement policy.
**Depends on**: F-272.
**Test gate**: differential. Word confirms independent occurrence order across
sections and the selected policy places formatted labels and endnote bodies.

### F-274, Note separators, markers, and restart policy (L)
Checked document and section policies select number format, start, restart,
and placement for both note families. Special-record mutation authors the
separator, continuation separator, and continuation notice in each
relationship-resolved note part while preserving unrelated XML. Body and note
markers share one formatted label stream, and custom marks leave numeric
numbering untouched. Footnotes reserve their page area. Endnotes flow from
their section or document boundary and carry to later pages. The pinned Word
build ignores `beneathText` even with a section override, while native layout
places the footnote beneath body text.
**Depends on**: F-269, F-272, F-273.
**Test gate**: differential. The pinned Word fixtures confirm document and
section end placement, section and page restart, custom marks, separators,
and carried-note notices. A native deterministic test records the pinned
`beneathText` divergence. Round-trip and atomic tests cover source XML and
invalid mutations without disturbing unrelated section numbering.

### F-275, Cross-story bookmarks, ranges, and annotations (L)
Create and mutate bookmark pairs, comment endpoints, permission ranges and
spelling or grammar proofing ranges in every valid story and nested block
control. The native inventory uses immutable story-qualified endpoints.
Comment pair removal retains the reference and definition as a point comment,
while moving the pair moves its reference run. Modern metadata, permission
policy and revision move ranges have separate owners.
**Depends on**: F-253, F-254.
**Test gate**: round-trip. Nested and crossing-invalid ranges are respectively
preserved or rejected atomically, and valid ranges retain exact endpoints.

### F-276, Complete fragment conflict and dependency policy (L)
Cross-document block fragment import supports every story owner and nested
block-control paragraph boundary. It closes style aliases and links, numbering
overrides, custom XML bindings and item properties, notes, comment threads,
paired markers, revisions, charts, diagrams, embeddings and reachable opaque
package extensions. External edges are retained without fetching. Compatible
imports publish one reopened candidate, with deterministic identity maps and
atomic integrity failures. Preserved opaque companions gain no new decoding
or rendering support.
**Depends on**: F-256, F-270 through F-275.
**Test gate**: regression. Importing a full-story fragment into a conflicting
destination remaps every dependency deterministically and leaves no dangling ID.

### F-277, Glossary and building-block creation (L)
Create, classify, update, insert, and remove glossary documents, AutoText,
building blocks, placeholders, and their related content through public APIs.
**Depends on**: F-237, F-253, F-276.
**Test gate**: round-trip. Public-created entries retain category, behavior,
content, relationships, and unsupported siblings after insertion and reopen.
Typed creation requires dependency-free content. Fragment-backed creation and
insertion share the complete dependency transaction. Last removal retains
one valid empty glossary part. Binding a placeholder updates an existing
control without changing its discriminator. Stale snapshots and duplicate
authored names fail atomically.

### F-278, General simple and complex field builder (L)
Provide typed and raw-instruction-safe builders for simple and complex fields,
nested instructions, ordered begin-separate-end runs, locks, dirty state,
format switches, and cached display content.
**Depends on**: F-260.
**Test gate**: round-trip. Every supported field shape reopens with identical
instruction semantics and ordered cached content.

### F-279, Pagination field materialization across stories (L)
Materialize PAGE, NUMPAGES, SECTION, SECTIONPAGES, PAGEREF, and related layout
fields in every supported story using one deterministic layout snapshot.
**Depends on**: F-263, F-271 through F-274, F-278.
**Test gate**: differential. Body, header, footer, note, and text-box field caches
match the pinned Word page and section values.

### F-280, Captions, sequences, and complete cross-references (M)
Author captions and labels, sequence fields, bookmark targets, and complete REF
number, position, hyperlink, and context switches.
**Depends on**: F-248, F-275, F-278.
**Test gate**: differential. Figure, table, equation, and numbered-heading
references match Word before and after insertion and renumbering.

### F-281, Indexes and tables of figures and authorities (L)
Author source markers and rebuild INDEX, TOF, and TOA result structures with
cached entries, page ranges, leaders, links, and preserved formatting.
**Depends on**: F-278 through F-280.
**Test gate**: differential. Mutating sources produces the same ordered entries
and page targets as the pinned Word update.

### F-282, Citations and bibliography authoring (L)
Create the native bibliography source and field-authoring foundation and the
measured citation and bibliography formatter subset, preserving producer
metadata, locale data and unavailable result caches with diagnostics. The user
moved the remaining full Word catalogue work to F-X192 in S91 before S90
publication. Document exact admitted inputs and verified output coverage,
including the 210-of-223 dense APA development checkpoint, without claiming
full catalogue parity. New formatting expansion belongs to F-X192.
**Depends on**: F-278.
**Test gate**: differential. Public source-built admitted citation and
bibliography controls match pinned Word identifiers, ordering, rich display
and package round trips. Outside the measured materialization boundary,
inspection and authoring retain source data. Unfinished recognized standard
branches fail the whole update atomically with an explicit error. Noncatalogue
branches retain caches with report diagnostics. Atomicity and XML preservation
gates remain mandatory.

### F-283, Complete numbering-aware navigation fields (L)
Schedule this unfinished story in S95.2. Close all numbering interactions across
TOC, STYLEREF, REF, PAGEREF, captions,
and document outline results, including table-cell headings and suppressed
paragraphs.
The native numbered outline retains accepted physical body, cell and nested
control owners using production cascade and numbering facts. Numbered STYLEREF
and relative REF share those facts and actual placement. Generated caches do
not supply source headings or copied sequence events. Explicit cache publication
retains its transaction and final-pagination boundaries. Native lifecycle
qualification remains part of the integrated gate, with unavailable historical
captures requiring fresh observation.
**Depends on**: F-248, F-279 through F-282, F-X200.
**Test gate**: regression. One multilevel numbered document stays consistent
across visible markers, navigation structures, references, and saved caches.

### F-284, Stable container-wide template grammar (L)
Define the public template grammar, formatting hooks, filters, nested paths,
whitespace behavior, errors, and structural expansion across every supported
story and container.
**Depends on**: F-253, F-261, F-276.
**Test gate**: regression. The versioned grammar renders the same payload across
body, table, header, footer, note, and text-box locations without run-split loss.

### F-285, Content control creation and lifecycle (L)
Create, inspect, edit, move, clone, and remove block, row, cell, paragraph, and
inline content controls with typed identity, alias, tag, lock, appearance, and
placeholder behavior. Public creation admits rich text and plain text. Other
kinds name F-286 and publish nothing. Python exposes insert, metadata, remove,
and unwrap on the frozen snapshot. Move and clone stay on the Rust facade.
WASM and CLI gain no method.
**Depends on**: F-253, F-254.
**Test gate**: round-trip. Each control location and property survives public
creation and mutation with no raw-XML fallback.

### F-286, Rich, repeating, and typed content controls (L)
Author rich text, plain text, date, checkbox, picture, combo, drop-down, group,
repeating-section, and repeating-item controls with type-specific validation.
`set_content_control_type` writes the payload onto an F-285 shell.
`repeat_content_control` clones the first repeating item. The last repeating
item stays through `remove_content_control`. Citation, equation, bibliography,
and document-part kinds stay refused.
**Depends on**: F-285.
**Test gate**: differential. Every supported control passes the pinned native
Word comparison under replacement and documented equivalent repetition.
Q24 retains clone-first append and last-item refusal as explicit intentional
differences from Word's addressed-item default and whole-section deletion.
Authentic Word provenance, version and input identity are mandatory. Locally
constructed expectations do not certify Word.

### F-287, Custom XML stores and data binding authoring (L)
`add_custom_xml_store` creates one custom XML item and its properties part
through the bibliography relationship sequence and returns the braced item id.
Schema URIs are stored in caller order and are not fetched.
`set_content_control_binding` writes the three `CT_DataBinding` strings onto an
existing shell. Value updates stay on the existing tag and bind setters.
Unknown item bytes stay byte-identical. Bibliography keeps its own item bytes.
**Depends on**: F-285.
**Test gate**: round-trip. Bound controls resolve, update, and reopen against the
correct store while unknown store content remains byte-identical.

### F-288, Legacy form field creation (M)
Create text, checkbox, and drop-down legacy fields with complete supported help,
status, default, formatting, size, list, and calculation properties.
**Depends on**: F-237, F-278.
**Test gate**: differential. Public-created legacy fields match Word and retain
their typed values through update and reopen.
Explicit reset defaults use `Field::legacy_form_default_value` and the batch
`Document::legacy_form_fields_with_defaults` query. The public
`LegacyFormFieldData` and `LegacyFormFieldInfo` structs keep their existing
fields. Python exposes the readonly `default_value` property. Implemented
queries and native evidence remain separate from final sprint acceptance.

### F-289, Modern Word form authoring (L)
Compose content controls, protection, editable ranges, validation, placeholders,
and repeating data into a high-level modern form surface without creating a
second document model.
**Depends on**: F-275, F-285 through F-288.
**Test gate**: differential. A dense source-built form behaves and renders like
the pinned Word form across fill, reset, protect, and reopen operations.

### F-290, Mail-merge package and data-source authoring (M)
Author mail-merge settings, field mappings, recipients, connection metadata,
and an offline embedded table. External retrieval stays policy-controlled and
offline by default. New mappings use the 30 Word canonical role ids in
MS-OI29500 2.1.369 order. Word reads field-map data by position and ignores
mapped names. A mapped slot uses type dbColumn. An unused slot uses type null.
Query returns logical mappings in that role order and omits null padding.
The story does not execute connectors or arbitrary MERGEFIELD aliases.
`mail_merge_package` is absent only when no recognized owned configuration
remains. Opaque residual XML is not a configuration. A package has at most
one recipient-data part. The source relationship type is `mailMergeSource`
for both data source and ODSO source, and header ownership follows the
qualified XML reference. The implemented package API stores and reads this
metadata without executing a connector. Final integrated acceptance remains
required.
**Depends on**: F-233, F-270, F-278.
**Test gate**: round-trip. A public-created merge package retains its declared
source and mappings, and unavailable external data returns diagnostics.

### F-291, Tracked insertion and deletion authoring (L)
Create tracked text, paragraph, table, row, cell, and nested-story insertions
and deletions with deterministic revision IDs, authors, dates, and accepted and
original views.
**Depends on**: F-253, F-254.
**Test gate**: differential. Authored revisions match Word structure and both
views across save, reopen, accept, and reject.

### F-292, Property revisions and move ranges (L)
Author paragraph, run, table, row, cell, numbering, and section property changes
plus paired move-from and move-to ranges in every valid story.
**Depends on**: F-275, F-291.
**Test gate**: differential. Property and move revisions match the pinned Word
result and resolve atomically in both directions.

### F-293, Complete comments and modern comment metadata (L)
Create comments, replies, people identities, resolved state, anchors, and modern
comment extension metadata across body, tables, notes, headers, footers, and
text boxes.
**Depends on**: F-275, F-291.
**Test gate**: round-trip. Every comment thread retains its range, identity,
reply graph, and resolved state after unrelated story edits.

### F-294, Permission ranges and protection integration (M)
Author editable ranges for users and groups, validate paired markers, and
compose them with document protection and supported form-filling policies.
**Depends on**: F-275, F-289.
**Test gate**: differential. Protected documents expose only the declared
editable ranges and reopen without repair.

### F-295, Comparison output as complete revisions (L)
Make comparison emit the complete tracked-revision surface, including property
changes, moves, tables, controls, fields, notes, drawings, and related stories.
**Depends on**: F-234, F-235, F-291 through F-293.
**Test gate**: differential. Pinned document pairs produce the same revision
semantics and accepted and original views as Word.

### F-296, Collaboration identity and deterministic time policy (M)
Provide one caller-controlled identity and clock policy for comments, revisions,
properties, signatures, fields, and generated package metadata. Deterministic
mode never reads ambient user or wall-clock state.
**Depends on**: F-244, F-249, F-291, F-293.
**Test gate**: regression. Identical explicit identity and time input produces
byte-identical collaboration packages across repeated runs.

### F-297, Complete Word drawing anchor and effect authoring (L)
Author inline and anchored pictures with page, margin, column, paragraph, and
character-relative positioning, wrap modes, distances, overlap, crop, rotation,
transform, fill, line, and supported effects.
**Depends on**: F-255, F-262.
**Test gate**: golden. A source-built drawing matrix matches the pinned Word
geometry and deterministic raster output.

### F-298, Shapes, text boxes, groups, and connectors in Word (L)
Expose shared DrawingML shapes, connectors, groups, and rich text boxes through
WordprocessingDrawing without requiring PresentationML or raw XML APIs.
**Depends on**: F-253, F-297.
**Test gate**: differential. Public-created Word shapes reopen with the same
geometry, text story, relationships, and rendering as the pinned reference.

### F-299, AlternateContent, VML, and SVG compatibility authoring (L)
Generate modeled modern choices and bounded transitional fallbacks for shapes,
text boxes, watermarks, SVG images, and legacy consumers. Callers select a
compatibility policy rather than supplying raw wrappers.
**Depends on**: F-298.
**Test gate**: round-trip. Each choice and fallback pair is schema ordered,
relationship complete, and accepted by the pinned modern and compatibility
oracles.

### F-300, Charts at arbitrary Word insertion points (M)
Insert and edit charts in body, cells, headers, footers, notes, and text boxes
where WordprocessingML permits them, with part-scoped embedded workbooks and
deterministic relationships.
**Depends on**: F-158, F-255, F-253.
**Test gate**: differential. Equivalent charts at every valid insertion point
match Word package ownership and rendering.

### F-301, SmartArt and diagram authoring in Word (L)
Create and edit Word SmartArt and diagram parts through the shared typed model,
including data, layout, colors, styles, relationships, and deterministic
fallback rendering.
**Depends on**: F-219, F-220, F-255.
**Test gate**: differential. A source-built Word diagram reopens, renders, and
retains unsupported extension data like the pinned Word reference.

### F-302, Embedded objects, icons, and alternative-format parts (M)
Attach, replace, extract, and remove embedded packages, object icons, linked
objects, and bounded alternative-format import parts without executing their
payloads.
**Depends on**: F-236, F-255.
**Test gate**: round-trip. Each safe object retains exact bytes and ownership,
and removal leaves no orphan relationship or content-type override.

### F-303, Drawing and embedded-content layout completion (L)
Lay out and render the supported drawings, text boxes, charts, diagrams, object
icons, and compatibility choices in every valid Word story with source
provenance and accessibility metadata.
**Depends on**: F-297 through F-302.
**Test gate**: golden. The combined drawing page matches the pinned Word oracle
at the declared geometry and SSIM thresholds.

### F-304, Typed package extensibility facade (L)
Create, inspect, replace, and remove arbitrary safe custom parts through typed
content types, owner-scoped relationships, explicit external-target policy, and
validated package publication.
**Depends on**: F-243, F-249, F-255.
**Test gate**: round-trip. A source-built extension graph remains reachable,
deterministic, and lossless while unsafe paths and relationship cycles fail.

### F-305, Attached templates, web extensions, and task panes (L)
Author and inspect attached-template relationships, web settings, web-extension
stores, task panes, content-add-in declarations, permissions, and external
resource locations without executing web content.
**Depends on**: F-270, F-304.
**Test gate**: round-trip. Every public-created declaration reopens with exact
ownership and reports unavailable execution explicitly.

### F-306, Executable compatibility attachment and signature rules (M)
Attach and remove caller-supplied VBA, ActiveX, OLE, and custom UI payloads as
opaque bytes, and define package and VBA signature retention, invalidation, and
re-signing requirements for every mutation.
**Depends on**: F-172, F-236, F-238, F-304.
**Test gate**: regression. Mutation policy never presents an invalid signature
as valid, never executes a payload, and preserves unrelated bytes exactly.

### F-307, Complete accessibility authoring and audit (L)
Author alternative text, decorative flags, table headers and scopes, reading
order, languages, accessible links, equations, controls, drawings, and document
metadata, then align DOCX audit results with tagged PDF semantics.
**Depends on**: F-266 through F-303.
**Test gate**: differential. A source-built accessible document passes the
declared Word and tagged-PDF structural checks with no false clean result.

### F-308, Fully modeled and losslessness diagnostics (L)
Report every modeled, preservation-only, unsupported, unsafe, orphaned, or
lossy feature by stable document and package path. A public-authored document
can assert that it contains no unexplained raw XML or unsupported semantics.
**Depends on**: F-264 through F-307.
**Test gate**: regression. Deliberate unsupported content produces one stable
diagnostic, while the complete source-built document reports fully modeled.

### F-309, Strict, transitional, and repair-free conformance (L)
Validate package relationships, content types, expanded names, schema child
order, required attributes, Strict and Transitional vocabulary, and supported
markup compatibility before publication.
**Depends on**: F-X077, F-308.
**Test gate**: differential. The complete generated corpus passes the strict
validator and pinned Word no-repair check, and one mutation per rule fails.

### F-311, Positioned paragraph frame placement (M)
Lay out and render the `w:framePr` positioned paragraph frame, including its
anchors, wrap mode, drop cap and the text that flows around it. F-264 models
and authors the whole property, so this story is layout and rendering only.
**Depends on**: F-264.
**Test gate**: golden. A page holding a margin-anchored frame, a page-anchored
frame and a drop cap matches the pinned deterministic geometry, and the
surrounding text wraps where Word wraps it.

### F-312, Run visual effect render projection (L)
Render the run effects F-265 models but does not paint: `w:outline` stroke-only
glyphs, `w:shadow` and `w:emboss` and `w:imprint` relief, the `w:bdr` character
border box, `w:kern` gating by size, and `w:fitText` horizontal segment
scaling. Needs new segment state in `oxml-layout` and the matching PDF backend
work, which is why F-265 stopped at the model.
**Depends on**: F-265.
**Test gate**: golden. One deterministic page carrying every effect matches its
pinned geometry, and a run with no effect is byte identical to the same run
before this story.

### F-310, Determinism, resource limits, bindings, and stability gate (L)
Close M24 with byte determinism, explicit memory and time bounds, cancellation,
complex-document performance budgets, native, Python, WASM, and CLI capability
classification, stable template grammar, final README refresh, and a separate
1.0 readiness decision.
**Depends on**: F-264 through F-309.
**Test gate**: regression. The complete milestone matrix has no unexplained
partial row, all required public surfaces pass their gates, and repeated output
is byte-identical under the declared deterministic inputs.

---

## Cross-cutting

### F-X001, rdocx-cli tests (M)
The published binary has one compiled-executable integration test for each of
its seven subcommands in a single test binary. Fixtures are constructed in
code. Text extraction preserves document order, and both render branches use
bundled-font deterministic output.
**Test gate**: all seven named command integration tests pass, and the text,
validation, and deterministic-render sensitivity mutations fail.

### F-X002, README example correctness (S)
The three root README Rust examples use `rust,no_run` and compile against the
current `rdocx` rlib without executing filesystem writes. They cover blank
authoring, read and mutation, and render and export through public facade APIs.
**Test gate**: `python3 scripts/readme_doctests.py` compiles all three examples.

### F-X003, Deduplicate the sample generators (S)
`generate_all_samples.rs` and `generate_samples.rs` overlap substantially.
**Test gate**: one generator produces every sample the harness needs.

### F-X004, Fix the shared temp path in the test suite (S)
`integration_test.rs` writes to a fixed, non-unique temp path shared across
concurrent runs.
**Test gate**: two concurrent `cargo test` runs both pass.

### F-X005, Tag rpptx-v0.1.2 (S)
Retain complete registry metadata after the immutable partial 0.1.0
publication, remove the CI-only tool dependency exposed by the 0.1.1 workflow,
prepare the exact incubating family at 0.1.2, and publish it through a newly
reviewed release tag before released rdocx consumers cut over.
**Depends on**: F-047 through F-050.
**Test gate**: all 12 incubating packages resolve from crates.io at 0.1.2 with
the expected owner, and the GitHub release targets the newly reviewed sprint
SHA.

### F-X006, Tag the expanded rpptx family (S)
Prepare the complete 14-package incubating family at 0.1.3, including
`oxml-cli-support` and `rpptx-cli`, then publish it only through
`/release rpptx-v0.1.3` after the command's separate final approval. The
complete family is published at 0.1.3. The immutable `rpptx-v0.1.2` tag and
its 12 published packages remain unchanged.
**Depends on**: F-143, F-144, F-145.
**Test gate**: all 14 incubating packages resolve from crates.io at 0.1.3 with
the expected owner, and the GitHub release targets the reviewed sprint SHA.

### F-X007, Integrate PR 25 and stable crate documentation (L)
Integrate Jon Stokes's PR 25 through the sprint branch, retaining contributor
credit in the GitHub merge record. The public Word facade gains custom list
definitions, per-paragraph numbering, composable hard line breaks and
hyperlinks, and fixed table-column widths. Rejected list updates remain
side-effect free, and fixed table geometry keeps the table width, grid, and
spanning cell widths consistent. Every stable crate has a package README that
states when to use it, links to its API documentation, and includes a current
example or a clear deprecation path. The README examples are compile-checked.
Typed numbering edits preserve unsupported attributes and child XML in schema
order across namespace aliases and collisions. Repeated tab stops carry public
source-occurrence provenance so edits, insertions, removals, and explicit
clears retain producer ownership in deterministic linear work. The public tab
parser tracks namespace scopes and accepts both empty and expanded tab-stop
elements. Preservation carriers extend one expanded-name `mc:Ignorable`
attribute without duplicating it, using the actual property ancestor scope
rather than a document-wide declaration list. Style, body, table-cell, header,
footer, footnote, and endnote paragraph properties retain established aliased
and default WordprocessingML parsing. Nested tab namespace scope has a normal
64-element depth bound. These public model additions set the stable release
boundary at 0.5.0.
**Test gate**: the merged PR's focused round-trip suite passes against current
`main`, the two rejected-state and table-geometry regressions pass, and every
stable crate README example compiles against its packaged crate. Numbering
round trips cover schema order, foreign namespace collisions, nested property
markup, provenance-only replacement, repeated occurrence ownership, explicit
clear carriers, namespace shadows, and expanded tab elements. The hash harness
remains 28 of 28. The gate also covers direct style and paragraph boundaries,
table cells, headers, notes, foreign same-local negatives, property-local
compatibility scope, and bounded deep tab aliases. Stable package archives stay
below 10 MiB, and the public migration examples compile.

### F-X008, Tag v0.5.0 (S)
The stable workspace package, nine internal pins, and eleven inherited
lockfile packages are 0.5.0 after F-X007. The exact seven stable crates.io
packages are published at 0.5.0 from the reviewed `v0.5.0` tag. The two Python
project versions and `rdocx-wasm` inherit 0.5.0 without gaining publication
authority. All 15 incubating manifests remain at 0.1.3, with exactly 14 in the
incubating crates.io family and `rpptx-wasm` unpublished. `publish.yml` runs the
exact stable and incubating metadata preflights before its patched 21-package
workspace dry run. No incubating, WASM, Python, or npm package is part of the
stable publication.
**Depends on**: F-X007.
**Test gate**: the stable metadata regression proves the workspace version,
nine pins, eleven lock entries, two Python versions, WASM literals, README
requirements, exact stable publication set, and unchanged incubating 0.1.3
state. The workflow contract, 12 README examples, 28-entry hash harness, exact
patched 21-package dry run, seven stable archive inventories, and `cargo deny`
pass. All seven stable packages resolve independently from crates.io at 0.5.0
under owner `mantissaman`, the GitHub release targets the reviewed sprint SHA,
and the PR 25 contributor credit and merge note remain visible on GitHub.

### F-X009, README coverage for every workspace crate (L)
Every one of the 26 Cargo workspace packages declares a README. Each document
states what the crate owns, when it should be used directly, its relationship
to adjacent packages, and provides a concrete Rust, CLI, Python, or JavaScript
example appropriate to that package. Internal and unpublished packages are
labelled honestly and gain no publication authority. The README runner checks
the exact workspace package set, required sections, manifest wiring, examples,
and archive inventory.
**Test gate**: `python3 scripts/readme_doctests.py` verifies exact README
coverage for all 26 workspace packages, compiles 26 Rust examples, validates
the CLI, Python, and JavaScript snippets, and proves all 21
publishable archives contain the byte-identical declared README.

### F-X010, Tag v0.6.0 (S)
Prepare the complete stable train at the next minor version, 0.6.0. The eleven
workspace-version packages move together, including the exact seven crates.io
packages and the four unpublished Python and WASM support packages. Stable
README dependency examples, metadata regressions, lock entries, Python project
versions, and WASM contract literals move to 0.6.0. The incubating train remains
at 0.1.3. The reviewed `/release v0.6.0` workflow publishes only the exact
seven stable crates after full verification, a clean microscope, a clean
sprint review, and separate immediate approval. No PyPI, npm, WASM, Python, or
incubating publication is authorized.
**Depends on**: F-X009.
**Test gate**: the stable release regression proves the eleven-package train,
nine internal pins, exact seven-package publication set, README requirements,
lock entries, Python project versions, WASM literals, and unchanged incubating
train. The exact 21-package dry run, README compilation and archive inventory,
28-entry hash harness, and supply-chain gate pass. All seven crates resolve at
0.6.0 under owner `mantissaman`, each crates.io README is present, and the
annotated `v0.6.0` tag targets the reviewed sprint SHA.

### F-X011, Tag rpptx-v0.2.0 (S)
The complete incubating train is published at the next minor version, 0.2.0. The
fourteen publishable `oxml-*` and `rpptx-*` packages move together with
unpublished `rpptx-wasm`, their root dependency pins, lock entries, README
dependency examples, source assertions, workflow regressions, and local WASM
package version. The completed stable train remains at 0.6.0. Incubating 0.2.0
was published only after full verification, a clean sprint review, and
separate immediate approval. `/release rpptx-v0.2.0` published only the exact
fourteen incubating crates. No npm, PyPI, Python, WASM, or stable package was
published.
**Depends on**: F-X010.
**Test gate**: the incubating release regression proves the fifteen-package
preparation group, fourteen internal pins, exact fourteen-package publication
set, README requirements, lock entries, source and workflow assertions, and
unchanged stable train. The exact 21-package dry run, README compilation and
archive inventory, 28-entry hash harness, WASM package gate, and supply-chain
gate pass. All fourteen crates resolve at 0.2.0 under owner `mantissaman`, each
crates.io README is present, and the annotated tag targets the reviewed sprint
SHA used by the successful GitHub release workflow.

### F-X012, Restore pinned CI toolchains (M)
Hosted CI installs the reviewed Poppler 26.01.0 rendering oracle from its exact
source archive and SHA-256 rather than a moving package-manager version. The
shared installer bounds download and streaming extraction resources, rejects
unsafe archive members and populated prefixes, builds only the three required
tools, and verifies each runtime identity. Test, MSRV, both Python binding rows,
and Presentation fidelity invoke it unconditionally before use. The WASM job
verifies the official Binaryen 125 Linux archive and exact
`wasm-opt version 125 (version_125)` release identity. Product code, package
versions, published artifacts, and rendering baselines remain unchanged.
Test and MSRV also install exact uv 0.10.2 through the reviewed official setup
action, isolate its cache, and run their corpus tests with an explicit 8 MiB
Rust test-thread stack.
They run on Ubuntu 24.04 and install LibreOffice 26.2.5.2 from the reviewed official Linux x86-64
archive before the full workspace suite. The shared installer verifies SHA-256
`2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`,
uses bounded streaming extraction, refuses populated prefixes, and checks the
exact reviewed build identity before the three `rpptx-chart` viewer gates run.
The installer also declares the exact Ubuntu runtime-library package set needed
to execute that official build.
**Test gate**: behavioral regressions execute every source, resource, runtime,
and prefix guard. Workflow mutations reject missing, conditional,
failure-tolerant, or successfully short-circuited installer steps and reject a
weakened Binaryen checksum or identity gate. They also reject uv action,
version, cache, or stack drift. The same contract rejects LibreOffice version,
checksum, bound, runtime, ordering, or consumer-step drift. Full verification
and a hosted pull-request CI run at the reviewed SHA pass with all 28 hashes
unchanged.

### F-X013, Footnote and endnote placement (M, split at design)
Carries the footnote half of the external PR 2 contribution, whose
anchored-drawing half was superseded by F-X007 and the M7 anchor work. Split
into three children at design time, when fixing endnote placement and splitting
oversized notes were both taken into scope. The parent closes when every child
closes.

### F-X013a, Footnote line advance (S)
Footnote text advances horizontally across the segments of a line rather than
drawing every segment at the same indent. A footnote assembled from several
runs, which is what any footnote carrying mixed formatting produces, no longer
collapses into an unreadable stack at a single x. The advance accumulates the
segment width that line breaking already computed, so the fix introduces no new
measurement.
**Test gate**: regression, named as a sentence describing the failure it
prevents. A footnote built from several differently formatted runs renders its
segments at strictly increasing x, and a single-segment footnote is unmoved. The
hash harness carries an expected delta for every baseline holding a
multi-segment footnote, stated and justified in the commit.

### F-X013b, Footnote reservation and splitting (L)
Pagination reserves the height a page's notes occupy before body content fills
the text area, so body text and the note area no longer overlap. A page reserves
the separator offset once and each distinct note referenced by a line placed on
that page once, which keeps a note with the page carrying its reference rather
than with the paragraph that owns it. A note too tall for the space remaining
splits at a line boundary and continues on the next page, so an oversized note
can neither starve a page of body content nor stall pagination. Notes are laid
out once into a shared height map that the reservation and the rendering pass
both consume, so a reserved height and its rendered height cannot diverge.
**Depends on**: F-X013a.
**Test gate**: regression, named as sentences describing the failures they
prevent. A page whose body fills the text area leaves the reserved note area
clear. A note taller than its remaining space continues on the following page
without repeating its marker. A page carrying two references to one note
reserves that note once. The hash harness carries an expected delta for every
baseline holding a note, stated and justified in the commit.

### F-X013c, Endnotes at the document end (M)
Endnote references stop rendering their note at the foot of the page that
carries the reference. Endnotes collect into a document-end sequence rendered
after the final body page in reference order, while footnotes keep their
per-page placement. The layout carries the two note streams distinctly rather
than a single identifier that a footnote and an endnote of the same number both
match, which today resolves to whichever the footnote part happens to define.
**Depends on**: F-X013b.
**Test gate**: regression, named as sentences describing the failures they
prevent. A document mixing footnotes and endnotes places each stream in its own
region. A footnote and an endnote sharing a number resolve to their own note
rather than both to the footnote. The hash harness carries an expected delta for
every baseline holding an endnote, stated and justified in the commit.

### F-X017, Notes broken to their own section's width (S)
A note is line-broken to the width of the section that references it rather
than to the final section's width. `NoteRegistry` is built once ahead of
pagination against one content width, which is correct for every document whose
sections share a page size and wrong for any that does not. Note positioning is
already per-section, so this closes the half F-X013b left open.
**Depends on**: F-X013b.
**Test gate**: regression. A document whose two sections differ in page width
breaks each note to the measure of the section holding its reference, and a
single-section document is byte-identical to before.

### F-X014, Kashida justification values (S)
`ST_Jc` accepts `lowKashida`, `mediumKashida` and `highKashida`, mapping each to
justified alignment instead of rejecting the value.

The consequence is larger than the alignment. `CT_PPr::from_xml` propagates the
rejection with `?`, and that error travels all the way out of
`CT_Document::from_xml`, so a document carrying one of the three Arabic
justification settings **fails to open at all**. This is a load failure, not a
layout inaccuracy.
**Test gate**: regression, named as a sentence describing the failure it
prevents. A document whose paragraph carries each kashida value opens and lays
out justified, and the existing rejection still holds for a genuinely unknown
string. The hash harness is unchanged, since no recorded baseline carries a
kashida value.

### F-X024, Move the theme adapter into rdocx-oxml (M)
`oxml-drawing` hosts `impl From<&CT_OfficeStyleSheet> for
rdocx_oxml::theme::Theme`, which is the single documented exception to the rule
that nothing in `oxml-*` depends on `rdocx-*`. That one edge makes the two
publication trains mutually dependent: `rdocx-layout` depends on `oxml-layout`
and `oxml-drawing` depends on `rdocx-oxml`, so neither train can publish first
once both carry breaking changes.

The adapter moves to `rdocx-oxml`, which the orphan rule permits because `Theme`
is local there and `CT_OfficeStyleSheet` is the foreign type. The edge inverts
to stable depending on incubating, the architecture rule loses its exception and
becomes absolute, and train-at-a-time publication works in one fixed order
forever: incubating, then stable.

`rdocx-oxml` gains a dependency on `oxml-drawing`, so a Word-only consumer now
compiles DrawingML. That is the accepted cost, chosen over deleting an adapter
that exists so `rdocx-layout`'s `LayoutInput.theme` does not churn when
PresentationML themes reach Word layout.
**Depends on**: F-X020.
**Test gate**: regression. The conversion produces the same `Theme` from the
same `CT_OfficeStyleSheet` as before the move, `cargo tree` shows no `oxml-*`
package depending on any `rdocx-*` or `rpptx-*` package, and the workspace
still builds with all 28 hashes unchanged.

### F-X022, Tag rpptx-v0.3.0 (S)
The complete incubating train moves to the next minor version, 0.3.0, because
S41 broke its public API rather than merely extending it. `oxml-layout` renamed
`TextSegment::footnote_id` and `GlyphRun::footnote_id` to `note`, changing the
type from `Option<i32>` to `Option<NoteRef>`, and added two fields to
`LineBreakParams`. Under semver a 0.x minor bump is the correct response.

The fifteen packages carrying an explicit 0.2.0 move together, their root
dependency pins, lock entries, README dependency examples and the local
`rpptx-wasm` version with them. Exactly fourteen are published: `rpptx-wasm`
stays unpublished. The stable train stays at 0.6.0 during this story, and its
pins on the incubating crates move to 0.3.0 so the later stable release can
resolve against a published 0.3.0.

This story prepares and, through `/release rpptx-v0.3.0`, publishes. Publication
happens only after full verification, a clean microscope, a clean sprint review
and separate immediate approval at the reviewed SHA. No npm, PyPI, Python, WASM
or stable package is authorized.
**Depends on**: F-X020.
**Test gate**: the incubating release regression proves the fifteen-package
preparation group, the fourteen internal pins, the exact fourteen-package
publication set, README requirements, lock entries and the unpublished
`rpptx-wasm` literal. The patched workspace dry run, archive inventory under
10 MiB, README compilation and `cargo deny` pass, and all 28 hashes stay
unchanged.

### F-X023, Tag v0.7.0 (S)
The complete stable train moves to 0.7.0, because S41 broke its public API.
`rdocx-oxml` added `note_type` to `CT_Footnote`, six fields to `CT_Anchor` and
four variants to `WrapType`, each of which breaks an exhaustive match or a
struct literal. `rdocx-layout` added fields to `ParagraphBlock` and
`AnchoredDrawing`. The `rdocx` facade's own public API is unchanged, and it
moves with its train regardless.

The eleven workspace-version packages move together: the exact seven crates.io
packages plus the four unpublished Python and WASM support packages. README
dependency examples, metadata regressions, lock entries, the two Python project
versions and the WASM contract literals move to 0.7.0. The incubating train
remains at 0.3.0.

`/release v0.7.0` publishes only the exact seven stable crates, after full
verification, a clean microscope, a clean sprint review and separate immediate
approval. No PyPI, npm, WASM, Python or incubating publication is authorized.
**Depends on**: F-X022. The stable crates depend on `oxml-layout`, so the
incubating train has to be resolvable at 0.3.0 on crates.io before the stable
train that pins it can publish. This is the reverse of the S39 order, where only
one train moved.
**Test gate**: the stable release regression proves the eleven-package train,
the nine internal pins, the exact seven-package publication set, README
requirements, lock entries, Python project versions, WASM literals and the
unchanged incubating train at 0.3.0. The patched workspace dry run, archive
inventory, README compilation and `cargo deny` pass, and all 28 hashes stay
unchanged.

### F-X025, /verify must run the release regressions (S)
`/verify --full` runs formatting, lints, the workspace suite, the hash harness,
the prose rules, the no-default-features path, the WASM targets, docs, packaging
and the supply-chain check. It does not run
`python3 -m unittest scripts.test_sprint_workflow`, which holds the release
family preflights that `.github/workflows/publish.yml` invokes by name as the
publication gate.

S42 demonstrated the gap rather than theorised it. F-X022 moved every version
carrier under `crates/`, passed the entire local gate, and still left the
incubating preflight and the `ci.yml` WASM literal asserting the old version. It
would have failed in CI at publication time.
**Test gate**: regression. A deliberately stale version literal in
`scripts/test_sprint_workflow.py` or a workflow file fails `/verify --full`,
and a clean tree passes it.

### F-X026, CI must run the release regressions too (S)
`/verify` step 6 runs `python3 -m unittest scripts.test_sprint_workflow` after
F-X025, so the release family preflights no longer run for the first time on a
tag. `.github/workflows/ci.yml` does not. Its `prose` job runs the sprint's other
two standard-library checks, `prose_check.py` and `sync_agent_skills.py --check`,
and not this one, so a contributor who does not run `/verify` can move a version
carrier and see a green pull request.

Filed by the S43 sprint review, `.claude/reviews/S43-sprint-review-pass-1.md`,
finding N1. It is narrower than the defect S42 hit, since F-X022 was authored
through the local gate, which is why it was not fixed inside F-X025.
**Depends on**: F-X025.
**Test gate**: regression. The module runs in a named CI job, asserted the way
the other job contracts are, and a stale version literal fails that job.

### F-X027, Wire the golden-PNG gate into something (S)
`scripts/golden_png_harness.py` generates deterministic PDFs, rasterises page one
at 150 DPI with the pinned Poppler oracle, and compares decoded pixels against
`scripts/golden_pixel_manifest.json`. `docs/hld/12-testing-strategy.md` describes
it in full. It appears in no `/verify` step and no CI job, so it runs only when
somebody remembers it, and a recorded manifest nothing checks is not a gate.

Filed by the S43 sprint review, finding N2. Pre-existing rather than caused by
S43. It surfaced because F-X021 went looking for what watches PDF output. The
story decides where it belongs, given that it needs `pdftoppm` and a pinned
Poppler build and so cannot sit in the same place as the hash harness.
**Depends on**: none.
**Test gate**: regression. A deliberate rendering change fails the gate wherever
the story puts it, and a clean tree passes it.

### F-X028, Repair the agent-facing documentation drift (M)
`CLAUDE.md` opens by stating that its instructions override default behaviour,
so an error in it propagates into every future session. Five claims in it are
false today, and two more sit in the command surface and the spec set.

`CLAUDE.md:159-170`, "Known defects being carried", lists three defects and says
"Do not 'fix' these". All three were fixed in M1. `MediaNamer::scan` takes the
maximum occupied suffix, `Document` holds `layout_cache` and
`deterministic_layout_cache`, and Caladea ships `LICENSE-Caladea` and
`NOTICE-Caladea` with `bundled_fonts.rs` correctly recording Apache 2.0. The
entry claiming a false licence notice ships today is the most serious, because
it tells an agent to leave a legal defect alone that does not exist.

`CLAUDE.md:15` puts the `rdocx-*` family on crates.io at 0.2.0. It is 0.7.0.

`CLAUDE.md:41` and `:163` place the bundled fonts at `crates/rdocx-layout/fonts/`.
They live in `crates/oxml-layout/`.

`CLAUDE.md:41`, `CLAUDE.md:60` and `docs/hld/10-bindings-spec.md:249` name a
`bundled-fonts` feature. No manifest defines one. Bundled fonts are compiled in
unconditionally and `system-fonts` is the optional feature, so the wheel-building
instruction in the bindings spec names a flag that cannot be set.

`docs/hld/15-build-and-toolchain.md:229-236` states in the present tense that
the shared-version group "is at 0.6.0", that the Python project and rdocx WASM
literals "are also 0.6.0", and that the incubating manifests are "prepared at
explicit version 0.2.0". The trains are at 0.7.0 and 0.3.0. This is the same
paragraph family F-X025 corrected two sentences of, found while confirming the
WASM publication position for F-X030.

`.claude/commands/verify.md:55-57` runs `cargo test -p rdocx-layout
--no-default-features` and tells the reader to rename the package when the
extraction lands. It landed. `CLAUDE.md`, `AGENTS.md` and the CI matrix all name
`oxml-layout`. Both invocations work and neither is a no-op, 87 tests against
62, so this is one gate document disagreeing with every other record rather than
a broken gate.

F-X025 corrected two instances of the same class in the spec set. These are the
third through the twelfth, which is what makes this a story rather than another
one-off patch. Three of them were found while doing something else, which is the
argument for the test gate below rather than another manual sweep.
**Depends on**: none.
**Test gate**: regression. A test asserts that every path, version and feature
name `CLAUDE.md` and `.claude/commands/verify.md` cite resolves against the
workspace, so the next stale claim fails the gate rather than surviving 40
sprints.

### F-X029, Path-filtered CI jobs (M)
`.github/workflows/ci.yml` defines thirteen jobs and no `paths` filter, so every
job runs on every change. A commit that touches only `docs/` currently runs the
workspace test suite, the MSRV suite, both WASM targets, the Python bindings,
the packaging archive build and the pinned-render fidelity job.

The filters that pay: `presentation-fidelity` needs the PowerPoint and shared
crates, `python-bindings` needs the binding crates, `supply-chain` needs the
manifests and the lockfile, `hash-harness` needs anything that can reach the
sample generator, and `prose` needs only tracked Markdown.

**The trap is required status checks, and this story exists to get it right.**
A job skipped by a `paths` filter never reports, so a required check waits
forever and the pull request can never merge. The fix is a gate job that always
runs and reports on behalf of the filtered set, rather than filtering the
required jobs directly. A story that adds filters without handling this converts
a slow pipeline into a stuck one.

Filters must also fail safe. A filter that is too narrow silently stops running
a gate, which is the same class of defect as F-X021 and F-X025: an instrument
reporting green because it never ran.
**Depends on**: none.
**Test gate**: regression. A test asserts, for each filtered job, a changed path
that must trigger it and a changed path that must not, so narrowing a filter by
mistake fails the suite. A docs-only change reports every required check.

### F-X030, Decouple the npm package versions from the Rust family version (S, archived)

**Archived without being started. Its premise was wrong.**

The story claimed that a JavaScript-only fix to `@tensorbee/rdocx-wasm` or
`@tensorbee/rpptx-wasm` could not ship without versioning a Rust family that had
not changed. There is no shipping. Neither package is published anywhere.

`scripts/test_sprint_workflow.py:1337-1349` asserts that the WASM CI job
contains none of `npm publish`, `npm login`, `npm adduser`, `npm token`,
`wasm-pack publish`, `NODE_AUTH_TOKEN`, `NPM_TOKEN`, `--registry`, `id-token:`,
`git tag` or `gh release`. The job packs a bundler tarball and install-tests it
locally, and that is the whole of it.
`docs/hld/15-build-and-toolchain.md` says the same in prose: registry
publication is "unconfigured and unauthorized", and no WASM or npm package
gained publication authority from any release.

So the version inheritance costs nothing. It would begin to cost something on
the day npm publication is authorised, and not before. Recorded in
`02-scope-and-non-goals.md` as a deliberate position rather than an oversight,
so the next reader does not refile this.

**Do not reopen this without first authorising npm publication.** If that
happens, the work is the version split plus the `ci.yml` assertions at
`scripts/test_sprint_workflow.py:1317-1319` and the lockfile package set the
stable preflight asserts.

### F-X031, Require the CI gate in branch protection (S)

F-X029 creates an always-reporting `ci-gate` that represents the result of the
path-filtered CI graph. S44 deliberately stops at the tracked workflow because
changing GitHub branch protection is an external repository mutation. S58 is
the reviewed operational boundary before the two depth milestones begin. Later
jobs continue to report through the same stable aggregate check.

In S58, inspect the reviewed workflow at the sprint head, confirm that
`ci-gate` is still the one stable aggregate check, and configure the repository
ruleset or classic branch protection to require that exact check. Do not remove
existing protections without an explicit reviewed decision. Bind the evidence
to the repository, branch pattern, ruleset or protection identifier, and the
reviewed sprint SHA.

**Depends on**: F-X029, F-X070.
**Test gate**: integration. A docs-only pull request reports a successful
required `ci-gate` while the filtered expensive jobs stay skipped, and a
selected failing job makes the required gate fail.

### F-X032, Expose complete Word layout results (S)

Expose the cached normal-font `WordLayoutResult` and an uncached caller-font
`WordLayoutResult` from `Document` so third-party renderers can consume
positioned pages together with the exact `FontData` and Word source map used by
layout. `layout` and `layout_with_options` return the shared accepted cache as
`Arc<WordLayoutResult>`, while tracked options remain uncached.
`layout_with_fonts` and `layout_with_fonts_and_options` return owned uncached
results. PDF, raster, and page access borrow the neutral layout from those same
bundles. No new layout engine or font-set cache is introduced.

**Depends on**: F-009, F-151, F-X037.
**Test gate**: regression. Every emitted glyph-run font id resolves to returned
font data, repeated default calls share the accepted layout cache,
caller-only family names and bytes appear in the owned result, and tracked
layout neither populates nor replaces the accepted cache. Public caller-font
options must expose different accepted and tracked revision projections.

### F-X033, Integrate PR 36 ordered body items (S)

Integrate Pedro Assumpcao's PR 36 through the active sprint branch while
retaining the contributor commit and GitHub pull-request record. The additive
native `Document::body_items` reader returns direct document-body children in
source order as paragraph, table, body-level content-control, or preserved
unsupported XML views. Existing recursive paragraph and table accessors retain
their current semantics. Self-closing modeled Word body children normalize to
the same typed ownership as paired elements, while foreign and unsupported
empty children remain raw. Python, WASM, and CLI surfaces remain unchanged.

The submitted checks ran against an older base. Retarget the pull request to
the integrated sprint branch, run current-base GitHub CI, and merge it with a
GitHub merge commit. Maintainer hardening and documentation remain separate
from the contributor commit.

**Depends on**: F-X038.
**Test gate**: integration. An in-code document with interleaved body
paragraphs, tables, content controls, and unmodelled XML opens through the
public facade and `body_items` reports every direct child once in exact source
order. Current-base GitHub CI, the submitted focused test, the full package
gate, and the unchanged hash harness also pass.

### F-X034, Reviewed release notes for every release (S)

Every release tag carries reviewed, human-written release notes rather than
only GitHub's generated commit summary. A canonical `/release-notes TAG`
ceremony reads the release plan, completed delivery records, relevant commits,
and contributor history, then prepares the versioned `CHANGELOG.md` section
with highlights, user-visible additions and fixes, compatibility or migration
guidance, and contributor credit. Its generated agent skill keeps the ceremony
identical across tools. The deterministic workflow CLI checks one exact SemVer
tag section with the complete ordered heading set and renders only its reviewed
body without changing the changelog. Missing, duplicate, semantically empty,
or placeholder sections fail. Raw HTML alone is not meaningful release text.
The publish workflow validates the same source before crates.io publication,
renders it once into runner-temporary storage, byte-compares a fresh render
immediately before GitHub release creation, and passes only that artifact to
`gh release create`.

**Depends on**: F-X025.
**Test gate**: regression. The custom command prepares complete notes from the
reviewed release record, its generated skill is in sync, release-note
extraction returns the exact versioned changelog section for both tag families,
missing or incomplete notes fail, validation precedes every crates.io publish
command, and GitHub can consume only the byte-identical reviewed artifact.

### F-X035, Tag rpptx-v0.4.0 (S)

Prepare and publish the complete incubating family at 0.4.0. This is the first
incubating release containing `oxml-chart`, which is required by the current
stable `rdocx` graph and was not published at 0.3.0. All 15 crates.io packages
move together, `rpptx-wasm` remains unpublished, and the reviewed release notes
name the chart addition, compatibility position, and contributors.

**Depends on**: F-X034, F-X037, F-X038.
**Test gate**: release. The incubating metadata regression, full verification,
22-package dry run, archive inventory, supply-chain gate, and unchanged hash
harness pass. After separate final approval, all 15 crates resolve from
crates.io at 0.4.0 and the GitHub release uses the reviewed notes at the exact
sprint SHA.

### F-X036, Tag v0.8.0 (S)

Prepare and publish the complete stable family at 0.8.0 after the incubating
0.4.0 dependency graph is available. The minor boundary covers the intentional
pre-1.0 low-level revision and field model changes plus the additive document
automation, complete-layout, and ordered-body facade APIs. Only the exact seven
stable crates publish. Python, WASM, npm, PyPI, and incubating publication stay
unauthorised. The reviewed release notes describe the new APIs, fixes,
compatibility boundary, and contributor credit.

**Depends on**: F-166, F-167, F-168, F-X032, F-X033, F-X035, F-X038.
**Test gate**: release. The stable metadata regression, full verification,
22-package dry run, archive inventory, supply-chain gate, and unchanged hash
harness pass. After separate final approval, all seven stable crates resolve
from crates.io at 0.8.0 and the GitHub release uses the reviewed notes at the
exact sprint SHA while PR 36 credit remains visible.

### F-X037, Trace Word glyphs to source paragraphs (M)

Carry format-neutral source spans through shaping, both line-splitting stages,
pagination, and positioned glyph output. `rdocx-layout` returns a typed
`WordLayoutResult` whose result-local side table resolves each source node to a
document, table, nested-table, header, footer, footnote, or endnote paragraph
path. Character ranges use Unicode scalar indices in the selected revision
projection. Generated markers, dynamically evaluated fields, and text whose
display transformation cannot preserve an exact source slice remain
unattributed rather than reporting a false location.

The existing `layout_document` functions retain their `LayoutResult` return
type and discard provenance. `layout_document_with_provenance` and
`layout_document_deterministic_with_provenance` return the Word-specific
bundle. The field model exposes its parsed-complex projection ownership through
`Field::projected_text`, so identical cached and literal text cannot shift a
later range. F-X032 exposes the bundle through cached and caller-font facade
paths, so external renderers receive pages, fonts, and source resolution
together.

This is an intentional low-level pre-1.0 source break for exhaustive
`TextSegment` and `GlyphRun` literals and belongs in the planned 0.4.0 and
0.8.0 release notes. It does not add rendering support for content that the
current layout engine skips.

**Depends on**: F-009, F-151.
**Test gate**: regression. Every attributed glyph run resolves to one exact
Word paragraph path and Unicode-scalar range whose projected text equals the
run text across ASCII, CJK, wrapping, tables, nested tables, headers, footers,
footnotes, endnotes, and both revision views. Both splitting stages preserve
contiguous ranges. Generated markers, evaluated fields, and non-bijective text
transformations remain unattributed. Caller-font and cached layouts carry the
same complete source map, packaged crates remain below 10 MiB, WASM checks
pass, and all 49 hash entries remain unchanged. The repeated-text field
regression proves that parsed complex fields advance projection offsets and
new simple fields do not.

### F-X038, Cache relayout work across document edits (L)

The normal-font path reuses the expensive work an interactive editor repeats
after every document mutation. Bundled plus system fonts are discovered once
per process. File-backed face bytes are shared by canonical file identity.
Shaping uses complete exact keys, and each document retains one synchronized
normal engine with a bounded paragraph cache. Deterministic and caller-provided
fonts remain isolated from the system snapshot.

Only context-independent body paragraphs reuse blocks. The complete context and
key compare styles, theme, embedded fonts, width, revision view, and typed
paragraph content. Numbering, drawings, fields, hyperlinks, media,
relationships, and other traversal-sensitive content bypass reuse. Diagnostics
and exact bounded font traces travel with each block. Whole-layout publication
is transactional, and cached scalar ranges are rebound to the current F-X037
result-local source nodes. Process, shaping, coverage, trace, pending, and
published caches have explicit entry and true retained-byte bounds. Poisoned
process locks recover without disabling later layout.

Issue 39 supplied the profiling, cache decomposition, and prototype. Credit
`@emptinessform` in both release families. The reported 1,144 ms to 101 ms
improvement is evidence, not a machine-independent CI threshold. Normal system
font discovery is a process-lifetime snapshot, so installing or replacing
system fonts requires a process restart. Deterministic and caller-font behavior
does not change.

**Depends on**: F-X037, F-X032.
**Test gate**: regression. A warm normal-font relayout equals a cold result in
pages, fonts, diagnostics, revision view, and resolved source provenance while
rebuilding only the changed safe paragraph. Complete context changes cannot
serve stale blocks, shaping keys never alias different content, process and
engine caches stay bounded and recover from poison, `Document` remains
`Send + Sync`, both WASM targets compile, and all 49 hashes remain unchanged.

### F-X039, Share layout payloads and transfer reusable engines (M)

Remove the remaining deep copies on the interactive layout boundary.
`FontData::data` shares immutable font bytes, and laid-out pages use shared
immutable page frames so a cached result or unchanged pagination tail can be
retained without copying its complete payload. These are intentional low-level
pre-1.0 type breaks. PDF, raster, page access, caller-font layout, and the
Word-specific provenance bundle continue to consume the same data.

An editor that rebuilds a `Document` for undo or redo can transfer reusable
normal-layout work through a checked ownership API. The receiving document
must validate the complete F-X038 context before reading any retained entry,
and a failed or incompatible transfer cannot replace its current engine or
publish a stale result. The design chooses the smallest explicit ownership
surface. It does not expose unchecked cache mutation.

Issue 39 supplied both measured proposals. Credit `@emptinessform` in the next
release notes that contain this work.

**Depends on**: F-X032, F-X037, F-X038.
**Test gate**: regression. Cloning complete layout results shares font bytes and
page frames, while PDF, raster, provenance, diagnostics, and visible output
remain identical. A transferred engine reuses safe work for the same complete
context, rejects or invalidates stale context, remains bounded and poison-safe,
and leaves deterministic and caller-font paths isolated. Both WASM targets,
package dry runs, and the hash harness pass unchanged.

### F-X040, Restart pagination and cache table blocks (L)

Make the reusable normal engine resume pagination from a safe checkpoint before
the first changed block and attach an unchanged tail when the complete pager
state and page boundary match. Checkpoints exist only where no paragraph,
footnote continuation, float, or other carried state crosses the boundary.
Initial reuse may fall back to a full pass for multiple sections or floating
drawings. Environment identity includes section geometry, headers, footers,
notes, styles, numbering, theme, fonts, media, revision view, and every other
input that can change page output.

Table blocks gain the same transactional, diagnostic-preserving, bounded cache
discipline as safe paragraphs. Tables with numbering or note references bypass
reuse until their traversal state can be represented exactly. Before either
cache is trusted, fix the existing paragraph-cache case where inserting an
earlier footnote reference changes a later marker number without invalidating
the cached block. That correctness regression remains independent of the
restart algorithm.

Issue 39 supplied the checkpoint, tail-splice, and table-cache prototypes plus
the footnote-marker observation. Credit `@emptinessform` in the next release
notes that contain this work.

**Depends on**: F-X038, F-X039.
**Test gate**: regression. Warm pagination after edits at the start, middle,
tail, and a page boundary equals a fresh engine in pages, fonts, diagnostics,
provenance, numbering, notes, fields, and outlines while rebuilding only the
bounded affected page range. Insertions, deletions, style and numbering edits,
footnote-marker renumbering, multi-section fallback, floating-drawing fallback,
failed layouts, and table cache bounds all have explicit cold-versus-warm
evidence. The hash harness remains unchanged.

### F-X041, Remove duplicated glyphs at break opportunities (M)

Make one stage own Unicode line-break segmentation and shaping. The current
Word conversion path slices shaped glyph arrays at break opportunities before
the shared line breaker reshapes them again. Approximate glyph slicing is not
valid for ligatures or other non-bijective shaping, and can place a boundary
glyph in both adjacent positioned runs even when the break is not taken.

Preserve exact text, spacing, source spans, hyperlinks, fields, note markers,
and formatting while ensuring every emitted chunk is shaped from exactly its
own text. The correction applies above PDF and raster output so third-party
renderers consuming `PageFrame` see the same fixed runs. Issue 23 and the
additional UAX 14 diagnosis came from `@emptinessform`, who should be credited
in the next release notes that contain the fix.

**Depends on**: F-030, F-104, F-X037.
**Test gate**: golden. Deterministic layout of spaces, hyphens, ligatures,
combining text, CJK, and untaken versus taken break opportunities emits each
source scalar and shaped glyph exactly once with contiguous provenance. The
reported `ttf-parser`, doubled-space, `financial`, and `allocated` cases are
covered through `PageFrame` and both built-in backends. The intentional sample
hash delta is isolated, explained, and reviewed.

### F-X042, Prove headers and footers in PDF output (S)

Close Issue 15 with an end-to-end public regression rather than another
model-only assertion. Author, save, reopen, lay out, and render a document with
default, first-page, even-page, inherited, and multi-section headers and
footers. Verify relationship resolution, `titlePg` selection, page-frame
placement, and final PDF text. If the fixture exposes a remaining drop, fix
only that path. If every case already passes, retain the regression as closure
evidence for the previously unreproduced report.

**Depends on**: F-168, F-X032.
**Test gate**: integration. A readable in-code package passes through the
public `Document` facade and produces the expected header and footer text on
each applicable page in both `WordLayoutResult` and deterministic PDF output.
Blank first or even variants do not borrow defaults, inherited variants remain
visible, unrelated package parts survive, and the hash harness is unchanged.

### F-X043, Reuse bundled-fallback caller-font layouts (M)

Expose a native Word layout path where caller fonts override the deterministic
bundled set and missing families resolve only from bundled fonts. Retain one
private reusable engine for this mode across edits and support undo or rebuild
transfer through an exact-context checked facade. Do not expose raw engine
take or set operations, and do not let this path observe the system-font
snapshot. PRs 40 and 41 supplied the concrete editor use case and prototype.

**Depends on**: F-X039, F-X040.
**Test gate**: regression. An incomplete caller set resolves bundled fallback
families while the strict caller-only path still fails, caller faces win for
the same family, compatible checked transfer records safe hits, font or
document context changes preserve both engines and reject reuse, staged
mutations retain valid work, warm and cold pages, fonts, diagnostics, and
provenance are equal, both WASM targets pass, and all hashes remain unchanged.

### F-X044, Scale paragraph-cache lookup for editors (M)

Remove editor-scale paragraph-cache thrash without weakening F-X040's exact
identity or traversal invalidation. Use a compact fingerprint only as a
prefilter before authoritative typed equality, avoid cloning the complete
paragraph key and linearly removing a hit, and size the bounded cache from
retained-memory evidence on the reported 700-paragraph workload. Optional
timing instrumentation must cost nothing when disabled. PR 41 supplied the
profile and prototype.

**Depends on**: F-X040.
**Test gate**: regression. Forced fingerprint collisions cannot alias typed
paragraphs, unsafe traversal content disables later reads, late failure
publishes nothing, entry and retained-byte bounds hold under eviction, a
700-paragraph warm edit avoids cache thrash, complete warm and cold outputs are
equal, disabled timing adds no runtime work, and the hash harness is unchanged.

### F-X045, Cache headers and footers transactionally (M)

Cache reusable header and footer layout blocks under the same transactional,
diagnostic-preserving, source-rebinding, and retained-memory discipline as safe
body blocks. Exact typed identity covers complete section geometry, referenced
parts, media bytes, revision view, font context, and provenance. First, even,
default, inherited, image, and watermark variants remain distinct. PR 41
supplied the optimization prototype, whose hash-only and unbounded form is not
accepted.

**Depends on**: F-X040, F-X042.
**Test gate**: regression. Safe header and footer blocks hit and replay exact
diagnostics, fonts, and provenance, while part text, media, watermark, page
height, variant, section, or context changes miss. Late failure publishes no
entry, combined entry and retained-byte ceilings hold, warm and cold outputs
are completely equal, and the hash harness is unchanged.

### F-X046, Reuse substituted pages exactly (S)

Retain bounded pristine and field-substituted page pairs so repeated PAGE,
NUMPAGES, and PAGEREF post-processing does not reshape an unchanged page. Reuse
requires exact total-page, bookmark-target, page-content, font, and pristine
page identity, and retained pairs count against the existing restart budget.
PR 41 supplied the optimization prototype.

**Depends on**: F-X040.
**Test gate**: regression. Stable PAGE, NUMPAGES, and PAGEREF pages reuse their
substituted frames, while page-count, bookmark, content, or font changes miss.
Field-free sharing is preserved, eviction respects entry and byte bounds,
complete warm and cold outputs match, and the hash harness is unchanged.

### F-X047, Attribute empty Word paragraphs (S)

Represent an empty Word paragraph with one zero-width empty text segment using
the resolved default font and a source span of `0..0`. This gives interactive
callers a caret target and correct line height without emitting a visible
glyph. Cover body, table, header, footer, footnote, and endnote stories while
keeping provenance and non-provenance layouts structurally compatible. PR 41
supplied the behavior prototype.

**Depends on**: F-X037.
**Test gate**: regression. Every supported empty Word story emits exactly one
zero-width attributed segment with resolved default metrics and the correct
source identity, non-empty paragraphs are unchanged, provenance and ordinary
layouts remain structurally equal, both backends render no new glyph, and the
deterministic hash harness remains unchanged.

### F-X048, Dense form table fidelity (L)

Close Issue 42 and PR 43 on the current hardened layout engine. Dense forms
must retain nested tables as recursively positioned table blocks, distribute
vertical-merge content across its exact grid span, honour `trHeight` rules,
and resolve table-style borders and paragraph properties through the applicable
`basedOn` and conditional style layers. Table-style properties stay
namespace-aware, schema-ordered, and byte-preserved through an unchanged
round trip.

Cell-anchored foreground drawings render in the cell coordinate space, while
`behindDoc` drawings join the page behind layer. Explicit `nil` cell borders
retain their normal suppressing meaning except at the exact outer-table edge
where the pinned Word oracle proves the reported compatibility behavior. Empty
paragraphs use paragraph-mark metrics without emitting glyphs, and the native
paragraph facade can append a run with the mark's direct run properties.

Do not merge or cherry-pick PR 43 as a stack. It contains the superseded PR 41
engine and cache surfaces, conflicts with current main, serializes table-style
properties outside schema order, uses local-name-only parsing for a typed
projection, undercounts new retained cache payloads, and has no focused tests
for the seven reported behaviors. Reimplement the useful behavior against the
transactional caches, exact context identity, source rebinding, and retained
memory ceilings completed in F-X040 through F-X047. Credit `@emptinessform`
for Issue 42, PR 43, the real receipt diagnosis, and the corpus measurements.

**Depends on**: F-X040, F-X045, F-X047.
**Test gate**: golden. A readable in-code dense form covers nested tables,
mixed grid spans, vertical merges, exact and minimum row heights, direct and
conditional table styles, outer and interior `nil` borders, 7pt empty cells,
and foreground and behind-cell anchors. It renders as one page in deterministic
PDF and raster output with the reviewed Word reference geometry, while
round-trip XML, provenance, warm-cold cache equality, transactional failure,
retained-memory bounds, both WASM targets, and the declared isolated hash delta
all pass.

### F-X049, Tag rpptx-v0.5.0 (S)

Prepare and publish the complete incubating family at 0.5.0 after the S52 and
S53 package, layout, and PDF work is complete. The minor boundary covers Agile
encryption, digital signatures, shared layout payloads, corrected shaping,
semantic tagged PDF, PDF/A output, and any shared redaction support. All 15
crates.io packages move together and `rpptx-wasm` remains unpublished.

The reviewed release notes cover only the incubating family. They identify the
issues and pull requests whose shared implementation is present, link the
records, and credit verified reporters and contributors, including
`@emptinessform` for PRs 40 and 41. Each included GitHub issue or pull request
receives a maintainer comment naming the release and the final implementation
boundary. Publication requires its own immediate approval at the reviewed SHA.

**Depends on**: F-172, F-173, F-174, F-175.
**Test gate**: release. The incubating metadata regression, full verification,
22-package dry run, archive inventory, supply-chain gate, WASM isolation, and
declared hash results pass. After separate final approval, all 15 crates
resolve from crates.io at 0.5.0 and the GitHub release body is byte-identical to
the reviewed `rpptx-v0.5.0` changelog section with verified contributor credit.

### F-X050, Tag v0.9.0 (S)

Prepare and publish the complete stable family at 0.9.0 after incubating 0.5.0
is available. The minor boundary contains the S52 encryption, signature,
layout correctness, editor reuse, and provenance work plus S53 signature
creation, accessible PDF, redaction, and dense-form fidelity. Only the exact
seven stable crates publish. Python, WASM, npm, PyPI, and incubating
publication remain unauthorized.

The reviewed release notes link and credit every included external report and
pull request. At minimum this release records Issues 15, 23, 39, and 42 plus
PRs 40, 41, and 43, with verified credit to `@mantissaman` and
`@emptinessform`. Each record receives a maintainer comment stating whether it
landed directly or through a hardened equivalent, naming the release, and
thanking the contributor. PR 43 remains open until F-X048 lands and then
closes as addressed rather than merged. Publication requires a new immediate
approval at the exact reviewed SHA.

**Depends on**: F-172, F-173, F-174, F-175, F-X048, F-X049.
**Test gate**: release. The stable metadata regression, full verification,
22-package dry run, archive inventory, supply-chain gate, binding and WASM
isolation, and declared hash results pass. After separate final approval, all
seven stable crates resolve from crates.io at 0.9.0 and the GitHub release body
is byte-identical to the reviewed `v0.9.0` changelog section with all verified
issue, pull-request, reporter, and contributor credit.

### F-X051, Honor caller-supplied font family aliases (M)

Make the `family` value supplied with a caller font a document-facing alias for
the font's embedded family name. Resolve an exact embedded family first, then a
caller alias, then the existing mapped and generic fallbacks. A label equal to
the embedded family adds no alias and leaves existing callers unchanged.

Expose byte-free alias mappings so many document-facing names can target one
loaded family without cloning the font bytes for every name. Alias identity
belongs to the reusable engine's font context. An unchanged mapping is a no-op,
while a changed mapping invalidates resolution-dependent state without
discarding unrelated valid work. Reimplement Issue 44 and PR 45 against the
current bounded cache and exact-context contracts, and credit `@emptinessform`
in the next release that contains the behavior.

**Depends on**: F-X043.
**Test gate**: regression. Multiple document-facing aliases resolve to the
intended caller font without repeated bytes, exact embedded-family requests
retain priority, and unmapped requests keep the existing fallback order.
Unchanged aliases reuse safe work, changed aliases miss the affected caches,
warm and cold pages, fonts, diagnostics, and provenance are equal, both WASM
targets pass, and the deterministic hash harness remains unchanged.

### F-X052, Restore interactive relayout performance (L)

Close Issue 46 on the hardened reusable layout engine. A generated 700
paragraph, 14 table mixed Korean and Latin document must no longer pay
whole-document `Debug` formatting, deep copies of unchanged page frames, or
full retained-context cloning on each body-only edit. Restart and paragraph
identities may use cheap stable prefilters, but exact typed equality remains
the collision authority. Unchanged raw and substituted page frames retain
their shared ownership across restart and tail attachment.

Checked transfer must accept a restored document whose body changed while its
styles, numbering, sections, related stories, theme, fonts, caller aliases,
revision view, and other retained-work inputs remain equal. It must still
reject every real context change without consuming either engine. The faster
path keeps transactional publication, complete invalidation, bounded memory,
diagnostic replay, current provenance, deterministic font isolation, and the
semantic `MarkedContent` structure introduced by F-173.

Use the Issue 46 workload and the `svg-poc-0.8` reference implementation as an
interleaved A/B oracle. The gate covers document load, a mid-document typing
edit, checked undo transfer, and table mutation through native and bundled
fallback paths. Credit `@emptinessform` for the report, measurements, and
candidate-build verification in the next release containing the correction.

**Depends on**: F-X039, F-X040, F-X043, F-X044, F-X045, F-X046, F-173,
F-X048, F-X051.
**Test gate**: regression. Instrumented tests prove a one-paragraph edit does
no whole-document debug serialization or deep copy of unchanged prefix and
tail pages, reports cache hits for every unchanged safe block, rebuilds only
the affected restart region, and accepts an exactly compatible restored-body
transfer. Warm output remains exactly equal to a fresh engine in pages,
structure, fonts, diagnostics, provenance, numbering, notes, fields, and
outlines. Retained and pending memory remain bounded, both WASM targets pass,
and interleaved release measurements for load, typing, undo, and table mutation
are no more than 1.25 times the reference on the same machine and workload.

### F-X053, Complete layout migration and contribution records (S)

Finish the documentation and GitHub record work left by Issues 44 and 46 and
PR 45. Amend the v0.9.0 compatibility section and its published GitHub release
body to state that `PositionedElement` is non-exhaustive and visible content is
nested under `MarkedContent`. External backends must recurse through
`MarkedContent::children` or use `oxml_layout::walk` when consuming
`PageFrame::elements`.

After F-X052 passes, close Issue 44 and PR 45 as addressed by the hardened
F-X051 implementation rather than merged, and close Issue 46 with both the
performance correction and migration-note evidence. Preserve authenticated
`@emptinessform` credit and all three record links for the next stable release
that contains F-X051 and F-X052. The confirmation comments on Issues 39 and 42
are acceptance evidence only. The note-operation regression is gone and the
dense-form corpus is covered, so neither closed issue gains duplicate scope.

**Depends on**: F-X051, F-X052.
**Test gate**: integration. The tracked v0.9.0 changelog section and published
release body are byte-identical after the compatibility correction, the note
names both supported recursive traversal choices, Issue 44 and PR 45 cite the
F-X051 implementation, Issue 46 cites F-X052 and the migration correction,
Issues 39 and 42 remain closed, and the next stable contribution inventory
retains the authenticated reporter and contributor credit.

### F-X054, Integrate PRs 47 through 52 (L)

Audit the six open reader contributions from authenticated contributor
`@pedroassumpcao` against current main, then land each supported outcome either
directly or as a hardened equivalent. PRs 47, 48, and 49 expose ordered cell,
run, hyperlink, and paragraph children without flattening nested tables,
content controls, revisions, fields, notes, comments, bookmarks, drawings, or
preserved XML. PR 50 exposes stable facts for unsupported body content without
inventing raw bytes for modeled constructs. PR 51 preserves producer-defined
numbering formats, and PR 52 rejects undecodable visible text instead of
silently substituting an empty string.

The ordered iterators must remain borrowed, source ordered, bounded by retained
document state, and namespace aware. Open-ended public item enums are
non-exhaustive before the eventual 1.0 boundary. Unsupported XML classification
must use the existing XML parser and in-scope namespace declarations rather
than a new ad hoc byte scanner. The PR 51 public `ST_NumberFormat` change
removes `Copy` and adds retained producer values, so its source incompatibility
is deliberate and must be explicit in the next pre-1.0 compatibility notes.
Every deviation
from an original patch is recorded with the reason, and all six direct
pull-request links and specific contribution outcomes remain in the v0.10.0
inventory.

Do not merge the GitHub pull requests merely to claim attribution. Integrate
the reviewed code through the repository lifecycle. After v0.10.0 publishes and
its body verifies, post one release-bound maintainer comment to each pull
request stating whether it landed directly or through a hardened equivalent,
thank `@pedroassumpcao`, and close the record without a merge if that is the
truth.

**Depends on**: F-X033.
**Test gate**: regression. Source-built documents prove exact direct ordering
for body, cell, paragraph, hyperlink, and run children across every supported
typed variant and preserved XML boundary. Prefix aliases, inherited namespace
scope, modeled unsupported facts, producer-defined numbering formats,
undecodable ordinary and deleted text, save and reopen equality, exhaustive
public documentation, and the unchanged legacy flattened accessors all pass.
The full stable API diff identifies the intentional PR 51 incompatibility and
no unreviewed breaking change.

### F-X055, Tag v0.10.0 (S)

The immutable v0.10.0 attempt prepared the exact seven-package stable family
after the M18 writers and F-X054, with the intentional pre-1.0 compatibility
boundary documented in its reviewed notes. The annotated tag was created at
the reviewed S56 SHA, and the workflow published `rdocx-opc` and `rdocx-oxml`.
Package verification then stopped at `rdocx-layout` because its source used a
shared layout API newer than the published 0.5.0 shared family. The other five
stable packages and the GitHub release were not published.

The v0.10.0 tag and two registry entries remain immutable. No contribution
notification or PR 47 through 52 closure is attributed to that partial
attempt. F-X056 publishes the required shared family at 0.6.0, and F-X057 owns
the coherent seven-package stable recovery at 0.10.1, including the reviewed
notes, contributor notifications, and authorized PR closures. Python, WASM,
npm, PyPI, and incubating publication were not authorized by the v0.10.0 tag.

**Depends on**: F-180, F-181, F-182, F-X051, F-X052, F-X053, F-X054.
**Test gate**: release. Preparation, full verification, package dry runs,
binding isolation, notes, inventory, and the declared hash result passed at the
reviewed SHA. Publication did not complete because the shared registry graph
could not verify `rdocx-layout`. The immutable partial result is the input to
the F-X056 and F-X057 recovery gates, not a completed stable-family release.

### F-X056, Tag rpptx-v0.6.0 (S)

The complete 15-package incubating family is published at 0.6.0 from reviewed
SHA `55fb2f54caf91d7dedc8936b4c7b116354590628` before the stable release retry.
The v0.10.0 publication attempt proved that the stable source
graph uses shared layout APIs added after the immutable 0.5.0 registry
boundary. `rdocx-layout` therefore cannot verify against crates.io until the
current shared family has its own reviewed release.

Move all 15 publishable incubating manifests, the 15 workspace pins, the
sixteenth preparation-only manifest, lockfile records, README requirements,
WASM metadata, CI
literals, release regressions, and the reviewed changelog section to 0.6.0.
Keep the stable family at 0.10.0 while this separate tag publishes. The exact
registry set, owners, annotated tag, release body,
and selected-family contribution notifications verify before F-X057 starts.

**Depends on**: F-X051, F-X052, F-X053, F-X054.
**Test gate**: release. All 15 incubating registry entries resolve at 0.6.0 from
the reviewed SHA, their owners match the authenticated registry inventory, the
GitHub release body is byte-identical to the reviewed notes, every included
external record receives its reviewed notification, and no stable 0.10.1
package publishes from this tag.

### F-X057, Tag v0.10.1 (S)

The stable workspace and all seven stable packages are published at 0.10.1
from reviewed SHA `ae0dcb162a7805e59e5890464b226765645ad547` after the
immutable partial v0.10.0 attempt. Stable workspace pins, binding metadata,
README requirements, CI literals, lockfile records, and release regressions
remain coherent at that version. Every shared dependency is pinned to the
verified incubating 0.6.0 family from F-X056.

The 0.10.1 notes describe the complete stable outcome and the v0.10.0 partial
publication accurately. The two registry packages already present at 0.10.0
remain immutable. All nine reviewed stable release comments are verified, and
PRs 47 through 52 are closed unmerged with their hardened-equivalent status.

**Depends on**: F-180, F-181, F-182, F-X051, F-X052, F-X053, F-X054, F-X056.
**Test gate**: release. All seven stable registry entries resolve at 0.10.1
against incubating 0.6.0 dependencies, their owners match the authenticated
registry inventory, the annotated tag targets the reviewed SHA, the GitHub
release body is byte-identical, all nine stable contribution notifications are
verified, and PRs 47 through 52 close with their reviewed hardened-equivalent
status.

### F-X058, Shared multilingual text substrate (L)

The shared layout family must own one complete text contract before stable Word
consumers can use language-aware hyphenation, complex-script shaping, or
bidirectional layout. Add conditional-hyphen opportunities, script and font
segmentation, cluster and offset preservation, complex-script line boundaries,
paragraph and run direction, and line-local visual ordering in the existing
incubating layout, drawing, PDF, and Presentation paths. Add the approved
deterministic multilingual fonts and legal files without changing legacy Latin
output. Stable Word property parsing, facade authoring, and final Word oracle
acceptance remain in F-198, F-199, and F-200.

**Depends on**: F-196, F-197, F-X061.
**Test gate**: regression. Shared deterministic tests prove conditional
hyphens, exact logical source spans, cluster-safe Arabic and Indic shaping,
Thai and CJK breaking, bidi visual order, searchable logical text, and
unchanged legacy Latin hashes.

### F-X059, Tag rpptx-v0.7.0 (S)

The complete 15-package incubating family is published at 0.7.0 after F-X058
from the annotated `rpptx-v0.7.0` tag at reviewed SHA
`1b076c16fb494fe47b054d761e061181a1ea0b15`.
Every incubating manifest, workspace pin, lock record, README requirement, CI
literal, release regression, and the unpublished `rpptx-wasm` preparation
carrier moves together. The stable family stays at 0.10.1, and the immutable
`rdocx-layout@0.10.1` registry graph continues to resolve
`oxml-layout@0.6.0`. The published 0.7.0 family is the registry boundary that
F-198, F-199, and F-200 must compile and run against.

**Depends on**: F-X058.
**Test gate**: release. All 15 incubating registry entries resolve at 0.7.0
from the reviewed SHA, their owners match the authenticated registry inventory,
the tag and GitHub release body match the reviewed evidence, every selected
external record receives its reviewed notification, and no stable package is
published.

### F-X060, Tag v0.11.0 (S)

The immutable v0.11.0 attempt prepared the stable workspace and exact
seven-package family at reviewed SHA
`25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4`. The annotated `v0.11.0` tag
targets that SHA. The release workflow published `rdocx-opc` and
`rdocx-oxml`, then stopped while verifying `rdocx-layout` because current
stable source uses `TextSegment.direction`, which is newer than the published
`oxml-layout@0.7.0` registry contract. The other five stable packages and the
GitHub release were not published.

No contribution notification is attributed to the partial attempt. Issues 53
and 54 and PRs 55 through 58 remain open. F-X068 publishes the required shared
family at 0.8.0, F-X069 owns the coherent seven-package stable recovery at
0.11.1 and its six leave-open notifications, and F-X070 owns the separately
approved post-recovery yank of the two incomplete 0.11.0 registry entries. The
v0.11.0 tag is never moved or deleted. Python, WASM, npm, and PyPI packages
remain outside publication authority.

**Depends on**: F-198, F-199, F-200, F-202, F-X059, F-X062, F-X063, F-X064, F-X065, F-X066, F-X067.
**Test gate**: release. Preparation and every local gate passed at the reviewed
SHA. Publication did not complete because the shared registry graph could not
verify `rdocx-layout`. The immutable partial result is the input to F-X068,
F-X069, and F-X070, not a completed stable-family release.

### F-X068, Tag rpptx-v0.8.0 (S)

The complete 15-package incubating family is published at 0.8.0 from the
immutable annotated `rpptx-v0.8.0` tag at reviewed SHA
`7f4414b0aeef1ec2cbae75fcb5aa96ab6dee6d70`. It supplies the additive
`TextSegment.direction` contract required by stable source after the immutable
0.7.0 shared release. All 15 registry entries resolve under sole owner
`mantissaman (Atul Sharma)`, the release body matches the reviewed notes, and
`rpptx-wasm@0.8.0` remains absent from crates.io. The stable family is published
at 0.11.1 and pins this shared boundary.

**Depends on**: F-200, F-X064, F-X065, F-X066, F-X067.
**Test gate**: release, passed. All 15 incubating registry entries resolve at 0.8.0
from the reviewed SHA, their owners match the authenticated registry inventory,
the annotated tag and GitHub release body match the reviewed evidence,
`rpptx-wasm@0.8.0` is absent, and no stable package publishes from this tag.

### F-X069, Tag v0.11.1 (S)

The complete stable recovery is published at 0.11.1 against the published
shared 0.8.0 family from the immutable annotated `v0.11.1` tag at reviewed SHA
`5a850ce9ae6c31f8365594ed2970193266f8b2a6`. Every stable carrier, internal pin, lockfile record,
Python metadata value, WASM contract literal, CI identity, README requirement,
release regression, and reviewed changelog section is at 0.11.1. Every shared
dependency remains pinned to 0.8.0. The release publishes exactly `rdocx-opc`,
`rdocx-oxml`, `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and
`rdocx-cli` in dependency order.

The published notes describe the partial v0.11.0 attempt accurately. The
selected contribution inventory credits authenticated `@emptinessform` for
Issues 53 and 54 and authenticated `@pedroassumpcao` for PRs 55 through 58.
Each record has exactly one release-bound thank-you and remains open.

**Depends on**: F-198, F-199, F-200, F-202, F-X062, F-X063, F-X064, F-X065, F-X066, F-X067, F-X068.
**Test gate**: release, passed. All seven stable registry entries resolve at 0.11.1
against incubating 0.8.0 dependencies, their owners match the authenticated
registry inventory, the annotated tag targets the reviewed SHA, the GitHub
release body is byte-identical to the reviewed notes, and all six leave-open
notification URLs verify.

### F-X070, Yank incomplete v0.11.0 packages (S)

After the complete v0.11.1 family verifies, remove the two incomplete v0.11.0
registry entries from ordinary dependency selection without rewriting release
history. After separate final approval, yank exactly `rdocx-opc@0.11.0` and
`rdocx-oxml@0.11.0`. The cleanup is complete. The annotated `v0.11.0` tag
remains immutable, no v0.11.0 GitHub release exists, and the other five 0.11.0
packages never existed. Complete coherent stable releases remain live and
unyanked. The cleanup changes no other registry version, tag, release,
notification, issue, pull request, or external contribution-record state.
Normal local sprint ledgers, progress notes, review artifacts, and handoff
records still advance through the feature workflow.

**Depends on**: F-X069.
**Test gate**: integration. crates.io readback reports the two incomplete
0.11.0 entries yanked, the other five absent, and all seven 0.11.1 entries
live and owned by the authenticated publisher. The immutable v0.11.0 tag still
targets the reviewed partial-attempt SHA and no v0.11.0 GitHub release exists.

### F-X061, Support staged dependency checkpoints in run-sprint (S)

`/run-sprint` detects when a later wave depends on an integrated and reviewed
F-ID that is not completed, then uses a resumable checkpoint before that
consumer. The route verifies and completes the dependency prefix, commits its
clean review evidence, records review at that resulting HEAD, reruns full
verification, and returns the same sprint state to implementation. A release
dependency extends that route with preparation, publication, and its separate
approval. Review evidence remains bound to the prefix, prepared release,
post-publication evidence, and final closure HEADs without a self-confirming
review loop. Resuming an existing run refreshes canonical title and size
metadata and discovers new F-IDs without discarding state, ownership, worker,
review, or verification facts.

**Depends on**: none.
**Test gate**: regression. The workflow contracts, A to B to C state regression,
and phase regression prove ordinary and release dependency checkpoints can
return to implementation before the final close-preflight without weakening
release approval or HEAD-bound evidence.

### F-X062, Reuse restart pagination with notes and headers (M)

The restart paginator admits documents with footnotes, endnotes, headers, and
footers when their retained context and body note-reference sequence are
exactly equal. It restarts only at note-clean page boundaries. Changed related
stories, changed note references, note-bearing tables, and other
traversal-sensitive content retain conservative full fallback. Endnote pages
append exactly once after a complete restarted body or arrive through an exact
cached tail. F-202 separately owns the 1,024-page capacity.

**Depends on**: F-202.
**Test gate**: regression. Source-built 700-paragraph note and header/footer
workloads retain bounded page work and exact warm-versus-fresh output, while
changed related stories and dirty note continuations invalidate reuse.

### F-X063, Avoid duplicate caller-font byte comparisons (S)

Issue 54 isolates a WASM relayout regression to a second exact comparison of
caller font bytes. `FontManager::load_additional_fonts` already performs the
authoritative ordered family-and-byte comparison. Normal warm relayout uses a
private font-elided retained-context comparison only after the font manager
reports that exact set unchanged. The retained context keeps exact bytes, and
checked engine transfer retains the complete ordered family-and-byte check.
Equal-length changed bytes invalidate both normal reuse and checked transfer.

**Depends on**: F-X052.
**Test gate**: regression. Five generated caller fonts totalling about 22 MiB
and 40 aliases perform zero repeated retained-context font-byte work on warm
layout, same-length changed bytes still invalidate reuse, checked transfer stays
exact, and warm output equals fresh output across positioned pages, font data,
diagnostics, outlines, provenance, and PDF bytes.

### F-X064, Accept whole-valued decimal table measurements (S)

PR 55 supplies the Word-produced `9345.0` compatibility case. The existing
signed integer projection uses one exact string parser for table widths, cell
widths, table indents, and default cell margins. It accepts integers and
decimals whose nonempty fractional portion contains only zeroes, then
checked-parses the integer portion into `i32` without floating point. Missing
values retain their existing default. Fractional decimals, exponent forms,
empty fractions, overflow, malformed input, percentages, and universal
measures fail explicitly rather than becoming zero. The latter two remain
unsupported union arms until a lossless public model is designed.

**Depends on**: F-X059.
**Test gate**: regression. Namespace-aware parser and canonical round-trip
tests cover every table-width site, negative lexical forms, unsupported union
arms, and the current Word corpus with 49 of 49 output hashes unchanged.

### F-X065, Expose tracked table grid changes (S)

PR 56 exposes the historical grid carried by `w:tblGridChange`. Recognize the
grid, active columns, and historical change by WordprocessingML namespace URI,
preserve exactly one change subtree after active columns in schema order, and
fail closed on a duplicate modeled change. Foreign same-local children remain
unmodelled with their exact bytes preserved. The active columns remain the only
layout grid. Native callers can query `TableRef::has_grid_change()`, while the
historical bytes remain inspection and round-trip data. The public low-level
grid fields are an intentional pre-1.0 exhaustive-literal source impact.

**Depends on**: F-X064.
**Test gate**: regression. Aliased and foreign namespace cases, duplicate
rejection, package save-reopen, and layout prove the historical grid is
preserved without changing active column widths or the 49 output hashes.

### F-X066, Classify legacy VML horizontal rules (S)

PR 57 adds a native reader classification for an unambiguous legacy horizontal
rule. Recognize a WordprocessingML `pict` containing exactly one VML `rect`
whose Office `hr` attribute is enabled by expanded namespace URI, not lexical
prefix. Accept the VML true forms `t` and `true`, and preserve and expose the
exact raw bytes. Numeric `1`, false, missing, malformed, foreign,
multiple-shape, visible-child, and ambiguous input remains `UnsupportedXml`.
Classification occurs once at the OXML parse boundary and records a compact
semantic flag in the existing raw-child position sidecar. Ordinary modeled
runs retain no namespace scope, and run equality includes the classification.
This story does not add layout or rendering support.

**Depends on**: F-X065.
**Test gate**: regression. Canonical and aliased positive cases, adversarial
foreign and ambiguous cases, public-literal and equality compatibility,
item-order preservation, package save-reopen, and
the current Word corpus pass with 49 of 49 output hashes unchanged.

### F-X067, Prime Word fidelity Cargo dependencies (S)

PR 58 at source SHA `c8fed1d1268fd765d602bac2da6524900c1c1cfd`
identifies that a cold hosted runner can reach the intentional locked offline
`rdocx` build before the complete Cargo graph is present. The Word fidelity job
runs exact `cargo fetch --locked` after its pinned Rust cache and before the
corpus harness. The harness remains locked and offline, so network preparation
stays explicit and render evidence cannot depend on an incidental warm cache.
Exact workflow order, cardinality, and mutation regressions harden the direct
submitted outcome. PR 58 remains open and unchanged. Contribution-hosted run
`33025657609`, Word job `98366252284`, proves the cold path and uploads both
required evidence files as one nonempty artifact.

**Depends on**: F-X064.
**Test gate**: regression. Workflow tests reject missing, unlocked, duplicated,
misplaced, or wrong-job dependency priming. The current pinned Word corpus and
PR 58 hosted run emit nonempty evidence, and all 49 output hashes remain
unchanged. The integrated hosted Word job remains a sprint-completion rider.

### F-X071, Integrate PRs 61 through 64 (L)

Adopt the current reviewed outcomes of PRs 61 through 64 as one current-tree
Word reader integration while retaining authenticated contributor credit.
The adopted contributor heads are `7c40c2e`, `fa48a39`, `60bc663`, and
`5cb5cba` from `@pedroassumpcao`, followed by separately labelled maintainer
hardening on the F-X071 worker branch.
Expose hyperlink and drawing safety facts, document and table completeness
facts, numbering and effective-formatting facts, and tracked insertion and
field facts. Harden the submitted table reader so retained XML carries every
required ancestor namespace binding and malformed row revision markers remain
in their schema slots. Harden effective numbering so the explicit or default
paragraph style identity is used consistently for style and numbering-level
resolution. Audit the bounded nested-revision and complex-field additions at
the exact adopted source SHA before integration. No reader fact may give typed
meaning to a foreign-namespace lookalike or weaken unmodelled XML preservation.

**Depends on**: none.
**Test gate**: regression. Focused namespace, schema-order, default-style,
revision-depth, and reader-projection fixtures pass after save and reopen, the
complete Word suites pass, and all 49 output hashes remain unchanged.

### F-X072, Keep paragraph caching across note references (M)

Keep paragraph-cache reads available after an otherwise safe paragraph that
contains a footnote or endnote reference. The cache key continues to include
the complete typed paragraph and revision view. Base cache-context reuse is
separate from exact footnote and endnote equality, so a changed reference ID or
note part invalidates each affected paragraph entry without poisoning later
safe paragraphs. Full restart and related-story caches still require exact note
parts. Fields, numbering, drawings, raw children, and other unsupported
paragraph-cache content remain conservative.

**Depends on**: F-X062.
**Test gate**: regression. A 700-paragraph document with one early footnote or
endnote reference records 699 paragraph-cache hits and one rebuild after a
later paragraph edit. Warm and fresh layouts are byte-for-byte equal. Changing
the reference or note part invalidates the required entry, and existing unsafe
content remains excluded.

### F-X073, Restart ordinary-prose pagination within the aggregate cache (L)

Permit restart-pagination records for ordinary multi-line prose, headings, and
keep-together paragraphs when the complete checkpoint state already represents
their effects. Continue to reject unrepresented numbering, drawings,
multilingual state, raw content, and other unsafe inputs. Field-bearing blocks
retain substitution pairs but receive no pagination checkpoints. Charge restart
records against the actual paragraph, table, header or footer, and restart
cache bytes under the existing aggregate budget rather than an independent
8 MiB ceiling. The existing entry caps and exact context fingerprints remain
fail-closed.

**Depends on**: F-202, F-X062, F-X072.
**Test gate**: regression. A 700-paragraph ordinary-prose document containing
a heading publishes a restart candidate larger than 8 MiB when the aggregate
remains below 64 MiB. Late edit, insert, delete, and undo layouts reuse bounded
work and remain byte-for-byte equal to fresh layout. A candidate above the
aggregate budget is rejected without changing output.

### F-X074, Tag rpptx-v0.9.0 (S)

The completed M21 PresentationML depth boundary is published as the exact
15-package incubating family at 0.9.0 from immutable annotated tag
`rpptx-v0.9.0` at reviewed SHA
`45b4f277ff5fd6d1b032e929c5dcee7fb9d2c550`. Every registry entry reports sole
owner `mantissaman (Atul Sharma)`, the GitHub release body is byte-identical to
the reviewed notes, and `rpptx-wasm@0.9.0` remains absent from crates.io.

The reviewed changelog covers collaboration, security, timing, media,
SmartArt, interchange, package variants, notes and handouts, animated export,
and bounded HTML and PDF import. The selected-family contribution inventory is
empty. Disposable CI proof PRs 59 and 60 have no shipped user outcome. PRs 61
through 64 and Issues 65 through 67 remain attributed only to the stable Word
family.

**Depends on**: F-213, F-214, F-215, F-216, F-217, F-218, F-219, F-220,
F-221, F-222, F-223, F-224, F-225, F-226, F-227, F-X068.
**Test gate**: release, passed. All 15 incubating registry entries, their owner,
the annotated tag target, release-body bytes, stable exclusion, and absent
`rpptx-wasm@0.9.0` verified after separately approved publication.

### F-X075, Preserve restart pagination across page-spanning paragraphs (M)

Keep the recorded pagination pass when an otherwise eligible ordinary-prose
paragraph spans a page boundary. A split continuation still creates no
checkpoint inside the paragraph. The next checkpoint may be recorded only
after the whole paragraph completes and the existing note, wrap, and resolved
state is clean. Numbering, drawings, raw XML, fields, unsafe tables,
backgrounds, multiple sections, dirty note state, and every other existing
restart exclusion remain fail-closed.

Remove the document-wide split veto that discards the completed recorded pass
and immediately paginates the whole document again. Retain the existing exact
context fingerprints, aggregate cache budget, checkpoint bounds, and
transactional publication. This is the hardened fix for Issue 67 and does not
add a mid-paragraph continuation model or a public API.

**Depends on**: F-X073.
**Test gate**: regression. A deterministic 175-paragraph source-built document
whose four-line paragraphs span 16 pages completes one recorded pass, retains
a restart record, and records no checkpoint inside a paragraph. Ten warm
middle edits produce 174 paragraph-cache hits and one rebuild, paginate only a
bounded affected page range, and equal fresh layout exactly. Late edit,
insert, delete, undo, note-bearing split, and displayed page-number footer
cases remain exact. Existing unsafe inputs still reject restart publication.
An interleaved release-mode comparison for the 175 and 700 paragraph native
and bundled-fallback paths is no worse than 1.25 times v0.11.1 and at most 0.75
times the pinned `0582da0` regression median. Each timing run first
authenticates the complete measured crate graph and exact injected harness by
content manifest. Reference runs also authenticate their pinned commit.

### F-X076, Tag v0.12.0 (S)

Publish the reviewed stable Word outcomes added after v0.11.1 as the exact
seven-package stable family at 0.12.0. Move every stable manifest, workspace
pin, lock record, README requirement, source assertion, CI literal, Python and
WASM metadata carrier, workflow preflight, and release regression in lockstep.
Keep the shared OOXML and PowerPoint family at its separately published 0.9.0
boundary. Python, WASM, npm, and PyPI remain outside publication authority.

The reviewed changelog credits Pedro Assumpcao for PRs 61 through 64 and
`@emptinessform` for Issues 65 through 67. Each outcome landed through the
maintained hardened equivalent. After successful publication and release-body
verification, post one specific thank-you comment to each record. Leave every
record's open or closed state unchanged, including open Issue 67.

**Depends on**: F-X074, F-X075.
**Test gate**: release. After one clean full verification and sprint review at
the exact prepared SHA, `/release v0.12.0` obtains separate final approval and
publishes exactly `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, and `rdocx-cli`. Every registry entry, owner, tag target,
release-body byte, selected-family exclusion, and all seven notification URLs
must verify before completion.

### F-X077, Share strict XML lexical validation (M)

The F-236 embedded scanner, the F-237 glossary parser, and the F-237
package-story scanner each carry independent XML 1.0 lexical validation for
declarations, literal characters, references, names, namespace declarations,
expanded attributes, and processing instructions. S68 sprint review pass 1,
`.claude/reviews/S68-sprint-review-pass-1.md`, finding S1, found that the three
copies make every security correction a three-site change and review burden.

Move the format-neutral checks to the lowest existing shared crate that all
three consumers already use. Keep owner-specific document roots, schema
positions, error variants, and diagnostic labels local. Do not add a new crate,
parser model, trait, generic, feature, or permissive recovery path.

**Depends on**: F-236, F-237.
**Test gate**: regression. The existing embedded, glossary, and package-story
malformed XML matrices all execute through one shared lexical validator, keep
their current error surfaces, and retain byte-identical mutation rollback.
Removing any shared declaration, character, reference, name, namespace, or
processing-instruction check makes at least one matrix fail.

### F-X079, Tag rpptx-v0.10.0 (S)

The shared strict XML lexical validator required by the stable M22 Word family
is published as the exact 15-package incubating family at 0.10.0. Every
incubating manifest, workspace pin, lock record, README requirement, source
assertion, CI literal, workflow preflight, release regression, and the
unpublished `rpptx-wasm` preparation carrier moved in lockstep. Stable source
dependency pins use the published shared 0.10.0 boundary without changing the
stable family version or granting it publication authority.

The exact `rpptx-v0.10.0` release notes come from the reviewed selected-family
diff and contribution inventory. That diff contains the shared lexical
validator, baseline-aware inline groups, and Word glossary package constants.
No external issue or pull request belongs to those selected changes, so the
reviewed contribution inventory is empty. Stable Word, Python, WASM, npm, and
PyPI packages remain outside publication authority. The immutable annotated
tag `rpptx-v0.10.0` targets reviewed SHA
`1e409c553b950eb8029e3e78e39ff775f18ba3ab`, and every registry entry reports
sole owner `mantissaman (Atul Sharma)`.

**Depends on**: F-X077, F-X080.
**Test gate**: release, passed. Exactly the 15-package incubating family is
published. Every
registry entry, owner, annotated tag target, GitHub release-body byte,
stable-family exclusion, absent `rpptx-wasm@0.10.0`, and applicable
contribution notification URL must verify before completion.

### F-X080, Restore CI release readiness (S)

Restore the hosted release gates that drifted from their reviewed inputs. The
`oxml-layout` package job must compare against every bundled Noto font and legal
file. The checksum-pinned Pandoc 3.10 installer must admit its reviewed
162,406,703-byte extracted archive under a 160 MiB ceiling. It skips without
materializing only the exact `pandoc-lua -> pandoc` and
`pandoc-server -> pandoc` aliases and rejects every other unsupported member
type. The Python adapter must compile exhaustively against the current native
`rdocx::Error` surface without exposing new MHTML or embedded-mutation methods.

**Depends on**: F-X077, F-239.
**Test gate**: regression. Mutation-sensitive workflow tests fail if any Noto
font or legal file is removed from the expected package inventory, if the
Pandoc extracted-size ceiling falls below the authenticated payload, if either
exact Pandoc alias changes, if another unsupported member type is accepted, or
if either current native error mapping is omitted. The reconstructed package,
Pandoc, and Python binding CI commands all pass before release preparation.

### F-X078, Tag v0.13.0 (S)

The immutable v0.13.0 attempt prepared the reviewed M22 Word-depth outcomes as
the exact seven-package stable family. The annotated tag targets reviewed SHA
`05332b17f481741e7d5ab4e39699c6d1536475af`. The workflow published
`rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`, and `rdocx-pdf`, then
stopped while verifying `rdocx`. F-238 added four public Word main content-type
constants after the immutable shared 0.10.0 release, so the packaged facade
could not compile against the registry `oxml-opc` API. `rdocx`, `rdocx-cli`,
and the GitHub release were not published.

The tag and five registry entries remain immutable. F-X081 publishes the
complete shared 0.11.0 family containing the F-238 constants. F-X082 publishes
the coherent seven-package stable recovery at 0.13.1. Python, WASM, npm, and
PyPI remain outside publication authority. Issue 69 is a valid post-release
paragraph-cache performance follow-up, but no Issue 69 change is part of this
partial attempt or its recovery releases.

**Depends on**: F-238, F-239, F-X077, F-X079.
**Test gate**: release. Preparation and every local gate passed at the reviewed
SHA. Publication did not complete because the registry `oxml-opc@0.10.0`
contract lacks the F-238 constants used by packaged `rdocx`. The immutable
partial result is the input to F-X081 and F-X082, not a completed stable-family
release.

### F-X081, Tag rpptx-v0.11.0 (S)

The exact 15-package shared OOXML and PowerPoint family is published at 0.11.0
from immutable annotated tag `rpptx-v0.11.0` at reviewed SHA
`0b6bd622f8a14189d7d1281d011f81319ef8ad2a`. The only selected source change
since `rpptx-v0.10.0` is the additive `oxml-opc` Word main content-type
vocabulary required by F-238. All publishable incubating manifests, workspace
pins, lock records, README requirements, source assertions, CI literals,
release regressions, and the unpublished `rpptx-wasm` preparation carrier move
in lockstep. Stable carriers remain at 0.13.0 for the separate recovery.

Preserve the immutable partial v0.13.0 tag and registry entries. Do not rerun
its workflow or publish a partial shared patch. Prepare exact selected-family
notes and an evidence-derived contribution inventory. Stable Word, Python,
WASM, npm, and PyPI packages remain outside publication authority.

**Depends on**: F-238, F-X079, F-X080.
**Test gate**: release, passed. All 15 shared and PowerPoint registry entries
resolve at 0.11.0 under the authenticated sole owner. The annotated tag,
byte-identical release body, stable-family exclusion, absent
`rpptx-wasm@0.11.0`, and empty contribution inventory verify. F-X082 is
eligible to start.

### F-X082, Tag v0.13.1 (S)

The exact seven-package stable family is published at 0.13.1 from immutable
annotated tag `v0.13.1` at reviewed SHA
`c391d12422c288be5db314bad8338dd08bb47d9a`. Every stable version carrier,
internal pin, lock record, README requirement, source assertion, CI literal,
Python and WASM metadata carrier, workflow preflight, and release regression
moved in lockstep. Every shared dependency is pinned to published 0.11.0.

The release notes describe the complete M22 outcome and the immutable partial
v0.13.0 attempt accurately. The registry-only proof packages normalized local
stable crates, then compiles `rdocx` while all shared crates, including
`oxml-opc`, resolve from the published 0.11.0 family. The local patched
workspace graph cannot mask this dependency boundary again. Python, WASM, npm,
and PyPI remain outside publication authority.

**Depends on**: F-238, F-239, F-X077, F-X081.
**Test gate**: release, passed. All seven stable registry entries resolve at 0.13.1
against shared 0.11.0 under the authenticated owner. The annotated tag targets
the reviewed SHA, the GitHub release body is byte-identical to reviewed notes,
selected-family exclusions hold, and every applicable contribution
notification verifies. The contribution inventory is empty, so no notification
was required.

### F-X083, Close confirmed Issue 67 and intake Issue 69 (S)

Record `@emptinessform`'s v0.12.0 confirmation that F-X075 fixed Issue 67,
close that issue without reopening completed implementation, and preserve the
three independent Issue 69 mechanisms, measurements, proposed commits, and
authorship as inputs to F-X084 through F-X086. The intake compares each offered
fork commit with current main rather than treating external code as trusted.

**Depends on**: F-X075, F-X076.
**Test gate**: regression. The existing page-spanning restart regression passes,
the Issue 67 closure cites the reporter's confirmation, and each live Issue 69
mechanism maps to exactly one pending story with authenticated credit.

### F-X084, Narrow note-part paragraph cache invalidation (M)

Changing a footnote or endnote part must keep paragraph-cache reads enabled for
ordinary body paragraphs and conservatively rebuild only entries whose
paragraph carries a note reference. Restart, header, footer, and note-page
caches keep their full-context safety gates. Review the independently offered
`@emptinessform` commit `4777a741` against current main and retain contributor
credit for any adopted implementation.

**Depends on**: F-X072, F-X083.
**Test gate**: regression. Warm and fresh layout stay element-for-element equal
after note text, insertion, and deletion changes, at least 698 of 700 unaffected
paragraphs hit the cache, and deterministic output hashes remain unchanged.

### F-X085, Memoize restart body identities once per layout (M)

One lazy tri-state candidate-identity memo supplies unchanged-body,
first-change, common-suffix, and restart-record publication work so each body
block serializes its exact identity at most once per layout. Fingerprint misses
remain serialization-free, exact identity bytes remain authoritative after a
fingerprint match, and publication moves computed bytes into retained entries.
Test-only counters measure the bounded transient slot and byte-capacity peaks.
Review `@emptinessform` commit `eff0ea0c` independently and retain contributor
credit for any adopted implementation.

**Depends on**: F-X075, F-X083.
**Test gate**: regression. A 715-block edit computes at most 715 candidate
identities, warm and fresh layouts remain equal, the cache stays bounded, and
the deterministic corpus hashes do not move.

### F-X086, Provenance-safe restart after body-length changes (L)

Allow sourced insert, delete, Enter, merge, and selection-delete edits to reuse
the unchanged prefix restart checkpoint even when body block count changes.
The complete restart-record context gate remains exact, while whole-body
unchanged detection remains length aware. Tail reuse requires absent provenance
or equal body length, so shifted source indices cannot leave retained
`SourceSpan` paths stale. The implementation adopts the complementary mechanism
from independently offered `@emptinessform` commits `9e48bc86` and `c8315b92`
as the enabling and provenance-safety inputs respectively, with authenticated
contributor credit retained.

**Depends on**: F-X075, F-X083.
**Test gate**: regression. A sourced insert and delete near block 640 of 700
recomputes at most three pages, matches a fresh layout element for element with
equal source paths, keeps only contiguous prefix page identity, and fails if the
full-pagination veto is restored or stale tail provenance is reused. Enter,
adjacent merge, and multi-block selection deletion carry the same bounded and
exact sourced contract.

### F-X087, Portable authored Word charts from PR 71 (L)

Integrate Kevin Brown's PR 71 contribution as a hardened S71 scope exception.
Authored Word charts remain editable across Microsoft Word and Apple Pages
without image flattening. Typed axis titles, explicit visible axes, value-axis
number formats, generic RGB palettes, category-point colours, pie and doughnut
legends, percentage labels, explicit doughnut holes, and optional numeric-cache
format omission serialize in schema order. A chart gains a
relationship-owned Office theme only when no valid theme exists. A related
theme is valid only when its target exists, has the theme content type, and
parses as a DrawingML theme.

The Word mutation stages the complete `Document`, typed theme state, package,
and deterministic identifier owner before one commit. `RgbColor` is re-exported
through `oxml-chart`, `rdocx`, and `rpptx`. Preserve Kevin Brown's contribution
credit and do not import the PR's stale S65 ledgers or its now-occupied F-X077
identity.

**Depends on**: F-158, F-245, F-249.
**Test gate**: differential. At the reviewed S71 implementation SHA,
source-built line, bar, pie, and doughnut documents save and reopen with exact
chart and editable-workbook semantics. The candidate has SHA-256
`54faeec0d56767577afa014564d56571c46d00df11c73baaa38889999a39b3f9`.
Microsoft Word 16.112.3 build
16.112.26083020 opens the exact candidate without repair. Pages Creator Studio
15.1.1 build 7044.0.273 renders the authored axes, colours, legend, percentages,
and doughnut shape, then exports a DOCX whose chart and workbook data remain
exact across all four editable workbooks. Each run records the export digest,
while the gate compares parsed semantics because Pages recalculates manual
chart layout coordinates and drawing extents between exports.
Malformed correctly typed themes fail or are replaced on a staged candidate
without changing retained source bytes. The hash harness remains unchanged at
49 of 49.

### F-X088, Verify and close Issue 69 after S70 fixes (S)

Verify the live Issue 69 report against the three completed S70 fixes and close
the issue only after their combined evidence passes on the reviewed S71 SHA.
F-X084 owns note-reference-aware paragraph cache invalidation. F-X085 owns the
once-per-layout restart identity memo. F-X086 owns provenance-safe restart for
insert, delete, Enter, merge, and selection-delete body-length changes. Preserve
`@emptinessform`'s report and offered commits
`4777a74167495a5116289e1f905dfd9ad4dbe807`,
`eff0ea0c28b5eaf08180b09b58e0c0f486b7433b`,
`9e48bc86876c294b8daa314e577e84b6fcd7ac97`, and
`c8315b92857c951146fc866cd044b214194a09a8` in the next stable release
contribution inventory.

The closure evidence credits the reporter, links the six focused regressions,
states that v0.13.1 remains affected, and says the fixes will ship in the next
stable release without promising a date. The reporter fork contains no
committed timing harness. The qualified temporary reconstruction uses exact
v0.13.1 and S71 source archives, release mode, deterministic bundled fonts, the
reported 700 four-line paragraph and 3 by 3 table workload, a 63-page prime,
three positions, warmup, and alternating measured rounds. It reports min and
median times plus paragraph-cache and page-layout work, and distinguishes direct
engine model mutation on macOS from the reporter's Windows editor environment.

**Depends on**: F-X084, F-X085, F-X086.
**Test gate**: regression, passed at reviewed S71 SHA
`667416b1b54968b1524d57232c44f73a175fd27a`. All six focused
deterministic-font regressions passed, the complete workspace gate passed, and
the 49-entry hash harness remained unchanged. F-X084 through F-X086 are present
on S71 and absent from v0.13.1. In the 21-sample-per-operation timing
comparison, all S71 medians are at most 22.269 milliseconds. Footnote insertion
falls from a 48.157 millisecond v0.13.1 median to 14.688 milliseconds on S71,
and footnote deletion falls from 52.845 to 22.269 milliseconds. Both note edits
change from zero paragraph-cache hits and 700 builds to 699 hits and one build.
The authenticated evidence credits `@emptinessform`, preserves all four offered
commits, records the environment caveats, and states that the fixes will be in
the next stable release without a date. Issue 69 is closed as completed at
<https://github.com/tensorbee/rdocx/issues/69#issuecomment-5592205748>.

### F-X089, Capability-led README family (L)

The root README is the product front page and every crate-local README uses the
same outcome-first treatment. The root page leads with the complete native
document workflow, a seven-row implemented-outcome summary, working examples,
and a dated, evidence-backed comparison with relevant Rust and cross-language
alternatives. Exact property boundaries remain available through the canonical
capability matrix without leading with internal delivery classifications or
sprint state.

Each of the 26 crate-local documents explains the result its consumer can
achieve, at least three implemented capabilities, when to choose it, its place
in the workspace, and one checked example in the consumer's language or command
surface. Deprecated shims and unpublished support crates remain labelled
accurately, but status warnings do not replace the value proposition. Claims
come from current public APIs, package metadata, compiled examples, and
reviewed official comparison sources. Volatile popularity, price, size,
memory, and speed claims remain excluded unless a dated reproducible
measurement is checked into the repository.

**Depends on**: F-242, F-X009.
**Test gate**: regression. `python3 scripts/readme_doctests.py` compiles all 23
Rust examples across the 21 Rust-library READMEs, validates the CLI, Python, and
JavaScript snippets, checks versions and local links for all 27 package
documents, verifies the capability-led section contract, resolves the approved
official comparison sources in focused network mode, rejects extra rows,
duplicate evidence, or a broadened uniqueness conclusion, binds security claims
to default-off Cargo features, and proves every one of the 22 publishable
archives contains its byte-identical declared README.

### F-X090, Accept part-local producer drawing identities (S)

Producer documents open when the same normalized `wp:docPr/@id` appears in
different physical XML parts. Imported drawing identities are validated within
each part, then their complete union remains occupied input to the
package-global authored allocator. Producer XML remains unchanged, and
normalized duplicates inside one part remain invalid.

**Depends on**: F-255.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/72>.
**Test gate**: regression. `cross_part_producer_drawing_ids_do_not_block_document_open`
opens a source-built package whose body and header reuse one valid drawing id,
preserves both parts through mutation and reopen, allocates a new authored id
outside the occupied union, and still rejects same-part normalized duplicates.

### F-X091, Serialize unused root default namespaces safely (M)

Classify a producer root default namespace by namespace-aware use rather than
rejecting every unknown default binding. An unused declaration does not block a
typed mutation or save. A used, ambiguous, or malformed binding still fails
closed and leaves the package unchanged. Add fallible
`Document::try_replace_text` for the CLI while retaining the legacy infallible
signature. Nested declarations shadow the root by lexical scope, unprefixed
attributes do not consume the default, and successful serialization refreshes
the cached namespace facts from the published main-story bytes.

**Depends on**: F-255.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/73>.
**Test gate**: regression. `unused_root_default_namespace_allows_atomic_save`
mutates and reopens the reported package shape while retaining its prefixed raw
producer element. Scope regressions cover same-URI and different-URI shadows,
undeclaration, unprefixed attributes, malformed input, and duplicate defaults.
A used inherited default namespace still fails atomically through native and
CLI paths without a panic or partial output.

### F-X092, Preserve logical reading order in generated PDFs (L)

Make the shared PDF backend expose complete logical lines instead of
run-fragmented extraction for large Word documents and PowerPoint
presentations. Emit rich runs with run-wide extraction geometry and preserve
logical source order across adjacent styled and bidirectional runs without
changing glyph paint order, pagination, or raster geometry.
The PDF-local line planner coalesces only adjacent runs with matching semantic
ownership and baseline plus contiguous source spans or logical indices. One
initial text matrix and relative glyph placement preserve exact paint, while
the first run carries the complete logical line and later runs suppress
duplicate extraction. Nested group transforms retain the same final page
coordinates.

**Depends on**: F-255.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/74>.
**Test gate**: regression. `large_word_and_presentation_pdfs_preserve_logical_reading_order`
builds large DOCX and PPTX inputs with uniquely numbered multiword lines split
across runs. Pinned Poppler 26.01.0 extracts each substantial line once and in
logical order while deterministic raster output remains unchanged.

### F-X093, Preserve drawings through document comparison staging (M)

Keep package-authoritative main-story XML and namespace ownership through
comparison, tracked-body construction, staged reopen, and accept or reject
postconditions. Valid body and related-story drawings remain complete while
text revisions are emitted. Invalid drawings still fail atomically.

**Depends on**: F-255.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/75>.
**Test gate**: regression. `document_compare_preserves_inline_drawings_through_staging`
compares source-built packages containing body and header drawings whose text
sibling changes. It preserves exact inline and anchored wrappers, namespace
bindings, extended `docPr` children, relationships, and media through compare,
save, reopen, accept, and reject at run, word, and character granularity. A
stable multi-unit run appears once, while a genuinely missing `wp:docPr/@id`
still rejects atomically.

### F-X094a, Expose Word collaboration and redline commands in rdocx-cli (M)

Expose shipped comment, revision-resolution, and comparison facades through
nested `rdocx comment`, `rdocx revision`, `rdocx compare`, and
`rdocx toc rebuild` commands. Comment additions use zero-based half-open body
paragraph and run coordinates. Revision filters select at most one id, exact
author, or paired inclusive RFC 3339 date range. Every mutation requires an
explicit output, publishes atomically, and returns exact schema-1 JSON when
requested. Revision inspection remains explicitly scoped to the main story,
while resolution and comparison declare their all-supported-story scope.

**Depends on**: F-148, F-150, F-234, F-235.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: integration. `cli_collaboration_commands_are_schema_stable_and_atomic`
lists and mutates comment threads, lists and resolves revisions, creates a
reopenable redline, verifies exact JSON, and proves invalid inputs publish no
destination.

### F-X094b, Structured CLI text and layout plus guarded replacement (L)

Add schema-1 JSON for rich accepted-view body text and deterministic body-item
layout. Text records carry a top-level body index and typed nested path. Layout
uses real point-space fragments for page-spanning items, empty tables, and
image-bearing paragraphs. `replace --expect N` guards the existing run-aware
replacement before publication.

The native sidecar keeps every direct body item addressable. Its public fragment
record exposes one-based physical and displayed pages plus point-space bounds,
while preserved unlaid items have an empty fragment list. Cached pagination
retains the same fragment result as fresh pagination without changing shared
positioned output.

**Depends on**: F-X032, F-X037, F-X047.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: integration. `cli_structured_text_layout_and_guarded_replace_preserve_exact_contracts`
checks nested text paths, run formatting, multi-page body fragments, empty and
image extents, and an exact-count mismatch that creates no output.

### F-X094c, Priority rdocx Python collaboration, comparison, layout, and TOC (L)

The Python `Document` exposes current comparison, main-body comments,
deterministic layout and page lookup, and TOC rebuild through precisely typed
immutable snapshots. Structural mutations advance binding revision once only
after success. Comparison, layout, TOC rebuild, and serialization release the
GIL. The surface does not complete the future all-story work in F-291, F-293,
F-295, or F-310.

**Depends on**: F-X094b, F-148, F-232, F-234, F-235.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: binding. `priority_word_operations_return_typed_snapshots_and_remain_atomic`
reopens a redline and comment thread, checks real layout fragments and TOC
counts, passes installed typing and stub checks, and proves native failures do
not mutate the Python document.

### F-X094d, rdocx Python sections, styles, rich stories, and hyperlinks (L)

After F-252, expose ordered section, style, rich header and footer, and
relationship-resolved hyperlink snapshots. Records retain source order, story
ownership, inheritance, and nested paths without adding a second document tree
or an untyped dictionary layer. Frozen Python records expose sections in EMU,
styles, physical stories, story items, all effective header and footer
variants, and hyperlinks. Native `StoryItemRef::links` returns existing
`LinkInfo` records and resolves each relationship through the checked story
owner. Native `Document::story_links` pairs the existing location and link
types in physical source order across nested and ancestor-owned items.

**Depends on**: F-252, F-255, F-X094c.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: binding. `word_structure_snapshots_preserve_order_ownership_and_types`
opens a three-section document with inherited and independent rich variants,
styles, and links, then verifies exact frozen records and installed typing
through save and reopen.

### F-X094e, rpptx Python rendering, comments, and notes (L)

Expose deterministic slide and speaker-note PDF and PNG rendering, notes text,
and the current modern comment-author, thread, reply, and ordered mutation
facade through precisely typed Python values. Rendering releases the GIL and
collaboration mutation retains native identity and atomicity. The native
facade adds one-slide and all-slide deterministic PNG conveniences over the
same resolved layout path. Python returns frozen `CommentAuthor`, `Comment`,
and `CommentReply` values, accepts native GUID and RFC 3339 strings, and
advances its global revision only after a successful collaboration operation.

**Depends on**: F-136, F-217, F-226.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: binding. `presentation_render_comments_and_notes_match_native_snapshots`
renders slide and notes outputs, reads notes text, mutates and reopens one
comment thread, passes installed typing, and proves invalid identities publish
no mutation.

### F-X094f, Prepare the version-aligned Python release paths (M)

Extend the reviewed release ceremony with independent Python distributions.
`py-rdocx-v0.13.2` selects `rdocx` at the matching stable source version, while
`py-rpptx-v0.11.0` selects `rpptx` at the matching incubating crate version.
Each release contains exactly six cp39-abi3 wheels and one source distribution,
uses trusted PyPI publication, publishes byte-identical GitHub release notes,
verifies registry ownership, and retains contribution evidence. Each project
uses its crate-local README as the Markdown long description and supplies a
specific summary, author, keywords, classifiers, and project URLs. The
artifact gate verifies that metadata and required installation, quick-start,
typing, and project-link guidance in both archive formats. The immutable PyPI
`rdocx 0.13.1` release remains available, so the corrected metadata ships as
0.13.2. A manual
`wheels.yml` run at the reviewed SHA is build-only and supplies the exact
artifacts for clean Python 3.9 and 3.12 install and runtime checks. Strict
typing and stub checks run under Python 3.12. After separate exact-tag
approvals, S72 published and verified `rdocx 0.13.2` and `rpptx 0.11.0` from
reviewed SHA `2b009243ed39ab66470d7484d490985368e865a8`. Both GitHub release
bodies match their reviewed notes byte for byte. The completed release gate
notified and closed Issues 72 through 76, contributor PRs 77 through 80, and
the unmerged verification PR 82.

**Depends on**: F-137, F-138, F-X094a, F-X094c, F-X094d, F-X094e, F-X096.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: release preparation.
`python_release_contract_rejects_partial_or_unapproved_publication` accepts
only the exact selected tag, distribution, six wheels, one source archive,
matching native crate version, trusted publisher, reviewed-note, verification,
and approval contract. Negative mutations reject every mixed, mismatched,
partial, or manual-publication path.
`python_release_contract_requires_complete_project_metadata` requires the
reviewed long description and project metadata in each source project and
built archive.

### F-X095, Integrate PRs 77 through 80 and restore deterministic CI (L)

Integrate the contributor outcomes from PRs 77 through 80 at pinned heads
`aed8f14d826e43fee52b6da75c47fe2c5d37645a`,
`4b8fd0d15b3920abf3e40f46f7752ade23589bf5`,
`8ede12b4102fb8bdc9421c22e059b7df010e2113`, and
`9a9d3e8eeab2a5b8f2088930beae50ebce918f23`. Preserve contributor credit while
reconciling the changes against the completed S72 section, story, binding, and
release work. Numbering reports only retained extra XML and attributes as
unmodelled. Document, self-closing body, revision, marker, field, table-cell,
and TOC reader facts remain namespace-aware, bounded, schema-ordered, and
stable through save, reopen, and repeated save. Pedro Assumpcao is credited for
the four pinned contributor heads in the sprint delivery record and release
notes.

Restore hosted Presentation fidelity by installing the exact reviewed
LibreOffice 26.2.5.2 and Poppler 26.01.0 builds on Ubuntu 24.04. Package-manager
LibreOffice is not an accepted substitute. Retain the existing S72 Python
binding repair for documents that carry an authored font table and prove both
reported hosted failures are absent on the integrated branch.

**Depends on**: F-X071, F-253.
**GitHub pull requests**: <https://github.com/tensorbee/rdocx/pull/77>,
<https://github.com/tensorbee/rdocx/pull/78>,
<https://github.com/tensorbee/rdocx/pull/79>, and
<https://github.com/tensorbee/rdocx/pull/80>.
**Test gate**: regression. The focused contributor reader tests cover narrowed
numbering evidence, strict body boundaries, missing revision authors, marker
child content, complex-field properties, cell margins, empty cells, nested
tables, and self-closing TOC coordinates. The CI contract test rejects a moving
LibreOffice installation or a non-Ubuntu Presentation fidelity runner. The
complete `rdocx-oxml`, `rdocx`, Python binding, workflow regression, and hash
harness gates must pass. Exactly seven `word/document.xml` hashes may change
for self-closing empty paragraphs, while every PNG and PDF fingerprint remains
unchanged.

### F-X096, Align Python distribution versions and release tags (M)

Publish each Python distribution at the version of its corresponding native
crate. The corrective stable source, native `rdocx`, `rdocx-py`, and the next
PyPI `rdocx` release use 0.13.2. The immutable PyPI `rdocx 0.13.1` release and
the complete crates.io 0.13.1 family remain available.
`rpptx-py` and PyPI `rpptx` use 0.11.0 with native `rpptx`, rather than
inheriting the unrelated stable workspace version 0.13.2. Package metadata,
wheel and source archive names, installed module versions, PyPI records, and
GitHub releases must all agree.

Use `py-rdocx-vX.Y.Z` and `py-rpptx-vX.Y.Z` as the Python tag namespaces. The
existing `v*` stable Rust and `rpptx-v*` incubating Rust tag families remain
unchanged, so neither Python tag can start crates.io publication. A tag build
may build the complete matrix, but its trusted publish job selects exactly the
one matching distribution. Manual dispatch remains build-only.

**Depends on**: F-X094e.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/76>.
**Test gate**: release preparation.
`python_release_contract_keeps_distribution_versions_independent` proves exact
crate and project version agreement, disjoint tag routing, six wheels and one
source archive per selected distribution, manual build-only behavior, and
fail-closed rejection of a version or family mismatch.

### F-X097, Preserve namespace-scoped drawings and complex fields in comparison (M)

Document comparison retains namespace bindings inherited by inline and anchored
drawing wrappers, including declarations owned by a story-part root or the
outer `w:drawing`. Run ownership maps physical complex-field XML onto the
modeled runs used by comparison, so a text edit next to a field does not fail
because the source contains more physical `w:r` elements than modeled runs.
Malformed drawings still fail atomically.

**Depends on**: F-X093, F-234.
**GitHub issues**: <https://github.com/tensorbee/rdocx/issues/75> and
<https://github.com/tensorbee/rdocx/issues/85>.
**GitHub pull request**: <https://github.com/tensorbee/rdocx/pull/106>.
**Test gate**: regression.
`comparison_preserves_inherited_drawing_namespaces_and_complex_fields` compares
body and header drawings with declarations at each reported ancestor scope,
then compares a changed paragraph containing a complex field. Save, reopen,
accept, and reject retain relationships, drawing identity, field structure,
and changed text without a count mismatch or partial mutation.

### F-X098, Preserve content-control type payloads (M)

Retain the selected content-control type element's attributes, ordered children,
and local namespace bindings while exposing the existing typed discriminator.
An unchanged type writes its preserved payload under the fixed output prefix.
Changing the public discriminator replaces that payload with the canonical
empty element for the selected type, without disturbing other `w:sdtPr`
children.

**Depends on**: F-253.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/84>.
**Test gate**: round-trip.
`content_control_type_payload_round_trips_until_type_changes` preserves reported
checkbox, date, combo-box, repeating-section, citation, and equation attributes
and child payloads, then proves an explicit type mutation emits only the
canonical replacement in schema order.

### F-X099, Expose direct body ownership for story items (M)

Keep the existing recursive `StoryItem.index_path` contract and add an optional
direct main-body owner index for items that can safely seed `RunPosition` and
other direct-body APIs. Native and Python snapshots expose the same value.
Nested items identify their containing direct body child rather than their flat
scan ordinal, and items outside the main story report no body owner.

**Depends on**: F-253, F-X094d.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/86>.
**Test gate**: binding.
`story_items_expose_safe_direct_body_owners` verifies direct paragraphs and
tables, nested fields, drawings, and content controls map to the correct body
child through native and installed Python surfaces. Header and footer items
remain unanchored, and the legacy recursive path remains unchanged.

### F-X100, Preserve explicit false table toggles (S)

Parse `w:tblHeader`, `w:cantSplit`, and `w:noWrap` through the shared OOXML
on-off vocabulary rather than element presence. Preserve explicit false values
as false and serialize them canonically, while retaining the existing meaning
of absent and bare true elements. F-258 builds its authoring setters on this
lossless low-level representation.

**Depends on**: F-253.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/83>.
**GitHub pull request**: <https://github.com/tensorbee/rdocx/pull/101>.
**Test gate**: round-trip.
`explicit_false_table_toggles_remain_false` opens every accepted false lexical
form for row header, split policy, and cell wrapping, then saves and reopens
without converting any value to true.

### F-X101, Honor run-level page breaks during pagination (M)

Carry forced page-break identity from ordered run content through line breaking
into Word pagination. A page break in its own paragraph or between two pieces
of text ends the current physical page immediately. Line breaks keep their
current line-only behavior, and column breaks remain separately classified for
the existing single-column limitation rather than silently becoming page
breaks. `LayoutLine::forced_break_after` carries the non-exhaustive
`ForcedBreakKind`, and overflow pagination chooses the earlier of its ordinary
fit boundary and the next page break.

**Depends on**: F-260.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/88>.
**Test gate**: differential.
`run_level_page_breaks_match_word_pagination` constructs both reported page
break shapes, compares their page count and page text with the pinned Word
layout oracle, and proves PDF, PNG, page fields, and TOC page targets use the
same resulting pagination.

### F-X102, Resolve header and footer pictures in their story scope (M)

Lay out inline and anchored header and footer drawings against the relationship
scope of the physical story part that owns them. Body, header, and footer
relationship identifiers may collide without sharing media. Missing or
external story relationships produce one stable diagnostic and never fall
back to a document-level image with the same identifier.

The layout input keys related images by selecting main-part relationship and
story-local relationship. Each header or footer lowers its paragraphs through
a scoped `MediaRegistry` view that shares the immutable media payload map.
Footer loading follows the same internal-target rules as header loading, and
ordinary drawing diagnostics are deduplicated by scoped relationship.

**Depends on**: F-255.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/89>.
**Test gate**: differential.
`header_and_footer_pictures_render_from_story_relationships` builds body,
header, and footer pictures with colliding local relationship identifiers,
then compares deterministic PNG and PDF output with the pinned LibreOffice
render and verifies the expected image objects and alternate text.
The deterministic sample baseline changes only for the feature-showcase PDF.
The existing 400 by 40 header logo is now present on page 11.

### F-X103, Accept standard TOC switches and report rebuild diagnostics (M)

Accept Word's argument-free `TOC \\z` switch as a retained display policy that
does not block rebuilding in paginated output. Keep the switch in the field
instruction. Return the ordered diagnostic messages produced by a TOC rebuild
through native and Python reports, with `diagnostic_count` derived from the
same collection rather than maintained separately.

**Depends on**: F-X098.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/90>.
**GitHub pull request**: <https://github.com/tensorbee/rdocx/pull/101>.
**Test gate**: binding.
`word_default_toc_switch_rebuilds_and_reports_ordered_diagnostics` rebuilds the
reported Word-default instruction, preserves its content-control payload,
exposes no diagnostic, and returns exact ordered messages for malformed and
unsupported controls through Rust, Python, typing, and save-reopen checks.
The native report owns `Vec<String>` diagnostics and derives its compatibility
count through an accessor. Python exposes the same messages as a tuple and
derives its count from that tuple-backed snapshot.

### F-X104, Render DrawingML picture transparency (M)

Model the bounded `a:alphaModFix` child of an embedded picture blip as a
validated zero-to-100000 amount while preserving unsupported sibling effects
and duplicate or foreign lookalikes in order. Namespace resolution accepts the
conventional prefix and aliases declared through the enclosing picture fill.
Propagate its effective opacity through slide, layout, master,
background, cached preview, SVG, PDF, and raster rendering without changing
opaque pictures or animation opacity.

**Depends on**: F-217.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/91>.
**Test gate**: differential.
`picture_alpha_mod_fix_matches_presentation_renderers` compares slide and
layout pictures at 30 percent opacity plus an opaque control against the pinned
LibreOffice raster, then checks round-trip XML, SVG group opacity, PDF alpha,
and repeated deterministic PNG output.

### F-X105, Separate slide-owned placeholders from master header flags (M)

Treat an occupied date, footer, or slide-number placeholder physically present
on a slide as slide content. Master and layout `p:hf` flags govern only latent
placeholders inherited from that same source part. An inherited layout or
master placeholder still requires its own enabling container and is suppressed
by an occupied slide placeholder of the same latent type. The pinned secondary
oracle records the decision while native PowerPoint confirmation remains
optional.

**Depends on**: F-226.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/92>.
**Test gate**: differential.
`slide_owned_latent_placeholders_ignore_master_header_flags` covers absent,
all-false, and slide-number-only master `p:hf` containers. Slide-owned numbers
remain, unoccupied inherited dates remain hidden, and deterministic Rust text
and raster visibility match LibreOffice 26.2.5.2 through Poppler 26.01.0.

### F-X106a, Expose indexed content mutation and counted replacement in Python (L)

Expose native direct-body paragraph insertion, content cloning and movement,
content-index lookup, counted run-aware literal replacement, and counted regex
replacement through the existing Python `Document` and content handles.
Inputs use existing body coordinates and live Paragraph or Table handles.
Popping returns an opaque reusable `ContentFragment` whose typed kind is its
only exposed content. Each successful structural mutation advances the binding
revision exactly once. Counted replacements advance it only for a nonzero
count, while a failed preflight leaves the document and all handles unchanged.

**Depends on**: F-254, F-X099.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/94>.
**GitHub pull requests**: <https://github.com/tensorbee/rdocx/pull/109>,
<https://github.com/tensorbee/rdocx/pull/111>.
**Test gate**: binding.
`python_indexed_content_mutation_is_counted_and_atomic` inserts, clones, moves,
locates, and replaces content split across runs, then verifies exact counts,
stale-handle behavior, typing, save-reopen output, and rollback on invalid
coordinates or replacement syntax.

### F-X106b, Expose paragraph and run formatting mutations in Python (M)

Expose paragraph style and numbering readers and setters, run `style_id`
readers and setters, and highlight read and write access through the existing
Python paragraph, run, and font handles. Highlight uses Word's named
`ST_HighlightColor` vocabulary in both directions, while run shading remains a
separate six-digit hexadecimal or `auto` property. Values use the established
typed enums and nullable property conventions. Setters preserve unrelated
ordered run content and do not advance the binding revision because formatting
mutation does not change structural identity.

**Depends on**: F-X106a.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/94>.
**GitHub pull request**: <https://github.com/tensorbee/rdocx/pull/108>.
**Test gate**: binding.
`python_paragraph_and_run_formatting_matches_native_facades` applies styles,
numbering, named highlights, and shading to mixed-content runs, checks
native-equivalent values after reopen, rejects invalid names without mutation,
and passes installed strict typing and stub checks.

### F-X106c, Expose story mutation, hyperlinks, revisions, fields, and XML in Python (L)

Expose story text replacement, default header and footer text, story-scoped
hyperlink creation, revision snapshots and the current accept and reject
filters, field-cache updates, and read-only story-item XML through typed Python
methods and frozen values. Reuse native staged operations and StoryId ownership.
Do not expose a mutable raw-XML injection path or create a second document
model. Cloning removes duplicate comment anchors, story replacement removes
only hyperlinks made empty by that replacement, StoryItem snapshots carry and
reject stale binding revisions, and content lookup returns direct paragraph
matches before enclosing controls while exposing every ambiguous coordinate.
The compatibility StoryItem constructor maps omitted or `None` XML to empty
immutable bytes.

**Depends on**: F-X103, F-X106b.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/94>.
**Test gate**: binding.
`python_story_revision_field_and_xml_operations_are_typed_and_atomic` mutates
body and related stories, creates a scoped link, resolves revisions through
every requested filter, updates field caches, reads exact item XML, and proves
runtime, typing, GIL, revision, clone-anchor, direct-match lookup, and
failure-atomicity contracts.

### F-X107, Clone and remove existing table rows (M)

Add native and Python table operations that clone an existing row before or
after a requested position and remove a row by index. Cloning retains row
properties, cell properties, nested content, and preserved XML while assigning
fresh identities where the package-wide identifier contract requires them.
The native document methods return the inserted boundary and removal result.
Python Table methods accept negative source indexes, insert after the source by
default, and stale structural handles once after success. Row header and split
setters accept true, false, and removal using F-X100's lossless toggle
representation. Root namespace attribute names are converted to the prefix
vocabulary expected by content fragments, so an unused default namespace never
becomes an illegal `xmlns:xmlns` declaration. Removing a vertical-merge restart
promotes the matching continuation below. Every invalid index, topology, XML,
or reopen result fails without mutation.

**Depends on**: F-258, F-X106a.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/95>.
**Test gate**: binding.
`table_rows_clone_remove_and_clear_through_native_and_python` clones a formatted
nested row, clears both row toggles, removes the source row, and verifies
schema order, identities, rendering, a Word-style root default namespace,
installed typing, and save-reopen output.

### F-X108, Replace an existing picture atomically (M)

`Document::replace_image` and `replace_image_for_story` replace image bytes
through an existing body or story-local relationship while preserving the
relationship identifier and drawing markup. A format change
allocates a correctly named media part, updates the selected relationship and
content types, and removes the old part only when it is no longer referenced.
Shared targets use copy-on-write behavior so replacing one occurrence does not
silently change another. Python exposes matching operations over `Story`
snapshots. Native and Python operations publish only a fully validated package
candidate.

**Depends on**: F-255, F-X106c.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/96>.
**Test gate**: binding.
`replace_image_preserves_drawings_and_story_relationship_ownership` replaces
shared and unshared body, header, and footer images across PNG and JPEG,
verifies copy-on-write targets and content types, and proves Rust and Python
failure atomicity through save and reopen.

### F-X109, Split text runs at Unicode character offsets (M)

Add a checked native and Python `split_run` primitive for direct-body
paragraphs. The offset is a Unicode scalar-value index in the run's visible
literal text, and zero or end offsets are defined no-op boundaries. A split
copies run properties, treats tabs, breaks, fields, drawings, references,
symbols, and raw children as zero-width, keeps them in original order, updates
hyperlink and marker spans, and returns the selected or new run boundary so
existing half-open `RunRange` comment APIs can anchor exact words without
changing the stable `RunPosition` struct.

**Depends on**: F-260, F-X106a.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/97>.
**Test gate**: binding.
`split_run_enables_exact_comment_ranges_without_losing_content` splits ASCII,
multibyte, hyperlink, field-adjacent, and mixed-content runs, anchors a comment
to the selected words, and verifies accepted text, formatting, marker ranges,
typing, rollback, and save-reopen order.

### F-X110, Control field updates on document open (S)

Model `w:updateFields` in settings with the shared Word on-off vocabulary and
preserve its schema slot. Native and Python accessors return `None` when the
element is absent or unmodelled and accept `None`, true, or false to remove or
set one modeled occurrence through the existing staged settings mutation
boundary. Duplicate and malformed forms remain byte-preserved and reject
ambiguous mutation. No field update operation changes this policy implicitly.

**Depends on**: F-244, F-X100.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/98>.
**Test gate**: binding.
`update_fields_on_open_is_typed_optional_and_schema_ordered` covers absent,
bare true, explicit false, namespace aliases, duplicate preservation,
set-clear-reopen behavior, Python typing, and rollback on relationship
allocation failure.

### F-X111, Attach portable CLI binaries to Rust releases (L)

Build `rdocx` for stable tags and `rpptx` for incubating tags on the six
reviewed Linux, macOS, and Windows targets. Package one executable, its crate
README, and the workspace licence per target-specific archive. The aggregate
job rejects missing, extra, empty, non-executable, or byte-mismatched members
and produces the complete SHA-256 checksum manifest. Asset validation completes
before crates.io publication can start, and the GitHub release waits for both.
Exact `cargo-binstall` metadata resolves the same selected-family archives.
The static musl binary disables system font discovery while retaining bundled
fonts. Python releases remain unchanged and carry no CLI assets.

**Depends on**: F-X095.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/100>.
**Test gate**: release preparation.
`rust_release_assets_are_complete_family_scoped_and_installable` mutation-tests
tag routing, runner and target coverage, archive names and contents, checksums,
failure propagation, release upload ordering, and `cargo binstall` resolution
for both CLI crates.

### F-X112, Publish the complete S73 package families (L)

Publish the next reviewed minor version for the stable rdocx family and its
matching PyPI `rdocx` distribution, plus the next reviewed minor version for
the incubating rpptx and shared family and its matching PyPI `rpptx`
distribution. Run four separate `/release` actions at one clean reviewed S73
SHA. Each action retains its existing exact family allowlist, requires a fresh
immediate approval, verifies registry ownership and artifacts, and posts a
human release result to every included issue and pull request. The stable Rust
release supersedes rather than backfills the unpublished 0.13.2 crate set.
The approved versions are stable Rust and PyPI `rdocx` 0.14.0 and incubating
Rust and PyPI `rpptx` 0.12.1, under tags `rpptx-v0.12.1`, `v0.14.0`,
`py-rdocx-v0.14.0`, and `py-rpptx-v0.12.1`. The failed immutable
`rpptx-v0.12.0` tag published no packages and created no GitHub release. The
recovery incubating tag publishes first because packaged stable crates require
the shared 0.12.1 registry family.

**Depends on**: F-257 through F-263, F-X097 through F-X111, F-X113 through F-X122.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/99>.
**Test gate**: release preparation.
`s73_release_contract_requires_four_version_aligned_families` proves exact
Rust and Python version agreement, the 7-package stable and 15-package
incubating allowlists, the four reviewed tags, and reviewed notes whose
linked issues, pull requests, and credited contributors equal the per-family
inventory required by the stories each family ships. The seven-file Python
artifact sets, selected-family CLI release assets, separate tag approvals,
registry owners, and posted notifications are proved by `/release`.
After comment verification, close each included issue or pull request that
remains open and is fully addressed by the published outcome. Leave already
closed records unchanged.

### F-X113, Preserve appended paragraphs in document comparison (M)

Correct terminal paragraph-mark ownership when comparison inserts two or more
paragraphs after the original final paragraph. Rejecting the staged revisions
must reproduce the original without an extra empty paragraph, and accepting
must reproduce the edited story. Start and middle insertions, intentional empty
paragraphs, fields, drawings, and unrelated package parts retain their current
behavior. A self-closing original final paragraph expands around the tracked
boundary instead of emitting paragraph properties as raw body content.

**Depends on**: F-X097.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/115>.
**Test gate**: regression.
`comparison_appends_multiple_terminal_paragraphs_without_residue` compares one,
two, and three appended paragraphs and proves exact accept and reject
reconstruction with no synthetic terminal content.

### F-X114, Rebuild TOC entries with document styles and geometry (M)

Resolve built-in TOC levels by `w:name="toc N"` so localized style ids remain
authoritative. Use an effective style right tab when present and otherwise
derive the page-number stop from the field's section text width. Numbering tab
suffixes become ordered `w:tab` run content rather than literal control
characters inside `w:t`.

**Depends on**: F-X103, F-263.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/116>.
**Test gate**: regression.
`rebuilt_toc_uses_localized_styles_section_tabs_and_structural_suffixes` covers
localized ids, effective style tabs, A4 margins, invalid geometry fallback,
unmodelled style-child preservation, and numbered headings through save,
reopen, layout, and pinned Word 16.113 comparison.

### F-X115, Preserve modern comment metadata and identity (M)

Use the standard Open XML commentsExtended content type, keep comment ids and
parent links stable through rdocx save and reopen, and add optional validated
RFC 3339 dates through native `add_comment_with_date` and
`reply_to_with_date` and the Python `date` keyword. New ids use the unused
nonnegative space. No date remains the deterministic default, and docs state
that third-party editors may independently renumber comments.

**Depends on**: F-X109.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/117>.
**Test gate**: regression.
`comments_keep_standard_content_type_ids_dates_and_threads` proves content
type, ids, dates, parents, replies, resolved state, Python typing, and no-op
sidecar preservation through reopen.

### F-X116, Make Python story reads linear and complete (L)

Materialize one story source and owner inventory per Python snapshot rather
than rebuilding the package for every item. Use the same accepted-view walker
for StoryItem text, Paragraph text, and Paragraph runs. Visible runs nested in
tracked insertions and inline content controls retain recursive source paths,
support checked mutation, and become stale after structural edits.

**Depends on**: F-X106c, F-X109.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/118>.
**Test gate**: performance.
`python_story_inventory_scales_linearly` counts complete-story traversals while
doubling paragraphs, table cells, and hyperlinks, and the companion binding
gate proves nested visible runs agree across every text view. Story-item and
hyperlink snapshots each build their source inventory once, and path-backed
nested run mutation is checked through save, reopen, and stale-handle failure.

### F-X117, Render transparent and large raster pictures safely (M)

Premultiply decoded RGBA before constructing tiny-skia pixmaps. Raise the
checked decoded-image ceiling to 64 MiB for the reported 4000 by 1500 and 2100
by 2100 assets while retaining the 16 MiB encoded ceiling and checked dimension
and allocation bounds. Supported images render in PDF and every raster format.
An image beyond the bound records its scoped rejection, reports one stable
diagnostic, and lowers as a visible bounds fallback instead of disappearing.

**Depends on**: F-X104.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/119>.
**Test gate**: raster.
`straight_alpha_images_composite_with_premultiplied_pixels` proves transparent
stored white and black pixels compose identically, while the large-image gate
covers both reported successful sizes and one explicit over-limit failure.

### F-X118, Make notes rendering and replacement safe (L)

Render a notes slide from its already known owning slide when the producer
omits the reverse relationship, while rejecting conflicting owners. Add staged
formatting-preserving notes text mutation and include notes in presentation
replacement counts. The CLI shares the guarded output policy, adds `--expect`,
rejects zero unexpected matches, and publishes only after validation.

**Depends on**: F-X105, F-X111.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/120>.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/121>.
**Test gate**: CLI.
`rpptx_replace_is_guarded_counted_and_includes_notes` covers existing outputs,
input equality, zero and expected counts, speaker notes, formatting, rollback,
and staged-file cleanup. The companion
`notes_render_without_a_reverse_slide_relationship` regression covers
Google-style notes relationship graphs and rejects conflicting owners.

### F-X119, Complete round-three Python authoring and inspection (L)

The Python bindings accept in-memory Word picture insertion with checked
StoryItem placement and source-compatible path-aware comment positions for
table-cell paragraphs. Presentation handles expose shape bounds and identity,
run font name, size and color, text autofit mode, and a Run text setter that
preserves formatting. The existing direct-body RunPosition remains source
compatible.

**Depends on**: F-X106c, F-X108, F-X109, F-X115, F-X118.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/121>.
**Test gate**: binding.
`python_round_three_authoring_and_inspection_is_typed_and_lossless` exercises
every added Word and PowerPoint operation from installed wheels, then verifies
typing, relationships, properties, save-reopen behavior, and failure atomicity.

### F-X120, Accept fractional DOCX line spacing values (S)

Accept producer-written plain decimal `w:spacing/@w:line` values and normalize
them to the nearest signed integer twip, with exact half values rounded away
from zero. Preserve the existing integer path. Reject exponent notation,
non-finite values, malformed decimals, and values outside the signed 32-bit
range without weakening any other spacing attribute. Save the normalized value
canonically as an integer.

**Depends on**: none.
**GitHub pull request**: <https://github.com/tensorbee/rdocx/pull/122>.
**Test gate**: regression.
`word_fractional_line_spacing_opens_with_nearest_twip_values` covers observed
producer values, exact positive and negative halves, signed boundaries,
canonical save and reopen, invalid forms, and deterministic layout.

### F-X121, Adopt PR 123 authored line-chart portability (S)

Adopt Kevin Brown's PR 123 at exact head
`1375b6342548e79ec17faa99ac76a57e4a1c5e9b`. Authored axis titles declare an
empty layout followed by non-overlay semantics. Authored line plots explicitly
disable chart-level markers and smoothing so viewers do not invent per-point
markers or smooth the path. Other chart families retain their existing output,
and source data, editable workbooks, and palette choices remain unchanged.

**Depends on**: F-X087.
**GitHub pull request**: <https://github.com/tensorbee/rdocx/pull/123>.
**Test gate**: regression.
`authored_charts_emit_portable_viewer_defaults` proves title child order,
explicit false line defaults, non-line exclusion, parse and rewrite stability,
and unchanged workbook and palette semantics. The pinned external oracle binds
the resulting four-chart DOCX to
`ab67b50393fc5258f7a3e9719344639d665feccc2615b13cab1915ea9a84566b`.

### F-X122, Recover the immutable rpptx 0.12.0 release attempt (M)

Treat CRLF and LF as equivalent representations when the Rust release
workflow validates packaged README and licence text. Keep exact archive names,
member sets, executable checks, nonempty payload checks, archive checksums, and
text content after newline normalization. The failed immutable
`rpptx-v0.12.0` tag remains at its reviewed commit with no registry packages or
GitHub release. Prepare the complete 15-package incubating Rust family,
unpublished WASM and binding carriers, stable shared dependency pins, and PyPI
`rpptx` together at 0.12.1. F-X112 then publishes `rpptx-v0.12.1` before the
unchanged stable 0.14.0 family and finishes with `py-rpptx-v0.12.1`.

**Depends on**: F-X111, F-X119, F-X121.
**Test gate**: release regression.
`rpptx_0_12_1_recovery_is_version_aligned_and_line_ending_safe` proves exact
0.12.1 version carriers and dependency pins, the revised four-tag S73 contract,
CRLF acceptance for packaged text, and rejection of changed text or weakened
archive validation.

### F-X123, Accept producer TOC style variants (S)

Rebuild TOC fields whose custom-style list ends in one producer-added comma.
Resolve duplicate style identifiers deterministically for TOC discovery without
weakening the strict public style-graph validator used by style mutations.
Report each duplicate choice while retaining the package's complete style XML.

**Depends on**: F-X114.
**GitHub issues**: <https://github.com/tensorbee/rdocx/issues/124> and
<https://github.com/tensorbee/rdocx/issues/125>.
**Test gate**: regression.
`toc_rebuild_accepts_trailing_style_separator_and_duplicate_style_ids` proves
both producer variants, deterministic first-definition lookup, diagnostics,
entry generation, save and reopen, and unchanged strict mutation validation.

### F-X124, Make content cloning linear and explicit (M)

Clone one direct body child with bounded linear package work rather than
repeated whole-story discovery and reopen cycles. Preserve transactional
identity freshening, relationship scope, and stale-handle behavior. Python
type errors name `source` and `destination` and state that the latter is a
direct body index. Source and destination resolve from one owned story
inventory, and direct-content scans reserve section-property inspection for
preserved nodes.

**Depends on**: F-X106a, F-X116.
**GitHub issues**: <https://github.com/tensorbee/rdocx/issues/126> and
<https://github.com/tensorbee/rdocx/issues/132>.
**Test gate**: regression.
`clone_content_scales_linearly_and_names_invalid_arguments` proves bounded
growth over source-built body sizes, unchanged cloned content and identities,
and exact Python errors for the two reversed-signature cases.

### F-X125, Compare table grid changes (M)

Represent a changed table grid as one tracked table deletion followed by one
tracked table insertion when a cell-level revision cannot reproduce both
grids. Acceptance yields the edited table and rejection yields the original,
while surrounding paragraphs and unchanged tables retain normal comparison.
The established row-marker representation removes the unused table shell
during revision resolution, so no new block wrapper grammar is introduced.

**Depends on**: F-X065, F-X097.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/127>.
**Test gate**: regression.
`comparison_tracks_changed_table_grids_as_table_replacement` covers gained,
lost, and resized columns, exact acceptance and rejection, revision inventory,
schema order, save and reopen, and unchanged row and cell comparison.

### F-X126, Preserve drawings through comparison acceptance (M)

Close required drawing bindings on the inline or anchor root in comparison-only
story projections while preserving the source ownership used by package
serialization. Text-only edits around unchanged body, header, and footer
drawings pass the acceptance and rejection postconditions without hiding a
real drawing change or changing story-local relationship ownership.

**Depends on**: F-X097, F-X102.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/128>.
**Test gate**: regression.
`text_only_comparison_with_body_header_and_footer_drawings_accepts_exactly`
uses source-built root-owned and run-owned namespace variants across run, word,
and character comparison. It proves exact acceptance and rejection, byte-exact
self-comparison, scoped relationships and media, normalized raw drawing
payload, complex-field survival, and changed-drawing sensitivity.

### F-X127, Collapse adjacent page break requests (S)

Treat a paragraph-level `pageBreakBefore` immediately after a run-level page
break as one page transition. Do not collapse either request across intervening
visible content, paragraph shading or borders, revision marks, drawing-clear
offsets, tables, section transitions, columns, or non-page breaks.

**Depends on**: F-X101.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/129>.
**Test gate**: differential.
`adjacent_run_and_paragraph_page_breaks_share_one_transition` proves two pages
for the combined case, unchanged single-break controls, separation boundaries,
and the pinned LibreOffice page and text result in deterministic font mode.

### F-X128, Preserve Word paragraph and revision identities (M)

Retain namespace-aware root attributes on modeled paragraphs, runs, and
section properties, including `w14:paraId`, `w14:textId`, and every `w:rsid*`
value. Serialize them in deterministic source order, preserve foreign and
unknown attributes, and keep authored identity allocation separate from this
lossless reader correction.

**Depends on**: F-X115.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/130>.
**Test gate**: round-trip.
`paragraph_run_and_section_identity_attributes_survive_noop_save` proves exact
modeled and foreign root attributes across save and reopen, aliases, edits,
public item filtering, schema order, and deterministic bytes. Focused unit
coverage rejects alias duplicates and pins authored `paraId` precedence.

### F-X129, Tolerate unmatched notes placeholders (S)

Ignore a notes-slide placeholder whose complete placeholder key has no notes
master match. Continue overlaying every matched placeholder and rendering
non-placeholder notes content. Keep ambiguous or multiply matched placeholders
as hard failures and return one ordered diagnostic for each skipped overlay.

**Depends on**: F-X118.
**GitHub issue**: <https://github.com/tensorbee/rdocx/issues/131>.
**Test gate**: differential.
`notes_pdf_skips_only_unmatched_slide_placeholder_overlays` source-builds the
Google Slides index variant and proves unchanged matched output, exact notes
text and geometry, and package-byte preservation. Its focused unit matrix
proves source-ordered diagnostics, ambiguity rejection, and the required
slide-image failure.

### F-X130, Show package depth, footprint, and speed (L)

Strengthen the root and all 26 crate-local READMEs after the S74 authoring
surface is complete. Each page must explain the depth that matters to its own
consumer, lead with concrete implemented outcomes, and connect specialist
crates to the complete Rust, CLI, Python, and browser families. The `rdocx-py`
and `rpptx-py` pages remain the long descriptions published on PyPI and must
make their native engine, typed API, local execution, rendering, review, and
package-preservation advantages immediately visible.

Add reproducible, dated evidence for package footprint and speed where the
measurement applies. Rust crate archive sizes, CLI release assets, Python wheel
and source archive sizes, installed footprint, large-document layout and PDF
throughput, and any Python boundary measurements must name the exact version,
platform, build mode, input, command, and statistic. Every uniqueness statement
must be bounded to a named, dated set of reviewed official sources. Do not make
an unqualified claim about every library, or compare timings produced by
different workloads or environments.

Extend the existing README validator and its mutation tests so capability,
comparison, size, and performance evidence cannot drift from the command that
produced it. All Rust examples still compile, Python and shell snippets still
match the installed surfaces, local links resolve, and every published crate,
wheel, source archive, and CLI bundle carries the intended README text.

**Depends on**: F-X089, F-264, F-265, F-266, F-267, F-268, F-269, F-270.
**Test gate**: regression.
`readme_depth_footprint_and_speed_claims_are_evidence_backed` proves the exact
27-page inventory, complete family and Python depth summaries, bounded official
comparisons, reproducible measurement provenance, checked examples, and
byte-identical packaged long descriptions.

**Tracked human action**: Python wheel and source-distribution sizes remain
absent until the six-platform `wheels.yml` job records them for the
`rdocx-py` and `rpptx-py` pages. The installed Python site-packages footprint
and Python boundary timing remain absent until each reviewed wheel is measured
with its pinned interpreter. CLI release archive sizes remain absent until the
selected-family `publish.yml` tag job records all six assets for the matching
CLI page. WASM bundle sizes remain absent until the pinned wasm-pack and
wasm-opt jobs produce the reviewed browser artifacts for the two WASM pages.
These are release or human measurements, not values a native workspace test
can reproduce, so F-X130 publishes no placeholder number for them.

### F-X131, Retain only the namespace declarations a root attribute uses (S)

F-X128 retains producer root attributes on `CT_P`, `CT_R`, and `CT_SectPr`, but
`capture_root_attribute_record` records every namespace declaration on the
source element, including declarations no retained attribute uses. The existing
alias machinery already materializes a binding onto each element that needs one,
so a declaration used only by a child is emitted twice, once on the modeled root
and once on the child. Retain a declaration only when a retained attribute uses
its prefix.
**Depends on**: F-X128.
**Test gate**: regression.
`a_section_root_retains_no_namespace_declaration_its_attributes_do_not_use`
proves the redundant declaration is gone, that a declaration a retained
attribute does use is still written, and that expanded-name precedence still
resolves.

### F-X132, Match a retained namespace owner by structure, not by identity (S)

F-X128 gave `CT_R` and `CT_P` root-attribute retention and F-X131 keeps the
declaration a retained attribute's prefix needs, so every retained paragraph
and run is written with a `w` binding that the document root already owns.
`prepare_staged_package` flushes the document part before
`canonicalize_drawing_ids` reads it back, so
`nested_modeled_namespace_owners` in `crates/rdocx/src/document.rs` sees those
redundant declarations and treats each retained element as a nested namespace
owner. Two runs that agree on every retained fact then own each other's
declaration equally, and `replay_nested_namespace_declarations` fails closed
with `cannot identify retained 'r' nested namespace owner after mutation`.
A declaration that rebinds
a prefix to the URI already in scope resolves no name differently, so it owns
no namespace and must not make an element an owner.
**Depends on**: F-X128, F-X131.
**Test gate**: regression.
`accepting_revisions_still_saves_when_runs_carry_revision_identities` proves a
redlined document whose runs carry `w:rsid` identities accepts, saves, reopens
and keeps every identity attribute, while a genuinely rebinding declaration on
two indistinguishable owners still fails closed.

### F-X133, Stop rebinding a canonical prefix on every retained element (S)

Retained producer attributes keep their bindings in the private record, while
serialization omits same-URI declarations already guaranteed by the part
root. The main document guarantees `r` and `mc`, and writes canonical `wp`
unless a retained root binding claims that prefix. Headers and footers follow
the same `wp` rule alongside guaranteed `r`, and note parts guarantee `r`.
The existing canonical `w` and `w14` behavior remains. A regression pins one declaration of every
guaranteed prefix at the document root with producer attributes, a new binding,
and an unmodelled child preserved through reopen. Standalone paragraphs and
comments retain local bindings.
**Depends on**: F-X131, F-X132.
**Test gate**: regression.
`a_retained_element_does_not_rebind_a_prefix_its_part_root_declares` proves
document serialization emits no redundant canonical declaration, retains the
producer attributes and unmodelled child, and writes a genuinely new binding.

### F-X134, Keep Python story hyperlink snapshots linear (S)

The owned story hyperlink projection batches story owners, items, and namespace
scopes, but hyperlink text extraction reparses the complete physical story from
byte zero for every link. A document with one hyperlink per paragraph therefore
does quadratic prefix work even though the public Python snapshot API promises
one bounded native inventory per accessor.

Collect hyperlink spans before materializing their public records, inventory
the exact namespace scope at every hyperlink start in one source pass, and
extract each hyperlink text from its namespace-complete bounded fragment. Keep
relationship resolution, physical source order, nested-owner isolation, and
deduplication unchanged.
**Depends on**: F-X116.
**Test gate**: regression.
`story_link_snapshots_do_not_rescan_story_prefix_per_link` counts skipped story
prefix work and proves that doubling a hyperlink-rich story does not restart a
full-document scan for each link. The installed Python
`test_python_story_inventory_scales_linearly` gate must pass on the hosted
macOS binding runner without relaxing its scaling bound.

**Delivered.** Story hyperlink snapshots now inventory each physical source's
item and hyperlink namespace scopes in bounded passes, then extract link text
from the namespace-complete hyperlink fragment. The native regression reports
zero repeated prefix bytes, the unchanged Python scaling test passed twenty
consecutive runs, and the complete installed Python suite passed 67 tests.

### F-X135, Integrate PRs 146 through 151 and resolve unassigned reports (L)

Integrate the exact reviewed heads of contributor PRs 146 through 151 after
F-X134 restores the hosted Python gate. Preserve the contributor's focused
commits where their behavior remains correct, reconcile the overlapping
PowerPoint facade and binding surfaces as one coherent API, and re-record
combined package measurements only after the integrated tree is final. PR 148
also requires the deterministic golden-PNG baseline that its current head did
not update even though its intentional layout change moves three reviewed
rasters.

Resolve open Issues 134, 135, 136, 139, and 140. Contributor PRs 153 and 154
cover the same reports and supply additional regression cases. Comparison must
track paragraph formatting when producer revision identities are present, treat
an empty paragraph-property shell as absence, and
accept a reordered drawing paragraph without printing raw model bytes in a
failure. A save must retain an authored `nil` border token and preserve every
unchanged modeled part and relationship part byte for byte. Issue 138 remains
open because PR 148 addresses only paragraph spacing, cell margins, and border
bands. It does not implement the reported line-height and row-splitting gaps.

**Depends on**: F-X134.
**Test gate**: regression.
`no_op_save_preserves_every_unchanged_part` proves DOCX and PPTX saves retain
the exact bytes of modeled parts and relationship parts whose typed state did
not change, while a targeted edit rewrites only its owned part and graph edges.

### F-X136, Fix table row breaks and footer-only pages (L)

Complete Issue 138 after the paragraph-spacing, cell-margin, and border-band
portion delivered by F-X135. A source-built page-height fixture places a
repeated-header table before an empty following paragraph and a heading with
`w:pageBreakBefore`. Correct row measurement and default row fragmentation
must not leave a page carrying only the footer. Splitting ordinary rows
preserves cell text order, border ownership, repeated headers, and body-layout
fragments. Rows with vertical merges remain whole until a safe merge-fragment
model exists. `w:cantSplit` moves a row whole when
it fits on the next page. Explicit exact row heights retain their clipping
contract. The private reporter document is unavailable, so the source-built
cases pin the reproducible boundary rather than claiming its exact 53-page
render has been reproduced.

**Depends on**: F-X135.
**Test gate**: regression.
`table_row_breaks_before_footer_only_page_and_repeats_header` asserts the
footer-only page is absent, every body line appears once in reading order,
and the table header repeats on the continuation page.

### F-X137, Package and CLI safety contribution wave (L)

Review and integrate PRs 174, 178, 182, 185 and 197 against the current
sprint branch. Cover output refusal, atomic saves, root namespaces, empty
comments parts and package class selection. Reconcile shared package and
CLI changes without losing unrelated producer XML.
**Depends on**: F-X136.
**Test gate**: regression. The Issue 156, 157 and 164 reproductions, failed
save cases and unchanged-part checks pass on the combined prefix.

### F-X138, Word story and content contribution wave (L)

Review and integrate PRs 177, 179, 180, 191, 195, 202, 210 and 211.
Keep the stated stacks 195 before 202 and 210, and 202 before 211. Align
story visibility, replacement counts, direct body coordinates, comment run
positions and compatibility text box copies.
**Depends on**: F-X137.
**Test gate**: regression. Issue 163 and 172 entry points and the Issue
160 content-control location matrix pass on the combined prefix.

### F-X139, Word identity and comparison contribution wave (L)

Review and integrate PRs 183, 184, 190, 193, 198 and 205. Reconcile
identity preservation and comparison semantics with F-X138's changes in
`rdocx-oxml` and `rdocx/src/comparison.rs`.
**Depends on**: F-X138.
**Test gate**: regression. Identity-only changes create no redline, the
TOC rebuild accepts producer run attributes, and comparison option cases
pass on the integrated Word model.

### F-X140, Rendering and layout contribution wave (L)

Review and integrate PRs 175, 188, 196, 199, 200, 206 and 207. Own the
single reviewed hash baseline update from PR 188. Resolve PR 196's failed
Presentation fidelity corpus gate and PR 206's MSRV and Test failures before
either can land.
**Depends on**: F-X137.
**Test gate**: golden. The deterministic Word and presentation fidelity
gates and the output-stability harness pass with every expected delta
labelled and reviewed.

### F-X141, Word Python contribution wave (L)

The Word Python binding, stub and native facade expose the contributed table,
section, style, field, replacement, story, rendering and core-property APIs
from PRs 176, 187, 194, 201, 203 and 212 on the S76 prefix. Section updates
use layout defaults for missing partners, native new documents initialize
common Word styles, and inserted TOCs rebuild after heading edits. A separate
reviewed hash baseline records the output changes after F-X140. The remaining
Issue 168 checklist stays with F-X147.
**Depends on**: F-X139, F-X140.
**Test gate**: integration. The contributed Python workflow, typing smoke
and Rust parity cases pass after save and reopen.

### F-X142, Presentation Python contribution wave (L)

Review and integrate PRs 173, 181, 189, 192, 208 and 209. PRs 208 and
209 follow 189. Reconcile the shared shape, table, layout and binding
files before running the Issue 169 deck workflow. Complete its remaining
shape hyperlinks, shape and text range comments, and built-in table style
rendering alongside the contributed APIs.
**Depends on**: F-X140.
**Test gate**: integration. Python and Rust deck operations round-trip,
validate and render with correct shape geometry and table structure.

### F-X143, Revision listing and CLI story contribution wave (M)

Review and integrate PRs 186 and 204. PR 204 follows PR 198 from F-X139.
Verify the revised CLI text and validation behavior against the S76 story
and comparison prefixes.
**Depends on**: F-X139, F-X141.
**Test gate**: integration. Revision listing spans all stories and the CLI
detects malformed related parts and undefined styles in its stated scope.

### F-X144, Identity and producer matrices across operations (L)

Complete Issues 157, 159 and 160 after the contribution wave. Preserve
the attached identity and producer matrix rows as executable acceptance
cases, add the missing content-control walker locations and verify
`mc:Ignorable` on every edited part.
**Depends on**: F-X143, F-X151.
**Test gate**: regression. Both full matrices and the `add_picture` column
pass with no comparison refusal, dropped identity or lost producer part.

### F-X145, Comparison options and redline completion (L)

Complete Issue 161 after the option-binding and comparison PRs. Carry
edited-side comments through the redline, compare rebuilt TOCs and place
markers correctly after changed text, including PR 205's refused case.
**Depends on**: F-X143, F-X144.
**Test gate**: regression. Python and CLI options agree and redline accept
reproduces edited text and comment state in every specified case.

### F-X146, Word line height and inline picture spacing (L)

Complete both parts of Issue 162 and the rich-line part of Issue 226,
reviewing PRs 222, 225 and 237 against the integrated rendering prefix.
F-X163 completes Issue 226's plain-line fit after PR 242. Measure text
pitch with bundled fonts and prevent proportional spacing from scaling a
tall inline picture's height.
**Depends on**: F-X140.
**Test gate**: golden. The four-family, two-size, two-spacing pitch matrix
and the Word-exported picture fixture meet pinned expected geometry.

### F-X147, Complete rdocx Python production checklist (L)

Complete Issue 168 after F-X141 and F-X153. Address rich section story authoring,
transactional counted replacement and rendering parity in the integrated
build, then finish style mutation and removal, run removal, an XML or
package write path and text-anchored comments. Record a reviewed scope
decision for any intentionally unsupported long-tail operation.
**Depends on**: F-X141, F-X145, F-X153.
**Test gate**: integration. Every Issue 168 checklist example runs from
Python, or has an explicit accepted scope decision and documented fallback.

### F-X148, Complete rpptx Python production checklist (L)

Independently exercise F-X142's and F-X156's complete Issue 169 checklist against the
production deck fixture and cross-viewer output. Record any integration gap
found after S77 and repair it through a scoped follow-up rather than claiming
the checklist was completed twice.
**Depends on**: F-X142, F-X156.
**Test gate**: integration. The complete deck chain round-trips, validates
and renders, with every checklist item either passing or explicitly scoped.

### F-X149, Word fixture and workflow acceptance gate (L)

Convert Issue 158's attached fixtures, two matrices and two workflows
into repeatable Word acceptance evidence. Keep deterministic font mode and
review every expected hash delta. Exercise DOCX, its CLI and Python binding
after the narrower Word issues are repaired.
**Depends on**: F-X144, F-X145, F-X146, F-X147, F-X152.
**Test gate**: differential. The report workflow and both matrices pass
against their pinned reference outputs.

### F-X150, Reconcile issue closure evidence (M)

Audit every acceptance criterion in the 22 open Issues 156 through 172,
215 through 217, 226 and 227 against the integrated S76 to S82 result.
Issue 158 closes last. Do not infer acceptance from a PR merge.
**Depends on**: F-X149, F-X158.
**Test gate**: integration. Each criterion has a passing artifact or an
explicitly unresolved issue, and the final sprint verify and review gates
pass before any closure action.

### F-X151, Word preservation and comparison PR intake (L)

Review PRs 214, 228, 229, 232, 233 and 239 after S77. Reconcile story
export, lexical validation, field text and TOC comparison without dropping
unmodelled XML.
**Depends on**: F-X143.
**Test gate**: regression. Producer, compare and package round trips pass on
the combined prefix with the hash harness unchanged or a reviewed delta.

### F-X152, Full-story CLI diff and count repair (M)

Review PR 236 and complete Issue 227 across cells and related stories.
**Depends on**: F-X143, F-X151.
**Test gate**: regression. The issue reproducer locates every changed story
paragraph and counts each changed paragraph once.

### F-X153, Word Python supplemental contribution (M)

Review PR 220 for anchored CLI comments, numbering and run removal. Reconcile
its overlap with F-X141 before finishing the Issue 168 checklist.
**Depends on**: F-X141, F-X144.
**Test gate**: integration. Python and CLI comment coordinates, numbering and
run removal pass after save and reopen.

### F-X154, Presentation text and preservation repair (M)

Review PRs 218 and 223 and complete Issues 215 and 216. Preserve unmodelled
`a:bodyPr` attributes and give text-frame line feeds authoring semantics.
**Depends on**: F-X142.
**Test gate**: regression. Unedited text bodies remain byte-equivalent where
required, and assigned line feeds survive layout, validation and rendering.

### F-X155, Presentation drawing API contribution (L)

Review PRs 219, 221, 224, 230 and 234. Replay PR 219 after 189 and PR 234
after 207. Reconcile hyperlinks, line ends, shadows, preset geometry and
connector style in schema order.
**Depends on**: F-X142, F-X154.
**Test gate**: integration. The authored effects and links reopen in
python-pptx, validate and match the pinned cross-viewer renders.

### F-X156, Presentation slide and table contribution (L)

Review PRs 231, 235 and 238. Replay PR 231 after 181 and reconcile scoped
replacement and slide import against F-X142's built-in table style resolution.
**Depends on**: F-X142, F-X155.
**Test gate**: integration. Imported and edited decks reopen, validate and
render without losing media, notes or relationships.

### F-X157, Complete deck-chain authoring checklist (L)

Finish all six Issue 217 items after their PRs, including scope decisions
where explicitly allowed. Check the complete chain in both Python APIs and
independent viewers.
The acceptance chain changes a shape's preset, sets its direct outer shadow,
suppresses a connector's theme effect, and adds a line end before importing a
slide with media and notes and replacing text at slide and frame scope. The
saved package is read by python-pptx 1.0.2, validated by the rpptx CLI, and
rendered by LibreOffice 26.2.5.2 and deterministic rpptx at 72 DPI. Each item
has a scoped raster comparison with the tolerance in the testing strategy.
No authoring operation needs a fallback. Pinned LibreOffice and structural
reopen provide viewer evidence where PowerPoint 16.104 is unavailable.
**Depends on**: F-X148, F-X155, F-X156.
**Test gate**: differential. Each item passes python-pptx reopen, `rpptx
validate` and LibreOffice versus rpptx rendering, or has a reviewed scope
boundary and fallback.

### F-X158, Presentation fixture and workflow acceptance gate (L)

Convert the attached deck workflow into repeatable acceptance evidence after
Issues 169, 170, 215, 216 and 217 are addressed.
**Depends on**: F-X140, F-X148, F-X154, F-X157.
**Test gate**: differential. The full deck workflow round-trips, validates
and matches pinned viewer outputs.

### F-X159, Word document validity baseline (M)

Review PR 240 after S82. Correct the package metadata and nested table output
that prevent a new document from opening in Word. Keep its intentional hash
change isolated and reviewed.
**Depends on**: F-X149.
**Test gate**: differential. Word opens the source-built document and nested
table, and the deterministic baseline changes only in the declared entries.

### F-X160, Tolerant style, drawing and measurement reads (L)

Review PRs 248, 249 and 250 after F-X159. Accept repeated style IDs and
multiple default table styles already present, duplicate drawing IDs within
one part, and decimal integer measurements without introducing new defects.
**Depends on**: F-X159, F-X147.
**Test gate**: regression. Both Issue 243 fixtures pass every style mutator,
and Issues 246 and 247 pass read, edit, save and reopen with the specified
rounding and drawing identity behavior.

### F-X161, Compact Word XML and namespace preservation (L)

Review PR 251 after F-X160. Serialize only changed parts compactly, declare
`w` once on the root and preserve unknown subtrees and producer attributes.
Own one labelled and reviewed baseline change for this sprint.
**Depends on**: F-X144, F-X160, F-X169.
**Test gate**: round-trip. Issue 245's edited package retains untouched
regions and valid namespaces, and its 20 declared hash entries reconcile.

### F-X162, Per-paragraph section width and pagination (M)

Review PR 241 after F-X161. Resolve section width at each paragraph rather
than using the final section's width for the whole document. Own one reviewed
baseline change for this sprint.
**Depends on**: F-X161.
**Test gate**: golden. Mixed-section paragraphs, tables and page boundaries
match the pinned Word render with deterministic fonts and declared hashes.

### F-X163, Plain-line trailing-space fit (M)

Review PR 242 after F-X146 and F-X162. Complete Issue 226's plain-line
symptom while preserving the rich UAX 14 correction from PR 222. Own this
sprint's labelled hash and golden pixel baseline change.
**Depends on**: F-X146, F-X162.
**Test gate**: golden. Both paths break at Word's reported widths without a
leading space or punctuation error, and all expected deltas are reviewed.

### F-X164, Added shape theme style (M)

Review PR 252 after PRs 207, 230 and 234. Emit an ordered `p:style` for new
shapes and verify its visible theme fill or line.
**Depends on**: F-X155, F-X169.
**Test gate**: differential. Source-built shapes reopen in python-pptx and
PowerPoint and render visibly in PowerPoint, LibreOffice and rpptx.

### F-X165, Python and CLI tracked revision view (M)

Review PR 256 after the S77 revision inventory. Expose accepted and tracked
views consistently in Python and the CLI, rejecting invalid selectors.
**Depends on**: F-X143, F-X150, F-X169.
**Test gate**: integration. Issue 253's old and new text appears in the
selected PDF and CLI output under pinned Poppler, with Python parity.

### F-X166, Picture and final-block comparison revisions (L)

Review PR 257, PR 258, PR 260 and stacked PR 261 in that order. Preserve
both image media and final paragraph properties through accept and reject.
Reconcile shared comparison and revision code against F-X145 and F-X165.
**Depends on**: F-X145, F-X165.
**Test gate**: differential. Issue 254's three picture cases agree with Word
for Mac, and Issue 255's five final-block pairs open in Word with correct
revision marks, content and properties in both directions.

### F-X167, Accepted-view exporters and readers (L)

Review PR 259 after PR 256 and the content-control exporter work, then PRs
262 and 263. Make HTML, Markdown, plain text and layout agree on deleted
paragraph and row visibility without losing retained content.
**Depends on**: F-X151, F-X165, F-X166.
**Test gate**: regression. Every accepted-view output agrees after save and
reopen, and the Python bindings gate that fails on PR 259 is green.

### F-X170, High-level Word style formatting API (M)

Address Issue 264 with convenient paragraph and run formatting methods on the
existing `StyleBuilder`. Keep `Document` style create, update, lookup, default
selection and removal as the management surface. Demonstrate a custom style
that sets alignment, spacing, indentation, font, size, bold and colour without
constructing `CT_PPr` or `CT_RPr` directly. Preserve the typed property escape
hatch and imported unmodelled XML. Save, reopen and resolve the effective style
against the authored paragraph and run.
**Depends on**: F-246, F-X169.
**Test gate**: round-trip. The convenient settings survive save and reopen,
apply to document content and match their typed property equivalents.

### F-X171, Pin LibreOffice for macOS Python acceptance (S)

The S84 main merge passed the local full and source-built Python gates, but its
hosted `Python bindings (rpptx)` job failed the Issue 158 deck workflow because
the macOS job installed Poppler without LibreOffice. Download the reviewed
LibreOffice 26.2.5 macOS aarch64 image with a fixed SHA-256, mount it for the
job, verify build 26.2.5.2, and expose `soffice` to the complete presentation
binding suite. Keep the Linux viewer jobs and their existing installer intact.
Add a workflow assertion that prevents this macOS oracle setup from silently
disappearing. After hosted `main` CI passes, reconcile the 32 S84 PRs and nine
issues using their already reviewed criterion evidence, with Issue 158 last.
**Depends on**: F-X168.
**Test gate**: workflow regression. The CI contract rejects missing or
bypassed pinned viewer setup, the exact DMG digest and mounted build identity
match the reviewed oracle, and the full local presentation binding suite
passes. The hosted presentation job and aggregate required CI gate must pass
at the S85 close boundary before any S84 contribution closure.

### F-X172, Unified release artifacts and provenance (L)

Address [Issue 266](https://github.com/tensorbee/rdocx/issues/266) by making
each new stable or incubating tag select its Rust crates, CLI binary and Python
distribution at one matching version. Build the six CLI archives, six
`cp39-abi3` wheels and source distribution for that family. Publish one GitHub
release containing all thirteen files and one `SHA256SUMS` covering each of
them. Generate GitHub build provenance attestations for every CLI archive, wheel
and source distribution. Verify each downloaded subject with `gh attestation
verify FILE -R tensorbee/rdocx`. Preserve the build-only manual wheel
preflight and PyPI trusted publishing with no long-lived token. Keep old
release tags immutable. Release new tags from the exact reviewed `main` merge
SHA after sprint close. Update the release command, workflow contracts, tests
and the build and toolchain HLD for the unified family contract.
**Depends on**: F-X111, F-X094f.
**Test gate**: release workflow regression. Mutation checks reject a missing
asset or attestation, a checksum omission, a mismatched family or version,
untrusted PyPI publication, premature GitHub release creation, and partial
registry publication reported as success. Manual dispatch remains build-only.

### F-X173, Prepare unified rpptx 0.13.0 family (M)

Prepare the minor release of the 15 shared OOXML and PowerPoint crates,
`rpptx-cli` and the `rpptx` Python distribution at 0.13.0 from the reviewed S86
source. Record version and internal pin changes, metadata, compatibility and
contributor notes under `rpptx-v0.13.0`. Complete local preparation at the
dependency checkpoint, then pass the full gate and clean review on the final
integrated S86 result. After `/close-sprint` merges the reviewed source
to `main`, `/release` verifies that main merge SHA, obtains separate final
approval, and checks every crates.io and PyPI version, asset, checksum,
attestation, release note, owner role and contributor notification.
**Depends on**: F-X172, F-X133, F-271, F-272, F-273.
**Test gate**: release preparation. The local dry run, metadata contract,
build-only wheel preflight, scoped verification and zero-finding microscope
pass on the prepared source. The integrated full gate and clean sprint review
follow both preparations at S86 closure. Publication follows `/close-sprint`
from `main` with its own final approval and hosted verification.

### F-X174, Prepare unified rdocx 0.15.0 family (M)

Prepare the minor release of the seven stable Word crates, `rdocx-cli` and the
`rdocx` Python distribution at 0.15.0 from the same reviewed S86 source.
Record version and internal pin changes, metadata, compatibility and
contributor notes under `v0.15.0`, including the S85 to S86 work and Issue
266's release changes. Complete local preparation with the selected-family
metadata contract, exact package dry run, host wheel and source distribution
rehearsal, scoped gate, and zero-finding review. The integrated S86 source then
passes the full gate and clean review. A six-platform manual build-only
rehearsal runs on the final pushed sprint SHA before closure. After
`/close-sprint` merges the reviewed source to `main`, `/release` verifies
that main merge SHA, obtains separate final approval for this tag, and checks
every crates.io and PyPI version, asset, checksum, attestation, release note,
owner role and contributor notification.
After both family releases pass, comment on Issue 266 with the published
`rpptx-v0.13.1` and `v0.15.0` numbers and links to their verified releases.
**Depends on**: F-X173.
**Test gate**: release preparation. The local dry run, metadata contract,
build-only wheel preflight, scoped verification and zero-finding microscope
pass on the prepared source. The integrated full gate and clean sprint review
follow both preparations at S86 closure. Publication follows `/close-sprint`
from `main` with its own final approval and hosted verification.

### F-X175, Refresh CLI archive evidence after release hardening (S)

The S86 main CI run found that the two CLI archive measurements still describe
the packages before the reviewed Windows CLI stack fix. Recompute the exact
`rdocx-cli` and `rpptx-cli` source archive footprints from the final tracked
tree and update their README rows and the enforced measurement inventory.
Keep all other release metadata and behavior unchanged. The focused S87 repair
is the reviewed source boundary for both prepared Issue 266 families.
**Depends on**: F-X174.
**Test gate**: release regression. The two CLI archive measurements match
fresh package builds on macOS and Linux, both README and release-regression
CI jobs pass, and the full gate remains green with 49 unchanged hash entries.

### F-X176, Repair unified release inventory and respin PowerPoint (M)

The first `rpptx-v0.13.0` tag passed platform builds and attestation checks,
then stopped before publication because `cargo publish --dry-run` did not leave
`.crate` files for the tag job's archive inventory. Retain that failed tag.
Check the unpacked package directories produced by the successful dry run,
and add a clean-target regression that rejects missing or extra selected
packages. Prepare PowerPoint 0.13.1 across all carriers and
notes, keep the stable Word family at 0.15.0, and prove both tag paths with
build-only rehearsal, full verification and a clean sprint review.
**Depends on**: F-X175.
**Test gate**: release regression. The clean-target package inventory check,
both family version and notes contracts, manual build-only artifact rehearsal,
hosted release regressions, full local gate and unchanged 49-entry hash harness
all pass before `/close-sprint S88` and fresh release approvals.

### F-X177, Accept unified fontdb source features (S)

Adopt [PR 269](https://github.com/tensorbee/rdocx/pull/269) from
`changjoon-park`, reviewed at `7949573b92fd988a682a11b8e674af85e5df622e`.
Cargo may enable `fontdb/memmap` through another dependency such as `usvg`.
The font source match compiles with that unified feature and reads the
additional source through `Database::with_face_data`. Retain the existing
binary and ordinary file paths and deterministic bundled-only construction.
No public API, dependency version or default feature changes.
**Depends on**: none.
**Test gate**: compilation regression. `cargo check -p oxml-layout --features
fontdb/memmap` fails on the claimed base and passes after the fix. Default,
no-default and no-default plus memmap checks and font tests pass, with all
49 hash entries unchanged. Record the contributor disposition for sprint close.

### F-X184, Safe comment ownership during content removal (L)

Resolve every criterion of [Issue 282](https://github.com/tensorbee/rdocx/issues/282), reported by `hadim`. Whole safely owned comment graphs removed with body, story, row or cell content lose their root, descendant replies and selectively owned commentsExtended, commentsIds and commentsExtensible entries. Partial cuts and comment-bearing detached fragments refuse atomically. Unprovable imported graph ownership refuses instead of discarding opaque XML. Add fallible native try_remove_content for Python and retain the legacy bool method with false on refusal and no panic. Raw all-story validation distinguishes an orphan from a reference-only comment or a linked reply. Preserve unrelated parts and source XML.
**Depends on**: F-X179, F-271.
**Test gate**: regression. `comment_removal_preserves_thread_closure` proves the reported failure before implementation and exact successful or refused behavior after save and reopen. Existing Rust, Python, typing and CLI entrypoints cover the complete issue criteria, with unrelated package members and opaque XML preserved. All 49 hash entries remain unchanged. Record reporter provenance and full acceptance for sprint close.

### F-X188, Preserve comment ownership when replacing or removing whole stories (L)

Resolve every criterion of [Issue 288](https://github.com/tensorbee/rdocx/issues/288), reported by `hadim`. Extend the completed F-X184 ownership transaction to default and first-page text header/footer replacement, raw and image replacement including even variants, and glossary entry removal. Complete safely owned threads and their selectively linked companions are removed together. Shared story targets retain comments while a section still references them, including reopened producer parts and different relationship ids reaching the same target. Preserve producer-owned parts and opaque references rather than deleting them to manufacture an ownership decrease. Add checked fallible native text setters for Python error propagation while retaining documented legacy wrapper conventions. Ambiguous ownership and partial cuts refuse before publication. Explicitly test whole endnote removal and all Issue288 variants rather than inferring acceptance from common helpers.
Issue [292](https://github.com/tensorbee/rdocx/issues/292) also requires
documented omission of comments from building-block creation, replacement,
fragment update, extraction and insertion. Filter qualified transferred
markers before comment dependency import. Isolate glossary-local ownership
from main comments even when numeric ids match, retain unrelated producer
parts and refuse ambiguous ownership atomically. Metadata-only updates and
general fragment behavior retain their existing contracts. Issue291 is
included as a separate field snapshot correction after this story. Issue281
remains open for F-184 and Issue264 remains excluded.
**Depends on**: F-X184.
**Test gate**: regression. `whole_story_removal_and_replacement_preserve_comment_closure` fails against the exact claimed Base for uncovered replacement and glossary routes, then passes source-built cleanup, shared-reference, companion, namespace and atomic refusal controls after save and reopen. Existing Rust, Python and CLI entrypoints cover the complete issue criteria. All 49 hash entries remain unchanged. Record reporter provenance and full acceptance for sprint close.

### F-X185, Expose comment anchor text and story location (M)

Resolve every criterion of [Issue 283](https://github.com/tensorbee/rdocx/issues/283), reported by `hadim`. Expose checked native and CommentRef anchor text and StoryRunRange, then frozen typed Python snapshots and CLI JSON locations. Accepted-view span text joins paragraphs with newline. Reference-only comments have no paired range and empty anchor text, while orphans have neither. Keep the seven-argument Python Comment constructor compatible through optional new defaults. Cover table cells, tracked insertions and block and inline goog_rdk wrappers.
**Depends on**: F-X184.
**Test gate**: regression. `comment_anchor_snapshots_match_accepted_story_spans` proves the reported failure before implementation and exact successful or refused behavior after save and reopen. Existing Rust, Python, typing and CLI entrypoints cover the complete issue criteria, with unrelated package members and opaque XML preserved. All 49 hash entries remain unchanged. Record reporter provenance and full acceptance for sprint close.

### F-X186, Move comment anchors without losing threads (M)

Resolve every criterion of [Issue 284](https://github.com/tensorbee/rdocx/issues/284), reported by `hadim`. Move a checked root comment's range and reference without replacing thread identity, metadata, replies or resolved state. Reuse staged story-range placement and support cross-story destinations where existing placement permits comments. Unknown ids, reply ids and unsupported destinations refuse atomically. Remove goog_rdk wrappers made empty solely by marker removal. Provide move-to-text with the existing add-comment body search scope and a CLI move command.
**Depends on**: F-X185.
**Test gate**: regression. `comment_moves_preserve_thread_identity` proves the reported failure before implementation and exact successful or refused behavior after save and reopen. Existing Rust, Python, typing and CLI entrypoints cover the complete issue criteria, with unrelated package members and opaque XML preserved. All 49 hash entries remain unchanged. Record reporter provenance and full acceptance for sprint close.

### F-X187, Scoped paragraph and cell text replacement (M)

Resolve every criterion of [Issue 285](https://github.com/tensorbee/rdocx/issues/285), reported by `hadim`. Add Python Paragraph.replace_text and Cell.replace_text with run-aware matching, optional expected count and document transaction atomicity. Expose Document.replace_text_at for a checked detached StoryItem to cover supported header, footer and note owners without giving snapshots an originating Document mirror. Preserve comment ranges, bookmarks and formatting outside a match. Keep unselected content and parts unchanged.
**Depends on**: none.
**Test gate**: regression. `scoped_text_replacement_preserves_unselected_content` proves the reported failure before implementation and exact successful or refused behavior after save and reopen. Existing Rust, Python, typing and CLI entrypoints cover the complete issue criteria, with unrelated package members and opaque XML preserved. All 49 hash entries remain unchanged. Record reporter provenance and full acceptance for sprint close.

### F-X190, Preserve cached text in multi-run complex field story snapshots (M)

Resolve [Issue291](https://github.com/tensorbee/rdocx/issues/291), reported by
Hadrien Mary (`hadim`) and contributed through
[PR287](https://github.com/tensorbee/rdocx/pull/287), semantic commit
`65365c4b2de49956f73c09bb326190e336db9bc8`. Bulk story snapshots and Python
StoryItem.text expose the stored result of a complex field spanning sibling
runs, matching existing direct story reads. Preserve one namespace context,
valid nested field and link coordinates, source XML, locations, ordering and
read-only revisions. No new public snapshot API. Run exclusively after
F-X188 and before F-282 resumes. Issue281 remains open for F-184 and Issue264
remains excluded.
**Depends on**: none.
**Test gate**: regression. `complex_field_story_snapshots_preserve_cached_text`
fails on the exact claimed Base through existing APIs, then passes actual
native and rebuilt Python controls across supported owners, nested fields,
links, aliases, namespace shadows and retained source bytes. All49 hash
entries remain unchanged. Record contributor provenance and full acceptance
before sprint close.

### F-X192, Complete the remaining Word bibliography catalogue (L)

Complete the unfinished Word bibliography catalogue scheduled for S98. F-282 delivers the native source-authoring foundation and its measured
formatter subset. This story owns the full pinned Word catalogue result
contract, without claiming the existing subset establishes complete parity.

Cover all seventeen schema source kinds, all twelve installed styles and all
223 documented ordinary bibliography locale selectors, plus any independently
established valid extensions. Retain the pinned style editions rather than
substituting newer editions. At the S90 development checkpoint, 210 of 223
dense APA locale grammars pass. Remaining selectors are1054,1105,1107,1108,
1109,1111,1112,1113,1115,1121,2117,2128 and2145. Existing captures are supporting
evidence. Completed own2117 and2145 controls and prepared Thai/Tibetan controls
do not count as implemented consumers.

Finish the remaining dense APA consumers, broader formatting across the other
eleven styles, sparse and repeated source properties, personal and corporate
contributors, multiple authors, missing author/date values and mixed scripts.
Complete multiple-source and empty-selection bibliography behavior, grouped
citations, locators, suppression, prefix/suffix, author/year disambiguation,
source and field locale precedence, application-default context and filtering.
Derive locale-aware sorting, numeric assignment and stable ties from
independent Word discriminators. Partial ICU or Core Foundation matches do not
establish the Word comparator. Do not invent an ASCII or en-US approximation.
A new collation dependency requires an approved design revision and exact
publication, MSRV, WASM, bundled-data, supply-chain and archive riders.

Keep exact source identities, opaque XML, rich run and paragraph properties,
IEEE table structure, physical field ownership and cache boundaries. Prove
source mutation, repeat update, save/reopen and atomic refusal across supported
stories. Equivalent namespace prefixes produce equal library results while
preserving imported XML and documenting the measured Word discrepancy.
Retain historical interrupted-update captures as cache-preservation evidence,
and use actual completed update and normal reopen captures for output claims.
Never generalize font coverage from a finite measured grapheme.

**Depends on**: F-282.
**Test gate**: differential. `full_bibliography_catalogue_matches_pinned_word`
checks exact source identities, instructions, selection, ordering, readable
text, rich properties and reopened packages across the complete independently
observed catalogue. No catalogue row is skipped, mapped by assumed language
alias or accepted from a cache fallback. All remaining branch and context
controls pass. Unsupported producer extensions still preserve caches and
metadata with diagnostics. Scoped verification, independent zero-finding
review, unchanged or explicitly reviewed harness output and all earned risk
riders are required before completion.

### F-X193, Preserve semantic mark highlights through nested HTML runs (S)

Assess Pedro Assumpcao's (`pedroassumpcao`)
[PR294](https://github.com/tensorbee/rdocx/pull/294) at immutable head
`63cd119dcb8fa9c367bf93141f79780db7f64886`. The HTML importer must carry the
semantic highlight of `mark` through nested inline children, including bold
and italic runs, while ordinary CSS backgrounds remain scoped to their own
elements. Unmarked siblings must not inherit the highlight. Preserve existing
inline formatting, importer behavior and producer XML on unrelated paths.

**Test gate**: regression. Nested bold and italic marked runs retain their
highlight after public import and save/reopen, with unmarked siblings and
ordinary CSS backgrounds as discriminating controls. Custom mark colours,
nested marks and child background precedence retain their distinct scopes.
Unrelated opaque sibling XML and package parts retain exact source bytes
through save/reopen. Assess the contribution
against current main and record actual scoped verification and clean review.

### F-X194, Complete Python table text and indentation APIs (M)

Resolve `hadim`'s Issues
[295](https://github.com/tensorbee/rdocx/issues/295) and
[300](https://github.com/tensorbee/rdocx/issues/300) through existing table and
binding types. PowerPoint `Cell.text_frame` returns the same public TextFrame
surface used by shapes, including paragraphs, runs, font properties,
vertical_anchor and word_wrap. Users can add paragraphs and set alignment,
spacing, size, name, bold and colour without replacing the editable table with
shapes. Expose Cell.vertical_anchor directly if the existing text-frame API
does not provide it. Retain ownership, stale-handle and merged-cell contracts.

Word `Table.indent` reads and writes Length values and accepts None to remove
`w:tblInd`. Support negative values in the checked Rust authoring path and
Python setter as well as positive and zero values. Respect the existing Word
alignment, margin and compatibility placement rules rather than introducing
a second positioning policy. Invalid inputs fail without partial mutation.

**Test gate**: regression. Reproduce both missing Python APIs on the claimed
base, then verify installed Python cell text-frame edits and signed table
indentation through save/reopen and actual PDF/raster output. Include two
paragraphs, merged-cell ownership, stale handles, removal of explicit indent,
negative/zero/positive values and existing alignment/margin controls. Use
existing integration entrypoints and public objects.

### F-X195, Honor caller fonts and variable styles across render paths (L)

Resolve `hadim`'s Issues
[296](https://github.com/tensorbee/rdocx/issues/296) and
[297](https://github.com/tensorbee/rdocx/issues/297) as one font-selection
contract across Word and PowerPoint measurement and output paths.

Provide consistent `--font-dir` handling on every CLI command that lays out
text, including raster/PDF convert, render and thumbnail. An output format
that cannot consume the option reports an error rather than silently ignoring
it. Every Python render and layout entrypoint accepts caller fonts and font
directories with the same semantics as Word Document.to_pdf. Cover Word page
PNG, SVG, selected/all-page rendering and PowerPoint slide/notes PDF,
selected/all-slide and all-notes raster rendering and text_layout. Fit checks,
PDF and previews use the same face selection and caller-font precedence.

Resolve variable fonts to the requested weight and italic/slant instance for
both layout metrics and output glyphs. Bold selects a matching named or axis
instance, normally wght700. Italic uses an appropriate face or ital/slnt axis.
Instantiate embedded outlines or apply variation during subsetting so PDF and
raster output reflect the selected style. Where no requested face or instance
is available, use a verified synthetic style, never regular glyphs under a
misleading Bold font name. Preserve static-font, bundled-font, fallback and
font-licence behavior. Any dependency change earns the packaging, MSRV, WASM,
archive and supply-chain riders in the approved design.

**Test gate**: regression. A caller-only family absent from system discovery
produces consistent CLI and installed Python Word/PowerPoint PDF, PNG and
applicable SVG/notes/fit results. Pinned variable-font controls independently
prove different regular/bold outlines and matching layout metrics, italic
selection and truthful embedded faces. Use deterministic font mode and
existing generated or appropriately licensed fixture conventions. Cover
invalid directories, unsupported formats, default/bundled/static fonts and
repeated rendering. Declare and independently review every expected rendering
delta before moving the exclusive hash baseline.


PR [315](https://github.com/tensorbee/rdocx/pull/315) entered the next rolling
refresh at immutable head `af388f255d3b178aaa0aef03313034b815c71eb2`. It closes
Issues296 and297 and belongs to existing F-X195, with Issue313 crediting that
same implementation. No new story is created. Read the complete changed
files, comments and reviews before claiming font work. Reuse accepted loader,
render-path and instance changes, then implement uncovered approved criteria.
Its stated omissions include synthetic italic, bundled variable instances and
caller-font handout/PDF-A paths. Its fallback on subset failure, directory
precedence and synthesis geometry must be reconciled against the approved
plan. Contributor gate claims are not local evidence. Archive measurements
will be generated from actual integrated source.

### F-X196, Validate visible paragraph content without false empty warnings (S)

Resolve `hadim`'s
[Issue298](https://github.com/tensorbee/rdocx/issues/298). CLI validation must
not classify a paragraph as empty merely because its text projection trims
to an empty string. Paragraphs containing a drawing, picture, object, chart,
equation, field, symbol, tab, break, footnote/endnote reference or a content
control containing visible content are nonempty. Align a qualified read-only
main-part content scan with the recursive body paragraph inventory. Exclude
direct body and direct nested cell tables and unsupported wrappers, and enter
control-owned tables. Properties, range markers, deleted content
and producer page-break hints remain empty. Do not mutate the document or
discard opaque source XML. Truly empty
paragraphs retain the existing warning and metadata checks remain intact.

**Test gate**: regression. The reported three-picture document has no empty
paragraph warning. Independent controls cover every listed content kind,
content controls, text/whitespace-only and truly empty paragraphs, with
unchanged title and author diagnostics and no document mutation.

### F-X197, Unify Python colour inputs and font colour formats (L)

Resolve `hadim`'s
[Issue299](https://github.com/tensorbee/rdocx/issues/299) across both Python
families. Every public colour-setting parameter or property accepts RGBColor
and six-digit hex strings with or without a leading hash. Scalar RGB getters
return RGBColor consistently, whose string form is the hex value. Font.color
returns a colour-format object supporting writable rgb and the available type
and theme_color model, so the python-docx/python-pptx font.color.rgb idiom
works. Retain direct font.color assignment as a compatibility shortcut.
Preserve existing fill, line, shadow and theme format semantics rather than
flattening theme colours into guessed RGB values.

Cover style creation/update, font colour, table and cell borders, cell and run
shading, and every other existing colour-taking public call in both families.
Invalid types raise TypeError naming the affected parameter, accepted forms
and actual type. Invalid hex values receive explicit validation errors and
failed assignments leave source state unchanged. Update existing stubs,
examples and typing/runtime checks with the approved compatibility decision.

**Test gate**: regression. Installed Word and PowerPoint bindings accept both
input forms at every inventoried setter, return consistent RGB values and
support the nested font.color.rgb idiom, direct assignment and existing theme,
fill, line and shadow APIs. Save/reopen, malformed inputs, unset colours,
stale handles and actual rendering distinguish correct formats from coercion
or source loss. Strict typing and stub checks agree with the runtime API.

Shared expanded acceptance from Issues303 and308 remains owned by F-X197:
one concrete PyO3 extractor in existing oxml-py-support serves both real
bindings, with surface-specific None/auto behavior retained. All scalar RGB
setters include legacy triples and strict optional-hash ASCII hex validation,
including parameter-qualified surrogate and channel overflow errors before
mutation. Unconditional live font formats use existing owner paths and move
direct scalar reads to font.color.rgb. Word theme edits retain cached RGB and
tint/shade, while native tint arithmetic stays unchanged. PowerPoint same-kind
edits preserve transforms and exact foreign/unmodeled children, and kind
changes retain only shadow alpha. Existing enum/export files expose the
supported MSO_THEME_COLOR slots. Finite brightness edits change only existing
luminance transforms and refuse booleans and nonfinite values atomically.
Installed regressions, deterministic theme/brightness pixels, strict typing,
stubs and package/source-preservation riders distinguish API acceptance from
rendering or compatibility evidence. Later parity stories reuse this result.

### F-X198, Preserve alternate text-box copies during scoped replacement (M)

Resolve `hadim`'s [Issue301](https://github.com/tensorbee/rdocx/issues/301), a
follow-up to Issue285. Positive scoped replacement of a checked text-box
paragraph refuses with an explicit error when another qualified Choice or
Fallback text-box representation shares its physical MC AlternateContent.
The same guard covers a selected table cell inside the text box and leaves
the complete live document and package unchanged. Count mismatches and zero
no-ops retain their existing outcomes before this guard.
Preserve run formatting, sibling text boxes, comments, namespace contexts and
opaque producer XML. Match physical alternate ownership rather than identical
text. Existing document-wide replacement semantics remain unchanged.

**Depends on**: F-X187.
**Test gate**: regression.
`scoped_text_box_replacement_refuses_alternate_copies_atomically` proves the
Issue301 Choice-only failure at the claimed Base, then atomic refusal with
complete byte preservation. Installed Python maps
refusal to RdocxError. Controls distinguish single-representation boxes,
multiple identical boxes, stale locations, count mismatch and unequal or
ambiguous alternates through save/reopen.

### F-X199, Model namespace-qualified empty table cell margins (S)

Adopt `pedroassumpcao`'s independent
[PR302](https://github.com/tensorbee/rdocx/pull/302) at immutable head
`7ac2851cb61b5e7b26427004517cc0cd4db0dbd9`, based on
`733ce6f7b2b816c5760692abf5c99d68b29ae897`. Qualified self-closing `w:tcMar`
with no attributes except namespace declarations must model the same empty
`CT_TblCellMar` as the expanded empty form. Ordinary cells therefore expose
`has_cell_margins()` without falsely reporting unmodeled properties.
Preserve real unknown attributes and foreign or rebound namespace lookalikes
as exact raw XML. Keep existing absolute cell-property schema slots and all
unrelated source content. No public API, dependency or release version change
is required. Assess updated contributor heads before implementing uncovered
criteria and record actual local evidence rather than contributor test claims.
Locally measured package archive records accompany the changed native source.
Preserve actual-host provenance, inventory counts, compression tolerances and
historical performance observations instead of copying foreign-host records.

**Depends on**: none.
**Test gate**: regression. Both self-closing spellings, expanded empty margins,
Word aliases and default namespaces produce the same typed value. Foreign,
rebound and genuinely attributed controls retain exact bytes. Existing typed
mutation keeps unsupported siblings in schema order. A source-built 7-by-3
table retains all 21 cell texts and correct reader facts before and after
save/reopen. Use existing unit and integration entrypoints and require the
deterministic hash harness to remain unchanged. Issues264 and281 stay excluded.

### F-X200, Complete Python Word formatting and list parity (L)

Resolve [Issue303](https://github.com/tensorbee/rdocx/issues/303) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Expose superscript, subscript, caps, spacing, strike, hidden, language/script fonts and RTL. Add paragraph tabs, borders, shading, outline and bidi, extended style properties, document defaults, list helpers, restart/start overrides and checklists. Reuse F-X197 for both bindings' colour contract. Plain instances share abstract-definition counts, while explicit override consumption and rearming retain concrete instance ownership. Opted-in section events defer base-start initialization without consuming an unused override on the first item. EPUB uses the selected native table context for counters, while ODT reports its bounded shared-continuation loss. Preserve concrete source identity, explicit restarts, bounded supported constructs and authoritative navigation results.

**Depends on**: F-X197.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue303 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X201, Complete Python section stories and page authoring (L)

Resolve [Issue304](https://github.com/tensorbee/rdocx/issues/304) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Repair inherited section geometry, orientation dimension swaps and incomplete pgSz validation. Enable even-page stories coherently. Add live header/footer paragraphs, tables, pictures and fields, first-page/link controls, footnote/endnote creation, deletion and numbering, run page/column breaks, page colour, watermarks and borders. Accept documented Length values without silent unit confusion.

Live section handles coexist with immutable snapshots. The six story slots
resolve ordinary rich elements in their physical owners, with checked styles,
fields, pictures, nested tables and note references. New section geometry does
not clone related-story references. DrawingML text and image watermarks retain
real text or image source and checked rotation and alpha. Source-qualified
replacement refuses changed opaque owners. Styled note references shrink and
raise once, preserving explicit position shifts. Native typeface substitution
and print-background preferences are reported separately from deterministic
repository paint.

The new DrawingML text watermark route admits only opacity 0.5 and refuses other requested values atomically. F-X219 retains the unfinished full-range DrawingML text opacity work now scheduled in S98 after the original user-approved S96 carry. Legacy VML text opacity and DrawingML image opacity remain unchanged.

**Depends on**: F-X206.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue304 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X202, Complete Word tables pictures and internal links (L)

Resolve [Issue305](https://github.com/tensorbee/rdocx/issues/305) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Reuse signed removable table indentation from F-X194. Add empty rows, nested tables, table look/style/layout, merge-aware column insertion/deletion and cell splitting. Expose picture options, run/cell pictures, identity-specific resize and picture inventory. Add bookmark/heading hyperlinks and tooltips. Complete synchronized scoped alternate text-box edits beyond F-X198's safe refusal, preserving physical owner identity and atomicity.

Checked table operations use logical grid topology and selected reference-column formatting. Picture is a factory-created physical snapshot, and Python relationship-wide sizing is explicitly named set_picture_sizes_for_relationship. Internal links validate the existing bookmark grammar and support checked related-story source paragraphs. Paired scoped edits publish every admitted physical counterpart together or refuse atomically.

**Depends on**: F-X194, F-X198, F-X206.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue305 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X203, Complete Python Word document operations and validation (L)

Resolve [Issue306](https://github.com/tensorbee/rdocx/issues/306) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Expose HTML, Markdown, all-story text, ODT/RTF/EPUB export and word/character/page counts. Add existing template rendering, document/fragment import with style policy, typed custom/app properties, validation parity, LaTeX/MathML equations, content-control list/value updates and track-revisions settings. Reuse visible paragraph validation from F-X196 and section validation from F-X201. Content-control creation and tracked-edit authoring retain their later roadmap contracts.

Also resolve [Issue332](https://github.com/tensorbee/rdocx/issues/332) in the
shared native Markdown emitter, with whitespace outside emphasis, adjacent
equivalent spans coalesced and a separator between lists and tables. Resolve
[Issue334](https://github.com/tensorbee/rdocx/issues/334) by giving newly
authored WordCompatible profiles exactly one compatibilityMode15 setting,
while preserving Minimal and imported settings. Separate reviewed behavior
commits and measured output deltas cover both native provider corrections.

Also resolve [Issue339](https://github.com/tensorbee/rdocx/issues/339) with
[PR342](https://github.com/tensorbee/rdocx/pull/342), keeping qualified run
tabs/breaks and accepted field visibility consistent across all native, CLI
and Python story text views. Staged replacement edits visible owned controls
without changing tab stops, instructions, hidden owners or opaque source.
Namespace aliases/defaults/rebinding, simple/complex field break parity,
all-story ownership, repeat/reopen and late atomic refusal are required.
Complete cross-paragraph field graphs remain coordinated with F-X211.

Qualified checkbox updates admit default or aliased Word2010 roots, preflight the selected XML-safe scalar and font, preserve foreign source and publish all selected controls atomically. Typed export diagnostics, immutable equation snapshots, live metadata and checked physical fragment boundaries share the native operation providers. Strict installed typing and complete provenance qualify the public surfaces.

**Depends on**: F-X196, F-X201, F-X212.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue306 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X204, Complete shared CLI one-shot editing commands (L)

[Issue307](https://github.com/tensorbee/rdocx/issues/307) and its PR330
contribution share native/Python/CLI ordered counts, atomic refusal and one-shot
publication. Word exposes maps, regex, combined fields, physical image
extraction/inventory, typed metadata, tag fill and deterministic/body counts.
PowerPoint exposes slide lifecycle, exact notes text/file updates, core
metadata, maps and the common read-only fit provider with typed snapshots.
Footer/theme commands retain F-X208 ownership.

**Depends on**: F-X202, F-X203, F-X205, F-X207.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue307 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X205, Complete Python PowerPoint text fill and export parity (L)

Resolve [Issue308](https://github.com/tensorbee/rdocx/issues/308) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Reuse cell text frames and colour formats. Add run baseline/spacing/language/script fonts, numbered bullets and bullet attributes, gradient/pattern/picture/transparency fills, background fills, table style/banding, core/app/custom properties, sections, media lifecycle, JPEG/TIFF, notes/handout PDF, PDF/A and ODP. Repair inherited template metadata and 16:9 placeholder geometry with reviewed deterministic behavior deltas.

**Depends on**: F-X194, F-X195, F-X197, F-X206.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue308 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. The regression gate also checks all54 source-derived pattern
programs, conditional PPT gradient interpolation, multilingual physical font
slots and the bounded colour-grid cache budget. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

Contribution PR318 at b1aa842bf531a65c70a49d4ef282033d26e64020 closes
Issue308 and belongs to this existing owner. Reconcile its full source with
this complete acceptance, retaining omitted App/custom properties and shared
prerequisite ownership. Contributor gate and template artifact values require
independent local derivation.

### F-X206, Complete PowerPoint shape authoring and live handles (L)

Resolve [Issue309](https://github.com/tensorbee/rdocx/issues/309) and
[Issue323](https://github.com/tensorbee/rdocx/issues/323) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Resolve the rotated/flipped/scaled connector-group
render defect without dropping resolvable siblings, with owner diagnostics.
Expose alt text/title/decorative, flips, group/ungroup preserving rotated/flipped positions, glued connector endpoints, cropped picture placeholder insertion, cell vertical anchor, internal run slide links, highlight/small caps, alignment/distribution and native placeholder names. Keep layout handles across slide insertion and slide handles across title text changes without weakening stale-path safety.

Native and Python operations use staged owner-scoped publication, prepared
returned-handle capture and first causal retirement. Word retains conservative
structural invalidation, while presentation items survive unrelated edits and
supported relocation. Cell anchoring owns tcPr independently of bodyPr. Run
slide links retain qualified source namespaces and prune only unused owned
relationships. Picture placeholder replacement preserves physical identity,
source metadata, rectangle geometry and crop. Grouping retains opaque member
order and promoted namespace closure, with atomic unsupported-composition
refusals. Caps/highlight use the selected physical font and owned colour source.
Recoverable member diagnostics count actual fallback or omission without
silently dropping siblings or hiding fatal setup failures.

**Depends on**: F-X194, F-X197.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue309 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X207, Persist PowerPoint autofit and implement fit_text (L)

Resolve [Issue310](https://github.com/tensorbee/rdocx/issues/310) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Persist computed normAutofit fontScale and line-space reduction, resize spAutoFit frames while keeping the anchor, implement fit_text and deck-wide refresh_autofit. Measure with the same caller fonts as rendering and expose fallback diagnostics. Compare stored values with pinned PowerPoint ground truth and verify rendered containment. F-X204 consumes this API for CLI fit.

Explicit refresh inventories slide, layout, master and notes owners and
intersects shared-owner contexts before staged publication. Normal fitting
requires separately reshaped Q and persisted S trials. Uniform fit uses the
largest permitted integer size without resetting unspecified run families.
Stored fractional rendering remains read-only. Shape-to-fit table cells and
unrepresentable anchors refuse atomically rather than claiming partial fit.

**Depends on**: F-X194, F-X195.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue310 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X208, Author PowerPoint footer fields themes and transitions (L)

Resolve [Issue311](https://github.com/tensorbee/rdocx/issues/311) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Implement layout/master hf-aware slide-owned footer/date/number placeholder creation, deck/per-slide controls, skip-title and add_field. Expose placeholder format and theme colours/fonts. Render uncached slidenum and firstSlideNum with separate displayed and physical numbering. Legal hf-only policy remains latent until slide-owned fields exist. Author the requested common transitions with direction, duration and advance controls, lifting the specific existing non-goal. Add footer CLI and inspection using Python-visible operations.

**Depends on**: F-X197, F-X206.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue311 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X209, Edit PowerPoint masters layouts and themes (L)

Resolve [Issue312](https://github.com/tensorbee/rdocx/issues/312) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Address shape edits and image relationships by owning slide/layout/master part. Add persisted master/layout backgrounds and shapes, master text styles, layout rename/duplicate/delete with usage refusal, showMasterSp, theme colours/fonts and apply_theme from deck, potx or thmx with optional masters/layouts. Verify logo inheritance, branded backgrounds and theme application without losing slide content. Selected import preserves destination theme root source and direct slide content, validates qualified opaque references and transports actual namespace/MC closure. Optional master import allocates fresh part/ID scopes, while explicit layout mapping retargets only the owning slide relationship.

**Depends on**: F-X205, F-X206, F-X208.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue312 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X210, Extend metric-compatible fonts and bullet fallbacks (L)

Resolve [Issue313](https://github.com/tensorbee/rdocx/issues/313) in S91,
part of [Issue314](https://github.com/tensorbee/rdocx/issues/314)'s everyday
coverage checklist. Reuse caller-font and real variable-instance work from
F-X195 for Issues296 and297. Map Georgia to original variable Gelasio and
Segoe UI to Selawik with explicit bounded advance, kerning and line-metric
limitations. Selawik italic uses declared synthesis, and Bold metric parity
is not implied. Aptos approximation diagnostics follow actual contributing
requests and physically authenticated selected faces, including warm replay.
Owned layout, byte-render companions, frame records, CLI and fit consumers
use the same successful output facts. Qualified Symbol/Wingdings fallbacks
preserve original source codes and use monochrome semantic Unicode coverage,
without original-outline parity or arbitrary private-use rewriting.
The exact approved3968-byte symbols subset includes U2610 and U2611.
Caller font directories are the supported optional-font route. No optional
font-pack crate, distributed pack, automatic fetch or runtime dependency is
introduced.

**Depends on**: F-X195.
**Test gate**: regression. Reproduce each missing operation or reported defect
through the installed public API on the claimed Base before implementation.
Exercise all Issue313 acceptance criteria through save/reopen, native
rendering where applicable, malformed/atomic refusal controls and preserved
unmodelled XML. Python surfaces require accurate stubs, strict typing and
stubtest. CLI surfaces require help/reference and existing integration tests.
Use deterministic fonts and pinned native oracles for output claims. Record
Google Docs/Slides import compatibility as best effort, never a CI guarantee.
Reuse prerequisite evidence for shared fixes and test the integrated behavior.
Require scoped verification and independent zero-defect, zero-smell review.

### F-X191, Ignore namespace declarations in numbering reader completeness (S)

Resolve the native reader false positive contributed by Pedro Assumpcao
(`pedroassumpcao`) through [PR293](https://github.com/tensorbee/rdocx/pull/293)
at head `fd112a7ac3333709f62746b7065c5cc1b4be0eed`. A namespace declaration
alone on a numbering instance, abstract definition or level does not make
`Document::numbering_level().has_unmodeled_properties` true. Retain all
namespace declarations and producer XML, while actual unknown attributes and
extra children remain flagged. Preserve the existing narrower reader fact
and richer numbering authoring completeness contracts. Adopt only the native
reader correction and focused tests, not contributor archive sizes or the
archive-platform override. Run in an exclusive implementation wave before
F-X189 release preparation. Issue264 remains excluded and Issue281 remains
open roadmap work.
**Depends on**: none.
**Test gate**: regression.
`numbering_level_namespace_declarations_do_not_report_unmodeled_properties`
fails through the existing native API on the exact claimed Base, then passes
all twelve owner/declaration combinations, retained namespace and producer
payload checks, real foreign attributes, rebound aliases, raw children and
save/reopen controls. All49 hash entries remain unchanged. Record actual
current package evidence and contributor provenance before completion.

### F-X189, Prepare Word 0.16.0 and PowerPoint 0.14.0 families (L)

Prepare the exact seven-package Word family and Python distribution at 0.16.0,
and the exact fifteen-package shared OOXML and PowerPoint family and Python
distribution at 0.14.0. The user approved both new versions and workflow
records after confirming that the previous unified releases are published.
Update current version carriers, internal pins, lockfile, policy assertions,
README examples and measured archives without altering historical tags or
release records. Derive separate reviewed family notes and authenticated
contributor inventories from the complete range since each previous tag,
including S89 and S90 accepted work. Exclude Issue264 and retain Issue281's
open roadmap disposition. The prepared source preserves unpublished binding and
WASM carriers, external tool versions and immutable historical release records.
The local worker completion proves preparation only. Final integrated full
verification, sprint review and hosted build-only artifacts precede close,
with separate family publication approvals afterward. Require current local package and binding evidence,
then the final reviewed sprint's hosted build-only rehearsal before close.
Publication follows sprint close through separate exact-SHA approvals under
`/release`, shared family first where the Word dependency graph requires it.
**Depends on**: F-278, F-279, F-280, F-281, F-282, F-X179, F-X180, F-X181, F-X182, F-X183, F-X184, F-X185, F-X186, F-X187, F-X188, F-X190, F-X191.
**Test gate**: release regression. `s90_release_families_match_reviewed_versions`
in the existing workflow test module checks both exact family allowlists,
version carriers and rendered notes. Actual22-package patched publication dry
runs, archive bounds and clean installed Python3.9/3.12 runtime and typing
evidence pass without publication. The integrated hash baseline is unchanged
by release preparation. Hosted evidence binds the final reviewed sprint SHA.

### F-X182, Honor direct table alignment (M)

Resolve every reported variant of [Issue 277](https://github.com/tensorbee/rdocx/issues/277), reported by `hadim`, through PR 279 at `93749ddce266103b62a61e78c3b8a6b692a94646`.
Direct table jc overrides table-style alignment, including left, center,
right, start and end. Retain authored indentation rules and preserve source
XML. Nested and bidi controls distinguish existing placement semantics.
**Depends on**: none.
**Test gate**: regression. Source-built reported variants match fresh pinned
Word relative positions and survive unrelated edits and reopen. Prove the
failure before implementation. Deterministic PNG/PDF hash deltas are declared,
separately attributed and independently reviewed, with unchanged source XML.
Record reporter provenance and full acceptance for sprint close.

### F-X183, Correct table margins and legacy positioning (M)

Resolve every reported variant of [Issue 276](https://github.com/tensorbee/rdocx/issues/276) and [Issue 278](https://github.com/tensorbee/rdocx/issues/278), reported by `hadim`, through PR 280 at `795b29d78d5c2ca49c1b414c9818de201fe36b4e`, stacked on PR 279.
Absent side margins resolve to zero consistently in measurement and layout.
Retain authored and inherited margins, with the separately qualified built-in
TableNormal native behavior and effects-only default fallback.
Pre-15 and absent compatibility modes position eligible top-level tables
using the measured first-cell left margin and resolved indent context. Preserve modern,
centered and nested placement, and distinguish bidi and floating contexts.
**Depends on**: F-X182.
**Test gate**: regression. Source-built reported variants match fresh pinned
Word relative positions and survive unrelated edits and reopen. Prove the
failure before implementation. Deterministic PNG/PDF hash deltas are declared,
separately attributed and independently reviewed, with unchanged source XML.
Record reporter provenance and full acceptance for sprint close.

### F-X180, Correct cell nil and none border precedence (S)

Adopt [PR 274](https://github.com/tensorbee/rdocx/pull/274) from `hadim`,
reviewed at `59d9de9426233da389369ee10f5c8bd32ca8f75c`, against every
criterion of [Issue 272](https://github.com/tensorbee/rdocx/issues/272).
A cell `nil` removes its edge on the outer boundary as on the interior.
A cell `none` is unset for layout and inherits the corresponding table edge.
Both direct table borders and style-derived borders follow the same rule in
painting and border bands. Preserve the distinct source tokens on save.
**Depends on**: none.
**Test gate**: regression. The seven reported direct and Table Grid variants,
including a single unmodified corner and first-row top suppression, match
pinned Word line sets. The contributed gate fails before the correction.
The deterministic dense-form golden changes only the identified top segment,
with unchanged row geometry. All 49 hash-harness entries remain unchanged.
Record contributor provenance and complete issue acceptance for sprint close.

### F-X181, Ignore page and column breaks inside table cells (S)

Adopt [PR 275](https://github.com/tensorbee/rdocx/pull/275) from `hadim`,
reviewed at `5ea2f5710d4bb13d0182bd74417fa5ca58522984`, against every
criterion of [Issue 273](https://github.com/tensorbee/rdocx/issues/273).
Inside table cells, including nested tables, page and column run breaks
produce no page transition, line break or space. Text remains on one line and
row height matches the input without the break. Ordinary line breaks remain
line breaks and body page breaks still advance the page. Field-result control
characters interpreted as page and column breaks follow the same cell rule.
**Depends on**: none.
**Test gate**: regression. Leading and inline page and column breaks in ordinary
and nested cells match pinned Word text positions and no-break row geometry.
Line-break and body page-break controls retain their existing behavior.
The contributed gate fails before the correction. Source XML survives reopen,
and all 49 hash-harness entries remain unchanged. Record contributor provenance
and complete issue acceptance for sprint close.

### F-X179, Correct multi-paragraph comment threads from PR 271 (S)

Adopt [PR 271](https://github.com/tensorbee/rdocx/pull/271) from `hadim`,
reviewed at `e22641a8f20a1c31d905a8c1b83171f548d2e230`, against every
acceptance criterion of [Issue 270](https://github.com/tensorbee/rdocx/issues/270).
Read and write thread and resolved metadata using each comment's last
paragraph id. Author one paragraph per newline through every existing comment
entry point. Preserve legacy parent links, unrelated producer XML and fragment
thread closure. Leave Issue 264 and F-X178 untouched.
**Depends on**: none.
**Test gate**: regression. Last-paragraph reading, reply and resolution XML,
multiline save and reopen, single-paragraph controls, missing ids, fragment
closure and removal all pass. The contributed gate fails before the fix.
All 49 hash entries remain unchanged. Record acceptance and contributor
provenance for GitHub reconciliation through sprint close.

### F-X178, Clearable direct run formatting setters (S)

Resolve Issue364 using the existing clearable-setter contract. The delivered
formatting providers cover double strike, caps, vertical alignment, spacing
and hidden. Implement only the remaining width-scale and position `_value`
setters taking `Option`, matching `set_bold_value`. Preserve the one-way
setters. Rust facade scope and any binding parity are confirmed in design.
Issue264 remains an open requirements discussion, with no broader Styles
implementation implied by this repair.
**Depends on**: F-X200.
**Test gate**: unit and round trip. Each property writes its explicit value,
removes only its own element on `None`, and adds no `w:rPr` when absent.
Retain explicit-false controls for boolean setters. Values survive save and
reopen, with all existing deterministic hashes unchanged.

### F-X168, Current issue and contribution closure evidence (M)

Reconcile the eight open issues and 31 open PRs from the 2 October GitHub snapshot
and Issue 264 plus draft PR 265 from the 3 October refresh
against their full criteria and the integrated S76 to S84 result. Record
contributor comments and manual viewer checks. After the reviewed sprint is
merged to main through `/close-sprint`, close every issue and PR only with
its complete acceptance and contributor disposition evidence. Issue 158
closes last after every child criterion passes.
**Depends on**: F-X149, F-X158, F-X160 through F-X167, F-X169, F-X170.
**Test gate**: integration. Every issue criterion and PR disposition has a
prepared evidence link or an explicit final integrated check. Complete the
remaining full verification and sprint review at the S84 boundary. After
`/close-sprint` merges that result to main, close all nine issues and 32 PRs
with links to the verified evidence.
PR 265's section-selection code overlaps the completed F-X162. Adopt its
distinct multipage mixed-orientation table and content-control regression on
the integrated implementation before recording a superseded PR disposition.

### F-X169, Reconcile live contributions and open issue contracts (L)

Audit the 31 open PRs and eight open issues from the 2 October GitHub intake
against integrated main. Classify each older PR as patch-equivalent, requiring
incremental replay, or needing a contributor rebase. Record the current stack
order, overlapping files, review state, latest head CI and issue coverage.
Keep open issues and PRs open until their full integrated acceptance evidence
and contributor disposition are recorded. This is a planning and evidence
story, with no source changes or baseline update.
**Depends on**: F-X160.
**Test gate**: integration. Every open PR and issue has a current disposition,
every stacked dependency has an order, and the sprint plan agrees with the
backlog and live GitHub inventory.

### F-X021, The hash harness should cover PDF output (M)
The output-stability harness records `page1.png` and three `word/*.xml` parts
for each of the seven samples, and no PDF. PDF is a first-class output of this
workspace, produced by a different code path from the PNG: `oxml-pdf` writes
glyph positions, embedded font subsets and compressed streams, none of which the
rasterised PNG exercises. That path can therefore drift with no gate noticing.

F-X020 demonstrated the gap rather than theorised it. A routine
semver-compatible dependency refresh changed all seven sample PDFs while every
PNG stayed byte-identical and the harness reported 28 of 28. The change was
benign, and it was found by hand rather than by the gate that exists to find it.

Recording a PDF byte hash directly would be brittle, since a PDF carries a
creation date and object ordering that need not be stable. The story therefore
decides what a stable PDF fingerprint is, likely extracted text plus page
geometry plus glyph positions, before recording one.
**Depends on**: none.
**Test gate**: regression. A deliberate change to the PDF writer moves the new
entries and leaves the PNG entries untouched, and a re-run with no change
reproduces every entry exactly.

### F-X020, Refresh the dependency lockfile (S)
Every semver-compatible dependency update outstanding at the start of the sprint
is taken, and its effect on rendered output is measured rather than assumed.
Sixteen updates are pending and none is a security fix: `cargo audit` reports
zero vulnerabilities across 152 dependencies and `cargo deny check advisories`
passes. Two of the sixteen, `font-types` and `zune-core`, sit in the font and
image decoding path, which is why the hash harness is this story's real gate
rather than a formality.

The `ttf-parser` unmaintained advisory, RUSTSEC-2026-0192, is unaffected. It is
allowlisted in `deny.toml` with a documented reason, and clearing it needs the
`fontdb` to `fontique` swap rather than a lockfile refresh.
**Test gate**: the full workspace suite and the hash harness. A delta is
expected only if a font or image dependency moved rendering, and any delta names
the dependency that caused it and is reviewed before the baseline is re-recorded.
A delta traced to no dependency in the rendering path blocks the story.

### F-X019, Paragraph-relative drawings in later blocks should wrap (M)
Text flows around a wrapping drawing anchored to a later paragraph even when
that drawing is positioned relative to its own paragraph rather than to the
page or a margin. F-X016 looks ahead only for absolutely framed drawings,
because a paragraph-relative one has no position until its own paragraph is
placed, and resolving that needs the paginator to run twice. No sample or corpus
document hits the gap today.
**Depends on**: F-X016.
**Test gate**: regression. A paragraph-relative wrapping drawing anchored to a
later paragraph pushes earlier text aside, and a document with no such drawing
paginates in a single pass exactly as before.

### F-X018, Unknown enumerated values should not fail a document open (M)
Nine value parsers in `rdocx-oxml/src/shared.rs` and `styles.rs` return an error
for any string they do not enumerate, and several are reached through `?` from
paragraph, table and numbering property parsing. A document using a
spec-valid value the model does not yet list therefore fails to open rather
than losing one property. F-X014 fixes the three kashida values because they
were reachable from a real contribution. This story decides the general rule,
which is that an unmodelled enumerated value falls back to the element's default
and the surrounding properties survive.
**Depends on**: F-X014.
**Test gate**: regression. A document carrying an unmodelled value for each of
the nine enumerations opens, keeps every sibling property, and renders with the
default for the unmodelled one.

### F-X015, Anchored drawing wrap and alignment model (M)
`CT_Anchor` carries the wrap mode, the four text-distance attributes and the
optional horizontal and vertical alignment children, and `AnchoredDrawing`
carries them into the layout model. `wrapSquare` and `wrapTopAndBottom` parse to
distinct wrap modes rather than collapsing into `None`, which is what the
currently parsed-but-unread `WrapType` does today. `distT`, `distB`, `distL` and
`distR` round-trip through the serialiser. A `positionH` or `positionV` that
names an alignment records that alignment alongside its offset. Placement and
rendering are deliberately unchanged, so this story adds only the model surface
that F-X016 consumes.
**Test gate**: round-trip. Wrap modes, the four distances and both alignment
axes survive a parse and serialise cycle, including a prefix-tolerant read. The
hash harness is unchanged, which is what proves the story is model-only.

### F-X016, Floating drawing placement and text wrapping (L)
An anchored drawing whose position names an alignment resolves against its
`relativeFrom` frame by that alignment rather than by a zero offset. Body text
flows around a `wrapSquare` drawing, reserving the frame width plus the relevant
text distance on the lines the drawing spans, and clears a `wrapTopAndBottom`
drawing by starting below it. Reserved width is taken from the drawing frame and
its `distL` or `distR`, not from a scan of image pixels, since pixel extents
describe `wrapTight` and `wrapThrough` rather than `wrapSquare`. Line breaking
gains a per-line width reservation that the paginator can vary once it knows
where on the page the paragraph landed.
**Depends on**: F-X015.
**Test gate**: golden. A paragraph beside a left-aligned square-wrapped drawing
breaks its lines at the reserved width, a right-aligned one reserves from the
line end, and a top-and-bottom drawing pushes the following text below its
bottom edge plus `distB`. Unwrapped and `wrapNone` drawings lay out exactly as
before, which the hash harness proves by leaving every baseline without a
wrapped drawing unchanged.

### F-X211, Complete Python script compatibility and safe XML access (L)

Resolve the uncovered acceptance in [Issue316](https://github.com/tensorbee/rdocx/issues/316),
part of Issue314. Keep the existing17 python-docx and7 python-pptx examples
as a subset of roughly100 idiomatic short first-draft scripts spanning
Issues303 through313. Each checks the saved file. Print per-area unmodified
correct-file success rates in CI and record the final integrated rates.
Failures receive a fix under their owning story or a deliberate divergence
whose error names the supported call.

Expose validated element.xml read and element.replace_xml(str) for paragraphs,
runs, tables, cells, section properties, shapes, text bodies, slides, layouts
and masters. Reject wrong qualified element types and invalid ownership,
relationships or schema order atomically. Preserve opaque XML and package
relationships. Extend stable handles across supported structural edits,
including Word hyperlink insertion/removal and table grid changes in both
bindings. When survival cannot be proved, name the invalidating call in the
error. Existing F-X194 and F-X206 own their already requested stability fixes.

Audit accepted ignored options, invalid units and inconsistent OOXML. Known
caller-font, internal-link and section repairs stay with F-X195, F-X202 and
F-X201. Add a public regression for each case and actionable errors for
differently spelled upstream calls. Do not implement those fixes twice.

Also resolve [Issue331](https://github.com/tensorbee/rdocx/issues/331) at the
shared native parser boundary. Accept up to256 nested qualified Word content
controls and reject deeper inputs with a recoverable named-limit error,
including native document open and atomic raw XML replacement. Preserve
opaque foreign XML, retain existing recursion budgets across owner changes
and prove accepted traversal/save/drop on normal consumer stacks. Dedicated
257/5000-depth refusal controls do not change the frozen corpus denominator.

**Depends on**: F-X200 through F-X210 for final corpus acceptance, with
source and prerequisite capability edges refined in the detailed design.
**Test gate**: regression. Installed isolated bindings run the full corpus
and assert saved package structure. All requested element kinds round-trip
opaque content and refuse wrong types atomically. Handles survive supported
edits, unsupported invalidations name the triggering operation, and each
silent failure has a failing-before/fixed-after public control.

#### Delivery children

F-X211 retains the complete issue acceptance and closes only after F-X211a,
F-X211b, F-X211c and F-X211d are completed. Splitting does not add issue scope or waive
any existing source, preservation, parser, lifetime or corpus requirement.
The original approved parent plan remains the shared acceptance contract.

#### S91 PR317 reconciliation

PR317 at head3d056f46e97e9cc1caf2969b5096eeb6775a5fa0 closes Issues295,
298,299,300 and301. Existing F-X194/F-X196/F-X197/F-X198 remain their
implementation owners. Cell tcPr anchor is F-X206, while successful paired
replacement is reusable input for F-X202 beyond F-X198 safety.
Review its42-file diff and uncovered criteria before source reuse. Additional
native accessibility behavior is not authorized merely by its contributor
hash-baseline record. Keep independent labelled behavior review and derive
all archive/baseline evidence from actual accepted source.

### F-X211a, Word checked XML owners and resource admission (L)

Complete the five Word element cases through concrete checked native physical
owners and installed public bindings. Require actual source namespace closure,
introduced property and identifier admission, owner-local causal retirement,
exact opaque preservation and equivalent no-ops. Resource admission and
ordinary-stack parse, clone, save, reopen and drop retain inclusive
weighted256 control/text-box and32 table boundaries. The Word implementation
supplies eager current-owner publication, property-child retirement and
source-qualified regression assertions. Final cross-family scripts and shared
compatibility residuals remain owned by F-X211c.

**Depends on**: F-X200, F-X201, F-X202, F-X206, F-X210, F-X222.
**Test gate**: regression. Five installed Word XML cases, bounded native and
consumer-stack parser lifecycle, saved/reopened ownership and atomic failure
controls, plus all applicable original parent scoped risk riders.

### F-X211b, PowerPoint checked XML owners (L)

Nine PowerPoint element cases use actual slide, layout and master owners,
checked native operations and installed public bindings. The five captured
shape variants retain their kind. Qualified roots, schema order and actual
relationship target/type checks precede candidate publication. Standalone XML
retains inherited opaque namespaces. Selected logical owners survive checked
replacement while unsafe property, body, paragraph and run descendants retire.
No-op, refusal and allocation failure preserve package bytes and live tokens.
Shared parser and token prerequisites are serialized by root. Final cross-family
script and causal-error reconciliation remains owned by F-X211c.

**Depends on**: F-X194, F-X200, F-X206, F-X209, F-X210, F-X214.
**Test gate**: regression. Nine installed PowerPoint XML cases, shape variants,
qualified ownership, saved/reopened bytes and atomic relationship/type/order
refusal, plus all applicable original parent scoped risk riders.

### F-X211c, Integrated Python compatibility and causal errors (L)

Compose both checked XML families with completed providers. Resolve shared
lifetime, invalid units, ignored options and actionable error residuals once
under their actual owner. Preserve the complete frozen100 bodies and their
Word49/PowerPoint51 allocation, Base17/100 evidence and honest per-area rates.
Qualify source-built paired bindings, saved files, the complete lifetime
contract and candidate-bound CI reporting. Bounded aliases reuse checked
native providers, and causal errors retain their first public invalidator.
Unknown explicit symbol diagnostics use the actual Word and PowerPoint
accepted-layout consumers, preserving original codes and physical-face
precedence without retaining probe history.
Table rows and columns collections retain their physical ancestor authority
through supported topology edits, while affected indexed members retire with
the first public cause. Both-axis controls retain save/reopen, refusal and
parent-removal obligations.

Named reviewed non-successes remain in the denominator without correct-file
credit. No corpus rewriting or provider duplication.

**Depends on**: F-X211a, F-X211b, F-X200 through F-X210, F-X215, F-X217,
F-X218, F-X220, F-X211d for final integrated capability acceptance.
**Test gate**: regression. Installed isolated bindings execute the immutable100
and saved-file assertions, actual per-area rates and named causal invalidation
errors, with all remaining parent native, render, publication and risk gates.

### F-X211d, Repair Word compatibility owners and aliases (L)

Resolve the Word failures exposed by the actual unchanged first100 candidate:
checked table/cell/section XML retains exact opaque attributes and adjacent
children, held sections and paragraphs survive supported header/note/equation
publication, section margin aliases consume canonical units, Cell.merge uses
checked physical grid ownership with atomic refusal for removed opaque cell
properties and borders, and the requested canonical built-in table style is
defined lazily before use without changing the default factory. Transfer the four reviewed Word binding/test
carriers from c once, without reimplementing them. All original parent
acceptance, conservative Run retirement, manual/foreign refusal, atomicity,
source preservation and frozen100 identities remain unchanged.

**Depends on**: F-X211a, F-X200, F-X201, F-X202, F-X203.
**Test gate**: regression. Current Word opaque-owner and alias/survival
regressions fail on the actual retained candidate and pass after repair, with
saved/reopened/source-qualified facts and all applicable parent scoped riders.
Final complete immutable100 qualification remains owned by c.

### F-X212, Align native accessibility with visible paragraph content (M)

Native `Document::paragraph_visibility` returns final displayed-content flags
in recursive facade paragraph order from one unpublished current-source
candidate. The shared qualified scanner validates physical owners and admitted
typed slots, including the first expanded content-control slot. Native audit
uses direct body paragraphs only, resets the streak for every other body item,
and preserves metadata, image-alt and heading findings. Source construction
or classification failure adds one indeterminate Warning and skips empty
streak findings instead of fabricating flags.

Accepted pictures, drawings, objects, symbols, tabs, breaks, note references,
equations and complete displayed field results interrupt empty streaks.
Parser-owned simple and complex results preserve accepted nontext, nested
owner identity, namespace provenance and current-cache consistency. Deleted
results, opaque producer hints, foreign lookalikes, properties, loose field
markers or instructions and soft hyphens alone do not prove visibility.
CLI validation shares the classifier and recursive inventory, with explicit
complete-field display and indeterminate-error handling. All operations
preserve live source, package bytes and unmodelled XML.

**Depends on**: F-X196.
**Test gate**: regression. Public picture, namespace, field/cache, deleted,
opaque, duplicate-slot, current-edit and failure controls distinguish genuine
empty paragraphs across native and CLI consumers. Nested child mutation,
refresh and save/reopen retain owner identity and bounded graph work.
Deterministic attribution permits only the three feature_showcase XML and PDF
bytes/pages keys, with 46 unchanged entries. Audit-only reversal must recover
all original 49 entries, prove exactly the issue-count and removed-finding
XML effects, and restore accepted source. Expected deltas require independent
review and a separate labelled behavior commit. Contributor baselines are not
accepted evidence.

## Rolling contribution reconciliation, 2026-10-09

PR [319](https://github.com/tensorbee/rdocx/pull/319), immutable head
`9a36bc966d1e12e4e9e2f38b777889bd687c5dd9`, closes Issue313 and attaches
to existing F-X210. Caller fonts and variable instances remain F-X195 and
PR315. Its proposed four modified static Gelasio faces and new Noto Sans
Symbols2 subset differ from the explicitly approved eight-file inventory.
Capture and assess their sources, license terms, metrics, glyph mappings,
italic selection and diagnostic lifecycle before selecting any contribution.
The existing asset approval does not authorize that different inventory.
Its twelve claimed rendering deltas are contribution evidence, not accepted
local baselines. Root retains exclusive baseline ownership.

PR [320](https://github.com/tensorbee/rdocx/pull/320), immutable head
`252acccbba9adab9264cfb109c4b1d720ed79d4a`, closes Issue303 and attaches
to existing F-X200. Its colours remain F-X197. Assess its tri-state setters,
tab ownership and enums, defaults preservation, list helpers and installed
binding behavior against the full approved contract before reuse. Its body
explicitly leaves native shared numbering counters unchanged.

Issue [321](https://github.com/tensorbee/rdocx/issues/321) therefore maps to
that same F-X200, which already owns the native shared-definition counter
change and its authenticated Word evidence. It adds no duplicate F-ID or
second implementation. Its explicit two-instance continuation, first-use
startOverride and sub-level restart criteria are covered by the same native
state contract and require local regressions and explained sample changes.
F-283 retains its F-X200 completion prerequisite. Do not close Issue303 or
Issue321 based only on the contributor's narrower binding patch.

The refreshed intake contains seven open PRs and23 open issues. Issues264
and281 remain excluded. Contributor test, measurement and baseline claims
never substitute for local verification.

## Next rolling refresh, 2026-10-09

PR [322](https://github.com/tensorbee/rdocx/pull/322), immutable head
`6b5e2e352047357299d1f366f807f4661037430e`, closes Issue305 and attaches
to existing F-X202. Its table grid edits, pictures, links and handle policy
are contribution input for the full existing contract. Its explicit nested
table exclusions and main-document-only picture inventory do not satisfy
all qualified owners automatically. Indentation remains F-X194. Successful
alternate-copy text replacement remains F-X202 after F-X198 safety, reusing
the separately captured PR317 slice. Stable-handle residuals remain F-X211.
No new F-ID is created, and contributor gates are not local evidence.

PR318 advanced from the retained assessed b1aa842 head to
`bfdd6ab732e8c0003ea8c6110e3c891fb098d5b3`. The immutable comparison shows
one new commit changing only the PowerPoint chart artifact SHA in a Word
native test. Complete changed source, Git OID/bytes/SHA and refreshed reviews
are captured independently. Keep the prior assessment, with this addition
attached to F-X205's template acceptance. Measure the actual locally authored
artifact after adopting approved geometry. Do not copy the contributor's
expected hash or rewrite immutable cross-viewer acceptance evidence.

This refresh has eight open PRs and23 open issues, with all scope linked to
existing owners. Issues264 and281 remain excluded.

### F-X213, Preserve TOC entry boundary paragraphs during rebuild

**Size**: M
**Depends on**: F-279
**Test gate**: regression

Resolve Issue340 and assess PR341 at immutable15e799e891cfab6a65d988f621e5081bfc27c1a6.
When a TOC begin/separator occupies its first entry and its end occupies its
last entry, rebuild keeps those markers in the actual first and last rebuilt
entries without adding empty entry paragraphs. Support a single-paragraph
TOC, changed entry counts and zero-entry rebuild followed by another rebuild.
Retain existing marker-only boundary behavior, source instruction, field
identity, paragraph properties, namespace closures and unrelated opaque XML.
Apply namespace URI ownership, preserve unknown paragraph attributes/children
verbatim, and remove started/ended style nodes as complete owned elements.
One-to-many expansion retains paragraph-local namespace closure on the moved
end run and every suffix sibling. Multi-paragraph empty rebuild removes all
obsolete entry styles while retaining protected owner properties, then accepts
a later positive rebuild. Duplicate or malformed owned properties refuse
without mutation. A rejected unsupported boundary publishes no package or
handle change.
Reuse the contributor source and tests once, correcting preservation gaps.
Native and installed Python regressions, repeated save/reopen, deterministic
layout and attributed hash checks must pass. This bounded repair remains S91
while the complete F-283 navigation contract stays parked at S95.2.

### F-X214, Bound recursive PowerPoint parsing and traversal

**Size**: L
**Depends on**: none
**Test gate**: regression

Resolve Issue344 with the public shared nesting limit32. The first recognized owner counts as1, depth33 is a checked refusal, and opaque XML cannot supply condition metadata or grow the selected namespace stack. One shared nesting limit covers group, AlternateContent and timing cycles. Parse each subtree once with linear work. CLI and Python return a named checked error past the limit without abort. Parsed and authored trees at the limit extract text and render safely, including layout and renderer walkers. Release tests prove each cycle at and beyond the limit, and at-limit open, extract and render complete in well under one second.

### F-X215, Create charts in Python and replace existing chart data

**Size**: L
**Depends on**: F-X197, F-X206
**Test gate**: regression

Both Python families expose native chart creation with concrete category and
shared-X data objects. Word returns a checked body paragraph and supports
same-document direct paragraph-relative insertion. PowerPoint returns the
owned graphic frame and exposes live unique-ID chart data, including nested
owners. Replacement stages caches, formulas and formula-owned workbook cells
together. Surviving series formatting, unselected workbook entries and shared
chart/workbook owners remain preserved through copy-on-write. Regression
acceptance covers bar/line authoring, replacement cardinalities, checked owner
and genuine returned-allocation refusals, exact no-op, installed typing,
qualified corpus cells, native viewer/Edit Data and data-specific renders.

### F-X216, Import complete slides and merge decks with source formatting

**Size**: L
**Depends on**: F-X206, F-X215
**Test gate**: regression

Resolve Issues347,348 and349 through checked all-slide or caller-ordered
selected import and deck merge. Preserve the destination-formatting default,
with opt-in source layout/master/theme inheritance and required global table
style and embedded-font closure. Copy chart/workbook/style/colour graphs,
embedded objects and previews, notes, classic and modern comments with full
author/person identities and required parent threads. Map selected or explicit
existing slide jumps, diagnose unmapped jumps and refuse unsupported internal
types before publication. Destination resources and source bytes remain intact.

Rust and Python expose checked insertion, optional selected section blocks,
ordered diagnostics and batch atomicity. Python preserves the old single-slide
return and adds explicit diagnostic and batch companions with all returned
objects allocated before publication. CLI import uses one-based selection and
insertion, caller order and a separate atomic output protecting both inputs.
Regression, installed, corpus and actual repair-free native open/save/reopen
plus source-format font/colour/geometry render acceptance are required. No
helper result, source hash or prospective feature state substitutes for them.

### F-X217, Import HTML and Markdown through public Word surfaces

**Size**: L
**Depends on**: F-X200, F-X203
**Test gate**: regression

Resolve Issues350 and351 captured at2026-10-11 18:56UTC. Expose existing HTML creation, file import and paragraph-relative insertion in Python, and HTML-to-DOCX CLI with diagnostics and stubs. Implement CommonMark plus GitHub tables creation in Rust, Python and CLI with headings, body paragraphs, emphasis, strong, code, strikethrough, nested numbering, tables, hyperlinks, inline pictures, quotes and code-block styles. Optional template import retains styles, headers and footers. Dropped constructs produce immutable ordered diagnostics. Paragraph-relative insertion validates physical owners and stages diagnostic allocation before publication. CLI local resources stay bounded to the input directory and output is atomic. Per-construct and supported Markdown-to-DOCX-to-Markdown controls, HTML structure/styles in installed Python and CLI, and repair-free native Word open/save/normal-reopen/render of the combined and template inputs are mandatory.

### F-X218, Export presentations as semantic HTML and Markdown

**Size**: L
**Depends on**: F-X205, F-X206
**Test gate**: regression, `fx218_installed_and_cli_semantic_exports` with native semantic/corpus controls, independent HTML/CommonMark validators, native PowerPoint source authority, fresh installed typing/stub checks and atomic publication

Resolve Issues352 and353 captured at2026-10-11 18:56UTC. Rust, Python and CLI export one title-headed section per slide in existing outline reading order. Preserve paragraph levels and emphasis, links, semantic tables, picture alt text and embedded or adjacent image assets. Include notes and skip hidden slides when requested. Markdown uses nested lists and pipe tables. HTML uses semantic sections and tables and validates as HTML. Corpus expected-output comparisons cover every construct and option. This is semantic text export, with no pixel-exact promise.

Detached results carry text, owned assets and ordered diagnostics. Python string methods reject detached image mode, while export companions and save methods preserve assets. CLI semantic flags reject rendering options. Complete-body image validation uses the crate-private native validator and bounded existing workspace decoder dependencies even without render/default-template features. Native PowerPoint source semantics and save/reopen establish source authority, while HTML and Markdown retain their own independent conformance gates.

### F-X219, Qualify full-range text watermark interoperability

**Scheduling**: S98 after the 2026-10-28 capacity reorganization. Originally carried unfinished from S95.11 to S96. No faithful shared
Word and pinned LibreOffice representation was established. Preserve the0.5-only
restriction and original failed endpoints. Reopen design before implementation.

**Size**: M
**Depends on**: F-X201
**Test gate**: regression

Restore the full zero-to-one opacity range for real DrawingML text watermark
authoring after the user-approved S91 temporary 0.5 restriction. Preserve
legacy VML behavior and qualified inactive source branches. Use a faithful
representation rather than reversing schema opacity operands, flattening
colour or converting text to pixels. Actual native Word and LibreOffice
open, render and reopen must show correct zero, intermediate and full-opacity
results. Own rendering, API getters, source ownership, strict replacement and
invalid-operation atomicity must agree. Retain S91's failed endpoint evidence
and remove the temporary named restriction only after these controls pass.


### F-X220, Embed fonts through both native and public package families

**Size**: L
**Depends on**: F-X195, F-X200, F-X205, F-X210
**Test gate**: regression

Resolve Issue360 with admitted PR361. Expose existing native Word font embedding, font records and face removal in Python/CLI, plus checked Word embedding settings. Add equivalent native/Python/CLI PowerPoint face embedding, listing/removal and flags, with actual embedded font parts/relationships and schema-ordered source-preserving font lists. Explicit caller authorization, exact license identity and valid bounded font input are required. Qualify PowerPoint-produced font format before writer bytes and real uninstalled regular/bold native PowerPoint font use. Native Word acceptance uses its documented one-font-per-name limit, retaining actual synthetic bold and native save changes as evidence. All four API face kinds and genuine own-renderer face selection remain required. F-X223 owns native Word multi-face interoperability in S95.11. Imported font graph preservation remains required independently. Unsupported compressed formats remain exact. Decode shared parts once, reject hostile lengths without huge allocation and preserve F195 caller override across render paths. Update existing consumer recipes and prepare concrete downstream skill guidance without falsely claiming an external submission. No new tracked font assets, product modules or dependencies are needed.

Current native and public contracts use checked source-owned transactions,
immutable snapshots and distinct caller/embedded tiers. Word checked listing
retains its compatible owned snapshot representation and reports actual stored
part facts through its indexed metadata reader. PowerPoint lists supported and
opaque stored formats explicitly and shares repeated decoded bodies. Python
binary inputs require real EOF. Both CLIs use licensed font commands and staged
outputs. The regression gate retains separate authentic producer/native-use,
installed, deterministic render, hostile-input and verified archive evidence.

### F-X221, Prepare Word 0.16.1 and PowerPoint 0.14.1 patch releases

**Size**: L
**Depends on**: F-X193, F-X194, F-X195, F-X196, F-X197, F-X198, F-X199, F-X200, F-X201, F-X202, F-X203, F-X204, F-X205, F-X206, F-X207, F-X208, F-X209, F-X210, F-X211, F-X212, F-X213, F-X214, F-X215, F-X216, F-X217, F-X218, F-X220, F-X222
**Test gate**: release regression

Prepare matching patch Rust/CLI/Python families after every included S91 repair completes. Update exact stable seven-package Word family to0.16.1 and exact fifteen-package shared/PowerPoint family to0.14.1, with corresponding binding projects, internal pins/lock and meaningful authenticated family release notes. Use existing unified allowlists and release pipeline. Build-only manual rehearsal emits both families across six CLI targets, while each pushed release tag selects one family. Archive validation checks ordered member count and uniqueness before constructing ZIP or TAR member dictionaries, preserving exact binary, README and licence payload checks. Qualify current-source archives below10MiB, exact installed wheel/metadata/runtime/typing on Python3.9/3.12, both WASM graphs and reviewed-SHA build-only hosted artifacts. Current registry/tag facts are checked, then rechecked at release. Shared/PowerPoint publishes before Word consumes it. Preparation does not publish. Final integrated verification/review and close precede publication from the exact reviewed main SHA under the user's explicit corrected patch publication authority. Actual roles, publisher configuration and final assets remain mandatory. Final addressed issue and PR reconciliation follows successful publication. Bibliography, navigation, full-range watermark and standalone authoring remain S92, and discussion264/281 stay open.

The complete binding CI suite prepares licensed independent font and synthetic-axis inputs from source. It retains all104 paint cases and the separate native-eight gate.

### F-X222, Restore linear Word story range inventory and empty owners

**Size**: M
**Depends on**: F-X203, F-X213
**Test gate**: regression

Resolve Issue362 and assess PR363 at immutable0b9a3cccc1b31bac7ec4108c923faa176b072c47. Build each story source once for complete range paragraph discovery, batch namespace scopes and scan each owner/control from its owned excerpt. Preserve physical identities, order, de-duplication, exact namespace-complete paragraph output, comment replay and text-box cache coordinates. Reuse one paragraph inventory for evaluator range pairing and bookmark name counts, retaining independent failure behavior. Valid self-closing story owners contribute zero paragraphs and do not break comments or page-field updates. Qualify actual large-table source-build counts, old/new semantic and byte equivalence across every story family, hostile ownership/offset refusals and measured release-mode scaling through native, installed Python and CLI field/TOC paths. Existing successful output and49 hashes remain unchanged. Empty-owner support is a separately labelled behavior repair. Complete inventory counters and evaluator reuse cover one owned source per operation. Independent range pairing and bookmark-name failures retain existing fallback behavior. Report doubled TOC costs and remaining per-bookmark and single-location scans separately from complete-inventory performance. Do not copy donor archive measurements.

### F-X223, Qualify native Word embedded multi-face interoperability

**Size**: M
**Depends on**: F-X220
**Test gate**: differential

Investigate native Word use of distinct embedded regular, bold, italic and bold-italic faces for an uninstalled family against the documented MS-OE376 one-font-per-name boundary. Preserve F-X220 API and renderer contracts. Require exact prepared font glyph and metric proof before and after native save and normal reopen, absent host installation, fresh family/key controls and recorded native normalization. A workaround must preserve document semantics and source ownership. Do not claim synthetic bold is genuine bold or rewrite family names without an approved design. This interoperability work belongs to S95.11. The bounded comparison includes proper embedBold-only with no embedRegular, authentic paired and four-role inputs, and sensitivity cases whose native outcomes are observed rather than predetermined. The historical regular-slot bold-body result does not replace the bold-slot discriminator. Genuine face attribution joins complete glyph/metric/paint evidence with actual host/source availability, source role graphs and native normalization. The concrete method is defined in HLD12. Preparation does not grant native acceptance, and a demonstrated native limit requires explicit acceptance disposition. The accepted outcome documents Word16.114 build26101123: proper Bold-only selection is genuine across initial, saved and reopened phases, while paired Italic uses Regular glyphs with synthetic slant and native save empties its Italic part. Fourteen unsupported PDF observations retain no face verdict. This completes the bounded investigation with an explicitly accepted native limitation, without claiming broader multi-face support or changing the package, API or own-renderer contracts.

### F-X224, Repair hosted release tool and stack routing

**Size**: M
**Depends on**: F-X221
**Test gate**: regression

Supply exact source-qualified HarfBuzz14.4.0 to every native consumer, retain the approved eight MiB native animation-test stack and accept exactly one annotated integer or single-fraction sprint tag at current main. Preserve all existing source, oracle, metadata, package, full-suite and publication checks. Requalify the still-unpublished Word0.16.1 and shared/PowerPoint0.14.1 families after the private parser repair, with actual clean-source builds, installed tests, complete hosted CI and exact-SHA build-only rehearsal. Do not move existing tags, reopen contributor intake or treat historical execution as current changed-source evidence.

### F-X225, Keep DrawingML text-style parsing within the default WASM stack

**Size**: M
**Depends on**: F-X214, F-X221
**Test gate**: regression
**Spec**: HLD04 package preservation, HLD06 shape tree and preservation, HLD10 WASM facade, HLD15 pinned WASM runtime.

Replace private inline list-style storage and avoid accumulated parser temporaries so the real PowerPoint WASM constructor and graph round-trip work within the existing1MiB stack. Preserve all public APIs and field types, XML namespace/opaque ownership, schema order, parser limits, deterministic outputs and versions. Prove the exact Node24.11.1 genuine Base failure and unchanged-test GREEN, a real populated caller, native getter/mutator/clone/serialization behavior, both WASM graphs and affected verified archives. No stack increase, new product file/module/dependency, new asset or public surface.


### F-X226, Release superseded relationships after checked PowerPoint XML edits

**Size**: M
**Depends on**: F-X211b, F-X215, F-X216
**Test gate**: regression
**Spec**: HLD04 checked PowerPoint XML transactions and targeted graph cleanup, HLD06 opaque preservation, HLD10 chart and owner authority.

Release only superseded known-role relationships owned by a checked XML edit after proving the complete actual owner no longer references them. Prune only targets made newly unreachable by this operation, preserving shared chart/workbook graphs, opaque references and unrelated producer orphans. Preserve debug save validation and atomic publication. Reproduce the original chart alias failure, then prove debug and optimized public behavior, exact sibling graph preservation and unchanged deterministic outputs. No new API, product file, dependency or contributor intake.

### F-X227, Execute the unchanged first-draft compatibility corpus in owner CI

**Size**: L
**Depends on**: F-X211, F-X221
**Test gate**: regression
**Spec**: HLD10 Python compatibility, HLD12 actual corpus qualification and CI, HLD15 pinned owner CI execution.

Deliver the original Issue316 and PR333 corpus-in-CI criterion, also included
by Issue314. Run all100 unchanged scripts with actual per-area correct-file
rates and complete source/artifact/input custody. Keep the17/7 documentation
subset, historical95/100, fixed denominator and five approved no-credit
non-successes distinct from new execution. Two missing native inputs need an
explicitly approved, independently qualified revision before use. A one-job
owner CI route must preserve private Office resources, require explicit
runner authority and demonstrate task-scoped admission or isolation. Retain
full hosted binding suites and debug save assertions. No new production API,
dependency, product file, font substitute or reopened contributor intake.

### F-X228, Qualify faster CI without weakening gates

**Size**: M
**Depends on**: none
**Test gate**: regression

Assess PR376 for faster CI with measured current-head results. Use dev opt-level1 for Test and MSRV and release-mode presentation corpus rendering. Narrow presentation routing only against actual Cargo and example dependency closure. Preserve debug assertions, overflow checks, existing binding opt-level2, owner-only dispatch controls, HarfBuzz14.4, animation stack setup, all suites and independent oracles. No donor timing claim counts as current evidence.


### F-X229, Repair comparison reconstruction correctness

**Size**: L
**Depends on**: F-X203
**Test gate**: regression

Resolve Issues368 through371 using the correctness subset of PR374. Support whole-paragraph hyperlink insertion and deletion, edits adjacent to tables, empty paragraph properties and trailing or line-break-only paragraph deletion. Accept-all and reject-all must reconstruct destination and source properties, relationships, tables and stories through save/reopen. Preserve opaque XML, atomic failure and bounded resources. Keep presentation grouping changes separate.


### F-X230, Group comparison changes for readable revisions

**Size**: M
**Depends on**: F-X229
**Test gate**: regression

Resolve Issue372 using the separately attributed grouping subset of PR374. Group word and character revisions for readable comparisons while preserving acceptance/rejection reconstruction. Do not cross fields, opaque XML, comments, bookmarks, hyperlinks or formatting boundaries. Review intentional output changes separately from the correctness repair and test mixed edits plus unchanged controls.


### F-X231, Preserve PowerPoint HTML import whitespace

**Size**: M
**Depends on**: F-X217
**Test gate**: regression

Assess advanced PR359 PowerPoint HTML import whitespace. Preserve spaces beside emphasis, repeated whitespace across span boundaries and NBSP through native and installed public entrypoints and save/reopen. Qualify the already equivalent Word changes with the advanced tests without duplicating implementation. Semantic HTML and Markdown export acceptance remains with F-X218. Report intentional deltas and unsupported input explicitly.


### F-X232, Render legal revisions and qualify checked XML repros

**Size**: L
**Depends on**: F-X201, F-X211a, F-X211d
**Test gate**: regression

Resolve Issue373 from PR375 uncovered slices. Add legal revision colouring and improved double underline without changing stored formatting. Include render settings in layout cache identity and keep native, CLI and Python parity. Review proposed public RevisionStyle fields and ParagraphBlock type changes for compatibility in design. Reuse existing checked-XML and namespace providers, qualifying the exact header/footer paragraph, run and ancestor-namespace repros. Require independent native paint evidence and separately attributed deterministic product and external raster deltas.


### F-X233, Repair paragraph border geometry

**Size**: M
**Depends on**: none
**Test gate**: regression

Qualify Issue366's paragraph borders using the original input when available. When it is unavailable, use the user-approved inferred15point sz120 and raw15 sz15 controls, retain the missing-original limitation and do not claim exact-original reproduction. Inspect saved w:sz and w:space and independently measure cropped geometry and paint. Existing size conversion already divides by eight, so distinguish width from clearance, clipping and joins before selecting a fix. Require save/reopen and independent native comparison. Never loosen raster thresholds to hide a delta.


### F-X234, Preserve supported LaTeX conversion structure

**Size**: M
**Depends on**: none
**Test gate**: regression

Resolve the conversion half of Issue365 for its exact supplied formula. Preserve pi, lim, to, infty and meaningful spacing with exact normalized OfficeMath structure and explicit loss diagnostics. Reuse existing fraction, radical and n-ary support. Retain bounded input and parser depth limits. Pin the actual independent Pandoc producer and retain the existing structural oracle contract. This is a bounded issue repair, not unrestricted TeX support.


### F-X235, Qualify equation layout against native Word

**Size**: L
**Depends on**: F-X234
**Test gate**: differential

Resolve the layout half of Issue365 after conversion completion. Qualify integral limits, fractions, radicals, scripts, baselines and glyph coverage for the exact formula. Retain the existing one-point geometry and declared150DPI SSIM contract. Pin the actual independent native Word producer for new fixtures and retain historical evidence identities. Deterministic font mode, save/reopen and reviewed product/raster changes are required.


### F-X236, Measure public setter inlining before changing policy

**Size**: S
**Depends on**: none
**Test gate**: regression

Resolve Issue367 through a real cross-crate non-LTO consumer benchmark and compiler evidence. Apply selective public-setter inline annotations only where measured benefit justifies them, or record an evidence-based disposition explaining no change. Preserve semantics, public API and resource bounds. Do not introduce a blanket inline policy or assert an unmeasured performance problem.


### F-X237, Restore linear Python table owner inventory

**Size**: L
**Depends on**: none
**Test gate**: regression

Resolve Issue378's v0.16.1 Python table open and edit regression. Batch qualified
physical cell inventory and index revision-route reconciliation without changing
owner authority, stale-handle causes, merged or nested cell identity, opaque
source preservation or atomic publication. Opening and one edit must perform
linear work in source cells and owner inventory. Qualify60,120 and240 cells and
the exact40-by6 fill with deterministic work counts and authenticated same-host
release0.16.0 timing comparison. Preserve no-table controls and full fresh Word
Python acceptance. A small additive hidden native batch API may serve the
existing cross-crate registry consumer. No new Python surface is required.

### F-X238, Restore python-docx Word names from the first-draft corpus

**Size**: L
**Depends on**: none
**Test gate**: regression

Resolve the Word half of Issue379. Every Word name the issue lists, including
`Document.add_section`, works as in python-docx 1.2.0 or fails with an error
that names the supported rdocx alternative. The three divergence scripts fail
with a message naming the supported way. The PR333 corpus runner at
`193819ca` runs each listed script unchanged and reports the results
separately from the F-X211 corpus. Wrap existing native behaviour. New document
semantics require a separate design.

### F-X239, python-pptx notes names and refusal of silent zero values

**Size**: M
**Depends on**: none
**Test gate**: regression

Resolve the PowerPoint half of Issue379 and its unit follow-up. Add
`Slide.has_notes_slide`, `notes_slide.notes_text_frame` and the
`MSO_THEME_COLOR_INDEX` alias. The five assignments that now save zero raise
before they change the document, and each message names the unit to use.

### F-X240, Python and CLI signing and signature verification

**Size**: M
**Depends on**: none
**Test gate**: regression

Resolve the signing half of Issue380. Python `sign` and `verify_signatures`
for documents and presentations, accepting DER and PEM keys and certificates,
and matching `sign` and `verify-signatures` CLI commands with a non-zero exit
for an invalid or partial signature. Published wheels and CLI builds enable
`digital-signatures` only after dependency, licence and size gates pass. A
test changes one signed part and sees verification name it.

### F-X241, Python and CLI form reading and filling

**Size**: M
**Depends on**: none
**Test gate**: regression

Resolve the form half of Issue380. Python content control and legacy form
field reading and filling by tag, alias or name, and `rdocx fill --values`.
A template with plain text, checkbox, drop-down and date controls plus a
legacy text field reads back its filled values after save and reopen through
both Python and the CLI.

### F-X242, Form control geometry and AcroForm field model

**Size**: L
**Depends on**: none
**Test gate**: regression

First of three Issue381 stories. Layout reports the page rectangle of every
content control and legacy form field. Each maps to an AcroForm text, button
or choice field with a unique name from its tag, alias or legacy name, a
tooltip from its alias or help text, and its current value without
placeholder text. No PDF output changes in this story.

### F-X243, Field appearance streams, full-font resources and conformance

**Size**: L
**Depends on**: F-X242
**Test gate**: differential

Write the F-X242 fields into the PDF behind an opt-in option. Every widget has
an appearance stream, the field font in `/DR` covers typed input beyond the
page subset, and tagged output places widgets under `/Form` structure
elements. `qpdf --json` lists the expected names, types, values and
rectangles, and the PDF/A and PDF/UA checks still pass with the option on.

### F-X244, Signature field, bindings and CLI option

**Size**: M
**Depends on**: F-X243
**Test gate**: regression

Close Issue381. A marked control becomes an empty `/Sig` field, and the option
reaches Rust, Python and `rdocx convert`. Output without the option is byte
identical to before. Signing the PDF in the library stays out of scope. The
manual Acrobat Reader and Preview fill and sign check is recorded as evidence.

### F-X245, Preserve chart data changes in Word comparison

**Size**: L
**Depends on**: F-X229, F-X230
**Test gate**: regression

Resolve [Issue382](https://github.com/tensorbee/rdocx/issues/382). Detect a
chart-data-only edit at both word and run granularity and record a revision.
A chart-data edit combined with adjacent text changes must also compare
successfully. After saving and reopening the redline, accept-all reproduces
the edited chart part, embedded workbook and text, while reject-all reproduces
the original. Run the exact supplied script, retain its unchanged-chart control
and independently inspect package relationships, chart caches and workbook
values. Preserve existing picture comparison, revision grouping, opaque XML
and atomic refusal guarantees. Qualify native and installed Python entrypoints.

### F-X246, Honor chart legend position and overlay during rendering

**Size**: L
**Depends on**: none
**Test gate**: differential

Resolve [Issue383](https://github.com/tensorbee/rdocx/issues/383) across Word
and PowerPoint chart rendering. Honor `c:legendPos` right, left, top, bottom
and top-right, defaulting to right when absent. When `c:overlay` is absent or
false, reserve space outside the plot so legend entries do not cover bars.
Qualify true overlay, absent legends and hidden-legend controls without losing
unmodelled chart XML or schema order. The exact supplied default, bottom and
left repros must produce distinct renders in both formats with no legend/plot
overlap. Compare independently measured plot and legend geometry and pinned
native Office renders in deterministic font mode. Attribute intentional hash
and raster deltas to this story. During design, explicitly decide the related
Python legend visibility, position and overlay request for both families and
record any residual before claiming full issue closure.

### F-X247, Preserve the common LaTeX inputs reported in Issue365

**Size**: L
**Depends on**: F-X234, F-X235
**Test gate**: differential

Complete the finite 73 input conversion matrix in
[Issue365's follow-up](https://github.com/tensorbee/rdocx/issues/365#issuecomment-6200740130).
Preserve its authenticated original script, all 73 distinct input rows and their
order. Cover the named Greek and mathematical symbols, relations, operators,
arrows, upright functions, operatorname/mathrm/text, binom, dfrac/tfrac and
spacing without losing atom/script ownership or meaningful text. Design any
forced fraction style, text whitespace or unsupported representation distinction
explicitly before implementation. This is a bounded extension, not unrestricted
TeX support, and does not reopen the completed S93 exact-formula contract.
The bounded conversion canonicalizes frac/dfrac/tfrac to stacked Bar fractions,
uses a parenthesized NoBar for binom, distinguishes upright functions from plain
math and normal text with existing properties, and preserves literal whitespace
and three distinct explicit space tokens. Forced TeX display/text style is not
preserved. Native font/atom corrections and fraction/short-display frames retain
their independently qualified geometry bounds.

Retain original output separately from a corrected observer whose only script
change makes the unconditional FAIL summary depend on actual losses. Require
zero unexpected diagnostics and independently inspected correct normalized
structure for all 73 inputs. Execute actual pinned Pandoc forward/reverse
qualification and native/render comparisons where new paint semantics require
it. Preserve existing bounds, schema order, opaque source, deterministic fonts,
historical oracle identities and applicable geometry/raster thresholds.
Qualify native and fresh installed Python save/reopen, real semantic mutation
failures and applicable publication, binding, WASM and hash riders. Issue365
remains open until this broader acceptance has evidence on main.


### F-X248, Validate identified paragraphs inside block content controls

**Size**: M
**Depends on**: F-X196, F-X211a, F-X222
**Test gate**: regression
**Spec reference**: `04-opc-and-packaging.md`, Checked Word operation source ownership, `10-bindings-spec.md`, Python Word document operations, and `12-testing-strategy.md`, Integrated Python compatibility qualification.

Resolve [Issue384](https://github.com/tensorbee/rdocx/issues/384). Valid paragraphs
inside block content controls retain `w14:paraId` and `w14:textId` without a
source/model validation-owner mismatch. Run the unchanged reporter reproduction
through `rdocx validate`, `Document.validate()` and `rdocx.validate_file`.
CLI succeeds with exit0 and Python raises no exception. Open, save and reopen
preserve both identifiers, namespace bindings and opaque control content.
Cover a Word TOC, nested controls and a valid unchanged ordinary paragraph.
Retain refusal of genuinely mismatched physical owners and malformed source.
Do not strip producer identifiers or bypass the owner check to accept a file.

### F-X249, Set typed content control values as Word does

**Size**: M
**Depends on**: F-X241
**Test gate**: differential
**Spec reference**: `04-opc-and-packaging.md`, Checked Word operation source ownership, `10-bindings-spec.md`, Python Word document operations, and `12-testing-strategy.md`, Binding tests.

Own the content-control-value group of [Issue385](https://github.com/tensorbee/rdocx/issues/385).
A successful fill removes the selected control's `w:showingPlcHdr`. Date fills
write the corresponding `w:fullDate` while respecting supported producer date
format metadata. Drop-down fills accept a declared value or display text, write
the declared display text and refuse an unknown item. Individual by-tag and
by-alias selectors raise `KeyError` when missing, with invalid values refused
before publication. The Python setters and `rdocx fill` agree after save/reopen.
Use the pinned reporter fixture and independent Word-produced typed controls.

This is new typed native behavior beyond F-X241's existing display-string
wrappers. Design exact date lexical/calendar admission, bound XML consistency,
ambiguous item handling and compatible missing-selector errors explicitly.
Preserve the distinct existing bulk bind contract, all-match selection, atomic
candidate publication, source ownership and unsupported-kind refusal unless
an explicit reviewed compatibility decision changes a named method. Do not
expand this repair into F-285 through F-289 control creation or full forms.

### F-X250, Complete everyday Word chart equation list and note APIs

**Size**: L
**Depends on**: F-X200, F-X202, F-X215, F-X247
**Test gate**: round-trip
**Spec reference**: `09-charts-spec.md`, The ChartML model and Python authoring and owned replacement, `08-rendering-spec.md`, Word OfficeMath layout and Authored Word page appearance and note references, and `10-bindings-spec.md`, Python Word formatting and checked list publication.

Own the Word chart read-back, display-equation, list-numbering and notes groups
of [Issue385](https://github.com/tensorbee/rdocx/issues/385). After save/reopen,
`Document.charts` lists document charts with their type, categories and owned
series names/values. Design ordering, supported chart families, source scope
and unknown imported-chart behavior against existing ChartML providers.
`Paragraph.add_equation(source, format=..., display=True)` produces
`m:oMathPara`, exposes the display flag on inspection and preserves the
compatible inline default and conversion diagnostics.

`add_numbered_list_item(..., restart=True)` starts a fresh count and following
items continue it. `Document.restart_numbering(paragraph, start=5)` assigns a
valid current paragraph to a new instance with the required start override.
`ListLevel(..., font="Arial")` writes the bullet-level font. Preserve numbering
identity, stable existing list behavior and current-document owner checks.
Footnote/endnote references carry the corresponding `FootnoteReference` or
`EndnoteReference` style. Public document removal by note id works after an
unrelated structural edit retires a held note handle, preserving other notes.
Reuse existing checked native removal rather than implementing a second path.

Run all four pinned reporter groups against fresh installed Python, inspect
expanded chart/math/numbering/note XML independently, and qualify affected
paint with existing deterministic/native controls. This does not reopen the
finite F-X247 conversion contract or complete F-283 navigation composition.

### F-X251, Refuse invalid Word values before publication

**Size**: M
**Depends on**: F-X201, F-X220, F-X238
**Test gate**: regression
**Spec reference**: `03-architecture.md`, Owned publication and causal paths, `10-bindings-spec.md`, Python Word document operations and Licensed embedded fonts in Python and CLI, and `12-testing-strategy.md`, Binding tests.

Own the five Word invalid-value cases and Word font CLI wording in
[Issue385](https://github.com/tensorbee/rdocx/issues/385). Saving a DOCX to a
`.pdf` path refuses before writing and names `to_pdf`. A duplicate tab stop at
the same physical position refuses without adding another stop. A page break
in a header refuses instead of storing content Word ignores. A malformed
language tag such as `not a tag!` refuses without mutation. An embedded font
face missing its `name` table refuses without creating a font record or part.
Use a licensed constructed-font control and preserve ordinary valid fonts.

`rdocx font` existing-output errors must not recommend unavailable `--force`.
Preserve input/output collision, no-overwrite and staged publication behavior.
Design the exact language/format/path admission and exception contracts,
including valid document variants and supported byte/stream routes, before
changing compatibility. Do not introduce a generic filename policy or change
unrelated setters. Run all five pinned tests and meaningful CLI refusal
controls with byte, output-absence and held-owner atomicity assertions.

### F-X252, Refuse newly invalid Word raw XML atomically

**Size**: M
**Depends on**: F-X211a
**Test gate**: regression
**Spec reference**: `04-opc-and-packaging.md`, Checked Word operation source ownership, `10-bindings-spec.md`, Python Word document operations, and `12-testing-strategy.md`, Integrated Python compatibility qualification.

Own the Word raw-XML group of [Issue385](https://github.com/tensorbee/rdocx/issues/385).
A checked paragraph `replace_xml` refuses a newly introduced unknown `w:`
child and a newly introduced foreign element not covered by applicable
`mc:Ignorable`. Refusal preserves document bytes, original physical owner,
held handles and unrelated package parts. Run the exact two pinned mutations
and valid edit controls, with independent schema/Word refusal authority.

Design a bounded admission rule for edited XML, not a whitelist of content the
renderer models. Preserve imported unmodelled XML verbatim, valid foreign
extensions, inherited namespace/compatibility scope and unchanged no-ops.
Retain schema child order, resource limits and existing checked-owner proofs.
No generic XML engine or second owner model is implied by this story.

### F-X253, Honor imported chart types and hidden master shapes

**Size**: M
**Depends on**: F-X209, F-X215
**Test gate**: differential
**Spec reference**: `06-presentationml-model.md`, The shape tree, `09-charts-spec.md`, The ChartML model and Python authoring and owned replacement, `08-rendering-spec.md`, The seam that makes this cheap, and `10-bindings-spec.md`, Live PowerPoint masters and layouts.

Own the PowerPoint chart-type and master-render groups of
[Issue385](https://github.com/tensorbee/rdocx/issues/385). Imported stacked
horizontal bars, stacked columns, marker-free lines and exploded pies report
the corresponding python-pptx chart type. Unknown types return `None` rather
than a falsely specific first enum for the plot family. Prove the four known
types independently, since the reporter test also allows `None` for them.
Preserve chart data, plot order and opaque XML through read/save/reopen.

A slide with `show_master_shapes=False` omits master-owned painted shapes from
PDF, PNG, `convert` and `render`, matching actual PowerPoint and LibreOffice.
Keep layout-owned shapes, slide-owned shapes, placeholders, inherited styling
and background behavior distinct. Qualify the pinned two-slide/master-logo
reproduction and true/absent flag controls with independent native exports and
negative controls that detect an erroneously retained logo. Existing pass
routing tests alone do not establish the requested master visibility contract.
Attribute intentional render deltas in a labelled behavior commit with an
exclusive deterministic baseline lease. Do not retune unrelated rendering.

### F-X254, Refuse invalid PowerPoint values XML and empty edits

**Size**: L
**Depends on**: F-X204, F-X206, F-X211b, F-X220
**Test gate**: regression
**Spec reference**: `06-presentationml-model.md`, Checked shape and relationship authoring, Validation and Checked XML owner replacement, `04-opc-and-packaging.md`, Checked PowerPoint XML transactions, and `10-bindings-spec.md`, Licensed embedded fonts in Python and CLI.

Own the PowerPoint invalid-value, raw-XML and `replace --map` groups plus both
PowerPoint font CLI points in [Issue385](https://github.com/tensorbee/rdocx/issues/385).
Saving a PPTX to `.pdf` refuses before writing. Slide dimensions outside1to56
inches refuse. Both `flip_h` and `flip_v` on table or chart graphic frames
refuse when PowerPoint cannot honor them. A table style must be built in or
actually present in the deck. Define exact valid boundaries and compatible
exception behavior without silently coercing values or discarding source.

Checked `element.replace_xml` refuses newly introduced unknown or misplaced
`p:` and `a:` children, including the pinned shape, geometry, text-frame and
slide owner cases. `validate()` flags those actual schema violations in
imported inputs. Refused edits preserve bytes, held physical owners and
relationships. Preserve opaque imported content and valid Ignorable extensions,
namespace scope, child order and the existing source-owned transaction model.

A replacement-map pair without `expect` that matches nothing refuses as the
single `-p` route does, without publishing output. Retain explicit `expect=0`
and valid multi-pair semantics. `rpptx font remove` with a missing name exits
nonzero without publishing a pretend removal, and existing-output errors do
not recommend unavailable `--force`. Preserve staged no-overwrite behavior.
Run the exact pinned parameterized tests, both frame kinds and flip axes,
valid schema and native controls, and separate CLI failure/atomicity controls.
Design the bounded admission contract and split this story if actual work
exceeds its estimate, without dropping any of these acceptance criteria.

### F-X255, Prepare Word 0.16.2 and PowerPoint 0.14.2 patch releases

**Size**: L
**Depends on**: F-X221, F-X248, F-X249, F-X251, F-X252, F-X253, F-X256, F-X257, F-X258
**Test gate**: release regression
**Spec reference**: `10-bindings-spec.md`, Packaging, and `15-build-and-toolchain.md`, Packaging, Publishing and Release process.

At the user's end-S95 release request, prepare the exact seven-package Word
Rust/CLI family and Python distribution at0.16.2 and the exact fifteen-package
shared/PowerPoint Rust/CLI family and Python distribution at0.14.2. Preserve
the existing version lines, unified allowlists and release workflow. Align
binding manifests, internal pins, lockfile, current version assertions and
metadata. Keep the unpublished PowerPoint WASM carrier at0.12.1 and preserve
historical release notes, registry records and performance observations.

Authenticate the contribution inventory since the actual published0.16.1 and
0.14.1 tags. Write family-specific notes with included issue/PR links, actual
contributor credit, direct or hardened-equivalent provenance and bounded
acceptance claims. Disclose the approved Rust rendering struct source changes,
ZIP9 public error type boundary and newly refused authored inputs. Patch
numbering does not imply full source compatibility. Do not claim whole
Issue385 or Issue386 closure, future form authoring, or deferred native interoperability.

Qualify all22 genuine locally patched publication dry-runs, exact archive
inventories below10MiB and newly measured current README rows. Inspect fresh
installed artifacts, metadata and priority runtime under clean Python3.9 and
3.12, with strict typing and stub checks under3.12. Retain both WASM graphs,
deterministic hashes and pixels. Run the actual both-family hosted build-only
rehearsal at the reviewed source SHA and authenticate its downloaded six wheels
and source distribution per family. Preserve separate actual observer outcomes
and immutable historical evidence. Preparation does not publish. Final full
verification, independent sprint review and close precede `/release` from the
exact reviewed main merge SHA, shared/PowerPoint first, with separate final
approval per family before tags or registry publication.

### F-X256, Refuse ambiguous Word numeric units and preserve line spacing

**Size**: M
**Depends on**: F-X239
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Python API shape and the relevant Word or PowerPoint checked operations, and `12-testing-strategy.md`, Binding tests and Integrated Python compatibility qualification.

Own Issue386's Word space_after=6, font.character_spacing=2,
update_section(margin_top=1440) and line_spacing=2 cases. Design the finite
bare-number and typed Length contract before changing semantics. Match the
pinned python-docx behavior or refuse before mutation with the supported
Pt, Twips or multiple-spacing route named explicitly. Section geometry is a
unit-interpretation issue even when a tiny EMU value remains nonzero. Preserve
intentional zero, None, supported signed character spacing, truncation,
absolute typed spacing and valid section dimensions. Cover finite/nonfinite,
bool/type, save/reopen, held-owner and before/after package controls. Do not
expand this into a generic unit engine or silently alter all integer setters.

The acceptance source is [Issue386](https://github.com/tensorbee/rdocx/issues/386), reported against main d44a69af. Pin its original body and script separately from the existing Issue385 twelve-case union. Its self-catching script exit status is not acceptance. Report actual row outcomes and independent saved-package observations. A reported Python3.14 environment is not a new runtime qualification claim. This story is included in active S95 at the user's explicit request and gates F-X255 final release preparation. Completion alone does not authorize issue closure.

### F-X257, Complete remaining Word names and legacy dropdown diagnostics

**Size**: L
**Depends on**: F-X238, F-X241, F-X252
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Python API shape and the relevant Word or PowerPoint checked operations, and `12-testing-strategy.md`, Binding tests and Integrated Python compatibility qualification.

Own Issue386's newly named Word compatibility surfaces: frozen section
left_margin guidance, enum.dml.MSO_COLOR_TYPE and color.type comparison,
TabStop.alignment assignment, custom_properties set/delete/membership, and
paragraph._p plus run._r and run._element diagnostics. Match the pinned
python-docx behavior or refuse before mutation naming the actual supported
route. Keep frozen snapshots distinct from live document.section handles.
Design existing string-return compatibility before selecting enum returns.
Private XML names must guide callers to existing xml/checked replacement
providers, without an lxml graph or a second owner model.

Repair the legacy FORMDROPDOWN string diagnostic so it names the actual
received type and supported nonnegative index/entry route. This is distinct
from F-X249's typed SDT value/display behavior and must retain atomic refusal.
Default List Bullet/List Number and one-dimension picture refusal already have
qualified actionable alternatives under F-X238. Do not duplicate that work or
claim new default-template numbering or native aspect-ratio semantics.
Retain the exact reproduction, independently observed returns and saved XML,
existing consumers/stubs, source/handle atomicity and positive provider controls.

The acceptance source is [Issue386](https://github.com/tensorbee/rdocx/issues/386), reported against main d44a69af. Pin its original body and script separately from the existing Issue385 twelve-case union. Its self-catching script exit status is not acceptance. Report actual row outcomes and independent saved-package observations. A reported Python3.14 environment is not a new runtime qualification claim. This story is included in active S95 at the user's explicit request and gates F-X255 final release preparation. Completion alone does not authorize issue closure.

### F-X258, Complete remaining PowerPoint names color and language contracts

**Size**: M
**Depends on**: F-X239, F-X211b
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Python API shape and the relevant Word or PowerPoint checked operations, and `12-testing-strategy.md`, Binding tests and Integrated Python compatibility qualification.

Own Issue386's slide.slide_id, color.type and theme_color enum comparisons,
enum.lang.MSO_LANGUAGE_ID and font.language_id, shape._element and
text_frame._txBody requests. Match the pinned python-pptx behavior or refuse
before mutation with the existing native route named. F-X239's import-only
color enum acceptance did not promise enum-valued returns. Design concrete
return compatibility and finite language/identifier mappings before changing
existing string consumers. Do not imply a full LCID/locale engine or introduce
a second XML owner or lxml object graph. Preserve imported identities, chart
and layout/master owners, source XML, saved/reopened values and current stubs.
Qualify every named row individually, with before/after bytes and held-owner
checks for assignments and explicit supported controls.

The acceptance source is [Issue386](https://github.com/tensorbee/rdocx/issues/386), reported against main d44a69af. Pin its original body and script separately from the existing Issue385 twelve-case union. Its self-catching script exit status is not acceptance. Report actual row outcomes and independent saved-package observations. A reported Python3.14 environment is not a new runtime qualification claim. This story is included in active S95 at the user's explicit request and gates F-X255 final release preparation. Completion alone does not authorize issue closure.


### F-X259, Repair reviewed consistency defects and strengthen acceptance gates

**Size**: M
**Depends on**: F-283, F-284, F-X254
**Test gate**: regression
**Spec reference**: `08-rendering-spec.md`, Numbering-aware heading placement, `03-architecture.md`, structured template evaluation, `02-scope-and-non-goals.md`, Modern DOCX capability matrix, and `12-testing-strategy.md`, Numbering-aware navigation regression and template controls.

Own B1, S1, S2 and N1 from `.claude/reviews/S95.11-S95.2-code-quality-evaluation.md`. Share unnumbered STYLEREF caller semantics in the existing numbering formatter, retaining REF full-context behavior. Compare multilevel n/r/w and t values across evaluation, actual positioned rendering, update and save/reopen. Preserve implicit and explicit template lexical scopes for empty, single and multiple records in both grammars, while still rejecting malformed syntax, wrong intermediate kinds and known missing root properties. Keep late-error rollback and opaque XML ownership intact. Simplify the reviewed XML traversal predicate without changing its acceptance grammar.

Reconcile completed capability rows and enforce their live-owner invariant during metadata adoption. Strengthen the canonical design, implementation, review, verification, integration and completion commands and their existing reference skills with consumer maps, measurable acceptance, discriminating cases and final-state checks. Regenerate adapters. Process changes remain separate from behavioral commits. No new public API, crate, module or dependency is required. Final sprint full/hash/pixel/rider acceptance and independent review remain due, without rewriting historical acceptance records.

The 2026-10-22 scope includes [Issue400](https://github.com/tensorbee/rdocx/issues/400)
and candidate [PR401](https://github.com/tensorbee/rdocx/pull/401) at
62124737c428a77e0c4fec52b56665c67f221e7c. TOC outline-derived entries must
respect an explicit heading range while outline-only, body-text suppression and
custom-style selectors keep their own semantics. Qualify native, Python and CLI
rebuild with inherited/direct outline sources, lower and upper range bounds,
and saved/reopened entries. Contributor-reported Word observations are not
workspace-native acceptance evidence.

### F-X260, Restore live handles and bound structural edit cost

**Size**: L
**Depends on**: F-X211c, F-X237
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Python handle ownership and checked XML mutation, `03-architecture.md`, owned publication, and `12-testing-strategy.md`, Binding tests and Integrated Python compatibility qualification.

Resolve [Issue388](https://github.com/tensorbee/rdocx/issues/388) and [Issue389](https://github.com/tensorbee/rdocx/issues/389). [PR393](https://github.com/tensorbee/rdocx/pull/393) at beb9c448d0d99636a799d95f718acccb9ea55646 stays a candidate and is not merged.

Word cell, row and table property setters keep held and fresh paragraph and run handles when every current owner relocates to one candidate address. A real content replacement still retires the removed owner. PowerPoint `add_run`, `add_field` and `add_paragraph` keep earlier paragraph and run handles. An already materialised paragraph collection is not refreshed.

The seven Issue389 operations edit one owner. They are PowerPoint `add_textbox`, alt text, `left`, `fill.solid` and `replace_xml`, plus Word `replace_xml` and `remove_content`. A visit counter fails when the edit reads an unrelated slide or body block. The release-build large-to-small ratio below 3 is recorded with the final sprint. Debug timings do not pass. `Document.remove_content_range(start, end)` removes a half-open direct body range in one transaction, returns the removed count, and refuses a reversed or out-of-bounds range without publishing. It does not recurse into tables. Shared media and a refused XML edit stay on the checked package transaction.

The local tests are the Issue388 handle tests and `issue389_local_edits_do_not_scale_with_unrelated_package`. Installed product suites, the unchanged hash harness and the release ratio remain final-sprint evidence. Neither issue closes from this story.

The 2026-10-22 scope also includes read-only handle creation from
[Issue398](https://github.com/tensorbee/rdocx/issues/398) and candidate
[PR399](https://github.com/tensorbee/rdocx/pull/399) at
34dfd42782de08bc331a5f6c819774720bb60067. Preserve all three reporter
large/small ratios below3, including cold calls, whole body/cell walks,
first/last table position and nested/related owner correctness after edits and
reopen. Per-state warm-up is not acceptance for a failing first-call script.
Derived current-state indexes must not become a second ownership authority or
survive a mutation with stale routes. Candidate
[PR397](https://github.com/tensorbee/rdocx/pull/397) at
d9721353314ba969c05bd9f886bdd24d15d51a8d contributes another Issue389
approach but leaves Word edit costs unresolved. Its range signature/error and
new public comparison APIs do not replace the approved contract. All candidates
remain unmerged and all new acceptance remains pending.

### F-X261, Resolve Word package preservation and inspection contributions

**Size**: L
**Depends on**: F-X238, F-X220, F-X241
**Test gate**: regression
**Spec reference**: `04-opc-and-packaging.md`, checked Word source ownership and embedded fonts, `10-bindings-spec.md`, Word properties and checked XML owners, and `12-testing-strategy.md`, Binding tests and Word checked XML qualification.

Python custom properties iterate, measure, pair and look up names on the existing class, in native part order. An absent part has length 0. Style edits keep unmodelled root children, including `w:latentStyles`, in schema order. Picture removal drops a media part only when the candidate package proves it unreachable. Shared targets, pop then reinsert, and caller-embedded unused media stay. A plain sfnt font part is copied as stored. Other accepted faces are deobfuscated. An unsupported content type names the part. `validate` names a header or footer relationship whose target is not that story part. `CommentRef::resolved` is `Option<bool>`. Python is `bool` or `None`. A missing `commentsExtended` entry is unknown. A library-written comment that still has an extension is unresolved. The CLI lists only an explicit true as resolved.

Checked `replace_xml` binds a root attribute whose prefix is declared on the same root, in either order. An unbound prefix refuses without publishing. `StoryItem.xml` carries the namespace declarations a paragraph snapshot already writes. `part_names` publishes the save-equivalent entry names in lexicographic order, each with a leading slash. That public order is distinct from `zip_entries` writer order. `part_xml` reads the same entries, including content types and relationship parts. Ordinary part bytes are borrowed. Missing, binary, and non-UTF-8 parts, including UTF-16, still refuse. Default paragraph format accepts `Length` for `space_before` and `space_after`. `line_spacing` on the default, paragraph, and style surfaces uses the composed Q10 rule. `Length` is exact. A non-bool int or float is a finite line multiple. `None` clears the line and the rule. An omitted value preserves both. Bool is `TypeError`. Non-finite input and a raw product outside the closed signed 32-bit range are `ValueError`. An exact stored-grid multiple normalizes back to its signed numerator. Other finite values truncate toward zero only inside that range. A same represented value preserves an absent line rule. `exact` and `atLeast` read as a length. `auto` and a missing rule read as the stored value divided by 240. Native `auto_line_spacing` is infallible exact-grid normalization and stays separate from the Python refusal. Off-grid values keep truncation, saturation, and NaN becoming 0. A story-item paragraph edit updates a qualified text box and its admitted alternate together, or changes neither, and then the handle is stale.

The constructed remove, reinsert and `set_header_image` sequence keeps a reachable image. The private unreduced control-owner document stays unqualified. Six reduced rows do not close Issue 390. PR395 and PR396 stay unmerged. Installed product suites, deterministic font render, both WASM graphs and the unchanged hash harness remain final-sprint evidence.

### F-X262, Resolve PowerPoint authoring and table-render contributions

**Size**: L
**Depends on**: F-X205, F-X206, F-X258
**Test gate**: regression
**Spec reference**: `06-presentationml-model.md`, shape and text models, `08-rendering-spec.md`, preset geometry and Tables, `10-bindings-spec.md`, Python PowerPoint surfaces, and `12-testing-strategy.md`, Binding tests and render fidelity controls.

Slide placeholders iterate by ascending idx. Layout and master placeholders iterate in document order. Idx lookup stays on `placeholders[idx]`. Percentage transparency, gradient position, bullet percent, and baseline use ties-to-even rounding. EMU, twips, time, and adjustments still truncate. `MSO_SHAPE.UP_ARROW` is 35. The generator adds the published `upArrow` geometry because the ECMA source duplicates `upDownArrow` and omits `upArrow`. Painted table rows grow for wrapped text, spans, margins, and unwrapped vertical text. Stored `a:tr/@h` stays. `TextFrame.clear` leaves one paragraph, keeps its properties, and drops runs. Empty text leaves no runs. An empty notes line keeps the first regular run. A cell-run address uses the slide hyperlink relationship. `target_slide` on a cell stays refused. Bullet size accepts a percentage float or a `Length`. `remove_style` removes `p:style` from a shape or connector. An unsupported kind raises and does not publish.

[Issue392](https://github.com/tensorbee/rdocx/issues/392) stays open until the final sprint evidence exists. Candidate [PR394](https://github.com/tensorbee/rdocx/pull/394) at e9f9fbe7b86388b8702af5757bdfa7e4e4fc6596 stays unmerged. Its functional parent f9d165e77392b57a58223046d8cd551220c61a3a could not be re-read from this repository. Installed suites, the pinned viewer, the PowerPoint pixel baseline, both WASM graphs, and the hash harness remain final-sprint evidence. This story owns any PowerPoint paint baseline. The Word 49 hash and seven-pixel baseline stay unchanged.


### F-X263, Prepare Word 0.16.3 and PowerPoint 0.14.3 patch releases

**Size**: L
**Depends on**: F-X259, F-X260, F-X261, F-X262, F-285, F-286, F-287
**Test gate**: release regression
**Spec reference**: `10-bindings-spec.md`, Packaging and CI, `15-build-and-toolchain.md`, Packaging, Publishing, Release process and CI job matrix.

The user requests both Rust and PyPI releases after S95.3 completion and all
verification, explicitly retaining Word0.16.x and PowerPoint0.14.x. Fresh
registry inventories select absent0.16.3 and0.14.3 across the exact seven and
fifteen published Rust packages and matching Python projects. Recheck absence
before publication. Align current manifests, workspace pins, lock records,
Python projects, examples, current assertions and family contribution notes.
Preserve immutable historical evidence, external dependency versions and the
separate nonpublished WASM boundary. Preparation adds no product behavior.

Require genuine verified22-package publication dryrun, ordinary source-bound
wheels and CLIs, clean Python3.9/3.12 priority runs, strict3.12 typing/stub
checks, both WASM graphs and unchanged output baselines. Actual build-only
wheels.yml execution must bind reviewed source, all deciding jobs and exact
unique artifacts. Inspect and test downloaded host wheels as actual producers.
Keep unresolved native, fidelity and private-input qualifications explicit.
The source-availability checkpoint may push only the reviewed clean sprint
ref after all named local checks and independent microscope pass. It does
not permit a merge, tag, upload or feature completion before hosted acceptance.

All seven prerequisites must reach their applicable actual completion or
approved genuine local prefix checkpoint before implementation. Drafting and
registry/carrier research do not satisfy that boundary. This release-preparation
story is the eighth S95.3 story under the user's added release request, an
explicit scope expansion beyond the normal seven-story planning cadence.
Only close-sprint closes the sprint. Subsequent release commands require the
separate exact-main-SHA final approval, shared/PowerPoint family first.


### F-X264, Repair Word defaults package inspection and comment projection

**Size**: L
**Depends on**: F-X260, F-X261
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Python Word formatting and checked list publication, and Word checked XML owner authority. `04-opc-and-packaging.md`, The package, and Part naming. `03-architecture.md`, Owned publication and causal paths. `12-testing-strategy.md`, Checked Word comment anchor projections, and Word checked XML qualification.

In progress. Not delivered. The current mechanism is the approved spacing, comment-snapshot, and package-view contract. Final Word evidence and the full acceptance gate remain pending. Donor tests, contributor CI and benchmarks are not local proof.

Default, paragraph, and style line spacing share the Q10 rule. `Length` writes exact spacing. A non-bool int or float is a finite line multiple. `None` clears the line and the rule. An omitted value preserves both. Bool and other non-numeric types are `TypeError`. Non-finite input, a raw scaled product outside the closed signed 32-bit range, and a huge int whose conversion overflows `f64` are `ValueError`. An exact stored-grid multiple normalizes back to its signed numerator, including when the raw product is one ulp away. Other finite values truncate toward zero only after that bound, so `1.0021` becomes 240. A same represented value preserves an absent line rule. A new numeric value writes `auto`. Native `auto_line_spacing` is infallible exact-grid normalization. An exact stored-grid multiple becomes its signed numerator. Truncation, saturation, and NaN becoming 0 stay on every other value. Python refusals stay off that native setter.

Accepted comment text is one enclosing non-cell physical flow. Table cells in that flow are joined by newlines. An independent nested text box is excluded. `comment_anchor_snapshot_results` lists one result per comment. `comment_anchor_snapshots` is the strict wrapper over that body. Read-only `anchor_issue` stays outside seven-field metadata equality. Creation and movement require one destination owner. Validation walks every admitted marker in physical source order first. Only then is a valid unique same-part cross-owner bookmark, comment, or permission pair omitted from story ranges, and its name or id stays occupied. A proofing pair that crosses owners is refused. Returned ranges keep the legacy order, and owner-local ordinals stay. The name uniqueness filter still applies when a checked pair fails and paragraph projection succeeded.

`part_names` and `part_xml` read `zip_entries`. `zip_entries` is the order `write_to` serializes. Public `part_names` sorts those names lexicographically and adds a leading slash. That public order is distinct from the writer order. Metadata bytes are owned. Ordinary part bytes are borrowed. An invalid unrelated part fails the inspection with the same save error.

A neighboring self-closing paragraph in a header, footer, footnote, or endnote is part of field inventory. Endnote `PAGE` output remains a separate unmeasured output commit. A body range across a table includes the cell paragraphs. A `PAGEREF` whose bookmark is in another owner, and a generated table page range across cells, stay a named limitation outside the gate.

[Issue402](https://github.com/tensorbee/rdocx/issues/402) and candidate [PR414](https://github.com/tensorbee/rdocx/pull/414) at 3d3d42db0f9c66f26efb4ab32f37864d8bfb131e stay unmerged. [Issue405](https://github.com/tensorbee/rdocx/issues/405) and candidate [PR412](https://github.com/tensorbee/rdocx/pull/412) at 38fefb1730b474cca5e34f2b7996fc448ac38993 stay unmerged. [Issue406](https://github.com/tensorbee/rdocx/issues/406) and candidate [PR409](https://github.com/tensorbee/rdocx/pull/409) at 7df1c7e0106ee04c91171e4be41771d8504993a1 stay unmerged.

### F-X265, Preserve fragment relationship targets and duplicate styles

**Size**: M
**Depends on**: F-X261
**Test gate**: regression
**Spec reference**: `04-opc-and-packaging.md`, Relationship types, and Picture extraction and one-shot publication. `03-architecture.md`, Shared Word document operations. `10-bindings-spec.md`, Word tables, physical pictures and internal links. `12-testing-strategy.md`, Word document-operation parity controls, and Word table, picture and paired-owner controls.

Local implementation is in the claimed worker. Final sprint verification, hash, installed typing, WASM and native Word evidence remain pending. Donor tests, contributor CI and benchmarks are not local proof.

[Issue404](https://github.com/tensorbee/rdocx/issues/404) and candidate [PR411](https://github.com/tensorbee/rdocx/pull/411) at 3c9c296c276f69d2b6cc8a1eab5674627709887f stay unmerged. Picture byte identity holds for both insertion orders, the default path and reuse turned off. Q17, answered 2026-10-23, qualifies A2 on the actual available unused-relationship reproducer. Exact inserted picture bytes hold through policies, save and reopen, including the best available copied-target and body-replaced scenario. Local RED then GREEN on those available cases is mandatory. The original unreduced input is unavailable and stays unverified. It is a reporter follow-up after a concrete tested fix, and it is not a final acceptance blocker. Donor image 9 and image 10 remain a reconstructed additional case and are not claimed equivalent. Returned image and hyperlink relationship IDs retain the intended targets through intervening staged edits. Save and reopen retain the referenced target bytes. Chart and template publishers are in scope for the same relationship rule. After an actual tested fix is available on a reviewed sprint ref, or in the proper close phase, the user authorizes a reporter comment that invites a test of the original image and a reopen if it still fails. This record does not send that comment.

[Issue408](https://github.com/tensorbee/rdocx/issues/408) and candidate [PR410](https://github.com/tensorbee/rdocx/pull/410) at 4070a7886d10bced0e6da60d28e351fbc9b8a831 stay unmerged. Equal style definitions keep the first definition and add no duplicate across repeated imports. When the source first definition differs from the target first definition, import exactly one renamed copy of the source first definition. A shadowed later definition never forces a copy. Shadowed dependencies are not imported. Default-flag validation applies. The target keeps duplicate style XML verbatim. Template, building-block and rich-merge consumers follow the same rule.

### F-X266, Bound Word body reads and structural edit cost

**Size**: L
**Depends on**: F-X260
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Python Word document operations. `03-architecture.md`, Owned publication and causal paths. `04-opc-and-packaging.md`, Checked Word operation source ownership. `12-testing-strategy.md`, Binding tests, and Word document-operation parity controls.

Pending future intent only. Donor tests, contributor CI and benchmarks are not local proof.

[Issue407](https://github.com/tensorbee/rdocx/issues/407) and candidate [PR415](https://github.com/tensorbee/rdocx/pull/415) at 0fbf3e96709198dbd01b9e810181109d106d9f68 stay unmerged. The original comparison is 100 paragraphs against 1000 paragraphs, each with six verbose formatted runs, on three original calls. On the actual same release host the ratio is at most 3 times max(small, 1 ms), with an operation-count proof. No concurrent compilation is part of the measurement. Assess every additional donor route: run, table and cell XML, body paragraph and table adds, inserts and remove-range paths, plus fallbacks. Preserve held handles, refusal, namespace, unknown bytes and the debug full-path comparison.

### F-X267, Bound PowerPoint mutation cost across public publishers

The bounded native remediation covers visibility ownership, native numeric aliases, all four wipe directions, a shared raster/SVG linear alpha mask with explicit unsupported PDF effects, and six native SmartArt text profiles. Canonical animation fixtures isolate plain paint and preserve all nine clocks and thresholds. Historical shadow falloff and expanded paint-domain failures remain the explicit F-X269 follow-up. Native capture, trusted pin adoption and final gates remain required.

**Size**: L
**Depends on**: F-X260, F-X262
**Test gate**: regression
**Spec reference**: `10-bindings-spec.md`, Packaging, Explicit PowerPoint text fitting, and PowerPoint footer, field, theme and transition authoring. `04-opc-and-packaging.md`, Checked PowerPoint XML transactions. `03-architecture.md`, PowerPoint checked XML ownership. `12-testing-strategy.md`, PowerPoint checked XML qualification, and PowerPoint footer and transition acceptance.

Local implementation stores part bytes in `Arc<Vec<u8>>` and reuses a trusted read baseline. Timing acceptance, the measured M21 digest and all 248 calls require actual selected-artifact evidence and current source applicability. Donor timings are not evidence.

[Issue403](https://github.com/tensorbee/rdocx/issues/403), candidate [PR413](https://github.com/tensorbee/rdocx/pull/413) at bdf22890de1762d025b89b83b549c56941ef7199 and stacked candidate [PR416](https://github.com/tensorbee/rdocx/pull/416) at 768f8c109bd43fb9c1f6ecdfe2d913382bcc0368 stay unmerged. The inventory is all 248 public mutating calls from the external pin 3374348a48164f198ef077af46df485d0a1c469d. 159 of those calls grow with the deck, not only the five cases already recorded. Record 5 and 20 slide timings and 10, 40 and 80 media-rich timings, plus counters. Whole-deck operations and fallbacks that are by design are documented as such. PR413 alone leaves 107 package-bound calls. PR416 is incremental only once, on top of PR413 head bdf22890de1762d025b89b83b549c56941ef7199. Q9 on 2026-10-23 allows a bounded Rust change from owned `Vec<u8>` part storage to `Arc<Vec<u8>>`. Python stays unchanged. Prepared releases retain Word 0.16.4 and shared/PowerPoint 0.14.4, with the Rust shared-buffer migration stated prominently for both families. Q13 adopts the three PR416 behavior changes only as separate declared reviewed behavior commits: untouched slide bytes and digest, `fit_text` local scope, and transition and footer owner scope. The contract is approved. The local implementation is reviewed. Digest and timing acceptance bind the selected artifact, all 248 calls, original workloads and disclosed safety fallbacks through exact current source applicability. The pins, the Q13 behavior labels, and the expected deltas stay. Preserve atomicity, errors, bytes and handle validity across new, reopened, in-memory, signed and embedded paths.

### F-X268, Bound Word owner inventory and unchanged package staging

**Size**: L
**Depends on**: F-X266, F-X267
**Test gate**: regression
**Spec reference**: `03-architecture.md`, Owned publication and causal paths. `04-opc-and-packaging.md`, The package, and Checked Word operation source ownership. `10-bindings-spec.md`, Python Word document operations. `12-testing-strategy.md`, Binding tests, and Word document-operation parity controls.

In progress. The intended scope remains all eight mechanisms. Atomic refusal is implemented for the four checked Rust insertion methods and the seven existing Python paths. Qualified body splices and the owner-piece prefix are implemented. Prepared publication installs the prepared owner indexes and piece cache. Returned paragraph and table owners use that inventory with the existing native fallback. Opened-package staging is implemented for the current opening only. A checked byte reopen skips ZIP staging only for literal `PackageReadLimits::UNBOUNDED` and a trusted byte-opened baseline. New, in-memory, Flat OPC, limited and hook paths keep the checked ZIP fallback until a later legitimate byte open. Current proofs cover that opening, the main frame, relationships and identifier categories. There is no historical opening chain. Debug counters record namespace and reopen input bytes, including nonzero changed-header identifier work. Direct body addressing is implemented in source. The shared span writer matches existing document bytes, including conditional `w14`. The address cache stores XML-only facts for the same head and tail and exact aligned child bytes. One covering body frame is eligible. A raw span that crosses a body child boundary uses the checked full reader. A split raw paragraph that stays inside one body frame still uses that fallback. Prefix and suffix reuse keeps occurrence and order. `MainBodyScans` is one retained address scan record. It is independent of the parsed opening record and exists for a new document that has no opening record. Both reopen routes carry that scan record only for an exact main-part frame and exact child bytes. There is no history chain. `direct_body_item_location` stays hidden. Python `body_locations` serves pop, insert, clone, move and insert_document, and the body arm of `TableLocation::native` is the other consumer. Debug `_main_body_block_scans` is cumulative, does not reset on read, and counts actual address-scan attempts and whole-fallback children, including failed attempts. Frame scans, whole serialization, hashes, copies and other validation stay separate costs. There is no constant-time total and no stable Python API. Piece guards for a field, a tab stop or a bookmark remain model-based. Addressing reads fresh serialization and does not copy those typed facts. Main-part child staging is implemented in source with separate current serialization and retained-source read records. `holds_field` excludes value-only reuse. Parsed normalizing source does not grant writer reuse. Exact part identity and normalized parsed output qualify retained read facts across reopen. A narrow parsed-frame proof preserves producer body-leading whitespace, while alternate roots keep the whole parser. Drawing-ID edits use the same current-part reconstruction and keep the checked whole parse when unavailable. Scoped replacement proof and child canonicalization are implemented in source. Physical selection, context-bound decoder facts, checked disjoint child edits and whole-path refusal controls preserve the original fallback. Whole mechanism qualification remains in progress. Not delivered. The design is approved on 2026-10-24. Codex root approved the refresh2 extension the same day. Codex root approved the refresh3 extension on 2026-10-24 after reading the full corrected contract and the four-file delta. The canonical HEAD is `d6cefc239ab7de0d4dae63e13d730da558d3a3a5`, plus the working draft. The extension is the same story and its seventh mechanism. It adds A10 and C7. The new helpers stay the captured signatures. The staging counter stays six units. The 73 flat calls, the 5 linear calls, and the 252-call catalogue stay the same. Q22 covers the added cache layers. No new question is opened and no timing promise is added. The prior approved scope stays approved. PR421 stays an unmerged candidate and is not accepted code. This record is that Root review. It is not a self-review. Remaining performance work is still planned. No final gate pass is claimed. F-X266 and F-X267 satisfy the approved local provider prefix. F-X268 follows the integrated F-289 and F-290 source with shared leases released. Final provider and sprint acceptance remain pending. The test gate stays regression. The 10 to 17 developer-day figure stays historical for the scope before PR422. The approved refresh4 extension records a provisional 12 to 21 developer-days. Reassess it after decomposition. It is a planning judgment, not a measurement, not a capacity promise, and not actual work. The 7 to 12 figure recorded at the first approval stays historical. None of these figures is added to the seven-story 27 to 41 figure.

Codex root approved the refresh4 extension on 2026-10-25. The baseline is `d1f42516eb3747442fe4304a6331eab946dc9a0e`. Fact cutoff is 2026-10-24T23:05:39Z. It is the same story and its eighth mechanism. It adds A11 and C8 and extends A10 counted work from 9 calls to 14. A10 stays independent of the timing gate. The new private helpers stay the captured signatures. The staging counter stays six units. The current flat set is 79. The historical Q23 set of 73 stays historical. The 5 linear calls, the 252-call catalogue, the original 407 gate, and the PowerPoint 248-call run stay. Q22 covers the added `decodes` fact. Q24 is answered. The prior approved scope stays approved. PR422 stays an unmerged candidate and is not accepted code. F-X268 is in progress, owner codex. F-290 is unaffected. No final gate pass is claimed.

The full picture inventory reuses existing actual source reads for ordinary picture-free direct main-body paragraphs and completely plain table blocks under exact current source, frame and namespace proofs. Direct paragraph spans match exactly. A table certificate covers only a complete TableCell owner and its paragraph strictly inside that same table. Raw/control/revision/grid/property sidecars, nested tables and nonplain paragraphs decline. All story and owner traversal, scope closure, picture validation and error ordering remain authoritative. Opaque, private, drawing-bearing, related and unproved cases keep the original parser. Actual warm-hit and real force-off controls accompany the unchanged final timing rider. No new retained cache or public API is introduced. Qualification remains in progress.

Canonical relationship and retained-source readback reconcile exact validated values into the current owned model instead of copying every equal private-free child. Mandatory source reads, complete frame replacement, resource validation and full rollback remain. Drawing no-op publication still requires exact source bytes and complete actual model equality. Original private and whole fallbacks, retained slot rejection, record publication and signature authority remain. These changes introduce no retained cache or public interface. Their actual mechanism evidence does not complete the timing or memory rider.

Field-free picture insertion can share an exact immutable story projection across direct-item selection and body boundary discovery, and share its reopened projection for the final returned location. Field-bearing sources keep their original scan schedule. Namespace vectors are shared only within a scanner, with ordered copy-on-write declaration updates and all original attribute checks. No persistent cache or public API is added.

The one regression gate covers native and Python atomic namespace guards, owner-inventory reuse, unchanged package staging, eligible in-memory reopen, direct-body addressing, main-part body-child staging, scoped replacement proof, and deterministic work controls. The same gate carries a final release rider for the pinned 252-call catalogue, the current 79 flat calls, the five linear calls, the original 407 gate, the F-X267 248-call run, and a memory measurement. The body-child calls are counted-work obligations inside that regression gate. The approved counted-work set is 14. The current flat set is 79. The historical Q23 flat set of 73 stays historical. The five linear calls stay. That rider is not a second test gate. Q22 measures each added cache layer, the `decodes` fact, and the combined footprint. No question is pending. F-289 and F-290 publication and related-owner proofs stay in force.

Q18, answered by the user on 2026-10-24, is "Allow caching with measured memory use (Recommended)". A retained parse and package cache needs actual memory measurements, documented fallback cost, and unchanged correctness. The donor claim of 68 MiB extra at 6 MiB main XML is unverified.

Q19, answered by the user on 2026-10-24, is "Refuse atomically in Rust and Python (Recommended)". Affected body insertions refuse before mutation when a retained namespace owner would make saving fail. Valid documents keep working. The behavior change is a separate labelled commit.

Q21, answered by the user on 2026-10-24, refines Q19. It is "Add checked Rust methods (Recommended)". The four existing infallible Rust methods keep their legacy behavior. The new methods are `try_add_paragraph(&mut self, text: &str) -> Result<Paragraph<'_>>`, `try_insert_paragraph(&mut self, index: usize, text: &str) -> Result<Paragraph<'_>>`, `try_add_table(&mut self, rows: usize, cols: usize) -> Result<Table<'_>>`, and `try_insert_table(&mut self, index: usize, rows: usize, cols: usize) -> Result<Table<'_>>`. Python uses a checked path for every affected alias. Old Rust signatures do not change to `Result`. The later design names existing index and error behavior. There is no general XML framework.

Q20, answered by the user on 2026-10-24, is "Use these bounds and report residuals (Recommended)". Every targeted call at 100 paragraphs is at most 1.5 times the 25-paragraph time plus 5 ms. The pinned 252-call catalogue is measured and each residual stays individual. The original Issue 407 gate of 100 versus 1000 paragraphs remains.

Q22, answered by the user on 2026-10-24, is "Apply the measured-memory policy (Recommended)". It extends Q18 to the reopen, identifier, and direct-body caches. Each layer and the combined footprint are measured on otherwise equal release builds. RSS deltas are not summed into a combined figure. No cap is chosen. The donor 68 MiB figure is not accepted. Codex root approved this extension on 2026-10-24.

Q23, answered by the user on 2026-10-24, is "Use the added flat and linear gates (Recommended)". The flat set is 73 calls under the Q20 bound. Five calls, `Cell.replace_text`, `Document.set_picture_size`, `Document.add_picture`, `Document.move_content_control`, and `Document.move_comment`, use a separate bound at 250 and 1000 paragraphs: large is at most 5 times small plus 5 ms. Under that historical Q23 contract, `Run.add_picture`, `Cell.add_picture`, and `Cell.merge` were 252 residuals. Codex root approved the refresh2 extension on 2026-10-24. Q24 moves `Cell.add_picture` into the current 79. The current residuals are `Run.add_picture` and `Cell.merge`.

Q24, answered by the user on 2026-10-25, is `Require all 79 flat-cost gates (Recommended)`. The six calls were not in the historical 73: `Paragraph.replace_text`, `Cell.replace_text`, `Document.replace_text_at`, `Document.add_picture`, `Cell.add_picture`, and `Document.add_hyperlink_to_story`. The current set is 76 catalogue keys plus the same 3 XML reads, under the existing 100 versus 25 bound of at most 1.5 times plus 5 ms. All 5 A9 linear calls stay at 250 and 1000, including the 2 that overlap, and all 252 stay reported. Codex root approved the refresh4 contract the same day at baseline `d1f42516eb3747442fe4304a6331eab946dc9a0e`. Q1 through Q23 stay answered.

Candidate [PR415](https://github.com/tensorbee/rdocx/pull/415) at 2839e22c14f96aeaeab5245581b5681509774559, candidate [PR417](https://github.com/tensorbee/rdocx/pull/417) at bdf49f0005d1d15078b76b5da83c5f3e2b8cce12, and candidate [PR418](https://github.com/tensorbee/rdocx/pull/418) at dba496fcd8b4d482995a23ab81d27ad2bfcbf5e7 stay unmerged. They are not accepted code. The historical PR415 pin 0fbf3e96709198dbd01b9e810181109d106d9f68 remains the F-X266 record. Candidate [PR419](https://github.com/tensorbee/rdocx/pull/419) at 9916bcdc01bc4acc065ae45afa8bd8ae07a2d30c and candidate [PR420](https://github.com/tensorbee/rdocx/pull/420) at 37851676af3e82a0500c4743ef86da417950d2c3 are own increments on that chain. They stay unmerged and they are not accepted code. Actions run 40092396584 on the PR419 head failed Word fidelity because the LibreOffice 26.2.5.2 download hit TLS CERTIFICATE_VERIFY_FAILED before a comparison. That log is donor evidence only. A green PR420 check is not a PR419 pass. A later run 40103891571 on the same PR419 head concluded success, including Word fidelity job 120901503513 and CI gate job 120914750403. That later run stays donor evidence and is not local acceptance. Candidate [PR421](https://github.com/tensorbee/rdocx/pull/421) at ab038a21bb88a38d46c8f640587c3a20b57f6770 is the own increment after the PR420 head. It has 7 commits and 10 paths. Its own diff SHA-256 is 50c390ff8038d69b95d19510122ab51d2fd3c37b30f5b5d79a89dc610f805e35. It stays unmerged and it is not accepted code. Its checks were incomplete at cutoff 2026-10-24T17:02:28Z. Candidate [PR422](https://github.com/tensorbee/rdocx/pull/422) at cf7bf55edfab1e0617f6fe35810355cb02126018 is the own increment after that PR421 head. The chain is linear and has 5 commits. The compare is ahead by 5 and behind by 0. The merge base is the PR421 head. The 5 paths are `README.md`, `crates/rdocx-py/tests/test_core.py`, `crates/rdocx/src/document.rs`, `docs/hld/03-architecture.md`, and `scripts/readme_doctests.py`. The own diff is 60325 bytes and its SHA-256 is ee0a8ce239fafea714c7bc1a49601fab4f2c864c3931b8ff4c98c3fb9dce189c. Comments and reviews are 0. Nineteen donor check runs completed with success or skip conclusions, and the donor CI gate succeeded. The legacy combined status endpoint records pending. That field is literal and is not acceptance. The candidate stays unmerged and it is not accepted code. Final own-union fidelity stays mandatory.

### F-X269, Qualify PowerPoint shadow falloff and wipe paint domains
**Sprint**: S98
**Size**: M
**Depends on**: F-X267
**Test gate**: regression

Pending. Native PowerPoint16.114.1 shadowed rectangle controls expose a static shadow falloff/foreground geometry difference and a soft-wipe effect-domain difference. The existing shadowed source, lossless native PDF, movies, strict sample failures and alpha profiles remain under `/private/tmp/s954-q26-supplemental-native-analysis-v1` and `/private/tmp/s954-q26-wipe-factor-analysis`. Plain animation qualification in S95.4 does not qualify those shadowed samples. Preserve the original thresholds and source identities. Compare native static shadow geometry and alpha falloff separately from the wipe's expanded paint domain, using independently captured native evidence and meaningful transformed/overlapping shadow controls. No fitted pixel offset, source-bound replacement oracle or changed foreground cutoff may stand in for the cause. The future design must establish exact paint bounds and effect composition before implementation. It must also qualify held entrance-wipe composition with subsequent opacity and exit operations. The native one-literal hold/remove experiment demonstrates that the held filter can mask later properties. S95.4 isolates that lifetime in its fixture and preserves both exact source captures rather than claiming a product repair for the held combination. Size M is provisional and requires reassessment at design. No implementation or acceptance credit is assigned here.


### F-X270, Prepare Word 0.16.4 and PowerPoint 0.14.4 patch releases

**Size**: L
**Depends on**: F-X263
**Test gate**: release regression
**Spec reference**: `10-bindings-spec.md`, Packaging and CI, and `15-build-and-toolchain.md`, Packaging, Publishing, Release process and CI job matrix.

Prepare the user's requested post-S95.4 Rust, CLI and Python releases within the retained Word0.16.x and shared/PowerPoint0.14.x lines. Select fresh absent0.16.4 and0.14.4 across the exact seven and fifteen Rust packages, matching Python distributions and remote tags. Preserve already published0.16.3/0.14.3 changelog sections exactly and place only S95.4 additions in the new family sections, with explicit approved Arc source-break migration and native, private-input and performance limits. Align current manifest carriers, internal pins, lockfile, package metadata, README evidence and current policy assertions without rewriting historical version controls or the separate nonpublished WASM boundary.

Preparation adds no product API or algorithm. Require actual current-source22-package verified dryrun below10MiB, accurate README inventories/examples, actual build-only wheels.yml jobs and exact downloaded CLI/wheel/sdist qualification, host CLI version checks, Python3.9/3.12 installed priority calls and strict3.12 typing/stub checks. Existing product/native/performance/memory evidence remains applicable only through an explicit reviewed version-only source and consumer analysis. All eight product stories' final criteria, including Cell.add_picture, remain mandatory and block completion. A reviewed clean canonical sprint source-availability checkpoint may push only this sprint ref to run the named hosted build-only jobs, with publication jobs skipped. Only close-sprint merges. Subsequent release commands require fresh mutable preflights and separate approval at the exact reviewed main SHA, shared/PowerPoint first.


### F-X271, Restore pinned viewer downloads and planning policy

**Sprint**: S96
**Size**: M
**Depends on**: none
**Test gate**: regression

Restore the exact LibreOffice 26.2.5.2 oracle for Linux and macOS CI without upgrading the viewer, weakening TLS/checksums, skipping gates or silently changing executable identity. Current stable URLs return 404before Test, MSRV, Word fidelity and rpptx installed tests. Official old-release HEAD candidates return200but are not authenticated archives. Verify actual platform bytes, safe bounded extraction and executable version against retained independent pins before switching retrieval. Preserve failed CI logs. Reconcile the existing hard-coded roadmap policy with the authorized whole-number sprint shift, and keep metadata/classification policies meaningful. Scope is tooling/policy only. Source scripts and workflow changes belong to later implementation, not the analysis intake. Require current-head CI successors to run the previously blocked gates.

### F-X272, Preserve existing revisions during document comparison

**Sprint**: S96
**Size**: L
**Depends on**: F-X229, F-X230, F-X245
**Test gate**: regression

Finish every criterion of Issue 425 using the current own commits of PR 427, not its stale base or donor archive measurements. Preserve unchanged existing revision owners byte for byte, allocate new ids without collisions and add the new reviewer's changes elsewhere across body and related stories. CLI counts only newly created revisions, with binding/stub behavior aligned. Prove author-filtered accept/reject reconstruct both inputs while retaining the first review. Cover no-op, ignored stories/comments, tables/controls, headers, notes, section changes, comment graphs, moved/deleted owners and late atomic refusal. PR 427 still refuses changes to a revision-bearing owner. Implement those required cases or explicitly meet the issue's documented-limitation plus usable-workaround alternative before closure. The contribution calls this Combine-like preservation, while Word Compare uses accepted-view semantics. Verify that compatibility fact before approving the bounded policy. Ignored stories retain their current id-seeding exclusions, so do not claim global id uniqueness from compared-story allocation. No blanket Word Compare parity follows from the supplied positive. Preserve S95.4 physical/live-owner and package-cache invariants. Full native/installed and composition acceptance stays due.

### F-X273, Repair header body clearance and complete TOC navigation

**Sprint**: S96
**Size**: L
**Depends on**: F-283, F-X264
**Test gate**: differential

Finish Issues 424 and426 using PR 428 as partial input. Keep variant-selection refactoring, geometry behavior and navigation behavior separately labelled. Clear body layout against the actual selected header/footer extents, including paragraph spacing and the logo/table reproduction. Pin primary semantics and independently qualify positive/negative margins, first/even/default/blank variants, flow versus floating table/frame owners, anchors, borders and notes with Word. Pinned LO is additional evidence, not a replacement for promised native behavior. Preserve furniture geometry and recalculate body height, page count and TOC page targets. Give every physical source paragraph qualified navigation facts while distinguishing placed versus unplaced vertical-merge continuation/deleted-row content. Rebuild all supplied TOC switch combinations through Rust, Python and CLI without stale-fact bypass. Include placed headings, nested tables, styled unplaced headings and adjacent STYLEREF/outline consumers. The donor spacing mismatch and no-outline-switch styled-heading failure remain explicit incomplete acceptance until repaired or separately dispositioned. Run actual current-source native geometry/cache/PDF/PNG and combined comparison/page-format controls. Never close on the partial donor fixture alone.

### F-X274, Honor section and field page-number formats

**Sprint**: S96
**Size**: M
**Depends on**: F-X273
**Test gate**: differential

Resolve uncovered Issue 423without inventing a contribution or silently adding a public authoring API. Read preserved pgNumType format alongside restart values and use a common section/field formatting policy in layout, header/footer rendering and explicit page-field cache updates. Before design approval pin ECMA 17.6.12 and ordinary PAGE field-format switch precedence, explicit absence/unknown-format handling, roman bounds and alphabetic rollover. Qualify decimal, lower/upper roman, lower/upper letters, repeated section restarts, all selected header/footer variants and explicit switches. The original four-page file must display i,ii,1,2in native Word, saved/updated caches, library PDF text and CLI output. Preserve unknown XML, existing numeric conversion and default decimal behavior. Prove geometry and rebuilt TOC targets remain coherent with F-X273 after save/reopen. Unsupported imported states stay preserved and diagnostic, never guessed as successful parity.

### F-X275, Prepare Word 0.17.0 and shared/PowerPoint 0.15.0 with strict SemVer gates

**Sprint**: S96
**Size**: L
**Depends on**: F-X270
**Test gate**: release regression
**Spec reference**: `10-bindings-spec.md`, Packaging and CI, and `15-build-and-toolchain.md`, Packaging and Release process.

Prepare the explicitly approved post-S96 Word0.17.0andshared/PowerPoint0.15.0Rust, CLI and Python families. Align current manifests, internal pins, lockfile, package identities, README evidence and current policy assertions. Preserve published notes and F-X270history. Enforce complete numeric compatible/yanked/equal published baselines with pinned native API checks and external compiler canaries. Initial new-minor applicability and migration limits stay explicit. Prepare the validated narrow0.14.3reporter workaround and a separately held optional only-OPC.4yank candidate, without promising a universal compatible .14patch. Add family-only notes with authenticated credits and supported/refused boundaries, without inventing PowerPoint product changes.

Require the actual final-source22-package verified dryrun below10MiB, current README measurements, reviewed build-only wheels jobs and exact downloaded artifacts, host CLI identities, installed Python3.9/3.12priority calls and3.12typing/stubs. All F-X271through274final criteria remain mandatory in one common integrated union. Formal start dependency is the completed F-X270 machinery, without using a local checkpoint to waive F-X271actual hosted acceptance. The existing validated F-X271source-availability route may expose only the clean reviewed canonical sprint source for actual normal CI and build-only qualification. Publication is separate after close-sprint and requires final exact-main approval for each family, shared/PowerPoint first. Implement the strict release gate in existing workflow/script/command carriers. No new product API, source file or npm publication is included.


The strict gate uses the existing `release-semver` subcommand, native producer
jobs, fixed signed packet subjects and independently acquired current sprint
authority. Execute, exact-row reuse and frozen evidence consumption are
separate modes. The existing wheels consumer requires completed CI and fresh
selected-family publication preflight. Local implementation or synthetic tests
do not satisfy the required actual three-host and final hosted acceptance.
