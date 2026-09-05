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

### Post-v1, M14 through M22

**93 stories, roughly 336 developer-days**, or about 67 weeks solo. By
milestone, in days: M14 28, M15 12, M16 31, M17 23, M18 26, M19 85, M20 27,
M21 60, M22 44.

M19 is 85 of those 336 and may supersede a recorded permanent non-goal, so it
is the one milestone that remains a business decision rather than a scheduling
one. M20, M21, and M22 run before M19. Stopping after M22 leaves every planned
Word and PowerPoint capability complete. The spreadsheet programme starts last
and proceeds only if F-184 confirms a material gap in the Rust ecosystem.

Three stopping and compression choices remain without reworking completed
milestones:

- **Stop after M22.** S69 completes every planned non-spreadsheet capability.
  M19 can wait without carrying any Word or PowerPoint work.
- **Stop M19 after S64.** That boundary leaves a loss-aware xlsx reader,
  writer, calculation engine, charts, worksheet features, and locally
  refreshable pivots. Power Query and Office Scripts remain preserved but not
  executed until the later sprints land.
- **Parallelise after the core model.** Rendering and distribution can proceed
  separately from the Power Query and automation runtimes once their shared
  workbook, calculation, chart, and pivot contracts are reviewed.

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
job collects the exact twelve wheels and two source distributions and receives
PyPI OIDC authority only for the `py-v*` tag namespace. Manual dispatch never
publishes.
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
S70 begins. M19 proceeds only if no credible maintained crate provides the
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
ecosystem at S70, state whether the combined lifecycle gap still exists, and
archive M19 if it does not. If it does, amend `02-scope-and-non-goals.md`, define
the boundary between `oxml-sml` as chart support and `rxlsx` as a library, and
publish the preserve, model, and execute classification for every advanced
feature in this milestone. Compare the planned boundary with `calamine`,
`rust_xlsxwriter`, `umya-spreadsheet`, `xls`, and any credible successor without
claiming that simple read or write support is a differentiator.
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
Convert the supported modern Word document surface to and from bounded MHTML
with safe resource resolution and stable diagnostics.
**Depends on**: F-178.
**Test gate**: differential. Source-built MHTML and DOCX conversions preserve
body order, formatting, tables, lists, images, links, and declared loss records.

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
All six root README Rust examples use `rust,no_run` and compile against the
current `rdocx` rlib without executing filesystem writes. The read example uses
the total indexed `row_count`, `row`, `cell_count`, and `cell` APIs.
**Test gate**: `python3 scripts/readme_doctests.py` compiles all six examples.

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
the complete typed paragraph and revision view, while cache-context reuse
continues to compare the exact footnote and endnote parts. A changed reference
ID or changed note part therefore invalidates the affected reuse boundary
without poisoning later safe paragraphs. Fields, numbering, drawings, raw
children, and other unsupported paragraph-cache content remain conservative.

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

Publish the shared strict XML lexical validator required by the stable M22
Word family as the exact 15-package incubating family at 0.10.0. Move every
incubating manifest, workspace pin, lock record, README requirement, source
assertion, CI literal, workflow preflight, release regression, and the
unpublished `rpptx-wasm` preparation carrier in lockstep. Update stable source
dependency pins to the published shared 0.10.0 boundary without changing the
stable family version or granting it publication authority.

Prepare exact `rpptx-v0.10.0` release notes from the reviewed selected-family
diff and contribution inventory. Stable Word, Python, WASM, npm, and PyPI
packages remain outside publication authority. After one clean full
verification and sprint review at the exact prepared SHA, `/release
rpptx-v0.10.0` obtains separate final approval immediately before external
mutation.

**Depends on**: F-X077.
**Test gate**: release. Publish exactly the 15-package incubating family. Every
registry entry, owner, annotated tag target, GitHub release-body byte,
stable-family exclusion, absent `rpptx-wasm@0.10.0`, and applicable
contribution notification URL must verify before completion.

### F-X078, Tag v0.13.0 (S)

Publish the reviewed M22 Word-depth outcomes as the exact seven-package stable
family at 0.13.0. Move every stable manifest, workspace pin, lock record,
README requirement, source assertion, CI literal, Python and WASM metadata
carrier, workflow preflight, and release regression in lockstep. Prepare exact
`v0.13.0` release notes from the reviewed public API and contribution diff.
Python, WASM, npm, and PyPI remain outside publication authority.

Before version preparation, inspect the reviewed dependency diff. If a shared
crate version or stable dependency pin moved, add and complete a separate
incubating-family release F-ID first. Do not mix the stable and incubating
families under one tag. After the M22 end gate, one clean full verification and
sprint review must cover the exact prepared SHA. `/release v0.13.0` then
obtains its separate final approval immediately before external mutation.

**Depends on**: F-238, F-239, F-X077, F-X079.
**Test gate**: release. Publish exactly `rdocx-opc`, `rdocx-oxml`,
`rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli`. Every
registry entry, owner, annotated tag target, GitHub release-body byte,
selected-family exclusion, and applicable contribution notification URL must
verify before completion.

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
