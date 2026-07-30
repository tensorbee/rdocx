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

This backlog is **150 stories**. Summed by size it is roughly **390 developer-days**:
about 50 at S, 60 at M, 38 at L and 2 at XL.

At a sustained solo pace that is **17 to 18 months**, not the nine to twelve
first estimated when the plan was written at phase granularity. The story-level
sizing is the more trustworthy number and this document is the source of truth
for it.

Two ways to compress, both available without reworking anything above:

- **A second developer.** M7 (DrawingML) and M8 (PresentationML) parallelise
  cleanly once M6 lands, and M12 (charts) is self-contained throughout. Two
  developers is roughly 9 to 11 months.
- **Cut a read-plus-render release at the end of M10.** That is 104 stories,
  roughly 270 days, or about 12 months solo, and it is the point at which the
  library becomes genuinely useful.

M12 alone is 12 stories and roughly 60 days, which is the single largest
discretionary block.

Front-loaded risk sits in M4, M5 and M9. M4 and M5 change a shipped renderer.
M9 is the correctness heart of PowerPoint and has no rdocx analogue.

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

**Goal**: `oxml-core` and `oxml-opc` exist, rdocx consumes them, and no
behaviour changed.

**End-of-milestone gate**: hash harness unchanged, and `OpcPackage` opens a real
`.pptx` in a test.

### F-013, Create oxml-core (M)
Move `units.rs`, `raw_xml.rs`, `xml_text.rs`, the generic half of `namespace.rs`,
`core_properties.rs`, `error.rs`, plus `crates/rdocx/src/length.rs`. Make
`xml_text` public. Consolidate the duplicate `local_name` and `get_attr` helpers.
**Test gate**: the moved tests pass unchanged in their new crate.

### F-014, New unit types (M)
`Centipoints`, `Angle` in 60000ths of a degree, `Percent1000`, `Length::mm`.
**Depends on**: F-013.
**Test gate**: round-trip assertions including `Angle::from_degrees(90.0).0 == 5_400_000`.

### F-015, rdocx-oxml becomes a facade (S)
The three-line re-export block. Zero call-site changes across 323 uses.
**Depends on**: F-013.
**Test gate**: `git diff --stat` shows only `lib.rs`, `namespace.rs` and
`Cargo.toml` modified plus five deletions, and the workspace tests pass.

### F-016, Length re-export (S)
Delete `crates/rdocx/src/length.rs`, re-export from `oxml-core`.
**Depends on**: F-013.
**Test gate**: workspace compiles with no call-site changes.

### F-017, App and custom properties (M)
`AppProperties` as a union struct with `Option` fields, plus `CustomProperties`.
Neither exists today.
**Depends on**: F-013.
**Test gate**: a Word `app.xml` and a PowerPoint `app.xml` each parse, leave the
other format's fields `None`, and round-trip without emitting them.

### F-018, Create oxml-opc (M)
Move `rdocx-opc` verbatim. Replace `new_docx` with `with_main_part` and
`ContentTypes::minimal`.
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
**Depends on**: F-018.
**Test gate**: workspace compiles, and `rdocx::Error::Opc` wraps the new type.

---

## Milestone 3, Media (about 1 week)

**Goal**: one crate owns everything about an image byte string, and rdocx uses it.

**End-of-milestone gate**: hash harness shows exactly one expected delta, the
sniffed content types.

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
`native_size(default_dpi)`, documented as 72 for python-docx parity against
Word's 96.
**Depends on**: F-024.
**Test gate**: a 96 dpi PNG probed at `default_dpi = 72` yields the expected EMU.

### F-027, rdocx adopts oxml-media (M)
Delete `image_extension`, `image_content_type`, `guess_image_content_type` and
the `image_counter` field. Rewire `store_image_part`.
**Depends on**: F-023, F-025.
**Test gate**: the hash harness delta is exactly the sniffed content types, and
each is individually justified in the commit message.

### F-028, add_picture_auto (S)
A new method inferring intrinsic size, rather than changing `add_picture`'s
signature which has call sites in five examples.
**Depends on**: F-026.
**Test gate**: a picture added with no explicit size matches the image's native
dimensions at 72 dpi.

---

## Milestone 4, Layout primitives (about 2 weeks)

**Goal**: the format-neutral layout types live in `oxml-layout` and can express
a rotated, clipped, gradient-filled shape.

**End-of-milestone gate**: hash harness unchanged. This is the milestone where
that matters most.

### F-029, Create oxml-layout (M)
Move `output.rs`, `font.rs`, `bundled_fonts.rs` and `fonts/`, `error.rs`. Move
`FontFile` into the crate and re-export from `rdocx-layout::input`.
**Test gate**: the moved tests pass, and `Document::load_fonts_from_dir` still
compiles unchanged.

### F-030, Decouple line.rs (L)
Replace the four docx imports with `TabStop`, `Align`, `TabAlign`, `Underline`
and `LineSpacing`. Add `wrap: bool`. Write
`crates/rdocx-layout/src/convert.rs`.
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
**Depends on**: F-032.
**Test gate**: a single-stop gradient degrades to solid at construction time.

### F-034, Path and Group arms (M)
Add both `PositionedElement` variants, `PageFrame::background`,
`LayoutResult::diagnostics`, and `#[non_exhaustive]` on both enums.
**Depends on**: F-031, F-033.
**Test gate**: `rdocx-layout` compiles with zero construction-site changes, and
the hash harness is unchanged.

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

**Goal**: `oxml-pdf` renders rotated, clipped, gradient-filled paths and nested
groups, and rdocx's output is bit-identical to before.

**End-of-milestone gate**: golden-PNG diffs of the whole sample corpus show zero
pixel changes.

### F-037, Create oxml-pdf (S)
Rename `rdocx-pdf`, rewire to `oxml-layout` and `oxml-media`, delete the
duplicated header parsers.
**Depends on**: F-029, F-024.
**Test gate**: the eight moved tests pass.

### F-038, Golden-PNG harness (M)
Render the sample corpus to PNG and compare pixels. Distinct from the hash
harness, and specifically for F-039.
**Depends on**: F-037, F-001.
**Test gate**: passes on an unmodified tree, fails on an injected one-pixel offset.

### F-039, Global CTM flip (L)
Replace the per-element Y flip with one `q 1 0 0 -1 0 H cm`. Text `Tm` becomes
`[1 0 0 -1 x y]`, images `cm [w 0 0 -h x y]`.
**Depends on**: F-038.
**Test gate**: golden-PNG diffs show zero changes across the corpus.

### F-040, Group rendering (M)
`q`, `cm`, optional clip via `W n`, optional `/ExtGState` for opacity, recurse,
`Q`.
**Depends on**: F-039.
**Test gate**: `q`/`Q` counts balance in the content stream for a three-deep
nesting.

### F-041, Path rendering (M)
`m`, `l`, `c`, `h` then `f`, `f*`, `S`, `B` or `B*`. Stroke state via `w`, `J`,
`j`, `M`, `d`.
**Depends on**: F-039.
**Test gate**: fill-only emits `f`, stroke-only `S`, both `B`.

### F-042, Rewrite the three collection passes on walk (M)
Font subsetting, XObject registration and link annotations.
**Depends on**: F-035, F-040.
**Test gate**: three tests, one per pass, each with the target nested inside a
group. This is the R3 regression gate.

### F-043, Gradient shading dictionaries (L)
Type 2 axial and type 3 radial, with a type 3 stitching function over type 2
exponentials, and a `/Matrix` so gradients rotate with their shape.
**Depends on**: F-041.
**Test gate**: a rotated linear gradient renders with its axis rotated, asserted
on sampled raster pixels.

### F-044, ExtGState alpha (S)
One state per distinct alpha. Fixes the existing dropped-alpha bug.
**Depends on**: F-039.
**Test gate**: a 50 percent alpha fill over white rasterises to the midpoint colour.

### F-045, Rasteriser: groups, paths, gradients, dashes, background (L)
Recursive transform walk, clip masks, tiny-skia gradients, and the dash pattern
that is currently discarded at `raster.rs:73`.
**Depends on**: F-040, F-041, F-043.
**Test gate**: a rotated rectangle at 72 dpi has a filled interior pixel and an
empty corner, and a dashed line has gaps.

---

## Milestone 6, rdocx 0.3.0 release (about 1 week)

**Goal**: the extraction ships.

**End-of-milestone gate**: `cargo publish --dry-run` passes for every crate and
the `.crate` sizes are under the limit.

### F-046, rdocx-pdf deprecation shim (S)
**Test gate**: workspace compiles, `rdocx::Error::Layout` wraps the new type.

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
Document the crate moves, the deprecations, and the `0.3.0` breaking changes.
**Test gate**: every renamed crate is named in the CHANGELOG with its replacement.

---

## Milestone 7, DrawingML (about 4 weeks)

**Goal**: `oxml-drawing` models enough of the `a:` namespace to describe any
shape a business deck contains.

**End-of-milestone gate**: every `a:txBody` and `a:spPr` in the deck corpus
parses, serialises and reparses to a structurally equal value.

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

### F-065, Theme read and write (L)
`CT_OfficeStyleSheet` including `a:fmtScheme`, plus `office_default()`.
**Depends on**: F-060, F-061.
**Test gate**: a corpus theme round-trips, and `office_default()` produces a
theme PowerPoint accepts.

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
part as opaque.
**Test gate**: all 50 decks round-trip byte-identically with no XML modelling.

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

**End-of-milestone gate**: the SSIM harness meets its target across the corpus.

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

### F-099, Bullets (M)
Character, auto-number with the eight common schemes, none, size, colour, and
the Wingdings codepoint table.
**Depends on**: F-098.
**Test gate**: a Wingdings `F0B7` bullet renders as a visible bullet glyph, not
a missing-glyph box.

### F-100, Autofit (M)
Stored `normAutofit` applied verbatim, `spAutoFit` trusted, `noAutofit`
overflowing without clipping, and the 2.5 percent ladder for the bare case.
**Depends on**: F-098.
**Test gate**: a stored `fontScale` of 62500 renders at exactly 62.5 percent.

### F-101, Vertical text (S)
Transposed layout wrapped in a rotated group, with `eaVert` degrading.
**Depends on**: F-098.
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
**Test gate**: at least 0.95 SSIM on at least 80 percent of slides, and 100
percent render without panic or dropped shape.

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
**Depends on**: F-106, F-028.
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
`ContentPath`, `PathSeg`, the revision counter, `StaleElementError`, the `Length`
pyclass, error mapping.
**Test gate**: a stale path raises the named error with both revisions in the
message.

### F-130, rdocx-py core (L)
`PyDocument`, lazy collections, paragraph and run handles.
**Depends on**: F-129, F-008.
**Test gate**: `doc.paragraphs[3]` held across `remove_content(1)` raises
`StaleElementError` rather than reading the wrong paragraph.

### F-131, rdocx-py formatting and tables (L)
Font and paragraph-format sub-handles, tri-state properties, tables.
**Depends on**: F-130.
**Test gate**: `r.font.bold` returns `None` when unset, not `False`.

### F-132, Python enums, units and exceptions (M)
The `IntEnum` shims, `Length` subclassing `int`, the exception hierarchy.
**Depends on**: F-129.
**Test gate**: `WD_ALIGN_PARAGRAPH.CENTER == 1` and `Inches(1) == 914400`.

### F-133, rdocx-py rendering with allow_threads (S)
**Depends on**: F-130.
**Test gate**: four concurrent `to_pdf` calls from a thread pool complete faster
than serial execution.

### F-134, Type stubs and py.typed (M)
**Depends on**: F-131.
**Test gate**: `mypy --strict` and `stubtest` both pass.

### F-135, python-docx parity suite (M)
**Depends on**: F-131.
**Test gate**: every documented python-docx example runs unchanged, and
round-trips through python-docx preserve content.

### F-136, rpptx-py (L)
The same machinery over `Presentation`.
**Depends on**: F-129, F-116.
**Test gate**: the python-pptx documented examples run unchanged.

### F-137, wheels.yml (M)
maturin, abi3-py39, the platform matrix, OIDC trusted publishing, the `py-v*`
namespace.
**Depends on**: F-134.
**Test gate**: wheels build for every target and install into a clean venv.

### F-138, PR-time Python job (S)
`maturin develop && pytest`.
**Depends on**: F-137.
**Test gate**: the job fails when a binding test fails.

### F-139, Rewrite rdocx-wasm (L)
Wrap `rdocx::Document`. Keep the JS method names. Add the `system-fonts`
feature.
**Depends on**: F-029.
**Test gate**: a document with images, headers and numbering round-trips through
`fromBytes` and `toDocxBytes` with every part intact. This is the R-class
regression gate.

### F-140, wasm CI job (S)
**Depends on**: F-139.
**Test gate**: `cargo check --target wasm32-unknown-unknown` and
`wasm-pack test --node` both run on PRs.

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
**Test gate**: `thumbnail` produces a PNG of slide one, and `outline` prints the
title and bullet tree.

### F-146, npm publication (S)
`@tensorbee/rdocx-wasm` and `@tensorbee/rpptx-wasm`, which have no publish path
today.
**Depends on**: F-140, F-142.
**Test gate**: `npm pack` produces an installable tarball for each.

---

## Cross-cutting

### F-X001, rdocx-cli tests (M)
The binary is published and has zero tests.
**Test gate**: one integration test per subcommand.

### F-X002, README example correctness (S)
The read example uses `table.rows()` and `row.cells()`, neither of which exists.
**Test gate**: README examples compile as doctests.

### F-X003, Deduplicate the sample generators (S)
`generate_all_samples.rs` and `generate_samples.rs` overlap substantially.
**Test gate**: one generator produces every sample the harness needs.

### F-X004, Fix the shared temp path in the test suite (S)
`integration_test.rs` writes to a fixed, non-unique temp path shared across
concurrent runs.
**Test gate**: two concurrent `cargo test` runs both pass.
