# 08, Rendering spec

Owners: `oxml-layout` for the types, `oxml-pdf` for the backends,
`rpptx-render` for the slide pipeline.

## The seam that makes this cheap

`crates/rdocx-layout/src/output.rs` is already 100 percent docx-free, and
`crates/rdocx-pdf` depends on the layout crate and nothing else in the
workspace. It consumes only:

```rust
pub struct LayoutResult { pages: Vec<PageFrame>, fonts: Vec<FontData>,
                          metadata: Option<DocumentMetadata>, outlines: Vec<OutlineEntry> }
pub struct PageFrame { page_number: usize, width: f64, height: f64,
                       elements: Vec<PositionedElement> }
```

**A slide is a page with a fixed size.** Font subsetting, ToUnicode CMaps, JPEG
passthrough, PNG inflate, soft masks, PDF assembly and the tiny-skia rasteriser
all carry over unchanged. That is roughly 1,667 lines the presentation side does
not have to write.

## Extending `PositionedElement`

The obvious approach is to bolt `rotation` onto `Image`, `gradient` onto
`FilledRect`, `arrowhead` onto `Line` and so on. That is about ten new fields
across five arms, breaks every construction site in `rdocx-layout`, and **still
does not nest**, because `a:grpSp` transforms compose arbitrarily deep.

**Add exactly two arms instead.** Rotation, flips, group transforms, clipping,
all 187 presets and `custGeom` collapse into them, and rdocx's five existing
arms are untouched, so rdocx's output stays bit-identical.

```rust
/// 2x3 affine, row-major: x' = a·x + c·y + e ; y' = b·x + d·y + f.
/// Maps 1:1 onto the PDF `cm` operator and onto tiny_skia::Transform.
pub struct Transform { pub a: f64, pub b: f64, pub c: f64,
                       pub d: f64, pub e: f64, pub f: f64 }

pub enum PathCommand { MoveTo(Point), LineTo(Point),
                       CurveTo { c1: Point, c2: Point, to: Point }, Close }
pub enum FillRule { NonZero, EvenOdd }
pub struct Path { pub commands: Vec<PathCommand>, pub fill_rule: FillRule }

pub struct GradientStop { pub offset: f64, pub color: Color }

/// Content-addressed media handle. Replaces `embed_id`, which assumed one
/// global relationship namespace and is therefore invalid for pptx.
pub struct MediaId(pub u64);

pub enum Paint {
    Solid(Color),
    Linear { start: Point, end: Point, stops: Vec<GradientStop>, extend: (bool, bool) },
    Radial { center: Point, radius: f64, focal: Point,
             stops: Vec<GradientStop>, extend: (bool, bool) },
    Tile { image: MediaId, tile: Rect, transform: Transform },
}

pub struct Stroke { pub paint: Paint, pub width: f64,
                    pub cap: LineCap, pub join: LineJoin,
                    pub dash: Option<Vec<f64>> }

pub struct PathElement { pub path: Path, pub fill: Option<Paint>,
                         pub stroke: Option<Stroke> }

pub enum Effect { OuterShadow { dx: f64, dy: f64, blur: f64, color: Color } }

pub struct GroupElement { pub transform: Transform, pub clip: Option<Path>,
                          pub opacity: f64, pub effects: Vec<Effect>,
                          pub children: Vec<PositionedElement> }

#[non_exhaustive]
pub enum PositionedElement {
    Text(GlyphRun), Line { .. }, FilledRect { .. }, Image { .. }, LinkAnnotation { .. },
    Path(PathElement),      // new
    Group(GroupElement),    // new
}
```

`PageFrame` gains `background: Option<Paint>`, and `LayoutResult` gains
`diagnostics: Vec<Diagnostic>`. Both become `#[non_exhaustive]` with a
constructor, once, at the 0.3.0 cut.

### Why `Group` is the whole design

A shape's `a:xfrm` rotation applies to its fill, its outline **and its text**
together, so a group is the semantically correct carrier. A nested `a:grpSp`
with `a:chOff` and `a:chExt` becomes a nested `Group` whose transform is:

```
translate(-chOffX, -chOffY)
  · scale(extX/chExtX, extY/chExtY)
  · translate(offX, offY)
  · rotate_about(rot / 60000, cx, cy)
  · scale(flipH ? -1 : 1, flipV ? -1 : 1) about the centre
```

Rotated text, rotated pictures, rotated gradients and clipped picture frames all
follow for free. Arrowheads lower into extra filled `Path` elements, so they
need no backend support at all.

## The recursion hazard

**Three passes in the PDF backend iterate `page.elements` flat and would
silently skip anything nested inside a `Group`:**

| Pass | Location | Symptom if missed |
|---|---|---|
| Font subsetting | `font.rs:34` | Grouped text renders with no font |
| XObject registration | `writer.rs:69` | Grouped images vanish |
| Link annotations | `writer.rs:99` and `:355` | Grouped hyperlinks are dead |

These fail **only for pptx content**, so rdocx's suite never catches them. The
mitigation is one helper in `oxml-layout` that flattens groups and accumulates
the transform:

```rust
pub fn walk(elements: &[PositionedElement], f: &mut impl FnMut(&PositionedElement, &Transform));
```

All three passes are rewritten on it, and each gets an explicit test.

## Four latent defects to fix

All are forced by pptx, and all improve rdocx:

| Defect | Location |
|---|---|
| Y is flipped **per element**, which is incompatible with nested transforms | `writer.rs:424`, `:454`, `:463`, `:479` |
| `set_fill_rgb` drops `Color.a` everywhere, in both PDF and text | `writer.rs:414` |
| `dash_pattern: _` means dashes are ignored in **all** PNG output today | `raster.rs:73` |
| Images keyed `Im{page}_{elem}`, no deduplication, and the full font dictionary is written into every page | `writer.rs` |

The last one matters at deck scale: a 200-slide deck would embed the master's
logo 200 times.

## The PDF backend

**One global CTM instead of a per-element flip.** Emit `q 1 0 0 -1 0 H cm` once
at the top of the content stream, then write everything in top-left, y-down
coordinates so group transforms compose naturally.

- Text: `Tm` becomes `[1 0 0 -1 x y]`, the negative `d` cancelling the outer flip
  so glyphs stay upright. Mathematically identical output.
- Images: `cm [w 0 0 -h x y]`.
- Link annotations live in `/Annots`, not the content stream, so that code is
  untouched.

**This is the single highest-risk change in the plan.** Gate it on golden-PNG
diffs of the existing `samples/` corpus, comparing **pixels, never PDF bytes**,
because the operator stream legitimately changes. Land it as its own reviewable
commit before any pptx code exists.

Then: `Group` becomes `q`, `cm`, optional clip via `W n`, optional `/GS gs` for
opacity, recurse, `Q`. `Path` becomes `m`/`l`/`c`/`h` followed by `f`, `f*`,
`S`, `B` or `B*` by fill, stroke and rule.

**Gradients** are the real work: `/Pattern cs /P scn`, a pattern dictionary of
type 2 whose `/Matrix` is the element-local transform so gradients rotate with
their shape, a type 2 axial or type 3 radial shading, and a **type 3 stitching
function over type 2 exponential functions**, one per stop interval. Stops are
sorted, deduplicated and clamped, and a single-stop gradient degrades to a solid
at build time. **Stop alpha needs a luminosity soft mask and is out of scope for
v1**: composite the colour, drop the alpha, record a diagnostic.

**Alpha** becomes one `/ExtGState` per distinct value, which also fixes the
existing dropped-alpha bug for rdocx.

## The rasteriser

A recursive walk carrying an accumulated `tiny_skia::Transform` rather than the
single page scale. `Group` pre-concatenates and builds a `Mask` from the clip
path, threading it into the currently always-`None` mask argument. `Path` maps
almost directly onto `PathBuilder`, `fill_path` and `stroke_path`, **which
finally wires up dashes**. Gradients map near-directly onto tiny-skia's own
`LinearGradient` and `RadialGradient`, making this the easiest part of the job.
The hardcoded white page fill is replaced by `PageFrame.background`.

Outer shadow renders as an offset silhouette: children to a scratch pixmap,
tinted, drawn at the offset, then the real children on top. A separable box blur
can land later, and the type already carries `blur`.

## Preset geometry

| Option | Cost | Coverage | Verdict |
|---|---|---|---|
| Hand-write the top ~20 | 1-2 days | ~85 percent by frequency, but the tail is *visibly broken* | insufficient |
| Generate from the spec's shape definitions | 4-6 days | ~100 percent | **chosen** |
| Port LibreOffice's table | 2 days | ~100 percent | **rejected: MPL-2.0 file-level copyleft, incompatible with MIT OR Apache-2.0** |

**The decisive argument is that the marginal cost over hand-writing is near
zero**, because `a:custGeom` uses the identical guide and path machinery. The
evaluator gets written either way, so the presets become data rather than code.
Presets additionally carry the `<a:rect>` text rectangle needed to place text
inside non-rectangular shapes.

Mechanism: an offline generator under `tools/gen-presets/` emitting a
**checked-in** generated file. Not a `build.rs`, because checked-in generated
code gives reproducible builds, no build-time XML dependency, a clean
`cargo publish` and a reviewable diff.

**Resolve the provenance and licensing of the source shape definitions before
writing the generator.** If the ECMA-376 accompanying files prove unusable, the
fallback is deriving the tables from the specification text, which enumerates
every preset's guides.

Fallback for an unknown preset, in order: use `a:custGeom` if present, otherwise
emit the bounding rectangle **and still lay out the text inside it**, recording
a diagnostic. The invariant is *never lose a shape, never lose its text, only
ever lose its silhouette*.

## Text in a shape

Slide text is not flow text. It is a fixed box with anchoring, insets, optional
wrap, and autofit.

`line.rs` moves to `oxml-layout` and is decoupled from its four docx imports
(`CT_TabStop`, `ST_Jc`, `ST_TabJc`, `ST_Underline`) in favour of owned types,
with `LineSpacing` replacing the stringly-typed `line_rule: Option<String>`, and
one new behaviour: `wrap: false` never breaks except on an explicit break.

The algorithm: resolve `a:bodyPr`, compute the content box from the preset's
text rectangle minus insets, resolve each paragraph through the nine-level
chain, build inline items, stack lines from the box top, then anchor.

**Vertical text** is laid out horizontally in a transposed box and wrapped in a
`Group` with a 90 degree rotation. `eaVert` upright stacking is not supported in
v1 and degrades to rotated vertical with a diagnostic.

**Bullets** become marker inline items, with `marL` and `indent` mapping onto
the existing left and hanging indent support. The Wingdings trap: `a:buChar`
codepoints are usually private-use `F0xx`, and `FontManager::map_font_name`
aliases Wingdings to Symbol, which renders garbage. A small codepoint-to-Unicode
table is applied **before** font resolution.

### Autofit

**PowerPoint stores its own computed answer in the file.** Trust it.

- `a:noAutofit`, the default: draw and overflow. **Do not clip.** Spilling looks
  less broken than truncation.
- `a:spAutoFit`: the stored `a:ext` is what PowerPoint last computed. Do not
  resize.
- `a:normAutofit fontScale="62500" lnSpcReduction="20000"`: apply verbatim. This
  is both cheapest and most faithful, because it reproduces exactly what the
  authoring application decided.
- Only a bare `<a:normAutofit/>` needs iteration, and then walk PowerPoint's own
  quantised 2.5 percent ladder rather than binary-searching a continuous scale,
  so the computed value matches what PowerPoint would have written. At most 31
  steps, typically one to three, and a shaping cache makes repeat passes nearly
  free.

## Performance

`Document` keeps normal-font and deterministic `LayoutResult` values in
separate `Mutex<Option<Arc<_>>>` caches. `render_page_to_png`,
`render_all_pages`, `to_pdf` and `layout_page` share the normal result. The
deterministic page renderer uses its own result, while caller-supplied font
layouts remain uncached because those fonts are not part of a stable cache key.

Every public document mutation and mutable-accessor entry point clears both
caches before changing or exposing content. Rendering every page through the
single-page entry point therefore performs one layout per font mode instead of
one layout per page. The presentation facade follows the same ownership model
when it is added.

```rust
pub fn layout_presentation(input: &RenderInput) -> Result<LayoutResult>;
pub fn layout_slide(input: &RenderInput, index: usize) -> Result<PageFrame>;
```

## The renderer's input

```rust
pub struct RenderInput {
    pub slide_size: (f64, f64),
    pub slides: Vec<SlideBundle>,
    pub media: HashMap<MediaId, MediaData>,   // deduplicated across the deck
    pub fonts: Vec<FontFile>,
    pub metadata: Option<DocumentMetadata>,
    pub default_text_style: ListStyle,
}

pub struct SlideBundle {
    pub slide: CT_Slide,
    pub layout: Arc<CT_SlideLayout>,   // ~5 layouts shared by 200 slides
    pub master: Arc<CT_SlideMaster>,
    pub theme:  Arc<Theme>,
    pub notes:  Option<CT_NotesSlide>,
    pub rels:   RelScopes,
    pub hidden: bool,
}

/// A slide, its layout and its master each have their OWN relationship
/// namespace: "rId2" means three different things.
pub struct RelScopes {
    pub slide: HashMap<String, ResolvedRel>,
    pub layout: HashMap<String, ResolvedRel>,
    pub master: HashMap<String, ResolvedRel>,
}
```

Blips are resolved to bytes **before** emitting, and keyed by content hash. This
is why `embed_id` cannot survive the port, and it gives free deduplication of the
logo that appears on every slide.
