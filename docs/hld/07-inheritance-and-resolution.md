# 07, Inheritance and resolution

Owner: `rpptx-layout`.

This is the highest-risk correctness work in the project. Getting it wrong
produces subtly wrong fonts, positions and colours on real templates, which
users notice immediately and cannot describe. It is scheduled as its own
milestone for that reason.

`crates/rdocx-layout/src/style_resolver.rs` is worth studying for its
`merge_from` cascade shape, but the structure differs: this walks a **part
hierarchy**, not a style graph.

## The output contract

```rust
pub struct ResolvedSlide {
    pub size: (f64, f64),                 // points
    pub background: Option<Paint>,
    pub shapes: Vec<ResolvedShape>,       // flattened and in draw order
    pub diagnostics: Vec<Diagnostic>,
}

pub struct ResolvedShape {
    pub group_transform: Transform,
    pub bounds: Rect,
    pub rotation_deg: f64,
    pub flip_h: bool, pub flip_v: bool,
    pub geometry: ResolvedGeometry,
    pub fill: Option<Paint>,
    pub line: Option<Stroke>,
    pub head_end: Option<ResolvedLineEnd>,
    pub tail_end: Option<ResolvedLineEnd>,
    pub shadow: Option<Effect>,
    pub content: ResolvedContent,
    /// Set when we fell back: unknown preset, unsupported SmartArt, chart, ink.
    pub unsupported: Option<&'static str>,
}

pub struct ResolvedLineEnd {
    pub kind: ResolvedLineEndKind,
    pub width: ResolvedLineEndSize,
    pub length: ResolvedLineEndSize,
}

pub enum ResolvedLineEndKind { Triangle, Stealth, Diamond, Oval, Arrow }
pub enum ResolvedLineEndSize { Small, Medium, Large }

pub enum ResolvedGeometry {
    Rectangle,
    Custom { paths: Vec<Path>, text_rect: Option<Rect> },
    BoundsFallback,
}

pub enum ResolvedContent {
    None,
    Text(ResolvedTextBody),
    Image {
        media: MediaId,
        src_rect: Option<CropRect>,
        placement: ResolvedImagePlacement,
        dpi: Option<f64>,
        rotate_with_shape: bool,
    },
    Table(ResolvedTable),
    Group(GroupElement),
}

pub enum ResolvedImagePlacement {
    Stretch { fill_rect: Option<CropRect> },
    Tile(ResolvedTilePlacement),
}

pub struct ResolvedTilePlacement {
    pub translation: Point,             // points
    pub scale_x: f64, pub scale_y: f64, // fractions
    pub flip: ResolvedTileFlip,
    pub alignment: ResolvedRectAlignment,
}

pub struct ResolvedTable {
    pub column_widths: Vec<f64>,
    pub rows: Vec<ResolvedTableRow>,
    pub right_to_left: bool,
}

pub struct ResolvedTableRow {
    pub height: f64,
    pub cells: Vec<ResolvedTableCell>,
}

pub struct ResolvedTableCell {
    pub text: Option<ResolvedTextBody>,
    pub row_span: u32,
    pub grid_span: u32,
    pub horizontal_merge: bool,
    pub vertical_merge: bool,
    pub fill: Option<Paint>,
    pub margins: TextInsets,
    pub left: Option<ResolvedTableBorder>,
    pub right: Option<ResolvedTableBorder>,
    pub top: Option<ResolvedTableBorder>,
    pub bottom: Option<ResolvedTableBorder>,
}

pub struct ResolvedTableBorder {
    pub stroke: Option<Stroke>,
    pub priority: u8,
}
```

Every theme reference, colour transform, inherited property and list-style level
is **already collapsed to a concrete value**. The renderer consumes this and
nothing else, and never sees a `p:` or `a:` type.

Line endpoint kinds and sizes cross the boundary as source-neutral values on
`ResolvedShape`. A missing kind or DrawingML `none` becomes no endpoint.
Missing width and length use the DrawingML medium default.

`ResolveCtx::resolve_slide` owns this boundary. It converts EMU coordinates to
points, resolves colours through the effective colour map and theme, evaluates
custom geometry in each path's local coordinate space, and returns text and
table content without part-tree lifetimes. Text bodies retain concrete insets,
anchor, wrap, direction, autofit, paragraphs, runs, paragraph spacing, and
bullets. Character and auto-number bullets both retain their independently
inherited font, colour, size, and choice values.

A run or field's direct `a:hlinkClick` does not enter the text cascade.
`resolve_slide_with_resources` resolves its relationship identifier through a
`ScopedHyperlinkTargets` map for the producing slide, layout, or master part,
then freezes only an external URI in `ResolvedRunStyle`. A missing relationship,
an internal slide jump, or an action-only hyperlink keeps its text and records a
stable source-scoped diagnostic. No hyperlink property is inherited from list,
paragraph, placeholder, or theme defaults.

An untyped field in an effective `sldNum` placeholder is normalized to the
`slidenum` field type. This includes a slide placeholder whose type comes from
its layout or master match. Typed slide-number fields retain that type across
the resolver boundary so the renderer can substitute the current page before
shaping.

Table resolution selects the explicit table style or the table style list's
default. A known Office built-in ID without a package definition resolves from
its built-in visual family and theme accent. A package definition of that ID
takes precedence, and an unknown ID follows the list default. A table
without an explicit ID also uses a known built-in default GUID when the list
has no matching package record. Style lookup
does not add a definition to the saved package. Resolution applies whole-table,
row and column bands, first and last columns,
first and last rows, then corner regions. Direct cell properties apply last.
The result carries concrete fills, text styles, margins, four borders, spans,
merge ownership, and right-to-left order, plus the style's background fill
sized to the whole table. A background effect is ignored with a stable
diagnostic. Border values retain region priority so the renderer can settle
adjacent edge conflicts after all cells resolve.

Table cell text always uses fixed-box layout. Cell autofit is ignored and
records one stable diagnostic. Unsupported diagonal borders, effects, and 3-D
cell properties remain visible through stable diagnostics while their XML
stays preserved by the DrawingML model. Unsupported concrete cell or border
fills also produce a stable diagnostic instead of disappearing silently.

`ResolveCtx::resolve_slide_with_media` additionally accepts `ScopedMediaIds`,
whose slide, layout, and master maps keep relationship namespaces separate.
Each flattened picture uses its producing source to resolve an embedded
relationship to `MediaId`. External links remain unsupported and produce a
diagnostic without network access. Missing picture placement defaults to
stretch. Tile translation defaults to zero, scale to 100 percent, flip to none,
alignment to top-left, and `rotate_with_shape` to true. Tile translation crosses
the boundary in points using 12,700 EMU per point. The same scoped resource
value records each resolved media item's package content type when format
support must be decided before the renderer boundary.

A media picture resolves its poster through this same source-scoped image
lookup. Media-aware timeline assembly can request a private diagnostic identity
made from the producing slide, layout, or master scope and the local shape id.
That identity selects only the owning slide poster diagnostic when fallback is
applied, then any remaining internal tags return to their exact legacy text.
Ordinary static resolution and the existing timeline path never enable the
identity. A valid poster remains `ResolvedContent::Image`. An unresolved poster
can freeze as a labelled backend-neutral group without changing inherited or
ordinary sibling diagnostics.

`ResolveCtx::resolve_slide_with_chart_resources` additionally accepts
`ScopedChartResources` and the caller's `FontManager`. The chart maps keep
slide, layout and master relationship identifiers separate. A parsed supported
chart freezes the `oxml-chart` result as `ResolvedContent::Group`. The group
contains only backend-neutral paths and shaped labels. Missing, external,
missing-target and invalid ChartML resources retain the producing scope and
relationship identifier in their diagnostics.

An unsupported chart choice uses its immediate typed fallback picture when the
embedded image relationship resolves in the same source scope. The outer
graphic-frame transform remains authoritative, while the nested picture
supplies crop and placement. Without a usable preview, the resolver freezes a
group containing the shaped `Unsupported chart` label and keeps bounds-fallback
geometry for its visible outline. Both fallback routes record the stable chart
unsupported category.

An OLE graphic frame retains its raw payload as the sole serialization source
and may project a standard fallback `p:pic` for static rendering. Layout uses
that projection only when the blip is embedded, its relationship resolves in
the producing slide, layout, or master scope, and its package content type is
PNG. The outer graphic-frame transform remains authoritative for placement.
The nested picture supplies image crop and placement only. A rendered preview
records that embedded OLE interactivity is not rendered. A missing preview,
linked blip, unresolved relationship, or unsupported media type keeps the
visible bounds fallback and the existing OLE diagnostic.

An ordinary shape carries picture fill separately from text or other content.
A direct `spPr` blip resolves in the flattened slide, layout, or master source
that owns the shape. A blip selected through the theme style matrix belongs to
theme relationship scope. Until that scope is present, it remains absent with
a precise diagnostic and cannot consume a colliding relationship ID from a
slide, layout, or master.

Each flattened leaf carries an accumulated `group_transform`. Nested group
transforms map child coordinates through `chOff`, `chExt`, `off`, and `ext`,
then apply the centre flips and then the rotation, as PowerPoint does. A
shape's own transform flips before it rotates in the same way, and the
renderer mirrors a flipped shape's text back so that a horizontal flip keeps
it readable and a vertical flip turns it a half turn. A leaf outside a group
carries `Transform::IDENTITY`.

Unrepresentable content remains visible as a bounds fallback with a stable
unsupported category and a diagnostic. Before resolution, the facade expands
the six exact pinned SmartArt layout programs into transient ordinary shapes
inside the producing scope. Other SmartArt remains in this fallback category.
The category also includes unsupported charts
without a usable cached image, OLE without a supported resolved
static preview, unknown graphic frames,
connectors with absent, unknown, or failed geometry,
image media pending relationship resolution, preset geometry pending
evaluation, and fill forms that the backend-neutral paint model cannot
represent exactly. A connector preset uses the same generated preset evaluator
as an ordinary shape and retains its transform, direct line, fill, and
arrowheads. Connector custom geometry reuses the same checked DrawingML path
evaluator as ordinary shape custom geometry. A horizontal or vertical
connector may have a zero extent on its
collapsed axis and remains a finite stroked path. A connector's line resolves
from the `a:lnRef` of its typed `p:style` and its direct `a:ln`, as an ordinary
shape's line does. A connector with neither keeps a visible default line and a
diagnostic. Explicit
`p:bgPr` fills and `p:bgRef` theme styles resolve to concrete background paint
before crossing the renderer boundary. If slide, layout, and master all omit
`p:bg`, the resolved background remains absent and the raster backend keeps its
white page default. A background fill form that the neutral paint model cannot
represent leaves the page on that default and records a specific diagnostic.
Linear gradients project the DrawingML angle onto the complete rectangular
fill bounds. Omitted or false `scaled` keeps the normalized `(cos x, sin x)`
direction. True `scaled` normalizes `(width cos x, height sin x)` first. The
axis endpoints are the minimum and maximum rectangle projections along that
direction, so non-square fills expose the complete stop domain. A background
ignores `rotWithShape="0"` because it has no enclosing shape rotation. Shape
fills with that value use the same concrete linear paint when their effective
own and group transforms are unrotated and unflipped. A rotated or flipped
shape continues to diagnose the independent rotation policy. Unsupported
circle, rectangle, and shape path gradients, tile rectangles, and each
non-none flip variant retain distinct diagnostic categories.

Notes export resolves the notes master before the notes slide overlay. Master
elements therefore remain behind slide thumbnails, metadata, and speaker-note
content. Placeholder inheritance uses the same index-first and compatible-type
fallback rule as PresentationML, but rejects multiple candidates or a source
placeholder with no unique master owner. Notes-master and notes-slide
relationships retain their original owners through transient id remapping, so
equal relationship ids cannot resolve media from the wrong scope.

**Freeze this contract when the resolver lands, or the resolver and renderer
tracks diverge.** It is versioned with the crate.

## Draw order

The single most common source of visibly wrong output. For each slide, in
order:

1. **Background.** Slide `p:bg`, else layout `p:bg`, else master `p:bg`. An
   absent background at all three levels produces no background paint.
2. **The master's non-placeholder shapes**, if the layout's `showMasterSp` is
   not `0`.
3. **The layout's non-placeholder shapes**, if the slide's `showMasterSp` is not
   `0`.
4. **The slide's own `p:spTree`**, in document order.

**Placeholders on the layout and master are templates and are never drawn.**
They contribute inherited geometry and formatting only. A renderer that draws
them emits "Click to edit Master title style" into production output.

Two further suppressions while flattening passes 2 and 3:

- A master or layout placeholder is drawn only if nothing further down the chain
  occupies it. A layout placeholder suppresses the master placeholder it
  inherits from, and a slide placeholder suppresses the layout one with the same
  `idx`.
- An occupied slide-level `dt`, `ftr` or `sldNum` is direct slide content and
  renders independently of layout and master `p:hf` flags. An occupied latent
  placeholder inherited from a layout or master requires a `p:hf` container on
  that same source part, and that container must permit its type. An absent
  attribute on a present container retains the schema default of enabled.
  Omitting the source container does not make template date and slide-number
  fields visible. Their text comes from the slide's own shape when present,
  otherwise from the deepest occupied layout or master shape selected by the
  shadowing rule below.

This logic belongs in the flattener, not the renderer.

`ResolveCtx::flatten()` returns a borrowed `Vec<FlattenedItem<'_>>`. The first
item is the effective background when one exists, followed by master, layout,
and slide shape-tree leaves in final draw order. Each shape item retains its
source, a reference to the selected `ShapeTreeChild`, materialized
child-coordinate scale, and the remaining backend-neutral rigid group mapping.
For a non-shearing group mapping, accumulated child-coordinate scale changes
the leaf anchor and extent before concrete geometry and text-box resolution.
It does not change absolute line widths, font sizes, or effect metrics. Nested
groups apply the inner mapping before the outer mapping. Group bounds do not
clip children. Recursive groups and the selected immediate `mc:Fallback` are
walked in document order. A zero child extent uses finite unit scale on that
axis and emits a stable diagnostic. A sheared or singular accumulated mapping
emits a distinct diagnostic and retains the affine fallback. The resolver does
not claim general shear support.

The background view identifies slide, layout, or master as its producer. An
explicit `p:bg` borrows its typed rendering projection while its raw subtree
remains the sole PresentationML serialisation source. It retains a reference to
the context's per-master colour map for concrete resolution. `p:bgRef` selects
the indexed normal or background format-scheme fill and substitutes its
reference colour for every `phClr`. Direct `p:bgPr` `phClr` uses the effective
`bg1` colour. The first theme background fill style is used only when selected
by an explicit `p:bgRef` index. A direct `p:bgPr` picture fill resolves its
embedded relationship in the selected slide, layout, or master producer scope.
A picture fill selected through `p:bgRef` belongs to the theme relationship
scope, which is not yet present in `ScopedMediaIds`. It remains absent with a
precise diagnostic. A colliding relationship ID in the master scope must not
satisfy it.

The layout `showMasterSp` controls only the master non-placeholder pass. The
slide `showMasterSp` controls only the layout non-placeholder pass. An absent
value means true. Ordinary master and layout placeholders remain templates and
are omitted. An occupied latent placeholder has nonempty field or run text.
Latent placeholders match by type across the level-specific indices used by
masters, layouts and slides. Direct slide placeholders stay in slide document
order and suppress the matching inherited type once. The deepest matching
occupied layout placeholder claims the inherited type before source visibility
is evaluated, so a hidden layout date cannot expose a stale master date. The
selected layout or master placeholder is emitted once when its own source
header-footer policy permits its type. Empty latent shapes fall back to the
deepest occupied layout or master match.

## The chains

Six independent chains. Getting them separately right is the difference between
"renders" and "renders correctly".

### 1. Position and size

Shape's own `a:xfrm`, then the matching layout placeholder, then the matching
master placeholder, then none. `ResolveCtx::effective_xfrm` returns an owned
clone from the first source that supplies a transform. A shape that resolves to
no extent is skipped rather than treated as an error. The public
`inherited_xfrm(placeholder, layout, master)` resolves the layout and master
part of that chain without a `ResolveCtx`. `effective_xfrm` and
`effective_picture_xfrm` fall back to it, and the facade's effective geometry
and layout change call it, so all of them share one matching rule.

### 2. Body properties

Body properties resolve independently per field. Resolution starts with exact
defaults of 91,440 EMU left and right, 45,720 EMU top and bottom, top anchor,
square wrap, horizontal text, no first-or-last paragraph edge spacing, and no
autofit. The master, layout, and slide shape then overlay their present
`p:txBody/a:bodyPr` fields in that order, so a later partial value does not
erase unrelated inherited fields. `spcFirstLastPara` accepts either XML boolean
spelling. An absent value means false.

### 3. The nine-level list style

Eight sources, merged in this order, later winning per property and per level:

1. `p:defaultTextStyle` in `presentation.xml`
2. The master's `p:txStyles`, selecting `p:titleStyle` for `title` and
   `ctrTitle`, `p:bodyStyle` for `body`, `subTitle` and `obj`, and
   **`p:otherStyle` for everything else including plain text boxes and table
   cells**
3. The master placeholder's `a:lstStyle`
4. The layout placeholder's `a:lstStyle`
5. The colour of the shape's `p:style/a:fontRef`, as the text fill of every
   level, so `lt1` gives an inserted shape light text over its accent fill as
   in PowerPoint
6. The shape's own `a:lstStyle`
7. The paragraph's `a:pPr`
8. The run's `a:rPr`

Within each list-style source, `a:defPPr` is applied to all nine levels before
the matching `a:lvlNpPr`. Paragraph fields and nested default character fields
overlay independently. Bullet colour, size, font and choice are independent
properties, while each fill and typeface slot is atomic. Effective values do
not inherit opaque XML or hyperlink actions. `cap="all"` and `cap="none"`
overlay as character properties through the same chain. Other capitalization
values remain unmodelled and preserved.

Each paragraph also resolves `a:endParaRPr` as a final character-property
overlay on the same inherited paragraph style. This terminal style supplies
the font, size, and metrics for an empty paragraph. It does not replace the
effective style of a nonempty run.

### 4. Style references

`p:style` carries numeric `a:lnRef`, `a:fillRef`, and `a:effectRef` values.
Zero selects no theme entry. Positive values use one-based indexing into the
corresponding `a:fmtScheme` list, and an out-of-range value returns
`ResolveError::StyleIndexOutOfRange` without clamping or panicking.

**`fillRef@idx` above 1000 means index `idx - 1000` into `a:bgFillStyleLst`.**

Resolution takes the format-scheme entry, substitutes every
`a:schemeClr val="phClr"` with the reference's own colour, then layers the
shape's explicit `a:spPr` fill, line and effects on top. Reference colour
transforms precede the placeholder colour's transforms. Fill and modelled
effects are atomic replacements. Line width, cap, fill, dash, join, head, and
tail values overlay independently, so omitted direct properties retain their
theme values. An explicit `a:noFill` replaces a referenced fill. A connector
resolves its line and effect references the same way, so its theme shadow is
drawn, and an `a:effectRef idx="0"` or an empty direct `a:effectLst` leaves it
without one. Its fill reference is not rendered yet.

Opaque effect children remain preserved rather than being claimed as resolved.
An opaque effect that still contains `phClr` returns
`ResolveError::UnresolvedPlaceholderColor`.

`a:fontRef@idx` is not numeric. It retains the typed `major`, `minor`, or
`none` collection selection. Typeface-token lookup within that collection is
the separate font-resolution step described below. Its colour is the fifth
text source of the list-style chain.

### 5. Colour

Covered in `05-drawingml-model.md`. The order is colour map, then theme, then
the transform stack in document order. The colour map is per-master and a dark
master inverts `bg1` and `tx1`.

### 6. Fonts

`a:latin@typeface` of `+mn-lt` resolves to the theme's minor font,
`+mj-lt` to the major font. `+mn-ea` and `+mn-cs` select the East Asian and
complex-script faces. The corresponding `+mj-ea` and `+mj-cs` tokens select
the major collection. A caller may supply an ISO 15924 script tag. The first
matching `a:font@script` entry in the selected collection overrides the base
face. A missing or unknown script falls back to the token-specific base face.
Ordinary typefaces and unknown theme-like tokens pass through unchanged.

Script detection belongs to later text shaping. The resolver accepts the
script explicitly because the current run model does not infer one.

## The resolver

```rust
pub struct ResolveCtx<'a> {
    pub theme: &'a CT_OfficeStyleSheet,
    pub color_map: ColorMap,               // clrMap, overridden by clrMapOvr
    pub master: &'a CT_SlideMaster,
    pub layout: &'a CT_SlideLayout,
    pub slide:  &'a CT_Slide,
    pub default_text_style: &'a CT_TextListStyle,
    list_style_cache:
        RefCell<HashMap<Option<PlaceholderKey>, EffectiveListStyle>>,
}

impl ResolveCtx<'_> {
    fn placeholder_chain(&self, sp: &Shape) -> (Option<&Shape>, Option<&Shape>);
    pub fn effective_xfrm(&self, sp: &Shape) -> Option<Xfrm>;
    pub fn effective_body_pr(&self, sp: &Shape) -> BodyPr;
    pub fn effective_list_style(&self, sp: &Shape) -> EffectiveListStyle;
    pub fn effective_text_properties(
        &self,
        sp: &Shape,
        paragraph: Option<&CT_TextParagraphProperties>,
        run: Option<&CT_TextCharacterProperties>,
    ) -> EffectiveTextProperties;
    pub fn effective_fill(&self, sp: &Shape) -> Option<Fill>;
    pub fn effective_line(&self, sp: &Shape) -> Option<Line>;
    pub fn effective_shape_style(
        &self,
        sp: &CT_Shape,
    ) -> Result<EffectiveShapeStyle, ResolveError>;
    pub fn resolve_color(&self, c: &ColorChoice) -> Color;
    pub fn resolve_typeface(&self, tf: &str, script: Option<&str>) -> String;
}
```

Built once per slide, because the layout and master chain is constant across a
slide's shapes. The presentation default, selected master style, master
placeholder and layout placeholder form a prefix memoised by
`Option<PlaceholderKey>`. `effective_list_style` clones that prefix before
applying the shape's own list style. Shapes that share a placeholder key can
therefore reuse inherited formatting without sharing their direct formatting.

## Testing this

Structural unit tests are necessary but not sufficient, because the failure mode
is "subtly wrong" rather than "wrong". The milestone therefore gates on visual
differential tests against decks whose correct appearance can be eyeballed, plus
a table of roughly 40 theme-colour and transform pairs sampled from real
PowerPoint renders and asserted to resolve to exact RGB values.

The executable visual differential pins python-pptx 1.0.2 and compares slide
size, ordered source shape kind, bounds and text over selected corpus decks.
Resolver-only records additionally retain backgrounds, group transforms,
concrete shape and run paint, diagnostics and unsupported categories. Latent
header-footer policy is asserted on the Rust result because python-pptx exposes
raw placeholder collections rather than native visibility. The native
PowerPoint acceptance record fixes the reviewed application build and original
corpus paths in the integration test.
