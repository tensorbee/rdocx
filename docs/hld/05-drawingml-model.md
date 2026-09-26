# 05, DrawingML model

Owner: `oxml-drawing`. The largest body of genuinely new type work in the plan,
because PowerPoint is roughly 90 percent DrawingML and rdocx has almost none.

## What already exists

Of `crates/rdocx-oxml/src/drawing.rs`, 863 lines, almost all is `wp:`
WordprocessingDrawing: inline and anchor wrappers, text wrapping, relative
positioning. None of that has pptx value.

The one transferable piece is `write_graphic_element`, roughly 60 lines, which
emits `a:graphic > a:graphicData > pic:pic` with `a:blip`, `a:stretch`,
`a:xfrm` with `a:off` and `a:ext`, and `a:prstGeom prst="rect"`. That is about
80 percent of a pptx `p:pic`, and it seeds the picture writer.

`theme.rs`, 335 lines, is pure DrawingML but read-only, has no `a:fmtScheme`,
and its tint and shade maths follow Word's convention. It **stays in
`rdocx-oxml`** for this phase. `oxml-drawing` writes its own canonical theme and
rdocx adopts it later through the `From` adapter.

Word picture authoring keeps the WordprocessingDrawing wrapper in
`rdocx-oxml` and resolves its `r:embed` through the OPC part that owns the
selected story. Authored picture fragments carry local standard `r` and `wp`
bindings when their surrounding story retains producer-shadowed prefixes.
Standalone footnote and endnote roots declare the standard relationship
namespace. The `rdocx` facade may replay a producer root binding, but it keeps
each authored picture namespace-complete and rewrites only the exact expanded
name attributes it owns.
The native run facade can append an inline picture from a relationship created
by `Document::embed_image`. The drawing stays in the caller-selected mixed run
position and carries the supplied point or EMU dimensions through the existing
WordprocessingDrawing wrapper. Relationship creation remains package-owned.

The document facade also authors pictures through `PictureOptions`.
`CT_Inline` and `CT_Anchor` own a typed `SourceRect`, and anchored pictures emit
schema-ordered relative positioning plus none, square, tight, through, top and
bottom wrapping. Tight and through output includes the required wrap polygon.
Text boxes use a WPS DrawingML primary branch with exact rotation and vertical
direction, plus a self-contained VML fallback whose shape type, text spacing,
and vertical flow agree with the selected option. Unsupported producer
`AlternateContent` remains opaque unless the facade authors that complete
fragment itself.

Cross-document body-fragment import treats each selected picture or chart
drawing as the root of a part-local relationship closure. It assigns fresh
package-global drawing identities, rewrites only schema-owned relationship
attributes in the exact retained wrapper, and copies reachable media, chart,
and embedded parts before publication. Repeated import therefore preserves the
producer drawing payload while preventing destination identity collisions.

## Modules

| Module | Contents |
|---|---|
| `order.rs` | Schema child-ordering helper |
| `color.rs` | Colour choices, the transform stack, resolution against a theme and colour map |
| `xfrm.rs` | `a:xfrm`, offset, extent, child offset and extent, rotation, flips |
| `geometry.rs` | `a:prstGeom`, `a:custGeom`, guides, adjust values, path lists |
| `fill.rs` | `a:noFill`, `a:solidFill`, `a:gradFill`, `a:pattFill`, `a:blipFill` |
| `line.rs` | `a:ln`, width, dash, cap, join, head and tail ends |
| `effect.rs` | `a:effectLst`, outer shadow modelled, the rest preserved |
| `shape_props.rs` | `a:spPr` and the group equivalent |
| `style_ref.rs` | `a:lnRef`, `a:fillRef`, `a:effectRef`, `a:fontRef` |
| `table.rs` | `a:tbl`, properties, grid widths, rows, cells, merge spans and continuation flags |
| `text/` | `a:txBody`, `a:bodyPr`, `a:lstStyle`, `a:p`, `a:pPr`, `a:r`, `a:rPr`, `a:t`, `a:fld`, `a:br`, bullets |
| `theme.rs` | `CT_OfficeStyleSheet`, read and write, plus `office_default()` |

## Two traps that are silent until PowerPoint refuses the file

**Child ordering.** OOXML schemas are `xsd:sequence`, so element order within a
parent is part of the contract. Emitting `a:ln` before `a:solidFill` inside
`a:spPr` produces a repair prompt, not a warning. Every writer in this crate
emits in schema order, and `OrderedRawChildren` keeps unmodelled siblings in
their original slots rather than appending them at the end.

**`a:t` whitespace.** Leading and trailing whitespace is significant and needs
`xml:space="preserve"`. Cheap to guard, and infuriating to diagnose later.

## Colour, the part everyone gets wrong

Resolution is three stages, in this order:

1. **The colour map.** `a:schemeClr@val` names a semantic slot such as `bg1` or
   `tx1`. That is mapped through `p:clrMap`, which lives on the slide master and
   may be overridden by `p:clrMapOvr` on a layout or slide. `bg1` usually means
   `lt1`, but **a dark master inverts it.** These are per-master mappings, not
   constants.
2. **The theme.** The mapped slot indexes `a:clrScheme`. Both `a:srgbClr@val`
   and `a:sysClr@lastClr` forms occur.
3. **The transform stack**, applied in document order.

```rust
pub enum ColorTransform {
    Tint(Percent1000), Shade(Percent1000), Complement, Inverse, Gray,
    Alpha(Percent1000), AlphaOffset(Percent1000), AlphaModulation(Percent1000),
    Hue(Angle), HueOffset(Angle), HueModulation(Percent1000),
    Saturation(Percent1000), SaturationOffset(Percent1000),
    SaturationModulation(Percent1000), Luminance(Percent1000),
    LuminanceOffset(Percent1000), LuminanceModulation(Percent1000),
    Red(Percent1000), RedOffset(Percent1000), RedModulation(Percent1000),
    Green(Percent1000), GreenOffset(Percent1000), GreenModulation(Percent1000),
    Blue(Percent1000), BlueOffset(Percent1000), BlueModulation(Percent1000),
    Gamma, InverseGamma,
}
```

**`LumMod` and `LumOff` are mandatory for v1.** Office themes apply them to
virtually every accent-derived colour. Getting them wrong makes an entire deck
the wrong shade, which users notice immediately and cannot describe.

Per ECMA-376 Part 1 section 20.1.2.3, `a:tint` and `a:shade` operate in linear
gamma, and `a:lumMod` and `a:lumOff` operate on HSL luminance. Neither is a
plain sRGB interpolation.

The transform vector is evaluated from left to right. Tint, shade, inverse,
and the absolute, offset, and modulation forms of the red, green, and blue
channels operate in linear light. Hue, saturation, and luminance transforms
operate on HSL components. Gray uses PowerPoint's encoded-channel weights of
0.2126 red, 0.7152 green, and 0.0722 blue. Final channel conversion rounds
exact halves upward. A zero-alpha result is transparent black.

`ResolvedColor` is the observable PowerPoint RGBA contract used by the shared
renderer. For partial alpha, PowerPoint's native shape clipboard PNG first
quantises each channel to an 8-bit premultiplied value and then exposes an
8-bit unpremultiplied RGBA pixel. Resolution reproduces that final round trip
so exact RGBA agrees with the pinned PowerPoint oracle. The focused partial
alpha regression prevents this transport-specific boundary rule from leaking
into the individual colour operations.

### Do not touch the Word path

`rdocx_oxml::theme::apply_tint_shade` takes Word's 0-255 byte convention and
does a naive sRGB interpolation. It feeds `rdocx-layout`'s colour resolution and
therefore every rendered PDF and PNG. **Correcting its maths during an
extraction would silently change every theme-derived colour in rdocx's output,
in a way indistinguishable from a migration bug.**

`oxml-drawing` therefore adds *separate*, spec-correct entry points:

```rust
pub fn apply_tint_shade_pct(hex: &str, tint: Option<Percent1000>, shade: Option<Percent1000>) -> String;
pub fn apply_lum_mod_off(hex: &str, lum_mod: Option<Percent1000>, lum_off: Option<Percent1000>) -> String;
mod color { pub fn srgb_to_linear(f64) -> f64; pub fn linear_to_srgb(f64) -> f64; }
```

The Word function keeps its name and behaviour, documented as "legacy, matches
Word's observed behaviour", so nobody unifies them by accident. Correcting it is
a separate, separately-reviewed change with its own reviewed hash delta.

F-265 gave the Word function its first production caller. `w:color/@w:themeTint`
and `w:color/@w:themeShade` are the same 0-255 byte convention, so run colour
resolution passes them straight into `apply_tint_shade` with its existing
signature. Making it live changed none of its arithmetic, and no call site on
the Word path reaches `apply_tint_shade_pct` or any other spec-correct entry
point. A unit test pins the values it computes, so a later correction has to be
a deliberate change to that test rather than a silent drift.

## Geometry

Preset and custom geometry share one evaluator, which is the decisive argument
for implementing presets properly rather than hand-writing the common twenty.
`a:custGeom` needs the guide machinery regardless, so once it exists the 187
presets are a data problem rather than a code problem.

```rust
pub struct Guide { pub name: String, pub op: GuideOp, pub args: Vec<GuideOperand> }
pub enum GuideOp {
    MulDiv, AddSub, AddDiv, IfElse, Abs, At2, Cat2, Cos, Max, Min,
    Mod, Pin, Sat2, Sin, Sqrt, Tan, Val,
}
```

The enum maps the 17 formula tokens `*/`, `+-`, `+/`, `?:`, `abs`, `at2`,
`cat2`, `cos`, `max`, `min`, `mod`, `pin`, `sat2`, `sin`, `sqrt`, `tan`, and
`val` with their DrawingML argument order. The environment owns guide names and
operands. It is seeded with `w`, `h`, `ss = min(w, h)`, the edges and centres,
the standard fractional width and height guides, and the standard circle-angle
constants. Declared adjust values fall back to their formulas and caller
overrides replace those defaults. Ordinary guides then evaluate in declaration
order. All arithmetic uses `f64`, and angles use 60000ths of a degree. Office
interoperability defines `mod x y z` as the Euclidean norm and `sqrt x` as
`sqrt(abs(x))`. Division by zero and every non-finite result are errors.

Input paths model move, line, cubic, close, and arc commands. Evaluation emits
only move, line, cubic, and close commands. **`a:arcTo` is flattened to cubic
Béziers inside the evaluator**, with each sweep split into segments no larger
than 90 degrees. Positive sweeps are clockwise in DrawingML's downward y-axis,
and the ellipse centre is derived so the arc begins at the current pen
position. A bounded segment count prevents untrusted guide values from forcing
an unbounded allocation. Neither renderer backend sees an arc.

Presets also carry an `<a:rect>` text rectangle, which is needed to place text
correctly inside a non-rectangular shape and would otherwise have to be invented.

The generated table and its provenance are covered in `08-rendering-spec.md`.

## Text body

```
a:txBody
  a:bodyPr     insets, anchor, wrap, vertical direction, autofit
  a:lstStyle
    a:defPPr?  paragraph and run defaults shared by every level
    a:lvl1pPr? through a:lvl9pPr?  overrides for one level
  a:p+
    a:pPr      alignment, level, indents, spacing, bullet
    (a:r | a:br | a:fld)*
      a:rPr    size in centipoints, bold, italic, underline, strike, spacing,
               baseline, typefaces, fill, hyperlink
      a:t      the text
```

Structurally parallel to `w:p` / `w:r` / `w:t` but a different vocabulary, so
this is a rewrite rather than a reuse. Two details that differ meaningfully from
Word: font size is centipoints rather than half-points, and line spacing is a
plain percentage rather than 240ths of a line.

Bullets are `a:buChar`, `a:buAutoNum`, `a:buNone`, with `a:buFont`, `a:buSzPct`
or `a:buSzPts`, and `a:buClr`.

`CT_TextBody` maintains at least one paragraph. Its minimal constructor creates
one empty paragraph, and whole-frame text replacement retains body properties,
the optional list style, and the first paragraph's formatting and end
properties while replacing the ordered text choices with one regular run.
Clearing text therefore leaves one empty paragraph rather than an invalid empty
body. Paragraph and run append operations preserve caller order. Fields and
line breaks remain in place unless the caller explicitly replaces that
paragraph's text.

Typed content transfer moves a non-empty body's paragraphs into another text
body without flattening runs, fields, bullets, or formatting. The source keeps
its body properties and list style and is left with one empty paragraph. A
destination that was empty adopts the moved paragraphs. Otherwise, they append
in their existing order. Preserved body-level children remain at reconciled
schema boundaries.

Paragraph properties, character properties, Latin font, and bullet values use
their existing typed models and schema-order writers. When paragraph properties
are absent, inserting them moves preserved boundary-0 content to the slot after
`a:pPr` and before the first run choice. Later raw boundaries do not move. This
keeps a preserved `mc:AlternateContent` run substitution after the newly
inserted properties without changing its bytes.

`CT_TextParagraphProperties::set_bullet` keeps one member per bullet group. A
typed colour, size, font, or choice removes a preserved `a:buClrTx`,
`a:buSzTx`, `a:buFontTx`, or `a:buBlip` of the same group, and clearing the
bullet removes all four. The writer applies the same rule to a bullet assigned
through the public field. `has_picture_bullet` reports a preserved `a:buBlip`
that no typed choice replaces. A typeface or bullet character that XML 1.0
cannot carry is refused, since the writer escapes markup only.
Line and paragraph spacing percentages accept the transitional
`ST_TextSpacingPercent` range, 0 to 13200000, so a python-pptx deck with more
than two lines of spacing opens. A typed `all_caps` replaces a preserved
`cap="small"` on write, so a character-property element never carries two
`cap` attributes, and `set_all_caps` drops the preserved value for good. `CT_TextBodyProperties` writes a preserved `a:prstTxWarp`,
`a:scene3d`, 3D, or extension child at its schema slot, so an autofit choice
added after parsing still precedes a preserved scene or extension.

The presentation facade projects direct shape offset and extent, non-visual id
and name, and the text body's explicit autofit choice through borrowed handles.
Regular-run handles project the direct Latin typeface, centipoint size, and
sRGB fill without reparsing XML. Run text replacement changes only `a:t`, so
typed character properties and retained foreign children remain attached to
the same run.

`a:pPr/@rtl` is a typed optional boolean direction input. Parsing accepts the
DrawingML attribute only in its unqualified schema form, while a foreign
same-local-name attribute remains opaque. Writing places the canonical typed
attribute with the other paragraph attributes and preserves unknown attributes
and children at their established boundaries. Resolution carries the typed
value beside the paragraph model into shared paragraph-wide bidi shaping, so
numeric and Latin spans follow the explicit base direction without changing
their stored logical text.

`CT_TextListStyle` types `a:defPPr` separately from its nine optional level
properties. The reader accepts any prefix, and the writer emits `a:defPPr`
before the ascending level properties with fixed `a:` prefixes. Unmodelled
siblings remain in their original positions around those typed children.

## Tables

`CT_Table` models optional table properties, the required column grid, and one
or more rows. Each row retains its stored height and ordered cells. Each cell
owns an optional `CT_TextBody`, defaults `rowSpan` and `gridSpan` to one, and
retains `hMerge` and `vMerge` independently. A merge origin is the
non-continuation cell whose row or grid span is greater than one. Continuation
cells stay explicit, including cells that continue a two-dimensional merge in
both directions.

`CT_Table::new` rejects zero counts, counts outside the DrawingML span range,
non-positive extents, and extents too small to give every row and column a
positive size. It creates a rectangular explicit cell grid. Width and height
division truncates toward zero, then assigns the remainder to the final column
or row so the stored grid sums match the requested frame extent. Every cell has
a minimal text body and default cell properties. Constructed tables enable
first-row styling and horizontal row banding, matching python-pptx 1.0.2.

A rectangular merge keeps every cell explicit. The top-left origin stores both
spans. Other cells in the top row retain the row span and set `hMerge`. Other
cells in the left column retain the grid span and set `vMerge`. Interior
continuations set both flags. Non-empty typed paragraph content moves to the
origin in row-major order, and each source keeps one empty paragraph. Splitting
is valid only at an origin, clears this pattern across its checked rectangle,
and does not redistribute the migrated content.

Table properties expose right-to-left order, first and last row and column
flags, row and column banding, and the optional table style id. Unsupported
cell properties remain raw XML at their schema boundary. The rendering subset
types cell margins, DrawingML fills, and the left, right, top, and bottom line
properties. Diagonal borders, effects, 3-D properties, and unsupported wrapper
forms stay opaque and carry unsupported metadata for the resolver. A modelled
fill or line form that the neutral paint model cannot render remains typed and
produces a stable resolver diagnostic.

`CT_TableStyleList` models the optional default style id and the ordered style
records. Each style can expose whole-table, band, edge, and corner regions. A
region models its cell fill or fill reference, text bold and italic state,
theme font reference, text colour, outer borders, and inside horizontal and
vertical borders. Producer wrapper elements such as `a:fill`, `a:tcBdr`, and
its edge children remain part of the modelled schema path. Empty wrappers and
unmodelled siblings retain their original form.

Readers accept any element prefix. Writers use fixed `a:` prefixes and schema
child order for the modelled subset. Table writers emit `a:tblPr`,
`a:tblGrid`, then `a:tr`, with cell text before cell properties. A caller that
extracts `a:tbl` from a larger XML part passes the ancestor namespace bindings
to `from_xml_with_inherited_namespaces`, which keeps opaque producer-prefixed
content namespace-complete in standalone output.

## Theme

`CT_OfficeStyleSheet` is read **and written**, because every `.pptx` requires
`ppt/theme/theme1.xml` and a master without one is invalid.

`a:fmtScheme` is modelled, unlike the current rdocx theme reader which omits it
entirely. Shapes reference it through `p:style`, and `07` covers how the indices
resolve. `CT_StyleMatrix.name` is optional because accepted producer themes in
the corpus omit `a:fmtScheme/@name`. Reading and writing preserve that absence
rather than inventing an attribute. A canonical `a:blip` with `r:embed` or
`r:link` declares the fixed relationship namespace locally, so a modelled fill
remains namespace-valid when its parent did not declare `r`.

`Blip` models the first DrawingML `a:alphaModFix` child as a bounded
`Percent1000` amount from zero through 100000. Reads resolve a conventional
`a` prefix or an alternate prefix declared on the effect, blip, or enclosing
picture fill. The writer emits one canonical `a:alphaModFix` in the raw-child
slot where the modelled effect occurred. Duplicate effects, foreign-namespace
lookalikes, unsupported siblings, and unmodelled attributes remain opaque and
ordered. A missing `amt` uses the schema default of 100000. Malformed or
out-of-range unqualified values reject the fill.

`office_default()` constructs the standard Office theme. It is the correctness
floor for a template whose master lacks a theme relationship. It is *not* how
`Presentation::new()` works, which uses a bundled binary template for the
reasons given in `06-presentationml-model.md`.

The native Word facade re-exports this same `CT_OfficeStyleSheet` for typed
theme authoring. `Document::set_theme` serializes it to the existing related
theme part or creates the missing package edge atomically. Word layout projects
the shared value through `rdocx_oxml::theme::Theme`. The legacy Word tint and
shade helper remains unchanged.

Callers with a concrete palette construct `ColorChoice` through
`ColorChoice::srgb(RgbColor)`. `oxml-chart` re-exports the same `RgbColor`, so
Word and PowerPoint chart facades share one colour value while the DrawingML
writer retains ownership of fixed-prefix `a:srgbClr` serialization.

## Preservation

Anything not in the list above is captured verbatim through
`oxml_core::raw_xml::capture_element` and re-emitted in place. That is not a
fallback, it is the design: parse what you render, preserve the rest.
