# 06, PresentationML model

Owner: `rpptx-oxml`, with the facade in `rpptx`.

## Parts

| Part | Type | Required |
|---|---|---|
| `/ppt/presentation.xml` | `CT_Presentation` | yes |
| `/ppt/slides/slideN.xml` | `CT_Slide` | zero or more |
| `/ppt/slideLayouts/slideLayoutN.xml` | `CT_SlideLayout` | at least one per master |
| `/ppt/slideMasters/slideMasterN.xml` | `CT_SlideMaster` | at least one |
| `/ppt/theme/themeN.xml` | `CT_OfficeStyleSheet` | one per master |
| `/ppt/notesSlides/notesSlideN.xml` | `CT_NotesSlide` | optional |
| `/ppt/notesMasters/notesMasterN.xml` | `CT_NotesMaster` | required if any notes slide exists |
| `/ppt/handoutMasters/handoutMasterN.xml` | `CT_HandoutMaster` | optional |
| `/ppt/authors.xml` | `CommentAuthorList` | optional, relationship-resolved |
| `/ppt/comments/commentN.xml` | `CommentList` | optional, one owner slide |
| `/ppt/tableStyles.xml` | `CT_TableStyleList` | conventional, not fatal if absent |
| `presProps.xml`, `viewProps.xml` | | conventional, not fatal if absent |

**Hard structural constraints.** Every slide has exactly one `slideLayout`
relationship. Every layout has exactly one `slideMaster` relationship. Every
master has a `theme` relationship and a `p:clrMap`. Every notes slide has
exactly one `notesMaster` relationship and exactly one `slide` relationship
back to its source slide. Every notes master has exactly one `theme`
relationship. A source slide has at most one `notesSlide` relationship. A deck
with zero slides is valid, and that is what a template is.

The presentation has at most one modern comment-author relationship. A slide
with modern comments has exactly one Microsoft comments relationship and one
matching `p188:commentRel` reference. One modern comment part belongs to one
slide. Author, comment, and reply ids are unique across the collaboration
graph, and every author reference resolves. Legacy ISO comment parts remain
opaque package content.

`CT_Slide` and `CT_SlideLayout` expose the presence-sensitive
`show_master_shapes` root attribute. An absent `showMasterSp` means true.
`CT_SlideLayout` and `CT_SlideMaster` expose their optional `p:hf` as
`CT_HeaderFooter`, with presence-sensitive slide-number, header, footer, and
date-time flags. Each absent flag means true. Boolean inputs accept both XML
boolean spellings and canonical output uses `1` or `0`.

The visibility attributes read through any PresentationML element prefix and
write at their fixed root locations. The modelled `p:hf` writes with the fixed
`p:` prefix in its schema position. Unmodelled root attributes, header-footer
attributes, and header-footer children retain their original payload and
relative positions.

`CT_Slide` also models the presence-sensitive root `show` attribute. Missing
`show` means visible, while the facade exposes its inverse as `hidden`. Boolean
input accepts both XML spellings and output uses fixed `1` or `0` values.
Unrelated root attributes and children remain preserved.

## Public facade

`rpptx::Presentation` opens a path or byte slice and owns the OPC package, the
typed presentation root, and the ordered slides resolved from `p:sldIdLst`.
Part names are never assumed. The main part, slides, and optional notes slides
are joined through their OPC relationships, including normalized relative
targets. Missing parts, missing or wrong relationship ids, external slide
targets, duplicate notes-slide links, and malformed typed roots return a
concrete facade error.

With the default-off `agile-encryption` feature, the native facade opens
encrypted paths and byte slices with default or caller-supplied
`PackageReadLimits`. It writes encrypted bytes through the shared fixed Agile
profile and publishes encrypted paths through a sibling temporary file and
atomic replacement. With the default-off `digital-signatures` feature, the
native facade exposes package signature reports, verifies a staged view of the
current presentation, and signs PKCS#8 private-key DER with X.509 certificate
DER. Signing commits only a candidate that verifies as cryptographically valid
with complete declared coverage.

The presentation and slide property surface is concrete and borrowed:

```rust
Presentation::slide_size(&self) -> Option<(Emu, Emu)>;
Presentation::set_slide_size(&mut self, width: Emu, height: Emu) -> Result<()>;
Presentation::core_properties(&self) -> Option<&CoreProperties>;
Presentation::core_properties_mut(&mut self) -> &mut CoreProperties;
Presentation::replace_text(&mut self, placeholder: &str, value: &str) -> usize;
Presentation::package_class(&self) -> Result<PresentationPackageClass>;
Presentation::to_bytes_as(&self, class: PresentationPackageClass) -> Result<Vec<u8>>;
Presentation::save_as_package_class(&self, path: impl AsRef<Path>, class: PresentationPackageClass) -> Result<()>;
Presentation::save_as_show(&self, path: impl AsRef<Path>) -> Result<()>;
SlideRef::hidden(&self) -> bool;
SlideRef::has_explicit_background(&self) -> bool;
SlideMut::set_hidden(&mut self, hidden: bool);
SlideMut::set_background(&mut self, fill: Fill) -> Result<()>;
SlideMut::clear_background(&mut self);
```

The native facade also owns the ODP conversion boundary. Import creates a fresh
presentation containing ordered slides, ordinary rectangle shapes and text
boxes, tables, embedded images, slide names, and speaker notes. Export projects
the same subset without mutating the source. Charts, transitions, media,
animation, SmartArt, and unsupported appearance semantics remain outside the
editable ODP subset and produce stable source or model path diagnostics.

The native facade also owns bounded HTML conversion through
`Presentation::from_html` and `Presentation::open_html`. Both return
`HtmlReadResult`, which contains a fresh editable `Presentation` and ordered
`HtmlDiagnostic` values. `HtmlImageResource` supplies caller-owned image bytes.
The supported model is HTML5 document or fragment structure, explicit absolute
CSS geometry, bounded type, class, id, descendant, and child selectors,
formatted text, tables without spanning semantics, images, and safe external
or relative hyperlinks. Unsupported implicit layout, browser layout, resource
fetching, scripts, transforms, table spans, and CSS produce stable DOM-path
diagnostics. Limit or projection failures return `Error::Html` without a
partial result. The surface is additive on the pre-1.0 facade and is available
with `default-template`.

The native facade also owns bounded PDF conversion through
`Presentation::from_pdf_bytes`, `from_pdf_bytes_with_limits`, and `open_pdf`.
`PdfImportMode` selects one preserved full-slide graphic or the editable text,
raster-image, nonzero path, and URI-link subset. Page points convert at exactly
12,700 EMU per point with truncation toward zero. CropBox origin and page
rotation are normalized before projection, and every page must have the same
effective size. Editable dashes require strictly positive PDF arrays at phase
zero or an exactly representable dash boundary. A zero member, interior phase,
or positive member that converts to a zero DrawingML stop produces an ordered
diagnostic and omits affected strokes until a valid dash state or
graphics-state restore. Unsupported safe operators and deterministic font
replacement produce ordered `PdfImportDiagnostic` values. Limit, parse, or projection
failure publishes no presentation. The additive pre-1.0 surface is available
with `render`.

Native collaboration, section, and master-setting access is also concrete and
ordered:

```rust
Presentation::comment_authors(&self) -> &[CommentAuthor];
Presentation::add_comment_author(&mut self, author: CommentAuthor) -> Result<()>;
Presentation::comments(&self, slide_index: usize) -> Option<&[Comment]>;
Presentation::add_comment(&mut self, slide_index: usize, comment: Comment) -> Result<()>;
Presentation::reply_to_comment(&mut self, slide_index: usize, comment_id: &str, reply: CommentReply) -> Result<()>;
Presentation::move_comment(&mut self, slide_index: usize, from: usize, to: usize) -> Result<()>;
Presentation::move_reply(&mut self, slide_index: usize, comment_id: &str, from: usize, to: usize) -> Result<()>;
Presentation::resolve_comment(&mut self, slide_index: usize, comment_id: &str) -> Result<()>;
Presentation::remove_comment(&mut self, slide_index: usize, comment_id: &str) -> Result<()>;
Presentation::sections(&self) -> &[Section];
Presentation::set_sections(&mut self, sections: Vec<Section>) -> Result<()>;
Presentation::notes_header_footer_mut(&mut self) -> Option<&mut CT_HeaderFooter>;
Presentation::handout_header_footer_mut(&mut self) -> Option<&mut CT_HeaderFooter>;
```

Native audio and video package access is concrete and keyed by slide index and
`p:cNvPr/@id`. `MediaInfo` reports `MediaKind`, embedded part metadata or the
exact linked target, poster relationship identity, bounded
`MediaPlaybackSettings`, and ordered `MediaDiagnostic` values. Additions accept
`MediaSourceInput` plus a required `MediaPoster`. `Presentation::media`,
`add_media`, `replace_media`, `extract_media`, and `remove_media` inspect and
mutate the complete owned package graph atomically. Linked media is never
fetched.

Native executable-content access is keyed by normalized source part and
relationship id. `EmbeddedContentInfo` reports OLE object, ActiveX control, or
VBA project kind, source and target parts, relationship identity, content type,
byte length, SHA-256, and absent, present, or invalidated signature state.
`Presentation::embedded_content` inventories the producing relationship graph
in stable order. `extract_embedded_content` returns exact stored bytes.
`replace_embedded_content` retains part identity and metadata, while
`remove_embedded_content` removes only the selected owner and newly unreachable
owned candidates. Both mutations require an explicit preserve-or-remove policy
for invalidated package and VBA signature evidence.

Callers supply GUIDs and RFC 3339 timestamps. Mutation validates identities,
authors, indices, section membership, relationship ownership, and occupied
part paths before committing a serialized and reopened candidate. Resolving
writes the `resolved` status on a thread's top-level comment, so a reply id is
unknown to it. Removing a top-level comment also removes its replies, and
removing a reply leaves its thread in place. A slide keeps its comment part,
relationship, and `p188:commentRel` reference after its last comment goes, so
the part holds an empty `p188:cmLst`. Moving a slide retains its producer
slide id. Removing a slide removes that id from section membership and removes
only collaboration content owned by that slide.

Core properties use the package-level relationship described in
`04-opc-and-packaging.md`. Read access does not dirty the source part. Mutable
access materializes a default model when absent and writes its relationship,
content type, and part on save.

`slides()` and `slide(index)` expose borrowed `SlideRef` handles in producer
slide order. A slide exposes its producer id, optional `p:cSld` name, immediate
z-order shapes, recursive visible text, and optional speaker-note text. Indexed
access returns `Option` and does not panic.

`replace_text` performs literal, non-recursive replacement across contiguous
regular runs in ordinary shapes, nested groups, and table cells. A match keeps
the first matched run's formatting and leaves an unmatched suffix in the last
matched run's formatting. Fields, breaks, and selected alternate-content
fallbacks are boundaries. An empty placeholder changes nothing and returns
zero. The returned count is the number of replaced matches.

The public facade also exposes total title and placeholder lookup and immutable
text-frame, paragraph and regular-run handles. Repeatable shape lookup covers
group children. Mutable shape, text and table handles have consuming nested
accessors so a path-based binding can retain one facade borrow across the
resolved operation without reaching into `rpptx-oxml`.

The owning facade edits the ordered slide collection through three atomic
methods:

```rust
pub fn remove_slide(&mut self, index: usize) -> Result<()>;
pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<()>;
pub fn duplicate_slide(&mut self, index: usize) -> Result<SlideRef<'_>>;
```

Every index is zero-based and must identify an existing slide. `to_index` is
the final index, and moving a slide to its current index changes nothing. A
duplicate is inserted immediately after its source. Collection changes keep
the facade records and `p:sldIdLst` synchronized.

Removal stages the complete change before replacing the live presentation. It
removes the selected slide id, presentation relationship, slide part,
relationship scope, and content-type override. An attached notes part and its
scope are removed with the slide. Matching `p:sld` entries are spliced out of
preserved `p:custShowLst` XML without changing its containers or unrelated
bytes.

Duplication also stages a complete graph. It allocates a new slide part,
producer slide id, presentation relationship, and destination relationship
scope. Internal targets are recomputed relative to the new part, external
target mode is retained, and numeric relationship ids are rewritten in typed
and preserved XML. Equal image bytes reuse the package-wide media part through
the normal `MediaStore`. Notes are copied to a new part whose slide
relationship points back to the duplicate. Custom-show membership is not
copied.

The table style part is typed for layout resolution. A caller may pass its
`CT_TableStyleList` to `ResolveCtx`, which uses either the table's explicit
style id or the part's default id. An absent part or unmatched id leaves direct
cell formatting and DrawingML defaults in force. Package parts outside the
typed model continue to use the normal preservation path.

`ShapeRef` normalizes the six typed shape-tree members to `ShapeKind`. Immediate
group members and the selected `mc:Fallback` view are exposed through child
iteration. Ordinary shapes return their text-body text. Table frames return
row-major cell text with tabs between cells and newlines between rows. Other
shape kinds have no direct text. `ShapeRef` equality is node identity. Two
handles compare equal only when they borrow the same underlying shape-tree
child, rather than when separate shapes happen to contain equal XML.
`ShapeRef::rotation` reads the rotation of the child's own transform and
returns `None` when the child has none, as a placeholder that inherits its
geometry does. `ShapeRef::placeholder_type` returns the explicit
`ST_PlaceholderType` token and `None` when the placeholder omits its type,
without applying an inherited or default type.

`slide_mut(index)` exposes a borrowed `SlideMut` handle. Its `shape(index)`
method retains read access, while `shape_mut(index)` returns a `ShapeMut` for an
immediate z-order child. `ShapeMut::child_mut(index)` recurses through group
children only. The selected `mc:Fallback` view remains read-only.

Position, size, rotation, and name setters support ordinary shapes, pictures,
graphic frames, groups, and connectors. Fill and line setters support ordinary
shapes, pictures, and connectors because those kinds own typed shape
properties. Adjustment mutation supports finite values on preset geometry.
Unsupported shape kinds and unsupported geometry return concrete facade
errors. Indexed access remains total and returns `Option`.

Ordinary shapes expose text mutation through behavior-bearing borrowed
handles:

```rust
ShapeMut::set_text(&mut self, text: &str) -> Result<()>;
ShapeMut::text_frame(&mut self) -> Option<TextFrame<'_>>;
TextFrame::paragraph_mut(&mut self, index: usize) -> Option<TextParagraphMut<'_>>;
TextFrame::add_paragraph(&mut self) -> TextParagraphMut<'_>;
TextParagraphMut::add_run(&mut self, text: &str) -> TextRunMut<'_>;
```

`TextFrame` also reads and replaces whole-frame text. Paragraph handles replace
text, paragraph properties, and bullets. Run handles replace text, character
properties, and the direct Latin font. The typed formatting values are
re-exported by `rpptx`. Structural append returns the newly inserted borrowed
item, and Rust's borrow rules prevent a live nested handle from being
invalidated by another structural mutation.

Whole-frame replacement creates a minimal body when needed and always retains
one paragraph. It preserves existing body properties, list style,
first-paragraph formatting, end properties, and placeholder metadata. Existing
fields and line breaks survive property-only edits. Explicit paragraph text
replacement removes its old run choices. Inserting absent paragraph properties
places them before preserved markup-compatibility run content, while later raw
boundaries and bytes remain unchanged. Unsupported shape kinds return the
normal contextual mutation error, and a shape without a text body returns no
text-frame handle.

Table graphic frames expose concrete borrowed `TableRef` and `TableMut`
handles through `ShapeRef::table` and `ShapeMut::table_mut`. Their cell access
is total and returns `Option`. Table handles expose row and column counts,
column widths, and the first-row, last-row, first-column, last-column,
horizontal-banding, and vertical-banding flags. Cell handles expose plain text,
typed text-frame mutation, direct fill, four optional margins, merge-origin and
continuation state, and span height and width.

Changing a column width uses a checked sum and synchronizes the graphic-frame
width. Merge accepts opposite rectangle corners in either order. It validates
the complete rectangle before changing state, rejects overlap with an existing
merge, migrates typed paragraphs in row-major order, and writes the DrawingML
origin and continuation pattern described in `05-drawingml-model.md`. Split is
valid only on a checked merge origin. It restores span one and clears
continuation flags without redistributing content. Fallible width, merge, and
split operations stage and serialize a table clone before committing it, so an
error leaves the table unchanged.

`SlideMut` also exposes the direct shape construction surface:

```rust
pub enum ConnectorType { Straight, Elbow, Curve }

pub fn add_textbox(
    &mut self, left: Emu, top: Emu, width: Emu, height: Emu,
) -> Result<ShapeMut<'_>>;
pub fn add_shape(
    &mut self, preset: &str,
    left: Emu, top: Emu, width: Emu, height: Emu,
) -> Result<ShapeMut<'_>>;
pub fn add_connector(
    &mut self, connector: ConnectorType,
    begin_x: Emu, begin_y: Emu, end_x: Emu, end_y: Emu,
) -> Result<ShapeMut<'_>>;
pub fn add_group_shape(&mut self) -> Result<ShapeMut<'_>>;
pub fn add_table(
    &mut self,
    rows: usize,
    columns: usize,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
) -> Result<ShapeMut<'_>>;
```

The owning facade adds pictures because media parts and relationships belong to
the presentation package rather than to a borrowed slide handle:

```rust
pub fn add_picture(
    &mut self,
    slide_index: usize,
    image_data: &[u8],
    image_filename: &str,
    left: Emu,
    top: Emu,
    width: Option<Emu>,
    height: Option<Emu>,
) -> Result<ShapeRef<'_>>;
```

With neither extent supplied, `add_picture` probes the image and uses its
native size with a 72-DPI fallback. With exactly one extent supplied, it infers
the other from the pixel aspect ratio and truncates toward zero. With both
supplied, it does not require intrinsic metadata. Unsupported bytes, missing
intrinsic dimensions, out-of-range inference, and an invalid slide index
return contextual errors without changing the presentation.

Picture insertion reuses equal media bytes package-wide and creates or reuses
an internal image relationship in the target slide's own scope. Package,
media-store, and relationship changes remain staged until picture construction
succeeds. The picture receives a tree-wide allocated id and deterministic name,
then its canonical `p:nvPicPr`, relationship-backed `p:blipFill`, and typed
`p:spPr` shell append at top z-order.

An ordinary shape has canonical non-visual properties, a typed transform,
preset geometry, and a minimal text body. `add_shape` keeps the string API but
accepts only names in the generated table of all 187 ECMA preset shapes. An
unknown name returns a contextual error without changing the slide. A textbox
uses `rect`, sets `txBox="1"`, has `a:noFill`, and contains `a:bodyPr`,
`a:lstStyle`, and one required paragraph. An empty group contains the required
`p:nvGrpSpPr` and `p:grpSpPr` shells, with no invented transform or members.

A constructed connector is free-standing and uses `line`, `bentConnector3`,
or `curvedConnector3` for `Straight`, `Elbow`, or `Curve`. Its transform offset
is the componentwise minimum endpoint, its extents are the absolute endpoint
spans, and its horizontal and vertical flips retain endpoint direction. A
horizontal or vertical connector may have one zero extent. A span that cannot
fit in the signed EMU representation returns a contextual error.

A constructed table uses a canonical `p:graphicFrame` with deterministic name
`Table {id}`, a typed transform, the DrawingML table URI, and a rectangular
`a:tbl` payload. The constructor rejects invalid counts and extents before tree
mutation. It appends at top z-order like the other borrowed slide constructors.
The table model owns truncating dimension distribution and assigns each
remainder to the final row or column so the grid matches the frame extent.

Every shape constructor, including the owning picture operation, rescans the
tree immediately before allocation, derives a deterministic producer name from
the allocated id, and appends at top z-order. The append operation shifts
raw-child boundaries at the old trailing position before adding the typed
member. Preserved schema-final content such as `p:extLst` therefore remains
after the new member, while all raw subtrees retain their bytes and relative
positions.

The ignored integration gate
`all_shape_constructors_open_in_powerpoint_without_repair` builds all four
forms in a tree with preserved schema-final extension content. The generated
deck opens without repair in pinned Microsoft PowerPoint 16.104, bundle
16.104.25121423.

The picture native-size comparison uses pinned python-pptx 1.0.2. The ignored
integration gate `added_picture_validates_and_opens_without_repair` confirms
that a generated picture deck validates and opens without repair in the same
pinned PowerPoint bundle.

The text mutation gate
`setting_text_on_placeholder_round_trips_and_renders` clears and then replaces
the same placeholder, saves and reopens the deck, resolves it through the
normal layout path, and compares the blank and changed PNG outputs. It uses
`layout_presentation_deterministic`, so the observed pixel change comes from
bundled or presentation-embedded fonts and never from discovered system fonts.
Placeholder type and `idx` remain unchanged through the mutation.

`to_bytes()` clones the source package, serialises the owned presentation,
slide, and notes roots back to their relationship-resolved part names, and uses
the deterministic OPC writer. Typed edits retain unmodelled attributes and
children in their raw slots and preserve schema child order. Parts outside
those owned roots remain the exact source bytes.

## `presentation.xml`

Carries `p:sldSz` (deck dimensions in EMU), `p:notesSz`, `p:sldIdLst`,
`p:sldMasterIdLst`, and `p:defaultTextStyle` which is the base of the text
inheritance chain.

Slide-size mutation rejects non-positive dimensions before changing the model.
It replaces only `p:sldSz/@cx` and `@cy`, preserving the producer size kind,
unmodelled attributes, unmodelled children, and schema position. An absent
`p:sldSz` is materialized with only the validated dimensions.

**`p:sldIdLst` order is slide order.** There is no separate ordering mechanism.
`CT_Presentation` records the original relationship-id order and reconciles raw
list boundaries against surviving relationship ids during serialization.
Comments and unmodelled children therefore remain anchored when slides move,
are removed, or are inserted.

**`p:sldId/@id` must be at least 256 and at most 2147483647, and unique.** A
value below 256 is a guaranteed repair prompt, and it is the single most common
defect in naive `add_slide` implementations.

`CT_Presentation` types the direct Office 2010 `p14:sectionLst` under the
required section extension URI. Sections retain ordered GUIDs, names, and
slide-id membership. Parsing accepts aliases and inherited default namespace
bindings. Dirty writing uses schema-ordered fixed-prefix children while
retaining unsupported attributes, direct events, descendants, and raw
boundaries. A clear that would strand direct raw list payload fails before
mutation instead of publishing invalid XML.

The PPTX, PPTM, POTX, POTM, PPSX, and PPSM distinction lives entirely in this
part's exact main content type. `PresentationPackageClass` maps those six
values without inspecting a path extension. Ordinary `save` and `to_bytes`
retain the opened class. `to_bytes_as` and `save_as_package_class` change only
the staged output override and leave the live facade unchanged.
`Presentation::save_as_show()` remains a compatibility wrapper for ordinary
PPSX output. A class conversion preserves executable payloads and relationships
and records retained package signature evidence as invalidated.

## Notes parts

A notes slide has this root sequence:

```
p:notes
  p:cSld       required
  p:clrMapOvr? optional
  p:extLst?    optional
```

A notes master has this root sequence:

```
p:notesMaster
  p:cSld       required
  p:clrMap     required
  p:hf?        optional, typed as CT_HeaderFooter
  p:notesStyle? optional, typed as CT_TextListStyle
  p:extLst?    optional
```

A handout master has this root sequence:

```
p:handoutMaster
  p:cSld       required
  p:clrMap     required
  p:hf?        optional, typed as CT_HeaderFooter
  p:extLst?    optional
```

All three roots reuse `CT_CommonSlideData`. They read namespace aliases, write
fixed `p:`, `a:`, and `r:` prefixes, and retain unsupported attributes,
elements, text, CDATA, comments, processing instructions, and document types
in their schema slots. Typed header-footer shells retain the same direct raw
events and expose presence-sensitive slide-number, header, footer, and
date-time flags.

`CT_NotesSlide::notes_text()` walks its shape tree in z-order, including
recursive groups and the selected rendering view of `mc:AlternateContent`.
It includes text only from a `p:sp` whose placeholder effective type is exactly
`body`. An absent `p:ph/@type` therefore qualifies. The placeholder matching
equivalence between `body`, `subTitle`, and `obj` does not broaden extraction.
Slide-image, slide-number, date, footer, header, and notes-master prompt text is
excluded.

Native notes-page export uses the presentation `p:notesSz` as the page size and
emits one page for every source slide, including a slide without a notes part.
Notes-master shapes form the page base. Notes-slide placeholders then match
their master placeholders by explicit index first and compatible type second.
Each match must be unique and consumed once. Missing, unmatched, or ambiguous
ownership fails before output. Slide-image placeholders receive the source
slide as a vector subtree. Header, footer, date, and slide-number placeholders
obey the typed master flags and stored text.

Audience handouts use the same relationship-resolved handout master and expose
the six fixed `HandoutLayout` choices: one, two, three, four, six, and nine
slides per page. The master remains below the slide thumbnails. Each thumbnail
is aspect-fitted, clipped, bordered, and numbered. The three-slide layout adds
five ruled note lines beside each thumbnail.

Each included `p:txBody` contributes run and field text in document order.
`a:br` contributes a newline and paragraph boundaries contribute a newline.
Empty bodies are skipped, and multiple nonempty body placeholders are joined
with a newline.

## The shape tree

```
p:cSld
  p:bg?          slide background
  p:spTree
    p:nvGrpSpPr  required, even though it is nearly empty
    p:grpSpPr    required
    (p:sp | p:pic | p:graphicFrame | p:grpSp | p:cxnSp | mc:AlternateContent)*
```

Document order is z-order. `p:spTree` missing its own `p:nvGrpSpPr` or
`p:grpSpPr` is a repair prompt, as is any `p:sp` without `p:spPr`.

`CT_Background` retains the complete captured `p:bg` subtree as its sole
serialisation source. Its read-only rendering projection distinguishes a
`p:bgPr` DrawingML fill from a `p:bgRef` style index and colour. Projection
parsing uses the namespace bindings inherited from `p:cSld`, so aliased
prefixes remain valid even when their declarations live on the part root.
Attributes, effects, unsupported siblings, and the original child order remain
inside the retained raw subtree and round-trip byte-identically.

Slide background authoring accepts a direct DrawingML `Fill` and writes it as
canonical `p:bg/p:bgPr` before `p:spTree`. Replacing an existing direct fill
preserves the captured `p:bg` and `p:bgPr` attributes and all raw siblings while
changing only the fill subtree. Replacing a `p:bgRef` or unsupported background
choice is rejected. Clearing removes only a direct `p:bgPr` background, so theme
references and opaque producer payloads remain untouched.

An ordinary `CT_Shape` owns its required `p:spPr` as a public boxed
`CT_ShapeProperties`. The allocation keeps recursive group parsing within the
normal test-thread stack even when a master contains deeply nested shapes with
large text bodies. Field access remains direct through `Box` dereferencing.
Modelled transform, geometry, fill, line, and effect children write with fixed
prefixes in schema order. Unsupported attributes and children remain in the
shape-properties ordered raw slots.

The optional `p:style` is typed as `CT_ShapeStyle` behind `CT_Shape::style()`.
Its required line, fill, effect, and font references parse with any bound
prefix and write with fixed `p:` and `a:` prefixes after `p:spPr` and before
`p:txBody`. Root attributes and unsupported children remain in their ordered
raw slots. The style payload shares the shape's existing allocation so typing
it does not expand the recursive `ShapeTreeChild` value.

```rust
pub enum ShapeTreeChild {
    Shape(CT_Shape),
    Picture(CT_Picture),
    GraphicFrame(CT_GraphicFrame),   // tables, charts, SmartArt, OLE
    GroupShape(CT_GroupShape),       // recursive
    Connector(CT_ConnectionShape),
    AlternateContent(Box<CT_AlternateContent>),
}

pub struct CT_AlternateContent {
    raw_xml: Vec<u8>,
    chart_choice: Option<Box<CT_GraphicFrame>>,
    selected_fallback: Option<Vec<ShapeTreeChild>>,
}
```

`CT_AlternateContent` selects ordered typed members only from its one immediate
`mc:Fallback` child. It resolves the branch by namespace URI and uses the same
shape-tree dispatch as `p:spTree` and `p:grpSp`. An immediate `mc:Choice` stays
opaque except for a read-only projection of its first chart-bearing
`p:graphicFrame`. Only the schema-positioned
`p:graphicFrame/a:graphic/a:graphicData` payload can select that projection.
Descendant extension payloads stay opaque. The projection resolves the
`c:chart@r:id` attribute by namespace URI and pairs it with the first immediate
typed `p:pic` fallback.
Other choices are not evaluated. No fallback is valid and produces no
selection. An empty fallback produces an empty selection, while more than one
immediate MC fallback is invalid.

The captured `raw_xml` subtree is the only serialisation source. The selected
fallback is a read-only rendering view and is never written back in place of
the producer XML, so choices, comments, processing instructions, attributes,
entities, whitespace, and fallback content remain byte-identical.

The graphic-data payload is exposed read-only. Typed table payloads retain a
dedicated mutable accessor, while chart and opaque payload bytes cannot be
mutated independently of their relationship projection.

SmartArt graphic data carries a typed `DiagramRelationshipIds` payload for the
data, layout, quick-style, and colour relationships while retaining the exact
`dgm:relIds` subtree. The five diagram part roots live in
`rpptx-oxml::diagram`: `CT_DiagramData`, `CT_DiagramLayoutDefinition`,
`CT_DiagramStyleDefinition`, `CT_DiagramColorsDefinition`, and
`CT_DiagramDrawing`. Readers are prefix-tolerant and namespace-aware. Untouched
parts serialize from their original bytes.

`CT_DiagramLayoutDefinition` and `CT_DiagramColorsDefinition` also expose
doc-hidden read-only render projections. The layout projection retains the
schema-owned nested instruction order and ownership for layout nodes,
iteration, choices, conditions, shapes, presentation mappings, algorithms,
parameters, constraints, and rules. The colour projection retains the
schema-owned colour choice kind, value, and ordered transforms. Parsing is
namespace-aware and schema-position-specific. Prefix aliases are accepted,
while namespace shadows and extension lookalikes remain opaque. These
projections do not expose raw XML or add a mutable object tree.

The data model exposes points, connections, supported node text, optional
background XML, and the schema-position-owned cached drawing id. Dirty writes
retain ordered raw events and unmodelled attributes while emitting `ptLst`,
`cxnLst`, `bg`, `whole`, and `extLst` in schema order. Checked edits reject
duplicate point or connection model ids. Layout, style, colour, cached drawing
id, and cached shape-count projections traverse only their schema-owned paths,
so same-namespace lookalikes in opaque or extension content remain raw.

`Presentation::smart_art` inspects graphic frames in the producing slide,
layout, and master relationship scopes. `set_smart_art_node_text` validates all
four frame roles and any present cached drawing role against their exact
internal relationship types, serializes and reparses only the selected data
part, validates the staged graph, and reopens before commit. A rejected edit
leaves the package unchanged.

Slide duplication remaps all four frame relationship attributes plus the
schema-owned cached drawing id and copies the complete bounded diagram graph.
`transfer_smartart_slide_from` additionally requires an explicit destination
layout, exactly one internal source layout, and a placeholder-free source
slide. Notes, comments, unsupported internal dependencies, relationship-owning
images, and graphs beyond 128 diagram parts reject during preflight. Cycles are
allowed and terminate through the visited-part map. Fresh part allocation and
content-aware image deduplication prevent cross-scope aliasing.

`CT_ConnectionShape` types the connector's required `p:spPr` and its optional
start and end connections. Each present `a:stCxn` or `a:endCxn` carries the
required unqualified shape `id` and connection-site `idx` as `u32` values.
Free-standing, start-only, end-only, and fully connected shapes therefore use
the same model. Unsupported connector locks, style, extensions, attributes,
and children remain in their ordered schema slots and round-trip without being
interpreted.

**`p:cNvPr/@id` must be unique within one `spTree`**, including inside nested
groups, preserved raw members, and every branch of `mc:AlternateContent`. A
`ShapeIdAllocator` retains typed recursive scanning and also performs a
namespace-resolved scan of the complete preserved tree. PresentationML aliases
are accepted, foreign `cNvPr` elements are ignored, and non-selected
compatibility choices still reserve their ids. Allocation starts at 2 because
the tree's own `p:nvGrpSpPr` takes 1, fills unused gaps, and reserves each
result.

## Placeholders

```rust
pub struct PlaceholderKey { pub ph_type: PhType, pub idx: Option<u32> }
```

Matching follows PowerPoint's actual rule, not the obvious one:

- Match on `@idx` when it is present on both sides.
- Otherwise match on `@type`, where **an absent `@type` means `body`**.
- `title` and `ctrTitle` form one equivalence class.
- `body`, `subTitle` and `obj` form a second.

**`idx` is the join key from a slide placeholder to its layout placeholder.**
Renumbering it severs position and formatting inheritance silently. Fresh `idx`
allocation is needed in exactly two other operations: adding a new placeholder to
a *layout*, and copying a slide between presentations whose layouts assign
different indices. Both are separate code paths.

`dt`, `ftr` and `sldNum` are **latent** placeholders. They are drawn from the
layout or master directly and are never cloned onto a new slide, because cloning
them produces duplicate footers.

Notes-page composition overlays a notes-slide placeholder only when its full
key resolves against the notes master. An unmatched notes-slide overlay is
ignored with one source-ordered diagnostic. Matched overlays, ordinary notes
shapes, and the required slide-image placeholder retain their existing
ownership rules. Ambiguous and multiply matched placeholders remain errors.

## Preservation strategy

This is the scope control for a format that is otherwise unbounded.
**Parse only what is rendered or edited. Preserve everything else verbatim**
through `oxml_core::raw_xml::capture_element`.

Preserved as opaque bytes in v1: `p:custShowLst`, legacy comments, ink,
`p:contentPart`, unsupported SmartArt algorithms and unmodelled diagram
content, and unmodelled OLE and ActiveX owner XML. The executable payload API
changes only relationship-owned package bytes or the complete validated owner
range. Untouched owner XML, previews, shared payloads, unrelated producer
orphans, comments, processing instructions, namespace bindings, and attribute
spellings remain exact.

Slide, layout, and master roots expose optional typed `p:timing` and
`p:transition` values. The timing projection covers parallel and sequence
containers, common time nodes, previous, next, start, and end conditions,
shape, slide, and time-node targets, paragraph builds, set and animate
behaviours, effects, motion paths, and unsupported timing nodes. Transition
values expose speed, Office 2010 duration, advance policy, effect parameters,
and Office 2015 morph metadata. Expanded-name readers select supported
`mc:AlternateContent` choices or their fallback. The complete captured subtree
remains the serialization source, and supported mutations replace only their
own attribute value bytes. Unsupported siblings, owner attributes, namespace
bindings, and relationship-bearing content therefore remain byte-identical.

Picture media projects only schema-owned direct children. Standard
`a:audioFile` and `a:videoFile` relationships are read directly from `p:nvPr`.
Office media is read only from the expected direct `p:extLst`, `p:ext` with the
Office media extension URI, and direct `p14:media` position. `r:link` takes
precedence if a producer supplies both link and embed. Trim values belong to
the `p14:media` picture extension, not to common timing nodes. Relationship
attribute replacement uses an in-scope relationships prefix or adds a fresh
nonconflicting binding. Unrelated extensions and nested lookalikes stay raw.

The timing projection also exposes concrete audio, video, common media, and
command values. Shape targets, volume, loop, display policy, trigger order,
and supported play, pause, stop, and numeric seek commands are typed. A
nonnumeric `playFrom()` or `seek()` remains `MediaCommandKind::Other` instead
of failing the timing parse. Timing IDs are allocated from namespace-aware
PresentationML nodes in schema-owned supported and unsupported positions while
foreign elements and raw wrappers are ignored.

Two narrow queries support the timeline resolver without adding a parallel XML
model. `ShapeTreeChild::non_visual_name` returns the selected child's cached
`p:cNvPr/@name`, including a selected chart graphic frame inside
`mc:AlternateContent`. `CT_Timing::condition_has_explicit_target` reports
whether one projected start or end condition carried any explicit target. The
second query distinguishes an absent target from an explicitly unsupported
target without changing the public `TimingCondition` fields. Both queries read
the existing namespace-aware projection, and the captured subtree remains the
serialization source.

Readers reject duplicate or out-of-order modelled timing children. One narrow
PowerPoint compatibility case accepts an attribute-free empty layout
`p:transition` immediately before `p:hf`, keeps both values typed, and writes
them in canonical `p:hf`, `p:transition` order. This is the producer shape in
the pinned `ArtisticEffectSample.pptx` corpus deck.

Modern Office 2021 comment authors, comments, threaded replies, and
`p14:sectionLst` are typed only at the fields callers inspect or mutate. Their
ordered raw sidecars retain unsupported anchors, attributes, namespace
bindings, direct events, children, extension lists, and text-body source bytes.
Readers resolve expanded names. Writers use schema order and a safe fixed
prefix, or fail before changing live state when a producer shadow cannot be
preserved.

An `mc:AlternateContent` model may inspect its selected fallback, an immediate
chart choice, or an immediate timing transition choice. The full captured
subtree remains opaque for serialisation and round-trips byte-identically.
Malformed choices are not parsed merely because they contain a modelled
descendant.

The gate on this is a full-corpus round-trip. Every modelled root parses,
serialises, reparses, and compares structurally. The expected package replaces
each modelled root with those exact canonical serialised bytes. After save and
reopen, every rewritten root must match its expected bytes, while every
unmodelled part must match its original bytes. Content types, relationships,
part names, and part counts must remain structurally unchanged.

Security staging serialises a modelled part only when its retained typed value
has changed. An untouched producer-shaped presentation therefore keeps the
exact signed part bytes, including lexical choices and namespace prefixes.
Relevant typed mutation stages new bytes without deleting retained signature
parts, so verification reports the preserved signature evidence as invalid for
the current presentation.

Executable-content mutations consolidate current typed roots into a staged
package before selecting ownership. OLE discovery accepts only schema-position
graphic frames in slides, loaded layouts, and presentation-listed masters.
ActiveX discovery follows a schema-position control reference through its
properties part to one binary relationship, and VBA discovery starts at the
presentation relationship. Compatibility paths that repeat one OLE
relationship collapse to one logical owner. Distinct relationship ids sharing
one graphic-frame removal range reject as ambiguous.

## Relationship remapping

`rpptx-oxml/src/relmap.rs`, roughly 140 lines:

```rust
/// Re-serialise `raw`, rewriting every attribute in the `r:` namespace whose
/// value matches /^rId\d+$/ through `map`. Preserves everything else
/// byte-for-byte, including comments and processing instructions.
pub fn rewrite_rel_ids(raw: &[u8], map: &HashMap<String, String>) -> Result<Vec<u8>>;
```

Modelled fields such as `CT_Blip::embed` are easy. But `r:embed`, `r:id`,
`r:link`, `r:dm`, `r:lo`, `r:qs` and `r:cs` also occur *inside preserved raw
blobs*. Without this function, a duplicated slide's SmartArt or embedded video
silently points at the source slide's relationships. The preservation strategy
that makes round-tripping possible is exactly what makes deep copy dangerous,
and this is the mitigation.

The typed `dgm:relIds` payload uses the same expanded-name rule and rewrites
its four existing relationship attributes in place, including inherited alias
prefixes. Diagram data separately rewrites only the unqualified `relId` on the
schema-owned `dgm:dataModel/dgm:extLst/a:ext/dsp:dataModelExt` path. It never
invents a drawing attribute or changes same-name content under an opaque
wrapper.

Media replacement does not apply that general raw rewrite to the complete
picture. It changes only the direct standard media relationship attribute and
the schema-owned Office media extension source. Duplication and cross-deck
copy still remap all retained relationship attributes. Their shape-id map also
updates exact PresentationML timing targets `p:spTgt/@spid` and
`p:inkTgt/@spid`, build targets on `p:bldP`, `p:bldDgm`, `p:bldGraphic`, and
`p:bldOleChart`, plus DrawingML connector endpoints. Foreign and qualified
lookalikes remain byte-identical.

## Adding a slide

**Synthesise, do not deep-copy.** This is what python-pptx does and it is
dramatically safer. A synthesised placeholder contains no `r:embed`, so **no
relationship remapping is needed at all** on the common path. It also carries no
stale `a:xfrm` overriding the layout, and no copied creation-id GUIDs.

The sequence:

1. Resolve the layout.
2. Allocate `/ppt/slides/slideN.xml` as `1 + max(existing suffix)`.
3. Build the `CT_Slide` shell. No copied name, no background,
   `p:clrMapOvr` containing `a:masterClrMapping`, and a `p:spTree` with its
   required `p:nvGrpSpPr` and `p:grpSpPr`.
4. Select cloneable layout placeholders: those with a `p:ph`, excluding the
   latent `dt`, `ftr` and `sldNum` types. Non-placeholder layout shapes are not
   cloned, they render through the layout pass.
5. Synthesise one minimal `p:sp` per placeholder. Copy `type` and `idx`
   **verbatim**. Empty `p:spPr` so position inherits. A `p:txBody` containing
   `a:bodyPr`, `a:lstStyle` and **at least one `a:p`**, because zero paragraphs
   violates `minOccurs=1`.
6. Create the slide's relationships: exactly one, to the layout, with a
   **relative** target computed rather than hardcoded, because a template may
   store layouts elsewhere.
7. Register the content-type override.
8. Add a `slide` relationship on `presentation.xml` and append a
   `p:sldId` with `id = max(existing).max(255) + 1`.
9. Return the slide handle.

Deep copy exists only for `duplicate_slide` and cross-deck copy, where it uses
`rewrite_rel_ids`, transfers media with content-hash deduplication, and assigns
fresh `p:cNvPr` ids across ordinary, grouped, and compatibility content. The
same map rewrites `a:stCxn` and `a:endCxn` endpoints so copied connectors target
the copied shapes. Slides and notes both use this narrow shape-tree behavior.

## Validation

`Presentation::validate()` returns issues rather than panicking, and runs
automatically under `debug_assertions` before `save`. It will save more
debugging time than any other three hundred lines in the project.

```rust
pub enum ValidationIssue {
    DuplicateShapeId { slide: usize, id: u32 },
    SlideIdOutOfRange { slide: usize, id: u32 },
    DuplicateSlideId { id: u32 },
    MissingContentTypeOverride { part: String },
    DanglingRelationship { part: String, r_id: String, target: String },
    UnreachableRelationshipTarget { part: String, target: String },
    EmptyTextBody { slide: usize, shape: usize },
    DuplicatePlaceholderIdx { slide: usize, idx: u32 },
    OrphanMedia { part: String },
    CustomShowReference { slide_id: u32 },
    MissingLayoutRel { slide: usize },
    MissingThemeRel { master: usize },
}
```

Collaboration and navigation validation also rejects duplicate ids, unknown
authors, shared comment parts, mismatched slide comment references, external
or wrong-type relationships, duplicate section ids, duplicate section slide
ids, and section membership that names no presentation slide. Facade mutations
stage serialization and reopen, so rejection leaves the live package and typed
model unchanged.

Media mutation validates slide and shape identity, poster bytes, safe filename
and MIME grammar, known container signatures, relationship kind and target
mode, schema-owned source locations, and staged structural reparse. Any failure
leaves the live package, slide model, relationships, and payload bytes
unchanged.

Executable-content mutation validates normalized source identity, unique
relationship ids and namespace-resolved owner attributes, exact relationship
types, internal targets, required parts and content types, complete signature
topology, schema-owned producer positions, and structural reparse. It commits
only after the staged presentation validates. Every inventory or mutation
failure leaves the live presentation and package bytes unchanged.

Mapped to the symptoms they prevent:

| Symptom | Cause |
|---|---|
| "PowerPoint found a problem and needs to repair" | `p:sldId/@id` below 256, or duplicated |
| "PowerPoint can't read this file" | missing content-type override for a new part |
| Slide opens blank, layout lost | relationship target with a leading slash, or the wrong relative depth |
| Repair | duplicate `p:cNvPr/@id` anywhere in one `spTree` |
| Repair | `p:txBody` with zero `a:p` |
| Repair | children out of schema sequence |
| Placeholder loses inheritance | `idx` was altered |

## The bundled template

`Presentation::new()` loads `crates/rpptx/assets/default.pptx` through
`include_bytes!`, behind a `default-template` feature. This differs from rdocx,
which builds an empty document from constructors, and the reason is arithmetic:
the Office theme's `a:fmtScheme` alone is roughly 250 lines of gradient and
effect XML, and each of the eleven standard layouts is another 120. Generating
that from Rust is about 2,500 lines of write-only code that nobody would ever
verify. python-pptx ships a binary template for the same reason.

Contents: 16:9 slide size, one master with the standard colour map and full
`p:txStyles`, the eleven standard layouts, a full theme, `presProps`,
`viewProps`, `tableStyles` defaulted to Medium Style 2 Accent 1, a notes master,
and **zero slides**.

The asset must live under the crate's own directory. A workspace-root `assets/`
compiles locally but is not included in the published `.crate`.

The PresentationML model crates remain version `0.0.0` with `publish = false`.
