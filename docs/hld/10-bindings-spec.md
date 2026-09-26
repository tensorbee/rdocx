# 10, Bindings spec

Owners: `oxml-py-support`, `rdocx-py`, `rpptx-py`, `rdocx-wasm`, `rpptx-wasm`,
`oxml-cli-support`, `rdocx-cli`, `rpptx-cli`.

## The PyO3 lifetime problem

A `#[pyclass]` must be `'static`. The facade is built on borrow handles:
`Paragraph<'a> { inner: &'a mut CT_P }`, plus consuming builders and
`Document::add_paragraph(&mut self) -> Paragraph<'_>` which holds the document
mutably borrowed for the handle's whole life. Python additionally requires that
`p = doc.add_paragraph("x")` stay usable across arbitrary later mutations,
including ones that reallocate the content vector.

References are out, categorically. Four options were weighed:

| Option | Verdict |
|---|---|
| **Index and path handles** re-resolving on every call | **chosen** |
| `Rc<RefCell<_>>` or `Arc<Mutex<_>>` in the core | rejected: rewrites every crate, pollutes the Rust API with borrow noise for users who never touch Python, and `Rc` is not `Send` so `allow_threads` is lost |
| Arena with generational ids | correct long-term, but converts the content vectors across every crate. Deferred |
| A separate owned mirror API | rejected: doubles the API surface, and "attach" reintroduces the identity problem |

### The chosen design

```rust
pub enum PathSeg { Slide(usize), Shape(usize), Body(usize), Row(usize),
                   Cell(usize), Para(usize), Run(usize) }
pub struct ContentPath { pub segs: SmallVec<[PathSeg; 5]>, pub revision: u64 }
pub struct RevisionCounter { current: u64 }

#[pyclass(name = "Document")]
struct PyDocument { inner: rdocx::Document, revision: RevisionCounter }

#[pyclass(name = "Paragraph")]
struct PyParagraph { doc: Py<PyDocument>, path: ContentPath }
```

The Rust API adds only total, index-based paragraph and run accessors needed to
re-resolve these handles. Read-only resolution stays on immutable paragraph
handles so it cannot clear the layout caches. Run setters and structural
mutations retain their required mutable resolution. No interior mutability
leaks into the core.
Aliasing is checked by PyO3's own `RefCell` on the pyclass, so a violation is a
clean `RuntimeError`, never undefined behaviour. Resolution is a handful of
vector index operations, negligible against FFI overhead.

The shared crate carries the Word path variants consumed by the rdocx binding
and the `Slide(usize)` plus repeatable `Shape(usize)` variants consumed by the
rpptx binding.

### The invalidation problem, handled loudly

An index path addresses a **position**, not an object. After
`doc.remove_content(1)`, a handle to paragraph 3 would silently read what used
to be paragraph 4. python-docx does not have this problem because it holds an
lxml element pointer that follows the element.

v0.1 therefore carries a **document revision counter**, bumped after every
successful structural mutation and captured by every handle at construction.
Failed and value-only mutations do not bump it. The shared crate reports a
concrete Rust `StaleElementError` on mismatch. The package binding maps that
domain error to its Python exception with the same revisions and message:

```
rdocx.StaleElementError: paragraph handle was created at document revision 4,
but the document is now at revision 5 (a structural change invalidated it).
Re-fetch it with doc.paragraphs[i].
```

**Loud failure beats silently wrong data.** There are no snapshot accessors that
keep working after invalidation.

v0.2 upgrades to lazily-assigned stable ids backed by `w14:paraId`, which OOXML
already defines for exactly this purpose, so they round-trip to disk and improve
DOCX fidelity as a side effect. Then a handle survives unrelated removals and
matches python-docx semantics, with no API change.

### Two supporting decisions

**Collections are lazy.** `doc.paragraphs` is a pyclass holding only
`Py<PyDocument>` and implementing `__len__`, `__getitem__` with negative and
slice support, and `__iter__`. `Document::paragraphs() -> Vec<ParagraphRef>` is
never called from the binding.

**Consuming builders are bypassed.** A `fn bold(mut self, val: bool) -> Self`
cannot back a Python property setter. The facade exposes 61 non-consuming
`set_*` twins: 24 on `Paragraph`, 19 on `Run`, and 18 across `Table`, `Row`, and
`Cell`. The existing builders delegate to them. The surface is additive, and a
borrowed nested handle can mutate without a rebind:
`doc.paragraph_mut(3).unwrap().add_run("text").set_bold(true)`.

Python paragraph text, run iteration, run indexing, formatting mutation, and
run splitting share the native accepted-view run order. Visible runs inside
inline content controls, insertions, and move destinations carry recursive
source paths. Deleted and move-source runs are absent. A successful structural
edit invalidates earlier path-backed run handles through the same document
revision check as direct handles.

The Rust facade also exposes the minimum automatic-hyphenation authoring
surface. `Document::set_auto_hyphenation` writes the Word document setting,
while `Run::language` and `Run::set_language` assign the direct `w:lang` value.
`Run::set_language_value` assigns or clears it. Omission remains off. These
additions do not create a parallel binding model or make language inference
part of the API contract.

The native pre-1.0 Word reader surface is additive. `Document` reports
document-background and section-layout completeness, resolves one numbering
level, and computes effective paragraph and run properties. Borrowed paragraph,
run, table, row, and cell handles expose hyperlink metadata, drawing and field
kinds, embedded or linked drawing relationships, ordered complex-field display
segments, numbering facts, table formatting, row grid offsets, horizontal
merge state, and unmodelled-content flags. `NumberingFormat`,
`ListLevelSuffix`, `NumberingLevel`, `DrawingKind`,
`DrawingRelationshipKind`, `FieldKind`, and `FieldDisplaySegmentRef` are
concrete native Rust values. Python, WASM, and CLI surfaces do not gain these
reader methods.

At the public low-level Rust boundary, `CT_RPr` includes the complete language
attribute set and its retained foreign attributes, while `LayoutInput` includes
the document automatic-hyphenation boolean. Full struct literals must provide
these fields. These are intentional pre-1.0 source breaks for the next stable
family. Established `TextSegment` construction and layout entrypoints retain
their existing shapes.

The shared native `ChartData` input includes optional category-axis and
value-axis titles plus a typed `Vec<RgbColor>` palette. `oxml-chart`, `rdocx`,
and `rpptx` re-export the same `RgbColor`. `ChartData::default()` supplies empty
optional styling fields for callers using update syntax. The added fields are
an intentional pre-1.0 struct-literal break and add no Python, WASM, or CLI
chart entrypoints.

**Threading.** `Document` remains `Send` and `Sync`. Its normal and
deterministic layouts live in separate
`Mutex<Option<Arc<WordLayoutResult>>>` caches. One private normal-font engine
lives behind a separate mutex and survives result invalidation, with a
compile-time regression gate preserving that threading contract.
`to_pdf`, `render_all_pages` and `to_bytes` run inside `py.allow_threads`, so a
Python thread pool genuinely parallelises work across documents. Concurrent
rendering of one document shares the immutable cached result after the first
layout for that font mode. That is a capability python-docx has no equivalent
for.

## Python API shape

**Drop-in compatibility is an explicit non-goal. Source compatibility for the
documented API is an explicit goal.**

python-docx's real-world surface is inseparable from lxml, and a large fraction
of production code reaches through `._p`, `._r`, `doc.element.body`, `qn()` and
`OxmlElement`. Promising drop-in means promising an lxml-shaped shadow API that
can never be delivered, and every gap then reads as a bug.

The compatibility promise is the completed binding surface, not every public
python-docx method. Its executable gate is an explicit seventeen-example
manifest from the python-docx 1.2.0 Working with Documents, Quickstart, and
Working with Text pages. Each entry records a stable v1.2.0 tagged source URL,
heading, exact source statements, declared transformation, and normalized
structural assertion. Sixteen entries use only the `docx` to `rdocx` import
substitution. The Quickstart held-row example additionally re-fetches
`document.tables[0].rows[1]` before its second cell assignment because the
first cell text replacement advances the global revision and stales the held
row. This is the minimal public compatibility adaptation and does not weaken
strict revision validation. A touch of `._p` raises a clear
`NotImplementedError` naming the attribute and its equivalent rather than an
`AttributeError` five frames away.

```python
from rdocx import Document, Inches, Pt, RGBColor, WD_ALIGN_PARAGRAPH

doc = Document("in.docx")
p = doc.add_paragraph("Hello")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(" world")
r.font.bold = True
r.font.size = Pt(18)
doc.add_picture("img.png", width=Inches(2))   # height inferred by oxml-media
doc.save("out.docx")
doc.save_pdf("out.pdf")                        # documented as an rdocx extension
```

- `font` and `paragraph_format` are themselves handles, so `r.font.bold = True`
  writes through the chain. They store only a document reference and content
  path, re-resolve on every operation, and become stale after a structural
  mutation.
- **Tri-state properties return `None` for inherit**, `True` or `False` when
  explicit. rdocx's `Option<bool>` already matches. Never collapse `None` to
  `False`. This applies to row header and split policy plus cell no-wrap state,
  whose low-level values preserve explicit false forms on round trip.
- `Length` is a pure-Python subclass of `int` and returns EMU, matching
  `docx.shared.Length`, with `.inches`, `.cm`, `.mm`, `.pt`, `.emu` and
  `.twips`. `Inches`, `Cm`, `Mm`, `Pt` and `Emu` are immutable subclasses, and
  `RGBColor` is an immutable three-channel tuple. Float constructors use
  `int(value * factor)`, preserving the truncation toward zero pinned by the
  Rust `Length`. The types are available at the top level and from
  `rdocx.shared`, while native-base inheritance stays outside the Python 3.9
  limited ABI.
- The bounded core enum inventory is pure-Python `IntEnum`:
  `WD_ALIGN_PARAGRAPH` and `WD_UNDERLINE` in `rdocx.enum.text`, plus
  `WD_TABLE_ALIGNMENT` and `WD_CELL_VERTICAL_ALIGNMENT` in
  `rdocx.enum.table`. All four are also top-level exports. Their checked
  integer literals cover the paragraph, run and table variants exposed by the
  S33 facade, including `WD_ALIGN_PARAGRAPH.CENTER == 1`. Underline codes use a
  total binding-oriented facade value accessor rather than expanding the
  published exhaustive Rust `UnderlineStyle` enum.
- The package layer owns `RdocxError(Exception)` as the base, with
  `PackageError`, `XmlError`, `StaleElementError` and `LayoutError` beneath it.
  OPC, I/O and missing-part failures map to `PackageError`, OXML failures map
  to `XmlError`, layout failures map to `LayoutError`, and the shared stale
  domain error maps to `StaleElementError`. `oxml-py-support` therefore remains
  independent of any Python base class.

The S33 formatting inventory is intentionally bounded to font name, size,
colour, bold, italic, underline and strike, plus paragraph alignment, spacing,
indentation, keep-with-next, keep-together, page-break-before and widow
control. Assigning `None` clears direct tri-state formatting. `Paragraph` also
exposes its style ID and its list numbering as a `(num_id, level)` pair, `Run`
its character `style_id`, and `Font` its named Word highlight and separate
shading fill. Highlight reads and writes `ST_HighlightColor` names through
`w:highlight`. Shading accepts six hexadecimal digits or `auto` through
`w:shd`. These setters also clear with `None`. The S33 table
inventory is lazy table, row, cell and nested paragraph lookup, table style,
alignment and width, plus cell text, width and vertical alignment. These
handles use `Body`, `Row`, `Cell`, `Para` and `Run` path segments and reach the
document only through the public `rdocx` facade.

`Table.clone_row(index, at=None)` accepts a negative source index, inserts
after that row by default, and returns the new live `Row` handle. The optional
`at` value is a zero-based insertion boundary. `Table.remove_row(index)` also
accepts a negative row index. Each successful operation advances the document
revision once, stales every earlier structural handle, and publishes only the
native staged result. Python index errors are rejected before mutation, while
native topology and serialization failures use the existing `RdocxError`
mapping. Removing the only direct row is rejected.

The Python `Document` also exposes the current native comparison, main-body
comment, deterministic layout, TOC rebuild, revision, counted replacement, and
field cache update operations. `RunPosition` and
`RunRange` are constructible frozen values for zero-based half-open run ranges.
`StoryRunPosition` and `StoryRunRange` are parallel frozen values whose
`StoryItem` snapshots can identify direct body or table-cell paragraphs.
`Document.add_comment` accepts either range form. The original direct-body
constructors and call shape remain unchanged.
`Comment`, `ComparisonDiagnostic`, `BoundingBox`, `LayoutFragment`,
`LayoutPage`, `TocRebuildReport`, and `Revision` are frozen typed snapshots.
`Document.revisions` lists main-document revisions with a snake_case `kind`,
while the accept and reject methods resolve revisions in every story and return
how many they resolved. `try_replace_text` and `replace_all_regex` return their
replacement counts. `update_fields` takes the native evaluation context as
keyword arguments, reads the wall-clock fields of `now` as given, and returns
the number of updated fields. Comments are
returned as a tuple in package order, comparison diagnostics are returned as a
tuple, and layout fragments are returned in body and page order. No operation
returns a borrowed native handle or an untyped dictionary.

Comparison, deterministic layout, TOC rebuild, path save, and byte
serialization release the GIL. A successful comment addition or reply advances
the document revision once. Comment resolution and removal advance it once
only when they update the document. Comparison and TOC rebuild compare their
serialized package state around the successful staged operation and advance
the revision once only when that state changes. Revision resolution,
replacement, and field updates release the GIL and advance the revision once
only when they report a nonzero count. A native error publishes no
candidate and does not advance the binding revision.

`Document.add_picture` accepts in-memory bytes, a safe filename, optional
paired EMU dimensions, and an optional checked `StoryItem` placement. Omitted
dimensions use native 72 DPI sizing. Omitted placement appends to the body.
The returned immutable `StoryItem` names the inserted paragraph at the new
document revision. A rejected image, filename, dimension pair, or stale path
leaves package bytes and binding revisions unchanged.

`rpptx` mirrors python-pptx through an unpublished mixed-layout `rpptx-py`
crate. `Presentation` owns the Rust facade and one revision counter. Lazy
layouts, slides, shapes, placeholders, text frames, paragraphs, runs, columns
and cells store only a presentation reference and `ContentPath`. The bounded
source-compatibility surface is the seven python-pptx 1.0.2 Getting Started
workflows. They change the import namespace and re-fetch through the public
path after each structural write, because strict global revision invalidation
intentionally stales every pre-write handle and collection. Pure-Python
`Length`, `Inches`, `Pt`, `RGBColor`, and the `MSO_SHAPE`, `MSO_SHAPE_TYPE`,
`MSO_CONNECTOR_TYPE`, and `MSO_FILL_TYPE` enumerations keep native inheritance
outside the limited ABI. `MSO_SHAPE` carries the 181 python-pptx
`MSO_AUTO_SHAPE_TYPE` members whose preset rpptx can author, each with its
preset name as `xml_value`. `UP_ARROW` is absent because the generated preset
table has no `upArrow`.

Presentation `Shape` handles expose optional `Length` values for left, top,
width, and height plus optional non-visual id and name.

`Presentation.slide_width` and `slide_height` read the optional `p:sldSz` as
`Length` values. Assigning one keeps the other, and a deck without `p:sldSz`
pairs the assigned value with the 4:3 size the renderer assumes for such a
deck. `Slide.slide_layout`
returns the layout the slide relates to, equal to the same entry of
`slide_layouts`, and `SlideLayoutCollection.index` returns its position.
`Slide.hidden` reads and writes `p:sld/@show`. `Slide.background.fill` is a
live `FillFormat` over the direct background fill that never changes the slide
when read, and `follow_master_background` reports and sets whether the slide
has no `p:bg`. `SlideCollection.remove` and `SlideCollection.move(from_, to)`
use the native staged slide operations and advance the revision once.

`Shape` geometry, `name`, and `rotation` are writable without a revision bump.
A missing partner coordinate becomes zero, as in python-pptx, and a negative
extent is a `ValueError`. `rotation` reads clockwise degrees normalized below
360 and writes them with round-half-even into the 60000-per-degree angle.
`shape_type` reports an `MSO_SHAPE_TYPE` member or `None`, and a placeholder
picture is a `PICTURE` even when it holds a video, as in python-pptx. `fill` and
`line` return live `FillFormat` and `LineFormat` views for ordinary shapes,
pictures, and connectors, and raise `ValueError` for other kinds. `FillFormat`
offers `type`, `solid()`, `background()`, and `fore_color`. A shape fill of
`a:grpFill` reads as `GROUP`, and setting a fill replaces it. `ColorFormat.rgb`
reads an sRGB colour as `RGBColor`, or `None` for any other colour, and writing
it keeps the transforms of an existing sRGB colour. Reading `LineFormat.color`
changes nothing, and assigning its `rgb` makes any other line fill solid,
pattern included. `LineFormat.width` reads zero without a width, removes the
width when `None` or zero is written, as python-pptx does, and rejects values
above the `ST_LineWidth` maximum. `adjustments` is a live
`AdjustmentCollection` of the effective preset adjustments, normalized so that
1.0 is 100000, and assignment truncates as python-pptx does. `xml` returns the
element serialized on its own as bytes. A picture's `image` is a frozen `Image`
snapshot with `blob`, `content_type`, and the python-pptx `ext`, and
`replace_image` changes only that picture through the native staged
replacement.

`ShapeCollection.add_shape` accepts a DrawingML preset name or an `MSO_SHAPE`
member. `add_connector` follows the python-pptx signature, `add_group_shape`
appends an empty group, and `add_picture` accepts a path, bytes, or a binary
file-like object, which is rewound first when it can seek. `remove` deletes one
shape of a slide with the relationships and parts only it used and advances the
revision once. Nested collections stay read-only.

`TextFrame.autofit`
reports `none`, `normal`, or `shape` when the body carries an explicit choice.
`Run.font` reads the run's direct Latin name, size, and sRGB colour, while the
`Run.text` setter replaces only that run's text and preserves its typed and
unmodelled properties.

The presentation binding exposes `to_pdf`, `render_slide_to_png`,
`render_all_slides`, `to_notes_pdf`, and `render_all_notes` through the native
deterministic facade. Every render call releases the GIL. A `Slide` exposes
optional speaker-note text as a readable and writable property. Assigning it
on a slide without notes creates the notes slide, and the notes master when
the deck has none, through `Presentation::set_notes_text`. A successful
notes assignment publishes the native staged mutation, advances the global
revision once, and makes pre-write handles stale. A rejected assignment leaves
package bytes and revisions unchanged. A `Slide` also exposes an ordered tuple
of frozen `Comment` snapshots. Each comment contains an ordered tuple of frozen
`CommentReply` snapshots, and the presentation exposes an ordered tuple of
frozen `CommentAuthor` snapshots.
Author, comment, and reply additions accept native GUID and RFC 3339 strings.
Comment and reply moves retain native final-position semantics. A successful
collaboration operation advances the global revision once. Constructor or
native validation failure publishes no candidate and leaves existing handles
valid.

## Native Word facade stability

The public `rdocx` facade is the common source for native, Python, WASM, and
CLI consumers. Custom lists are created with `Document::add_list_definition`
from up to nine `ListLevel` values. Each value can select any standard
`ListNumberFormat` plus start, level text, suffix, alignment, indentation,
marker run properties, legal numbering, restart, template code, and tentative
state. The compatibility method still ignores entries after Word's ninth
level. Paragraph numbering stores an explicit list ID and a zero-based level
from 0 through 8. Its in-place setters return `false` without mutation for a
larger value. `Document::set_list_level` can redefine an existing level without
rebuilding the document. A rejected redefinition is side-effect free.

Native Rust also exposes owned `NumberingDefinition`, `NumberingInstance`, and
`NumberingLevelOverride` values plus fallible inspect, create, update, remove,
and whole-graph validation operations. These operations preserve imported
style links but do not mutate them independently.
`Document::link_style_to_numbering` and
`Document::unlink_style_from_numbering` are additive pre-1.0 native Rust APIs.
They atomically mutate the paragraph style and effective numbering level for
one exact style, `numId`, and level tuple. Python, WASM, and CLI bindings gain
no numbering graph authoring surface.

F-248 also adds public fields to the pre-1.0 native Rust projections.
`ResolvedNumbering` exposes `number_current`, `number_level`,
`number_level_without_text`, `number_level_has_ancestor`, `number_full`,
`number_full_without_text`,
`number_context`, `number_suffixes_without_text`, `num_id`, and the hidden
`relative_to` method. `CT_PPr` exposes
`num_ilvl_raw`, `num_id_raw`, `num_pr_extra_attributes`, `num_pr_extra_xml`,
`numbering_revision_xml_positions`, and `numbering_revision_position` so typed
numbering edits can retain exact producer XML. `NumberingState` exposes
`restart_after_section_break`, and `CT_Numbering` exposes
`restarts_after_section_break`. These additions can break exhaustive struct
literals in pre-1.0 Rust consumers. They add no Python, WASM, or CLI surface.
`WordLayoutResult` also exposes `paragraph_numbering`,
`document_paragraph_numbering`, and `document_body_paragraph_numbering` as
result-local native Rust numbering lookups. The latter two remain hidden from
generated documentation but are still additive public Rust APIs.
`WordBodyLayoutFragment` and `WordLayoutResult::body_layout_fragments` add a
public pre-1.0 point-space extent lookup for each direct main-body item. The
record uses one-based physical and displayed page numbers. Preserved unlaid
content resolves to an empty slice, while an invalid body index resolves to
`None`.

Native Rust also exposes `WordPackageClass` for DOCX, DOCM, DOTX, and DOTM.
`Document::package_class` reads the exact main-part override.
`to_bytes_as` and `save_as_package_class` select an output class on a staged
copy without removing executable or opaque parts. `from_flat_opc_bytes`, its
limits overload, `open_flat_opc`, `to_flat_opc_bytes`, and `save_flat_opc`
provide bounded strict Flat OPC interchange through the same `Document` and
`OpcPackage` owners. These are additive pre-1.0 native APIs. Python, WASM, and
CLI bindings preserve opened class identity through their existing saves but
gain no selector or Flat OPC entry point.

Native Rust also exposes the additive pre-1.0 `WordCreationProfile` enum and
`Document::new_with_profile`. `Minimal(WordPackageClass)` preserves the compact
source-built graph. `WordCompatible(WordPackageClass)` owns the standard blank
support parts, and `Document::new()` selects its DOCX form. Macro-capable
profiles select package identity without manufacturing executable content.
Python, WASM, and CLI construction continues through `Document::new()` and
therefore receives the compatible DOCX default without a new selector surface.

Native Rust re-exports `StyleType`, `TableStyleRegion` and
`ConditionalTableStyle`. `StyleBuilder` authors paragraph, character, and table
styles with inheritance, reciprocal links, next styles, UI flags, base
properties including a table style's own row and cell properties, and
conditional table regions carrying all five property layers. `add_style` is
fallible in the pre-1.0 API. `set_style`, `remove_style`,
`set_default_style`, and `validate_style_graph` use the same `Result` boundary.
Builder clear operations remove optional links, UI metadata, base properties,
and conditional regions during a staged update, and
`remove_conditional_table_style` removes exactly one region while its siblings
survive. `Style::conditional_table_styles` returns typed
`ConditionalTableStyle` values rather than the OXML region type, so a caller
can name the return value without taking on schema order.
`StyleBuilder::conditional_table_style` takes a `TableStyleRegion` in place of
the earlier `&str`, which makes an invalid region unrepresentable. That is a
breaking change within the 0.x series, landing in the unreleased 0.14.0. Every
string the earlier form accepted is expressible as a variant, so the typed form
replaces it rather than sitting beside it.
`Table::set_look` writes the six booleans and the equivalent legacy `w:val`
bitmask together, `Table::clear_look` removes the selection, and the checked
`set_row_band_size` and `set_column_band_size` reject a band of no rows or
columns. `Paragraph::set_conditional_formatting` and
`Paragraph::conditional_formatting` select conditional regions through the same
`TableConditionalFormatting` shape a row and a cell already use.
Python, WASM, and CLI retain style package and render behavior without new
style mutation entry points.

Native Rust re-exports `CT_OfficeStyleSheet` and adds the concrete
`FontDefinition`, `EmbeddedFont`, `EmbeddedFontKind`, and
`FontEmbeddingLicense` values. `Document::set_theme`, `theme`,
`set_language_defaults`, `fonts`, `set_font`, `remove_font`, `embed_font`, and
`remove_embedded_font` form the additive pre-1.0 authoring surface. Embedding
requires caller bytes, an explicit authorization value, a nonempty exact
license identity without XML-normalized whitespace, and a valid OOXML font
key. Python, WASM, and CLI gain no new binding in this story.

Native Rust re-exports `CoreProperties`, `AppProperties`, `CustomProperty`,
`CustomPropertyValue`, `Twips`, and the bounded settings value types. `Document`
provides borrowed readers, staged setters, selective removals, and whole-part
removals for these property families. Document variables and compatibility
settings use stable string keys. Default tab stop retains exact integer twips.
`Document::update_fields_on_open` and `set_update_fields_on_open` read, set, or
remove the `w:updateFields` toggle that asks Word to update fields on open.
Python exposes that toggle as the `Document.update_fields_on_open` property and
maps absent or unmodelled producer forms to `None`. Mutating a duplicate or
malformed form reports the existing XML error without changing package bytes.
This is additive pre-1.0 native and Python API. WASM and CLI do not gain new
entry points and otherwise receive preserved package behavior.

Native Word mutations share one private document identifier owner. Existing
method signatures stay unchanged, but fallible operations can report imported,
preserved, overflow, and pending-collision errors before publication. Save and
byte serialization use a staged clone and assign authored relationship,
bookmark, comment, drawing, numbering, part, and content-type identities in
final recursive document order. Producer drawing definitions are checked for
uniqueness per physical XML part, then their package-wide union remains occupied
for globally fresh authored drawing allocation. `CT_Inline` and `CT_Anchor`
expose their parsed `doc_pr_id` on the pre-1.0 Rust model so callers no longer
receive an invented constant for a drawing. Python, WASM, and CLI gain the
deterministic behavior through the native facade without adding binding
methods.

Native Rust also exposes the concrete non-exhaustive `StoryKind` and
`StoryItemKind` enums, owned `StoryId` and `ContentLocation` paths, borrowed
`StoryItemRef` views, and the concrete `StoryError` resolution failures.
`Document::stories` returns body, cell, text-box, header, footer, ordinary note,
and comment owners in deterministic order. `Document::story_items` returns
paragraph, table, content-control, field, drawing, and preserved-node items in
source order. Its `xml` result is borrowed for exact package-backed subtrees and
owned for typed body or comment sources and namespace-complete complex-field
projections. `StoryItemRef::direct_body_index` adds the safe direct body owner
coordinate without changing the recursive `index_path`. Python frozen
`StoryItem` snapshots expose the same optional integer. Items outside the main
story and final section properties expose no coordinate.
`Document::set_story_text` resolves a checked operation-scoped
location against a staged package and publishes only a serialized and reopened
candidate. These additions are native Rust APIs on the pre-1.0 `rdocx` crate.
Python exposes story and story item snapshots, including each item's `xml` as
bytes. `Document.set_story_text` takes a `StoryItem` snapshot and resolves its
story kind, part name, owner index, item kind, and index path against the same
document revision. The compatibility constructor accepts omitted or `None`
XML and materializes it as empty immutable bytes. Cloned fragments omit Word
comment anchors, and story replacement removes only links whose visible
content the replacement emptied. Body text lookup prefers direct paragraphs
over enclosing controls, while `find_content_indices` returns every matching
body coordinate in that priority order. WASM and CLI gain no story traversal
or mutation entry point.

Native Rust also exposes the owned `ContentFragment` value and
`ContentLocation::end`. Paragraph, table, and block content-control
constructors create fixed-prefix fragments, while removal can return one exact
preserved direct child. `Document::insert_content`, `remove_content_at`,
`clone_content`, and `move_content` accept only canonical actual-item locations
or the explicit end boundary. The end boundary is after final direct content
and before body section properties. Moves stay within one story owner. Clones
freshen document identities, and relationship-bearing fragments require their
unchanged owner scope. The Python `Document` binds direct-body
`insert_paragraph`, `pop_content`, `insert_content`, `clone_content`,
`move_content`, and handle-based `find_content_index`. Paragraph and Table
sources must be live direct children of the same Python document. Coordinates
are zero-based insertion boundaries, and a popped `ContentFragment` exposes
only its typed kind and remains reusable because insertion clones the native
value. `try_replace_text` and `replace_all_regex` return exact native counts.
Successful structural mutations stale handles once, successful replacements
stale them only when their count is nonzero, and every rejected operation
leaves package bytes and handle revisions unchanged. The native and Python
changes are additive pre-1.0 surfaces scheduled for `rdocx` 0.14.0. WASM and
CLI gain no corresponding surface. Python resolves clone and move source and
destination locations from one owned story inventory. A rejected clone names
`source` when it is not a Paragraph or Table handle and names `destination` as
a direct body index when it is not an integer.

Native Rust also exposes owned `DocumentFragment` and non-exhaustive
`FragmentConflictPolicy` values. `DocumentFragment::from_range` captures a
nonempty half-open main-body selection and can include final body section
properties only when explicitly requested at the body end.
`Document::import_fragment` inserts at a checked main-body location and applies
caller-selected equivalent reuse independently to styles, numbering, and
related parts. The import remains package-authoritative and transactional.
These are additive pre-1.0 APIs in the published `rdocx` crate. Python, WASM,
and CLI gain no fragment-import surface here.

Native Rust also exposes fallible story-scoped relationship operations on the
same pre-1.0 `Document` facade. `add_picture_to_story` and
`add_hyperlink_to_story` append namespace-complete paragraphs to a checked
`StoryId`. `add_hyperlink_relationship_to_story` allocates one external link
without inserting content. `validate_internal_relationship_for_story`,
`image_data_for_story`, `replace_image_for_story`, and
`hyperlink_url_for_story` resolve only through the story owner's relationship
set and reject stale owners, missing identifiers, wrong types, wrong target
modes, and missing internal targets. These additions do not add a trait,
generic parameter, or WASM or CLI surface. `Document::replace_image` and
`replace_image_for_story` preserve the selected relationship identifier and
drawing XML in a staged mutation. Shared targets use copy-on-write, while a
format change receives a deterministic canonical extension and matching
content type. The old media part is removed only when no relationship still
references it. Python exposes `Document.image_data`, `Document.replace_image`,
and `Document.replace_image_for_story`. A successful replacement keeps live
content handles valid because it changes no modeled content location.
`StoryItemRef::links` inventories modeled links in source order and returns the
existing `LinkInfo` values after checked owner-scoped resolution.
`Document::story_links` pairs those records with their existing
`ContentLocation` owners and merges nested and ancestor-owned links by physical
source position. Python exposes `Document.add_hyperlink_to_story` for a `Story`
snapshot, resolved the same way as a story item, and `Paragraph.add_hyperlink`,
which adds a main-document link relationship and returns the new hyperlink run.

Native Rust also exposes concrete borrowed `SectionRef` and `Section` handles.
Each handle reports its zero-based document ordinal, schema-final ownership,
and configured geometry while retaining access to the complete section
properties. Both handles read page size, orientation, margins, gutter,
equal-width columns, page-number start, header and footer distance, title-page
state, and break type. The mutable handle adds checked setters for every value,
normalizes page dimensions when setting orientation, and rejects invalid or
out-of-range inputs before changing any field. `Document` adds `section_count`,
`sections`, total `section` and `section_mut` lookup, and fallible staged
`insert_section` and `remove_section` operations. Its older final-section
geometry convenience setters remain infallible and unchecked.

Native Rust also exposes non-exhaustive `HeaderFooterKind`, the existing
`HdrFtrType`, and owned `SectionStory`. `Document::section_story` resolves one
effective default, first, or even header or footer and reports its source
section and inherited state. `create_section_story`, `link_section_story`,
`inherit_section_story`, `unlink_section_story`, `replace_section_story`, and
`remove_section_story` are staged fallible operations. Rich content remains
addressed by the returned `StoryId` through the common story API. The facade
also exposes `even_and_odd_headers` and `set_even_and_odd_headers`, while first
story creation enables section `titlePg`. These are additive pre-1.0 native
Rust APIs. Python exposes immutable inspection snapshots, and only the default
header and footer text setters `Document.set_header` and `Document.set_footer`
as mutation entry points. WASM and CLI gain no corresponding binding surface.

`CT_SectPr` adds typed page-number start and raw child-position state, while
`PageFrame` adds `displayed_page_number` beside its physical `page_number`.
These model and handle additions are additive APIs on the published pre-1.0
Rust crates, though exhaustive struct literals can require new fields. The
published `CT_SectPr.header_refs` and `footer_refs` types remain
`Vec<HdrFtrRef>` with the complete native vector surface. Python, WASM, and CLI
gain no section mutation entry point and retain their existing package and
render behavior.

Python `Document.sections`, `styles`, `stories`, `story_items`,
`header_footer_variants`, and `hyperlinks` return tuples of frozen typed
records. Section lengths are integer EMU values. Story records retain kind,
normalized part name, and owner index. Item and hyperlink records retain tuple
index paths, while variant records retain section, source section, inheritance,
kind, and default, first, or even selection. Hyperlink URLs are resolved by the
native checked story API. Their order comes from the story-wide native
projection, including when a nested control precedes a link owned by its
ancestor item. Returned records are detached snapshots, so later document
mutation cannot alter an earlier result. Each `story_items` or `hyperlinks`
accessor materializes one native source and owner inventory for the complete
tuple instead of rebuilding it per returned record.

`Document::rebuild_toc()` is an additive pre-1.0 native Rust operation. It
updates only supported existing main-story TOC fields with deterministic
bundled-font page targets and returns `TocRebuildReport` with entry and newly
allocated bookmark counts plus exact retained-field diagnostics in physical
source order. `diagnostic_count()` is derived from the owned diagnostic
collection. A document without a TOC is unchanged and returns empty counts and
diagnostics. `rdocx-cli toc rebuild` publishes the validated result to an
explicit output and reports the counts through a schema-1 main-story record.
Python exposes the same operation and returns diagnostics as an immutable tuple
with a derived `diagnostic_count` property. WASM does not expose this operation.

The native facade re-exports the concrete OfficeMath tree from `rdocx-oxml`.
`Paragraph::equations`, `Paragraph::equation`, and their read-only equivalents
borrow inline and display equations in source order. Mutable paragraphs add
`equation_mut` and `add_equation`, while `ParagraphItemRef::Equation` keeps the
mixed paragraph item stream ordered. `Document::math_properties` and
`Document::set_math_properties` expose document-wide defaults from the
relationship-resolved settings part. Equations, expressions, arguments, and
document-wide math properties expose a read-only `has_unsupported_content`
query so layout and conversion can diagnose retained content without exposing
the preservation sidecar. The model and accessors are additive on the pre-1.0
Rust surface. Python, WASM, and CLI bindings remain unchanged.

The native facade also exposes `Document::legacy_form_fields` and
`Document::set_legacy_form_field_value` for text, checkbox, and drop-down
legacy fields. Form identity is the normalized source-part path plus its
source-order ordinal within that part. `Document::building_blocks` and
`Document::replace_building_block` expose owned supported projections of
existing relationship-resolved glossary entries, identified by glossary part
and source-order ordinal. Both mutation paths validate a staged package,
reopen it, and commit only after the selected identity and typed value survive.
They do not create entries, execute fields, or expand AutoText. Python, WASM,
and CLI bindings remain unchanged.

The native document renderer copies those defaults into the concrete optional
`rdocx_layout::LayoutInput::math_properties` field. This field addition is a
pre-1.0 source break for native callers that construct `LayoutInput` with a
struct literal. It adds no wrapper, trait, binding method, or command-line
surface.

The additive native conversion surface is `equation_from_mathml`,
`equation_to_mathml`, `equation_from_latex`, and `equation_to_latex`. Imports
return a normalized `MathArgument`, exports return canonical text, and both
carry ordered `MathConversionDiagnostic` records. The boundary accepts a bare
argument, so document-wide and display-paragraph properties stay with their
owners rather than being silently projected. Format attributes and expression
properties that cannot survive the conversion are diagnosed. The surface is
native Rust only and remains additive before 1.0.

Native Rust callers can import RTF through `Document::from_rtf_bytes` and
`Document::open_rtf`. These additive pre-1.0 APIs return an `RtfReadResult`
that carries both the converted `Document` and every stable diagnostic for
content that was safely dropped. Native Rust callers can export the supported
subset through `Document::to_rtf_bytes` and `Document::save_rtf`. The byte API
returns an `RtfWriteResult` with the serialized bytes and every stable
diagnostic for content that could not be represented. The path API serializes
before file I/O, publishes through the shared atomic replacement path, and
returns the same diagnostics after a successful save. The reader and writer
cover text, formatting, tables, lists, and PNG or JPEG images. Unsupported
destinations, visible formatting drops, and lossy writer inputs are reported
instead of hidden. Malformed RTF returns the facade `Error::Rtf` variant
without exposing a second document model. Python, WASM, and CLI surfaces do
not gain RTF entry points implicitly.

Native Rust callers can import HTML5 documents and fragments through
`Document::from_html` and `Document::open_html`. Both return an
`HtmlReadResult` containing the converted `Document` and ordered
`HtmlDiagnostic` values with a DOM location, optional CSS property, and stable
message. Invalid UTF-8, resource-limit violations, and unrecoverable projection
failures return `Error::Html` without publishing a partial document. The path
constructor caps reads at 64 MiB even if the file changes after its metadata is
read. The importer supports source-ordered paragraphs, runs, nested lists, and
spanned tables plus the bounded inline and embedded CSS subset. It does not
fetch external resources. Python, WASM, and CLI surfaces gain no HTML import
entry point and retain their existing methods and error contracts.

Native Rust callers can also insert rich HTML into an existing Word story with
`Document::insert_html_fragment(&ContentLocation, &str,
&[HtmlImageResource])`. The additive `HtmlImageResource` value contains an
exact source key, caller-owned bytes, and a filename. The operation accepts
body, cell, header, and footer destinations supported by the story API. Its
`HtmlFragmentInsertResult` returns the refreshed `StoryId`, the half-open direct
child range, and ordered `HtmlDiagnostic` values. Links and images are scoped
to the physical destination part. Unsupported or unresolved markup is
diagnosed, while invalid locations, malformed resources, bound failures, and
package failures leave the document unchanged. This surface remains native
Rust only.

Native Rust callers can import and export bounded MHTML through
`Document::from_mhtml_bytes`, `Document::open_mhtml`,
`Document::to_mhtml_bytes`, and `Document::save_mhtml`. Concrete
`MhtmlReadResult`, `MhtmlWriteResult`, and `MhtmlDiagnostic` values expose the
converted document or bytes and stable path-aware loss records. Malformed,
ambiguous, unsafe, or over-limit MIME returns contextual `Error::Mhtml` without
publishing a partial result. Export is deterministic and a path save is atomic.
These are additive native pre-1.0 APIs. Python, WASM, and CLI surfaces gain no
MHTML entry point and retain their existing method and error contracts.

Native Rust callers can import OpenDocument Text through
`Document::from_odt_bytes`, `Document::from_odt_bytes_with_limits`, and
`Document::open_odt`. Each returns an `OdtReadResult` containing a fresh
converted `Document` and ordered `OdtDiagnostic` values with stable source
paths. Archive, XML, style, or projection failures return `Error::Odt` with an
optional package part and byte offset without publishing a partial document.
The limits overload applies caller-supplied archive entry, part, and total
expansion bounds. Text, formatting, lists, tables, and supported images are
projected into the existing Word model. Safe lossy skips are diagnosed.
Python, WASM, and CLI surfaces gain no ODT import entry point and retain their
existing methods and error contracts.

Native Rust callers can export OpenDocument Text through
`Document::to_odt_bytes` and `Document::save_odt`. The byte method returns an
`OdtWriteResult` containing deterministic package bytes and ordered
`OdtDiagnostic` values with stable document paths. The path method serializes
completely, stages and syncs a sibling file, then publishes through the shared
portable replacement primitive. A failure cannot truncate an existing
destination. Export does not mutate the source document. Python, WASM, and CLI
surfaces gain no ODT export entry point and retain their existing contracts.

Native Rust callers can import and export OpenDocument Presentation through
`Presentation::from_odp_bytes`, `from_odp_bytes_with_limits`, `open_odp`,
`to_odp_bytes`, and `save_odp`. Read and write result values carry ordered
`OdpDiagnostic` records. Read failures publish no presentation, and path writes
serialize fully before atomic replacement. This is an additive pre-1.0 native
surface. Python, WASM, and CLI bindings gain no ODP entry point.

Native Rust callers can import HTML5 documents and fragments as presentations
through `Presentation::from_html` and `Presentation::open_html`. The additive
pre-1.0 surface exposes concrete `HtmlReadResult`, `HtmlDiagnostic`, and
`HtmlImageResource` values plus `Error::Html`. Conversion returns a fresh
editable presentation only after save, reopen, and validation. Python, WASM,
and CLI surfaces gain no presentation HTML entry point and retain their
existing methods and errors.

Native Rust callers can import PDF pages through
`Presentation::from_pdf_bytes`, `from_pdf_bytes_with_limits`, and `open_pdf`.
The additive pre-1.0 surface exposes concrete `PdfImportMode`,
`PdfImportLimits`, `PdfImportDiagnostic`, and `PdfImportResult` values plus
`Error::PdfImport`. Conversion returns a fresh presentation only after save,
reopen, and validation. Python, WASM, and CLI surfaces gain no PDF import entry
point and retain their existing methods and errors.

Native Rust callers can export EPUB 3 through `Document::to_epub_bytes` and
`Document::save_epub`. The byte API returns `EpubWriteResult`, which carries
the bounded deterministic publication and ordered location-aware
`EpubDiagnostic` values. The path API serializes first, stages beside the
destination, publishes atomically, and returns the same diagnostics. Outline
roots define spine items, content before the first root becomes front matter,
and a document without headings produces one item. Metadata uses stable title
and author fallbacks, while supported text, lists, tables, hyperlinks, and
images retain the established outbound HTML semantics. List identity, restart
values, nested levels, no-marker levels, and standard Roman and letter marker
formats remain distinct. An interrupted ordered list with the same numbering
identity continues from its next value, while nested numbering restarts for a
new parent. A numbered heading remains a heading inside its list item and owns
the navigation anchor. Custom marker text, marker styling, marker alignment,
and list semantics inside a table cell are diagnosed when EPUB list semantics
cannot preserve them. Supported image descriptions become XHTML alternative
text. Heading and navigation labels use only bounded direct projected runs.
Only structurally validated byte-sniffed PNG, JPEG, and GIF media referenced by
surviving body drawings is packaged. Extension fallback is forbidden, and SVG
is diagnosed and omitted. Drawing names, extents, preserved drawing XML,
alternate drawings, preserved Word text spacing, and simplified column breaks
receive stable loss diagnostics. Explicit no-underline formatting remains
non-underlined. Non-basic underline styles, patterned, foreground, or invalid
paragraph, run, and table-cell shading, style-derived deep headings, final
section properties, and document backgrounds are diagnosed at their source
locations. A paragraph without an explicit style receives a diagnostic when
document defaults or its effective default paragraph style carry active
paragraph formatting, or when active run formatting affects direct projected
text. Revision, change, and raw-only default state does not produce noise.
Preserved deleted text reports both spacing normalization and revision
flattening exactly once. Each dropped modeled item,
raw item, field semantic, relationship occurrence, or simplified property has
an ordered source-location diagnostic. This includes dropped named paragraph
style effects, reduced heading levels, and unconsumed document metadata. Typed
and raw views of one revision wrapper produce one diagnostic even when Word and
foreign namespace aliases share a paragraph boundary. Paragraph-local namespace
shadows override document-root bindings during that correlation. Python,
WASM, and CLI surfaces gain no EPUB entry point and retain their existing error
contracts.

The native Word facade provides additive `Document::try_replace_text` beside
the legacy infallible `replace_text` method. The fallible method stages the
replacement and publishes it only after namespace-safe serialization succeeds.
The command-line `replace` operation uses this boundary, reports the stable
serialization error, and creates no partial output. Python and WASM bindings
gain no corresponding method.

The native presentation facade provides the same staged boundary through
`Presentation::try_replace_text`, with exact counts across slides and speaker
notes. `rpptx replace` refuses an existing destination, including its input,
accepts `--expect` for an exact count, rejects an unexpected zero count, and
publishes only after the count and serialization succeed. `--expect 0` is the
explicit zero-match opt-in. Its stable stdout reports the count, literal source
and replacement values, and final output path.

Tagged PDF is an implementation detail of the existing deterministic and
normal PDF methods. Word layout now carries source semantics to the shared PDF
backend, but the native method signatures, returned byte type, binding method
names, and error contracts do not change. Python, WASM, and CLI consumers gain
no semantic-tree API. Presentation PDF methods continue to pass an untagged
layout with no structure tree.

Native Rust callers can request `PdfA2b` or `PdfA3b` through
`Document::to_pdfa_deterministic` and
`Presentation::to_pdfa_deterministic`. Both methods return the PDF backend's
typed conformance error through the facade error enum. These methods are
additive on the native pre-1.0 facades. Python, WASM, and CLI method names and
dependency selections remain unchanged.

The pre-1.0 shared layout API carries semantic types, `MarkedContent`, and
informative `Figure` variants through existing non-exhaustive enums. The image
variants stay unchanged. The `InlineItem::Group` and `LineItem::Group` variants
add an optional baseline field, which is a source break for direct variant
construction. `None` retains the established top-aligned behavior. A finite
baseline aligns nested output with surrounding text and is normalized before
pagination. The figure variant lowers to the one backend-neutral
marked-content carrier rather than creating a second PDF ownership
representation. A backend that consumes `PageFrame::elements` must recurse
through `MarkedContent::children` or use `oxml_layout::walk`.
Wildcard matches remain source compatible but must not discard an unrecognized
container, because visible content can be nested below it.

When native callers enable the default-off `agile-encryption` feature,
`Document::open_encrypted`, `Document::from_encrypted_bytes`, and the bounded
bytes variant open password-protected OOXML through the shared package layer.
`Document::save_encrypted` and `Document::to_encrypted_bytes` write the shared
fixed Agile profile after staging a cloned document and package. A failed save
does not mutate the live document, and the file API publishes through a
sibling temporary file. These additive native APIs are unavailable without
the feature. Python, WASM, and CLI manifests do not enable the feature, so
their API and dependency graphs remain unchanged.

When native callers enable the default-off `digital-signatures` feature,
`Document::verify_signatures` directly returns the shared package verification
reports. `Document::sign` accepts PKCS#8 private-key DER and X.509 certificate
DER on native targets. It flushes typed state into a staged document, asks the
shared package layer to sign and verify the complete graph, and commits only
the verified candidate. The additive APIs distinguish cryptographic validity
and complete declared coverage from certificate-chain trust. They do not
expand Python, WASM, or CLI surfaces and those dependency graphs remain
unchanged.

The native `Presentation` facade exposes the parallel default-off security
surface. `open_encrypted`, `from_encrypted_bytes`, and the bounded bytes
variant use the shared package reader. `save_encrypted` and
`to_encrypted_bytes` use the shared fixed Agile profile, with sibling-file
atomic publication for the path method. `verify_signatures` stages current
typed presentation state before returning shared reports. Native `sign`
accepts PKCS#8 private-key DER and X.509 certificate DER, then commits only a
candidate whose signature is cryptographically valid and has complete declared
coverage. Relevant mutation retains signature infrastructure for inspection
but makes verification report it as invalid. These additive pre-1.0 APIs do
not expand Python, WASM, or CLI surfaces, and their manifests select neither
security feature.

Native callers rebuilding one Word document from another can call
`Document::transfer_reusable_layout_from`. The method moves the source's normal
layout engine only when the complete private retained-work context matches. A
rejected transfer preserves both engines, a successful transfer preserves both
completed result caches, and no unchecked engine accessor becomes public. This
is an additive native Rust method. Python, WASM, and CLI surfaces gain no
transfer method.

Paragraph mutation supports explicit hard breaks and hyperlinks backed by a
document relationship. Native tables expose the additive `TableWidth`,
`TableLayout`, `TableBorderEdge`, `TableLook`, `TableCellMargins`, and
`TableBorderRef` values. Checked setters cover width modes, indentation,
layout, shading, aggregate and individual borders, default cell margins, style
look, and complete active-grid replacement. Matching borrowed readers expose
the stored typed values after reopen.

Grid mutation keeps the table width, every active column, and every covering
cell width consistent. A cell with `gridSpan` receives the sum of its covered
grid columns, and row-level leading and trailing omissions constrain coverage.
Negative or zero column widths, invalid spans or coverage, and overflowing
totals are rejected without mutation. The earlier unchecked compatibility
setters remain available. These are additive pre-1.0 native APIs. Python,
WASM, and CLI do not gain new table-property methods, but their owned
`rdocx::Document` remains package-preserving when native code uses the new
operations.

Native rows and cells also expose the additive `RowHeight`, `CellBorderEdge`,
`CellTextDirection`, and `TableConditionalFormatting` values. Row handles have
checked height plus direct header, split, alignment, and conditional-region
setters. Cell handles have checked width, border, margin, and shading setters,
typed alignment and direction, explicit or absent wrapping, conditional
regions, and checked nested-table construction. Matching borrowed readers
expose every stored typed value after reopen.

Operations whose validity depends on neighbouring cells live on `Table` and
take checked row and cell indexes. Grid omissions reconcile only untouched
empty edge cells. Horizontal spans consume or restore only untouched empty
cells. Vertical continuations require an equal grid range in the immediately
preceding row. Each operation validates a cloned complete table before
publication. These are additive pre-1.0 native APIs. Python, WASM, and CLI gain
no row or cell methods in this story.

Row cloning and removal depend on package-wide identities, so the additive
native operations live on `Document` as `clone_table_row(table, source,
insert_at) -> Result<usize>` and `remove_table_row(table, row) -> Result<bool>`.
They count nested tables in the same document order as `table` and
`table_mut`, retain relationship scope, and validate the complete changed
table before publishing a serialized and reopened candidate. The Python Table
methods resolve their live table path through these operations. WASM and CLI
gain no row mutation entry point.

Native Word table inspection includes additive
`TableRef::has_grid_change()`. It reports whether the low-level grid preserves
one historical `w:tblGridChange` and does not expose that historical snapshot
as active layout input. `CT_TblGrid` carries public historical and unmodelled
raw preservation fields. Full literals written against the earlier pre-1.0
shape must initialize those fields or use `Default`. This intentional low-level
Rust source impact does not add Python, WASM, or CLI methods.

Native Word callers can inspect comments through `Document::comments` and
author threads through `add_comment`, `reply_to`, `resolve_comment`, and
`remove_comment`. The additive native `add_comment_with_date` and
`reply_to_with_date` methods accept an optional validated RFC 3339 timestamp.
The Python `add_comment` and `reply_to` methods expose the same value as the
optional `date` keyword. Omission writes no date and remains deterministic.
Returned ids keep naming the same comment or reply after rdocx save and reopen,
although third-party editors may renumber them. `RunPosition` and `RunRange`
define top-level paragraph run
boundaries with an inclusive start and exclusive end. `Document::split_run`
splits one direct-body run at a Unicode scalar offset of its literal text, so
such a boundary can fall inside what was one run. The second part keeps the run
properties and the enclosing hyperlink. Tabs, breaks, fields, drawings,
references, and preserved raw children have zero width. Zero and the literal
text length return the existing boundary without mutation. Interior success
returns the new continuation index. Python exposes the same method and advances
the binding revision only when a continuation is created. `CommentRef` exposes
comment metadata, text, parent identity, and resolved state without permitting
part-local mutation. `rdocx-cli comment` lists, adds, replies to, resolves, and
removes comments. Add ranges use explicit zero-based, half-open body paragraph
and run coordinates. Every mutation publishes a complete validated document
to an explicit output. Python and WASM keep their package-preserving owners.

Native Word callers remove one exact non-empty literal with
`Document::redact_text`. The returned `RedactionReport` separates Word story,
metadata, chart-cache, and embedded-workbook replacement counts. The method is
additive before 1.0 and commits only a reopened, relationship-valid candidate
whose inflated outer and nested package entries contain no UTF-8 or UTF-16LE
trace. Python, WASM, and CLI surfaces gain no redaction method and continue to
preserve a document already redacted through the native facade.

Native Word callers use `Document::bookmarks` for immutable `BookmarkRef`
summaries and `Document::add_bookmark` for atomic insertion over the existing
top-level half-open `RunRange`. A summary exposes an optional id, name, range,
current text, and marker issue. Insertion validates the Word name and both
boundaries, rejects duplicate or producer-reserved names, and returns the
allocated nonnegative id. The shared recursive `Field` model retains the
complete `REF` and `PAGEREF` instruction, target argument, cached display,
dirty state, source form, and producer XML. These additions are native Rust
APIs only. Python, WASM, and CLI consumers keep their existing surface and
preserve the typed content when they save the owned document.

Native Word callers evaluate fields with `Document::evaluate_fields` and an
explicit `FieldEvaluationContext`. `FieldDateTime` supplies deterministic civil
time. Caller maps supply merge values and included text, including
`source#bookmark` keys for bookmark-scoped includes. Each `FieldEvaluation`
records a snapshot-local document-order index, original instruction, stored
display, and a `FieldOutcome` that is resolved text, pagination deferral, a
structured `TocField`, `TcField`, `MailMergeControl`, or `BarcodeField`, or a
stable stored-display fallback. The explicit context optionally supplies
one-based merge record and output sequence numbers. Formula results remain
text and use the existing formatting switches. Unformatted decimal formulas
use Word-compatible stable decimal display. Structured TOC outcomes retain
optional sequence identifiers and their page-number separators. Structured
outcomes never materialize a generated cache. Evaluation is read-only. It
never reads the ambient clock or filesystem and never changes field caches.
The new public
types are additive. The new public `FieldEvaluationContext` fields are a
pre-1.0 source break for native callers that construct the context with a
struct literal. The new `FieldOutcome` variants are also a pre-1.0 source break
for exhaustive native matches. Python, WASM, and CLI surfaces gain no evaluator
methods and continue to preserve the same package content. Python exposes
`Document.update_layout_backed_fields() -> LayoutBackedFieldUpdateReport` and
the count-only `Document.update_page_fields() -> int` wrapper. The frozen owned
report carries separate PAGE, NUMPAGES, and PAGEREF counts plus an immutable
diagnostic tuple. Both operations release the GIL while native deterministic
layout runs and advance the document revision only when a cache is written, so
handles taken earlier then raise `StaleElementError`. The new `field_source`
member of `oxml-layout`'s `TextSegment`, `GlyphRun`, and
`MultilingualGlyphRun` is a pre-1.0 source break for native struct literals.

Native paragraph item inspection reports whether comment-range and bookmark
marker source elements contained child elements or visible text. Complex-field
display segments expose their effective direct run properties in source order.
Both are additive pre-1.0 Rust reader facts. Python, WASM, and CLI surfaces do
not gain new methods, and their existing exhaustive consumers preserve the new
variants without changing output.

Native Word callers opt into cache materialization with
`Document::update_fields`, `Document::save_with_field_updates`, or
`Document::to_bytes_with_field_updates`. The facade stages the full evaluation
before mutation, updates resolved displays, and marks retained displays dirty.
Existing `save` and `to_bytes` methods continue to preserve intentionally stale
caches and producer dirty spellings. These methods are additive native Rust
APIs. Python, WASM, and CLI surfaces gain no field update methods and continue
to preserve updates already made through their owned `Document`.

Native Word callers merge flat records with `Document::mail_merge` or
`Document::mail_merge_sections`. Each record is a
`BTreeMap<String, String>`. Separate mode returns one complete validated
document per record. Section mode returns one document with record bodies in
input order and a next-page boundary after every non-final record. Empty input
is rejected. Missing merge values become empty text only inside these two
methods. A record-varying merge field in a referenced header, footer, footnote,
or endnote rejects section mode because it combines main-body stories only.
Both methods are additive on the pre-1.0 native Rust facade. Python, WASM, and
CLI surfaces gain no merge methods and continue to preserve documents already
merged by native code.

Native Word callers opt into advanced merge with `MailMergeData`,
`MailMergeRecord`, and `MailMergeValue`, plus owned image and formatter result
types. `Document::mail_merge_rich` returns one reopened document per top-level
record. `Document::mail_merge_sections_rich` combines those validated bodies
through the same section assembly contract as flat merge. Whole-block merge
fields define nested paragraph and row regions. Text values use the established
field switches, images use exact `Length` dimensions, and DOCX fragments import
their internal relationship closure only from a field-only top-level
paragraph. The optional `FnMut` formatter receives lexical source and ordered
field context and may replace text and run properties. Invalid markers, value
kinds, dimensions, XML text, fragments, relationships, or callback results fail
atomically. These types and methods are additive native Rust APIs. The flat
methods and Python, WASM, and CLI surfaces remain unchanged.

Native Word callers render templates with
`Document::render_template(&serde_json::Value)`. Scalar tags use
`{{ path.to.value }}` syntax and may cross ordinary Word run boundaries.
Dedicated marker paragraphs and rows use `{% for item in path %}` with
`{% endfor %}`, or `{% if path %}` with `{% endif %}`. Blocks nest within one
container. Loops require arrays and introduce lexical variables. Conditions
treat false, null, zero, empty strings, empty arrays, and empty objects as
false. Other JSON values are true. Structural generation is limited to the
main body and its tables, while other stories retain scalar rendering. Missing
paths, malformed markers, invalid scalar leaves, invalid numbering references,
and crossed container boundaries fail without mutation. One row loop may own
several adjacent template rows. Each iteration retains table banding, grid and
merge properties, and preserved row and cell XML. Repeated list items retain
one source numbering identity and level, so their sequence continues across
iterations. The existing method remains additive on the pre-1.0 native facade.
Python, WASM, and CLI surfaces gain no template method and continue to preserve
a document rendered by native code.

Native Word callers can also inspect content controls through
`Document::content_controls` and the tag or alias lookup methods.
`ContentControlRef` exposes immutable metadata and display text. Direct setters
update every matching tag or alias, while `bind_content_controls` applies a
string map with tag precedence and alias fallback. Bound values update their
custom XML datastore and displayed text atomically through the
package-preserving facade. These methods are additive native APIs. They do not
implicitly add Python, WASM, or CLI methods, and the existing binding surfaces
remain unchanged.

Native Word callers inspect direct body order through
`Document::body_items`. Each `BodyItemRef` borrows one paragraph, table,
body-level content control, or preserved unsupported XML child. It does not
flatten control content, and it does not change the recursive semantics of
`paragraphs()` or `tables()`. The API is additive on `rdocx` only. Python,
WASM, and CLI surfaces gain no ordered-body method and continue to preserve a
document opened and saved through their existing owners.

Native Word callers inspect direct order below the body through
`CellRef::items`, `ParagraphRef::items`, `HyperlinkRef::items`, and
`RunRef::items`. The non-exhaustive borrowed item enums expose every supported
typed child and each unsupported raw subtree at its original boundary.
`UnsupportedXmlRef` separately reports qualified name, local name, namespace
URI, and whether child content exists. Raw bytes are available only for an
actual preserved raw subtree. Existing flattened accessors remain unchanged,
and Python, WASM, and CLI gain no corresponding methods.

`RunItemRef::LegacyHorizontalRule` identifies the narrow run-level
WordprocessingML `pict` form containing one enabled VML horizontal rule. Its
borrowed accessor returns the exact preserved subtree bytes. Classification is
additive on the existing non-exhaustive Rust enum. Python, WASM, CLI, layout,
and rendering surfaces remain unchanged and continue to preserve the raw XML.

Native Word callers inspect tracked changes through `Document::revisions`.
Each immutable `RevisionRef` exposes the revision id, author, optional
timestamp, and `RevisionKind`. Results recursively cover the main document
body in document order, including tables, cells, and content controls. The
facade reads a typed projection while serialization continues to use the
captured raw WordprocessingML subtree. `rdocx-cli revision list` exposes this
main-story projection with an explicit scope field. Python and WASM load and
save paths preserve the revision XML without a revision inspection method.

Native Word paragraph handles expose
`Paragraph::add_run_inheriting_mark(&mut self, text)`. The method appends one
run whose direct run properties clone the paragraph mark properties, then
returns the ordinary mutable run handle. It is additive on the pre-1.0 Rust
facade. Python, WASM, and CLI surfaces gain no method and retain their existing
package-preserving behavior.

The same mutable `Run` handle exposes additive `add_tab`, `add_break`,
`add_picture`, `add_field`, and `add_symbol` methods. `add_break` accepts the
existing typed line, page, and column inventory. `add_picture` consumes a
relationship already created by `Document::embed_image`. `add_field` rejects
an instruction without a field name before mutation. `add_symbol` stores one
Unicode scalar as text. `set_text` remains the explicit replacement operation,
while formatting setters retain the complete ordered content sequence. These
methods are additive on the pre-1.0 native Rust facade. Python, WASM, and CLI
gain no implicit surface.

`add_symbol` keeps that meaning. `add_symbol_char(font, char_code)` is the
separate method that produces `w:sym`, and `add_special_character` produces
`w:cr`, `w:noBreakHyphen`, `w:softHyphen`, and `w:ptab`. `RunItemRef` gains
`Symbol`, `SpecialCharacter`, and the read-only `LastRenderedPageBreak`, which
has no authoring counterpart because the element is a producer hint.

The same handle authors every `EG_RPrBase` member. `set_slot_font` and
`set_slot_theme_font` take a `RunFontSlot`, and each sets one form and clears
the other for that slot alone. `set_font_hint` is independent and no font
operation clears it. `set_color_theme` authors the theme colour with its tint
and shade and leaves `w:val` as the literal Word caches beside it. Two shipped
setters change behaviour as a correction rather than a deprecation.
`set_font` and `set_font_value` write all four script slots and now clear all
four theme attributes, and `set_color_value` clears the theme colour, tint, and
shade. Before this, Word resolved the theme attribute the caller had left in
place and the authored value silently did nothing. `rdocx::RunProperties` is a
re-export of `CT_RPr`, so its added public members are a pre-1.0 minor source
break on the native facade. Python, WASM, and CLI gain no method.

The low-level `rdocx-layout::TableCell` payload is source-ordered
`Vec<CellBlock>`, with the present paragraph and recursive table variants. The
additional merge-span and cell-margin fields expose renderer input rather than
a second authoring surface. `rdocx-oxml::CT_Style` similarly exposes preserved
table-property bytes, typed table properties, conditional table-style
projections, and schema-positioned extra XML. These are intentional pre-1.0
Rust source breaks. Existing facade and WASM method names do not change.

Native Word callers inspect document protection through the borrowed
`Document::document_protection` accessor. `ProtectionMode` distinguishes
read-only, comments-only, forced tracked changes, and forms-only intent.
`DocumentProtection` also reports the recorded enforcement and formatting
flags, provider type, algorithm class and type, algorithm SID, spin count,
hash, and salt. The accessor reports metadata only. It does not verify a
password or enforce access control. This additive Rust API does not add
Python, WASM, or CLI methods. Those surfaces remain unchanged and preserve the
relationship-resolved settings part when they save their owned document.

The low-level revision and field storage is an intentional breaking pre-1.0
Rust boundary. `RunContent` adds `DeletedText` and replaces the narrow
`FieldType` payload with the recursive `Field`, `FieldInstruction`,
`FieldArgument`, and `FieldSwitch` model. `CT_R`, `CT_P`, `HyperlinkSpan`,
`CT_PPr`, `CT_RPr`, `CT_SectPr`, `CT_TblPr`, and `CT_TrPr` add required
preservation or revision fields, including ordered raw-child sidecars.
`CT_TcPr` also adds an ordered raw-child sidecar that retains external
namespace bindings declared only on the property owner or enclosing cell.
Only WordprocessingML children advance its schema insertion boundary, so a
foreign same-local-name child remains in its source slot. Serialization keeps
`w:textDirection` before preserved `w:tcFitText` and `w:vAlign`. This sidecar
assigns absolute schema slots to the unmodelled standard `w:hMerge`, `w:tcMar`,
`w:hideMark`, `w:headers`, `w:cellIns`, `w:cellDel`, `w:cellMerge`, and
`w:tcPrChange` children. This sidecar is part of the intentional pre-1.0 0.8
low-level Rust source break. Existing exhaustive matches and full struct
literals must be updated or moved to the provided constructors. The workspace
and its exact seven-package stable family are published at 0.8.0, not as a 0.7
patch. Earlier immutable registry versions remain available.
The additive `rdocx::Document` facade and
unchanged Python, WASM, and CLI surfaces do not inherit this low-level source
break.

The low-level layout boundary also adds `source: Option<SourceSpan>` to the
exhaustive public `TextSegment` and `GlyphRun` structs. Existing external
struct literals must supply `None` when they do not own an exact source range.
`rdocx-layout` adds `WordStory`, `WordSourcePath`, and `WordLayoutResult`, plus
normal-font and deterministic provenance entry points. Node ids resolve only
through the result-local Word source table, and ranges use Unicode scalar
indices in the recorded revision view. The existing layout functions keep
returning `LayoutResult`. The `rdocx::Document` facade consumes the provenance
entry points through additive native accessors, while Python, WASM, and CLI
surfaces remain unchanged. The exhaustive literal change is published in both
the incubating 0.4.0 family and the stable 0.8.0 family.

Native callers resolve tracked changes through `accept_all`, `reject_all`, the
exact-author pair, the inclusive RFC 3339 date-range pair, and the id pair.
Each method returns the number of modeled revision elements resolved. Shared
ids select every matching placement, author matching is case-sensitive, and
missing dates do not match a date range. Invalid bounds and malformed selected
changes return an error before mutation. Resolution covers the main document,
headers, footers, comments, normal footnotes, endnotes, and nested text boxes.
`Document::revisions` remains main-story-only. These eight methods are additive
on `rdocx::Document`. `rdocx-cli revision accept|reject` exposes the all-story
resolution boundary with mutually exclusive id, exact-author, or paired date
selectors. An omitted selector resolves all modeled revisions. Python and WASM
continue to preserve the resulting document when they save it.

Native callers generate tracked changes with `Document::compare`, supplying an
edited document, author, and RFC 3339 timestamp. The additive
`ComparisonDiagnostic` value reports stable formatting-only locations and
messages without turning those differences into revisions. Comparison rejects
existing modeled revisions and unsupported structural shell differences, and
it commits only after accepting and rejecting staged copies reproduce their
respective package-wide modeled baselines. `Document::compare` keeps its
source-compatible whole-run default and delegates to the additive
`compare_with_options` method. The concrete `ComparisonOptions` value selects
`Run`, `Word`, or `Character` granularity and left-biased ignores for
formatting, textual whitespace, fields, comments, and any public
`ComparisonStoryKind`. The non-exhaustive story enum names the main, header,
footer, comment, text-box, footnote, and endnote categories. The comparison
surface covers relationship-resolved stories, fields, and nested text boxes.
The native facade stages the main story from package-authoritative XML and
preserves exact unchanged drawing wrappers even when sibling text in the same
paragraph, table, cell, or control changes. Accepting and rejecting the result
retain the drawing payload, relationship graph, and media bytes.
Detached inline and anchor wrappers retain only the inherited namespace
bindings they use and that are not already carried by the story root. Dirty
typed inputs recover matching package drawing payloads before serialization.
Complex fields map every physical source run to one modeled comparison owner,
and sibling fields from one physical run share that owner.
It emits same-story moves and supported run, paragraph, table, and section
property revisions. Diagnostic locations retain the actual story identity and
stable owner path. `rdocx-cli compare` exposes the source-compatible whole-run
comparison with explicit author, RFC 3339 timestamp, and output. Python and
WASM preserve comparison output when they save their owned document.

Native Word rendering exposes `rdocx::RevisionView` and the concrete
`rdocx::RenderOptions`, whose default selects the accepted view. Additive
option-taking counterparts cover PDF bytes and files, single-page and all-page
raster output, page layout, deterministic rendering, and caller-supplied font
paths. The existing methods keep their accepted default. Python, WASM, and CLI
surfaces do not implicitly expose the selector and retain their existing
rendering behavior.
Native selected-image rendering adds zero-based page-list entry points that
share `rdocx::RasterFormat`, `rdocx::RasterOptions` and
`rdocx::RasterOutput` with `oxml-pdf`. The existing PNG methods remain
source-compatible opaque defaults. Python exposes the same image controls as
keyword-only `render_pages` arguments, keeps zero-based page indices, releases
the GIL for rendering, returns `list[bytes]` for PNG or JPEG, and returns one
`bytes` value for TIFF.

Native Word SVG adds `SvgDiagnostic`, `SvgRenderResult`, and four additive
`Document` methods. `render_page_to_svg` and
`render_page_to_svg_with_options` reuse normal layout. Their deterministic
counterparts reuse bundled-font-only layout. Every method takes a zero-based
page index and returns `None` beyond the laid-out document. The result contains
self-contained searchable SVG plus layout-first, path-specific lowering
diagnostics. Python, WASM, CLI, Presentation, and public `oxml-pdf` APIs do not
gain SVG methods or values.

Native renderers obtain the complete positioned output through
`Document::layout` and `Document::layout_with_options`. Accepted calls return a
shared `Arc<WordLayoutResult>` from the normal-font cache, including pages,
font bytes, revision view, and the result-local Word source map. After a
mutation, the retained normal engine may reuse bounded context-independent
paragraph and shaping work while rebuilding the completed result. Tracked calls
stay uncached and use a distinct revision-view paragraph identity.
`Document::layout_with_fonts` and
`Document::layout_with_fonts_and_options` return owned uncached bundles whose
font mapping contains the exact caller-provided bytes selected for shaping.
They construct a caller-only engine and cannot observe the normal process font
snapshot. `Document::layout_with_fonts_and_bundled_fallback` and its
option-taking counterpart return the same owned result shape while retaining a
private reusable deterministic-base engine. Caller faces override bundled
faces, missing families resolve from the bundled inventory, and system fonts
remain unavailable. Differing caller labels act as aliases automatically on
the existing strict and bundled-fallback paths.
`Document::layout_with_fonts_aliases_and_bundled_fallback` and
`Document::layout_with_fonts_aliases_and_bundled_fallback_and_options` add
explicit byte-free aliases to the owned bundled-fallback result paths.
`Document::transfer_reusable_bundled_fallback_layout_from_with_aliases` moves
private work only when the exact caller-font bytes, bounded aliases, and other
retained inputs match. Rejection preserves both engines. Deterministic calls
remain isolated on the bundled-font-only path. The built-in PDF, raster, and
page accessors consume their existing paths. These additions are pre-1.0 native
Rust APIs and do not add Python, WASM, or CLI methods.

Native Word callers measure a checked paragraph or table with
`Document::measure_content(&ContentLocation, Length, RenderOptions)`. Width must
be positive. The owned `ContentMeasurement` reports fractional
`height_points` and ordered `oxml_layout::Diagnostic` values from a fresh
deterministic production engine. Invalid, stale, unsupported, or mismatched
locations return an error. The document bytes and normal, deterministic, and
caller-font cache state remain unchanged. This is an additive pre-1.0 native
Rust API. Python, WASM, and CLI gain no measurement surface.

Native Word callers author watermarks with `Document::set_text_watermark` and
`Document::set_image_watermark`. Text uses fixed Word-like defaults of 468 by
117 points, 315 degree rotation, `D9D9D9`, Calibri, and 50 percent opacity.
Image callers provide positive width and height, while rotation stays zero and
opacity stays at 50 percent. Both methods replace one API-owned watermark in
every active default, first, and enabled even header variant atomically. These
methods are additive on the native pre-1.0 facade. Python, WASM, and CLI gain no
watermark methods and continue to preserve watermarks already authored through
their owned `Document`.

The same native facade exposes `add_picture_with_options`,
`add_text_box_to_story`, and `set_text_watermark_for`. `PictureOptions` selects
crop, size, inline or floating placement, relative axes, wrap, distances,
z-order, and behind-text behavior. `TextBoxOptions` selects bounds, rotation,
direction, and fill or outline colors while the method appends one checked
story paragraph. `TextWatermarkOptions` selects geometry, typography, opacity,
and one default, first, or even header variant. Invalid dimensions, crops,
rotation scaling, story kinds, or missing header variants fail atomically.
Selecting an even variant does not enable even and odd headers. These are
additive pre-1.0 native Rust APIs. Python, WASM, and CLI gain no corresponding
surface.

The public low-level `VmlWatermark` projection and the added paginator section
and header-selection fields are part of the intentional pre-1.0 Rust source
break for the next stable family. They expose renderer input, not a second
authoring surface. Opened header XML remains the serialization authority, and
callers should use the native `Document` methods for mutation.

The pre-1.0 shared layout surface provides multilingual text types for native
renderer producers. Direction, script, clusters, logical source ranges, and
two-dimensional glyph positions are available through the existing rich
values. `TextSegment` includes a required `direction` field so exhaustive
external literals must provide `TextDirection::Auto` when no override exists.
This is an intentional pre-1.0 Rust source break. PowerPoint exposes
resolved paragraph directions through `ResolvedSlideTextDirections` and
sibling resolver and renderer entrypoints. The sidecar leaves the exhaustive
`ResolvedParagraph` shape and all established entrypoints unchanged. Python,
WASM, and CLI surfaces gain no multilingual authoring method. Both WASM graphs
retain their host-font-free target contract while consuming the same bundled
fallback inventory transitively.

Word layout emits the same existing `MultilingualTextSegment` and
`MultilingualGlyphRun` values for paragraphs containing complex scripts. This
activates the existing rich-layout surface for native Word consumers without a
new entrypoint, binding method, or dependency. Low-level Word callers gain
`CT_PPr::bidi`, `ind_start`, and `ind_end`, plus `CT_RPr::rtl` and the paragraph
raw-position sidecar required for exact unknown-child replay. Exhaustive public
Word property literals must add the new fields or use `Default`. These are
intentional pre-1.0 Rust source breaks. Consumers that inspect positioned
elements handle the existing multilingual variant for both Word and
Presentation results.

The stable Rust numbering family includes complete standard number-format
enums and the numbering preservation model. `CT_Lvl`, `CT_AbstractNum`,
`CT_NumLvl`, `CT_Num`, and `CT_Numbering` expose typed and raw XML state so
producer extensions survive typed mutations. `ST_NumberFormat::Other(String)`
retains producer-defined tokens, so the enum is not `Copy` and `to_str` borrows
its value. Completing the exhaustive `ListNumberFormat` enum and extending
`ListLevel` require exhaustive matches and full struct literals to be updated.
Owned format tokens and level properties remove the former `Copy`
implementations from both public types. `NumberingDefinition::levels` uses
explicit `NumberingDefinitionLevel` entries so sparse imported `w:ilvl` values
remain addressable. Full low-level struct literals must add the preservation
fields or use the existing constructors. These are intentional pre-1.0 source
breaks. Python, WASM, and CLI consumers continue through the
package-preserving facade and do not construct these low-level structs.
Existing Python import error mapping uses the generic `RdocxError` exception
and gains no new exception type.

The same intentional low-level pre-1.0 boundary includes retained document
background children, linked drawing relationship ids, numbering style and
producer identifiers with raw namespace context, table, border, row, and cell
raw sidecars, row revision positions, grid offsets, horizontal merge state,
table-property exceptions, and insertion paragraph projection. Exhaustive
`rdocx-oxml` struct literals must provide the new fields or use existing
constructors. These preservation fields do not create new Python, WASM, or CLI
surface.

`CT_TabStop` also exposes `source_occurrence: Option<usize>`. Parsed numbering
tabs use this provenance to retain producer XML on the same occurrence after
an edit, insertion, or removal. New tabs carry `None`, and semantic equality
continues to compare only alignment, position, and leader. The public
`CT_Tabs::from_xml_with_prefixes` parser accepts the in-scope WordprocessingML
prefixes and tracks nested namespace shadows. Paragraph-property namespace
context stays in one internal projection used by numbering, style, body,
table-cell, header, footer, footnote, and endnote readers, so `CT_PPr` does not
expose a partially contextual parser. Established aliased and default
WordprocessingML inputs remain accepted outside numbering.

The Word table facade gains an additive advanced-geometry surface. `Table`
gains checked `set_float_position`, `set_overlap`, `set_bidi_visual`,
`set_cell_spacing`, `set_caption`, and `set_description`. `Row` gains checked
`set_width_before`, `set_width_after`, `set_cell_spacing`, and `set_hidden`.
`TableRef` and `RowRef` gain the matching readers. Each checked setter
validates before publication, so an invalid value leaves the document bytes
unchanged. The public types are `TableFloatPosition`, `TableAnchor`,
`TableFloatX`, `TableFloatY`, `TableTextDistance`, and `TableOverlap`, and the
two alignment payloads reuse the existing `DrawingHorizontalAlignment` and
`DrawingVerticalAlignment` because the vocabularies are identical. The lowered
`TableBlock` gains `bidi_visual` and `TableRow` gains `offset_left`, which are
additive fields on pre-1.0 native Rust projections rather than new binding
surface. Python, WASM, and CLI consumers gain no advanced table surface here.

## Native PowerPoint collaboration and navigation

The native pre-1.0 `rpptx::Presentation` facade exposes ordered modern comment
authors, comments, threaded replies, sections, and mutable notes-master and
handout-master header-footer settings. `CommentAuthor`, `Comment`,
`CommentReply`, and `Section` are concrete values. Callers provide stable GUIDs
and RFC 3339 timestamps, and mutation returns the ordinary facade `Result`
without creating an allocator, clock, trait, generic, or builder.

The additive methods are `comment_authors`, `add_comment_author`, `comments`,
`add_comment`, `reply_to_comment`, `move_comment`, `move_reply`, `sections`,
`set_sections`, `notes_header_footer_mut`, and `handout_header_footer_mut`.
They remain native Rust only. Python, WASM, and CLI consumers gain no
collaboration or navigation methods and continue to preserve these package
parts through their existing `Presentation` owner.

The low-level `rpptx-oxml` model adds the approved `comments` module and
extends existing presentation, notes, slide, relationship, and content-type
models. This is an additive semver change for the published pre-1.0
`rpptx-oxml` and `rpptx` crates. It adds no production dependency or feature
flag. Unsupported modern comment XML and all legacy comment parts remain
preserved, so consumers do not need a parallel raw authoring API.

## Native PowerPoint SmartArt model

The published pre-1.0 `rpptx` facade exposes concrete `SmartArtInfo` and
`DiagramPart<T>` values through `Presentation::smart_art`. The five concrete
diagram part instantiations expose bounded data-model points and connections,
layout family evidence, quick-style labels, colour labels, and cached drawing
shape counts. Missing, external, wrong-type, malformed, and parsed resource
states remain explicit rather than collapsing into an optional raw payload.

Native callers edit supported node text with
`Presentation::set_smart_art_node_text`. They may copy one placeholder-free
SmartArt slide between presentations with
`Presentation::transfer_smartart_slide_from`, supplying an explicit
destination layout index. Both operations validate relationship roles and
stage the complete package change before commit. Transfer is intentionally
bounded to one source layout, the five SmartArt relationship types, and
relationship-free internal images.

The published pre-1.0 `rpptx-oxml` crate exposes the concrete `diagram` module,
and `oxml-opc` exposes the diagram relationship constants. These additions are
native Rust APIs only. Python, WASM, and CLI consumers gain no SmartArt methods
and continue to preserve presentations already edited or transferred through
the native owner. No production dependency, feature flag, trait, dynamic
dispatch, generic parameter, or builder is added.

The low-level diagram definitions include doc-hidden read-only layout and
colour render projections for the native renderer. They expose typed nested
instruction ownership and typed colour choices with transforms, not raw XML or
mutation. F-220 adds no facade, binding, `rpptx-layout`, or `rpptx-render`
public surface.

## Native PowerPoint notes and handout export

The published pre-1.0 `rpptx` facade exposes `HandoutLayout::{One, Two, Three,
Four, Six, Nine}` and four render-feature methods:

```rust
Presentation::to_notes_pdf_deterministic(&self) -> Result<Vec<u8>>;
Presentation::notes_page_pngs_deterministic(&self, dpi: f64)
    -> Result<Vec<Vec<u8>>>;
Presentation::to_handout_pdf_deterministic(&self, layout: HandoutLayout)
    -> Result<Vec<u8>>;
Presentation::handout_page_pngs_deterministic(
    &self,
    layout: HandoutLayout,
    dpi: f64,
) -> Result<Vec<Vec<u8>>>;
```

These additions are native Rust APIs only. Python, WASM, and CLI surfaces add
no notes or handout methods and continue to preserve the underlying parts. No
new public surface is added to `rpptx-layout`, `rpptx-render`, or the OXML
crates. The additive facade API is reviewed through the pre-1.0 release gate.

## Native PowerPoint executable-content inventory

The published pre-1.0 `rpptx` facade exposes concrete `EmbeddedContentKind`,
`EmbeddedSignatureState`, `EmbeddedMutationPolicy`, and
`EmbeddedContentInfo` values. `Presentation::embedded_content` inventories
relationship-owned OLE, ActiveX, and VBA payloads without parsing or executing
them. `extract_embedded_content` returns exact stored bytes.
`replace_embedded_content` and `remove_embedded_content` use the normalized
source part and relationship id as identity and commit only a validated staged
package. The explicit mutation policy either retains invalidated package and
VBA signature evidence or removes only its validated infrastructure.

The published pre-1.0 `oxml-opc` crate adds Transitional and Strict OLE and
control relationship constants plus ActiveX binary, VBA project, and legacy and
Agile VBA signature constants. These are additive native Rust APIs. Python,
WASM, and CLI consumers gain no executable-content methods and continue to
preserve these payloads through the existing presentation owner. No feature,
trait, generic parameter, dynamic dispatch, wrapper identifier, crate, or
binary fixture is added.

## Native Word executable-content inventory

The published pre-1.0 `rdocx` facade exposes the concrete
`EmbeddedContentKind`, `EmbeddedSignatureState`, `EmbeddedMutationPolicy`, and
`EmbeddedContentInfo` values. `Document::embedded_content` returns stable
source-part and relationship identities, normalized target parts, resolved
content types, exact byte lengths, SHA-256 hashes, and signature state for
relationship-owned OLE, ActiveX, and VBA payloads.
`extract_embedded_content` returns exact stored bytes.
`replace_embedded_content` and `remove_embedded_content` stage, validate,
serialize, reopen, and re-inventory the complete package before commit. Their
explicit policy either preserves signature bytes as invalidated evidence or
removes only validated package and selected VBA signature infrastructure.

This is additive native Rust API in the existing facade. Python, WASM, and CLI
consumers gain no executable-content methods and retain their existing opaque
round-trip behavior. No public OXML API, feature, trait, generic parameter,
dynamic dispatch, wrapper, crate, or binary fixture is added.

## Native PowerPoint media model

The published pre-1.0 `rpptx` facade exposes concrete native Rust media values:
`MediaInfo`, `MediaLocation`, `EmbeddedMediaInput`, `MediaSourceInput`,
`MediaPoster`, `MediaPlaybackSettings`, `MediaPlaybackTrigger`, and
`MediaDiagnostic`. `MediaKind` is the concrete audio or video discriminator.
The facade methods are:

```rust
pub fn Presentation::media(&self, slide_index: usize) -> Result<Vec<MediaInfo>>;
pub fn Presentation::add_media(
    &mut self,
    slide_index: usize,
    kind: MediaKind,
    source: MediaSourceInput<'_>,
    poster: MediaPoster<'_>,
    left: Emu,
    top: Emu,
    width: Emu,
    height: Emu,
    settings: MediaPlaybackSettings,
) -> Result<ShapeRef<'_>>;
pub fn Presentation::replace_media(
    &mut self,
    slide_index: usize,
    shape_id: u32,
    source: MediaSourceInput<'_>,
) -> Result<()>;
pub fn Presentation::extract_media(
    &self,
    slide_index: usize,
    shape_id: u32,
) -> Result<Option<Vec<u8>>>;
pub fn Presentation::remove_media(
    &mut self,
    slide_index: usize,
    shape_id: u32,
) -> Result<()>;
```

Embedded sources require bytes, a safe filename, and an explicit safe content
type. Linked sources retain their exact external target and are never fetched.
Add requires a validated poster image. Mutations preserve raw XML, schema
order, relationship ownership, shared payloads, shape identity, geometry, and
failure atomicity. Unknown safe media types remain opaque, extractable, and
diagnostic.

The published pre-1.0 `rpptx-oxml` picture and timing modules expose concrete
media projections. Trim start and end belong to the Office picture extension.
`CommonMediaNode` does not carry trim fields, and `CT_Timing::add_media` accepts
only timing-owned volume, loop, display, trigger, and target values. The
published pre-1.0 `oxml-opc` crate adds audio, video, and Microsoft media
relationship constants. The dependency-free `oxml-media` crate adds safe MIME
and container-signature classification plus non-image media naming.

These are additive pre-1.0 native Rust APIs. The timing signature and common
media value exclude trim because the Office picture extension owns it. No
Python method, WASM method, CLI option, production dependency, feature flag, or
decoder surface exists.

## Native PowerPoint timing model

The published pre-1.0 `rpptx-oxml` crate exposes concrete timing and transition
values through its `timing` module. `CT_Slide`, `CT_SlideLayout`, and
`CT_SlideMaster` carry optional `CT_Timing` and `CT_SlideTransition` fields.
Callers can inspect supported containers, conditions, targets, builds,
behaviours, effect parameters, transition policy, and morph metadata. Bounded
mutation methods change one common-node duration, transition speed, or existing
morph option atomically while retained unsupported XML remains the
serialization source.

The low-level model also exposes exactly two additive queries used by the
timeline resolver:

```rust
pub fn ShapeTreeChild::non_visual_name(&self) -> Option<String>;
pub fn CT_Timing::condition_has_explicit_target(
    &self,
    node_id: u32,
    end_condition: bool,
    index: usize,
) -> Option<bool>;
```

The published `rpptx-layout` crate adds `TimelinePosition`,
`EvaluatedShapeState`, `EvaluatedTransition`, `EvaluatedFrameState`,
`ResolvedShapeIdentity`, `ResolvedTimelineSlide`, and `evaluate_timeline`.
`rpptx-render::timeline` lowers an evaluated slide and composes ordinary and
morph transitions. The native facade adds one deterministic entry point:

```rust
pub fn Presentation::render_timeline_deterministic(
    &self,
    slide_index: usize,
    position: TimelinePosition,
    outgoing_slide_index: Option<usize>,
) -> Result<DeterministicTimelineFrame>;
```

`DeterministicTimelineFrame` returns the composed `PageFrame`, the exact
`EvaluatedFrameState` used for that page, and ordered diagnostics. Invalid
slide indices and non-finite evaluated state fail closed. This remains an
additive pre-1.0 native Rust API. It adds no Python method, WASM method, CLI
option, production dependency, or feature flag. Existing static render methods
do not enter the timeline path. Unsupported timing behaviours remain explicit
raw nodes rather than acquiring a second authoring surface.

The published pre-1.0 `rpptx-layout` crate also exposes
`MediaPlaybackPhase` and `EvaluatedMediaState`. The published pre-1.0 `rpptx`
facade adds `MediaFallbackPolicy`, `DeterministicMediaTimelineFrame`, and one
media-aware deterministic entry point:

```rust
pub fn Presentation::render_media_timeline_deterministic(
    &self,
    slide_index: usize,
    position: TimelinePosition,
    outgoing_slide_index: Option<usize>,
    fallback_policy: MediaFallbackPolicy,
) -> Result<DeterministicMediaTimelineFrame>;
```

The nested result retains the existing `DeterministicTimelineFrame` and adds
ordered playback states with stable shape id, phase, source position,
normalized volume, and loop status. `PosterFrame`,
`DeterministicPlaceholder`, and `Fail` make every approved poster policy
callable. This is additive pre-1.0 native Rust surface. It adds no Python,
WASM, or CLI method, feature flag, production dependency, generic, trait, or
codec decoder. Existing static and timeline entry points retain their exact
diagnostic strings and results.

The published pre-1.0 `rpptx` facade also exposes the concrete native animation
values `AnimationTransition`, `GifLoopBehavior`, `AnimationFormat`,
`AnimationSegment`, `AnimationExportOptions`, and `DeterministicAnimation`.
The entry point is:

```rust
pub fn Presentation::export_animation_deterministic(
    &self,
    segments: &[AnimationSegment],
    options: AnimationExportOptions,
) -> Result<DeterministicAnimation>;
```

Segments declare slide index, positive duration, fixed click count, and either
no transition source or an explicit outgoing slide. Options declare bounded
frame rate and pixel dimensions, animated GIF loop behavior or Motion JPEG AVI
quality, and the existing `MediaFallbackPolicy`. The result carries the encoded
bytes, exact output timestamps, and ordered diagnostics. The facade uses one
prepared media-aware timeline assembly for the complete export and writes one
opaque frame at a time through capped pure-Rust encoders. This additive native
surface adds no Python, WASM, CLI, trait, generic, builder, wrapper, feature
flag, system codec, subprocess, or binary asset.

## Packaging

**maturin, mixed Rust and Python layout**, so type stubs and enum shims have a
home. `python-source = "python"`, `module-name = "rdocx._rdocx"`,
`features = ["pyo3/extension-module"]`. The rpptx package uses the parallel
`rpptx._rpptx` module name.

**abi3-py39.** One wheel per platform rather than one per interpreter version,
so roughly 6 wheels instead of 48. The cost is marginally slower attribute
access and no free-threaded build under abi3. Start abi3-only and revisit only
if profiling shows attribute overhead matters.

Matrix: `manylinux_2_28` x86_64 and aarch64, `musllinux_1_2` x86_64, macOS
x86_64 and arm64, Windows x86_64, plus an sdist.

Two traps specific to this workspace:

- **`fontdb`'s `fontconfig` feature is useless on musl and Windows.** Gate it
  per-target.
- **Bundled fonts are always compiled into wheels.** The optional
  `system-fonts` feature adds host discovery, but a bare manylinux container
  still has the bundled fallback inventory needed for `to_pdf()`. Roughly 4 MB
  per wheel is a fair trade for deterministic fallback text.

Each mixed package ships a hand-written native-extension stub beside its
extension module and a `py.typed` marker at package root. The stubs describe
concrete lazy handle and collection types, integer and slice overloads, typed
iteration, path-like inputs, byte outputs, optional values, bounded enum inputs,
and concrete Length returns. Native handles and collections are factory-only,
so their stubs reject direct construction just as the extension types do. The
pure-Python units, enums, and exception hierarchies remain inline typed rather
than duplicated in package-level stubs. Exact `mypy==2.3.0 --strict` smoke
checks and `stubtest` against freshly installed wheels keep the declarations
honest. Do not auto-generate them from PyO3.

**Distribution names `rdocx` and `rpptx`**, import names identical. The binding
crates are `publish = false`, because a cdylib has no business on crates.io.
Each Python project uses its crate-local `README.md` as a Markdown long
description. The metadata also carries a distribution-specific summary,
author, keywords, Python and topic classifiers, and project URLs. The README
provides installation, compatibility, capability, quick-start, typing, and
project-link guidance on PyPI. Wheel and source-distribution validation checks
the embedded metadata and the required README sections before publication.
Full-description comparison normalizes platform CRLF to LF and ignores terminal
newline count. All prose and other metadata remain exact.

The Rust package trains remain separate. The exact 15-package shared OOXML and
PowerPoint workspace family is published at 0.12.1 from immutable annotated tag
`rpptx-v0.12.1` at reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b`. The stable workspace is published
as the exact seven-package 0.14.0 family from immutable annotated tag `v0.14.0`
at the same reviewed SHA. Every selected registry entry and its sole owner are
verified, and the stable archives require shared 0.12.1.
The immutable v0.13.0 tag at
reviewed SHA `05332b17f481741e7d5ab4e39699c6d1536475af` published five
low-level stable packages, then stopped because packaged `rdocx` required the
newer `oxml-opc` Word main content-type constants. `rdocx`, `rdocx-cli`, and
the GitHub release remain absent. The immutable v0.11.0 attempt at
reviewed SHA `25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4` published only
`rdocx-opc` and `rdocx-oxml`. It created no GitHub release and posted no
contribution notifications. The complete seven-package recovery is published
at 0.11.1, and all six reviewed leave-open notifications are posted. The
immutable PyPI `rdocx 0.13.1` release remains available with its short
summary. PyPI `rdocx 0.13.2` carried the corrective Python metadata, and its
crates.io train is superseded rather than backfilled. S73 published four
version-aligned releases at one reviewed SHA. The stable source, `rdocx` Python
project, and `rdocx-wasm` track 0.14.0 for `v0.14.0` and `py-rdocx-v0.14.0`.
The shared and PowerPoint source, `rpptx` Python project, and unpublished
`rpptx-wasm` crate track 0.12.1 for `rpptx-v0.12.1` and `py-rpptx-v0.12.1`.
The failed immutable `rpptx-v0.12.0` workflow published no registry packages
and created no GitHub release. Because packaged stable crates require the
shared family, `rpptx-v0.12.1` was published before `v0.14.0`. Each tag used
its own immediate final approval.
Every binding and WASM crate remains unpublished on crates.io. Neither Rust
release gives binding, WASM, npm, or Python package publication authority. Every
later release still requires its selected-family gate and a separate final
approval at the reviewed SHA. Complete coherent stable releases remain live and
unyanked. After separate immediate approval, the incomplete `rdocx-opc@0.11.0`
and `rdocx-oxml@0.11.0` entries are yanked. Their package bytes, every other
version, the immutable v0.11.0 tag, and GitHub release state remain unchanged.

## CI

`wheels.yml` on **`py-rdocx-v*` and `py-rpptx-v*` tag namespaces**, separate
from `publish.yml` on `v*` and `rpptx-v*`, so a Rust release does not rebuild
Python wheels and a binding-only fix does not force a crates.io release.
Publishing uses PyPI trusted publishing via OIDC, with no long-lived token in
secrets. Manual dispatch builds `rdocx` and `rpptx` across the six declared
targets. A tag build runs only the selected distribution cells. Each package
produces one source distribution and uploads each matrix product independently.
Every native wheel is
installed into a fresh environment for its compatible pytest, exact
`mypy==2.3.0 --strict`, and `stubtest` gates. Each musllinux wheel is installed
in a fresh Python 3.9 Alpine environment and runs the same package parity suite
as the native cells.

The build jobs have only repository read permission. A separate publish job
depends on all wheel and source-distribution jobs, selects exactly six wheels
and one source distribution for the tag's project, and receives
`id-token: write` only for a `py-rdocx-v*` or `py-rpptx-v*` tag event in the
`pypi` environment. Manual dispatch builds and tests both distributions but
cannot publish them. Every external action and the maturin tool version are
pinned to reviewed immutable versions.

Each Python release family contains one distribution at its native crate
version. After the reviewed SHA is pushed but before either Python tag is
created, a manual build-only run at that SHA must produce twelve `cp39-abi3`
wheels and two source distributions across both projects. The selected six
wheels and one source distribution must name the selected project and version
in both filenames and embedded metadata. The selected distribution is
installed under clean Python 3.9 and 3.12 environments for the priority runtime
gate. Exact `mypy==2.3.0 --strict` and stubtest run under Python 3.12 because
that mypy version requires Python 3.10 or newer. The tag-only publish job uses
PyPI trusted publishing. Successful publication is not complete until the
selected project version, all seven files, authenticated owner or maintainer
roles, the exact reviewed GitHub release body, and every contribution
notification are verified.

The current reviewed Python releases are `rdocx 0.13.2` from
`py-rdocx-v0.13.2` and `rpptx 0.11.0` from `py-rpptx-v0.11.0`, both at SHA
`2b009243ed39ab66470d7484d490985368e865a8`. Each PyPI release contains the
exact six platform wheels and one source distribution, exposes its complete
crate-local README and project metadata, and passes clean Python 3.9 and 3.12
runtime checks plus Python 3.12 strict typing and stub checks. Their GitHub
release bodies match the reviewed changelog text byte for byte.

**A PR-time job that builds the wheel and runs pytest is mandatory.** The
absence of exactly this job for wasm is why `rdocx-wasm` rotted.

The rdocx parity suite pins and asserts `python-docx==1.2.0` before comparison.
It writes the approved S33 content and direct formatting with each producer,
opens both outputs with both readers, and directly compares normalized public
paragraph, run, table, cell, unit and enum records. It compares no ZIP or XML
bytes. Relative float line spacing remains distinct from absolute `Length`
spacing in those records. An explicit table style is checked after each saved
output is reopened by both readers. The suite commits no binary fixture and
keeps python-docx out of runtime package dependencies.

## WASM

### The rdocx wrapper

```rust
#[wasm_bindgen]
pub struct WasmDocument { inner: rdocx::Document }
```

`fromBytes` delegates to `Document::from_bytes`, and `toDocxBytes` delegates to
`Document::to_bytes`. The facade therefore flushes modeled changes into the
original package. Images, headers, footers, numbering, settings, themes, font
tables, notes, properties, content types, relationships, and opaque parts stay
in the package rather than being reconstructed by the binding.

The constructor, `fromBytes`, `addParagraph`, `addHeading`,
`addBoldParagraph`, `addTable`, `getText`, `paragraphCount`, `toDocxBytes`,
`toPdf`, `toHtml`, `toHtmlFragment`, `toMarkdown`, and `replacePlaceholder`
names remain stable. `toPdf` delegates to the normal `Document::to_pdf` facade
and returns its bytes directly. `Document::open`, `save`, and a second
deterministic PDF alias stay absent because browser callers supply and receive
bytes and the WASM profile already excludes host font discovery.

The `system-fonts` feature is default-on in `rdocx-layout` and `rdocx`, which
preserves native behavior. `rdocx-wasm` disables `rdocx` defaults, while the
bundled font data remains unconditional. The wasm32 graph therefore excludes
host font discovery without inventing a second bundled-font feature.
The crate-local sRGB2014 profile is compiled into `oxml-pdf` and introduces no
host API or runtime dependency. The native PDF/A methods are not exported by
either WASM wrapper.

Caller-font alias setters and alias-aware layout or transfer methods remain
native Rust APIs. Neither WASM wrapper exports a new method, and its host-font
free dependency graph is unchanged.

The R-class regression constructs a document with an image, header, and
numbering, then checks the complete part, relationship, and content-type graph
through `fromBytes` and `toDocxBytes`. The same contract is an inline
`wasm-bindgen-test` for Node. The Node test reflectively calls those generated
JavaScript members and crosses the `Uint8Array` boundary in both directions.
A second inline Node test calls generated `addParagraph` and `toPdf` members,
then requires a complete PDF with a Type 0 font, an embedded TrueType stream,
and the bundled Carlito base font. Pull-request CI target-checks the wrapper
with the locked workspace graph and runs both tests in Node.

`rpptx-wasm` owns one `rpptx::Presentation`, never a mini-model. Its default
profile exposes the constructor, `fromBytes`, `toBytes`, `slideCount`, and
`addSlide`. It includes the bundled default template but no renderer, PDF
backend, rasteriser, or host font discovery. The `render` feature adds only
`toPdf` and selects the facade's deterministic renderer. The optimized default
artifact must remain below 1,000,000 bytes after deterministic gzip.

Modern presentation package-class inspection and output selection remain
native Rust APIs. Python, WASM, and CLI callers continue to preserve the
source main content type through their existing byte or path save operations,
but they gain no package-class selector in this milestone.
Pull-request CI target-checks the default wrapper with the locked workspace
graph and runs its package-preserving inline test in Node.

The npm package names are `@tensorbee/rdocx-wasm` and
`@tensorbee/rpptx-wasm`. Both use the bundler target, their Rust package
versions, and release output optimized by exact wasm-opt 125 with `-Oz`,
`--enable-bulk-memory`, and `--enable-nontrapping-float-to-int`. Pull-request
CI creates local tarballs with `npm pack`, installs each tarball into a separate
fresh consumer, and checks the installed WASM, JavaScript glue, public
TypeScript declaration, and module import. This is an installation gate only.
The job has no npm publication, registry authentication, token, OIDC, release,
or tag authority.

## CLIs

`rpptx-cli` extends the seven-command `rdocx-cli` surface with `inspect`,
`text`, `convert`, `diff`, `replace`, `validate`, `render`, `thumbnail`, and
`outline`. It uses clap derive and `serde_json` for `--json`.

`inspect` reports the file, slide and layout counts, slide size, core metadata,
and each slide's identity, hidden state, and shape count. Its JSON form uses the
shared schema-1 envelope. `text` emits slide text in presentation order.
`convert` produces deterministic PDF, PNG, JPEG or TIFF output. Multi-slide PNG
and JPEG output uses one-based filename suffixes and renders one slide at a
time, while TIFF writes one multi-page stream. `diff` compares slide text with
longest-common-subsequence semantics and rejects matrices above one million
cells. `replace` delegates to the facade's literal, formatting-preserving text
replacement. `validate` is dispatched separately so its exit status carries the
verdict. `render` uses deterministic fonts and the shared one-based range
grammar for image output.

PNG rendering is limited to eight million pixels per slide for both `convert`
and `render`. A zero-slide PNG conversion fails without creating output.
The exact validation gate corrupts one relationship and requires a nonzero exit,
then requires every verified pinned corpus deck to exit zero without skips.

`thumbnail` renders slide one with deterministic fonts at exactly 320 pixels
wide and preserves the rendered page aspect ratio. Its output defaults through
the shared extension helper. `outline` prints each slide title once, followed
by non-title text paragraphs in recursive shape z-order. Tables use row-major
cell order, paragraph levels add two spaces of indentation, empty text is
omitted, and embedded paragraph breaks become spaces.

Shared range parsing, output-path defaulting, and JSON envelope rules live in
`oxml-cli-support`. Ranges are positive, one-based, comma-separated values and
inclusive ranges. Parsing sorts and deduplicates the result, and rejects more
than 100,000 requested values before expansion. The output helper replaces or
adds only the requested extension. The envelope accepts an object without a
caller-supplied `schema` field and adds the reserved top-level
`{"schema": 1, ...}` contract.

`rdocx-cli` uses the shared envelope for inspect JSON and the shared path helper
for convert defaults. General image conversion uses one-based `--pages` ranges.
`render --page` remains the zero-based legacy single-page selector and is
mutually exclusive with the one-based `render --pages` range. Both flags select
against the same deterministic layout snapshot that is passed to the shared
raster backend. The legacy `--page 0` default PNG path and single-line stdout
remain unchanged. The `text` command emits paragraphs and table cells in
document order through the facade plain-text representation. `text --json`
emits schema-1 accepted-view paragraphs with a zero-based direct body index,
typed zero-based nested path, direct style and numbering, text, and ordered
runs. Run formatting is null when no direct run properties exist. Otherwise it
contains nullable direct bold, italic, strike, underline, font, point size,
colour, highlight, language, and character style fields. `layout --json` uses
bundled deterministic fonts and reports every direct body item. Its point-space
fragments carry one-based physical and displayed page numbers, and preserved
unlaid items retain an empty fragment list. `replace --expect N` checks the
run-aware replacement count before staged publication. A mismatch creates no
output and leaves an existing destination untouched. Both the selected page
and all-page `render` paths use bundled deterministic fonts. The compiled
surface also includes nested comment thread commands, main-story revision
inspection, all-story filtered revision resolution, whole-run comparison, and
TOC rebuild. Every new mutation requires an explicit output and publishes
through the shared staged output set. Their schema-1 records state `main` or
`all-supported-stories` scope. Revision selectors are mutually exclusive, and
RFC 3339 start and end bounds must be paired. The complete compiled surface is
covered by one integration binary, with fixtures constructed in code and no
command-only test dependency.

Rust release tags also distribute the selected CLI as prebuilt archives.
Stable `v*` tags carry only `rdocx`, and incubating `rpptx-v*` tags carry only
`rpptx`. Each family has native archives for GNU Linux x86-64 and arm64,
static musl Linux x86-64, macOS Intel and arm64, and Windows x86-64. Every
archive contains exactly the executable, its crate README, and the workspace
licence. The two CLI manifests expose cargo-binstall metadata that resolves
these target-specific archives without selecting source compilation or an
unreviewed quick-install source. Native builds retain the CLI default system
font feature. The static musl build disables host discovery and retains bundled
fonts.
