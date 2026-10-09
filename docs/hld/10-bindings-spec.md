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
source paths. Deleted and move-source runs are absent. Paragraph text also
reads the runs inside smart tags, inline custom XML elements, and simple fields,
which the run handles do not address. A successful structural
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
reader methods. The native numbering level reader reports actual extra
attributes and children on its instance, definition and level. Retained XML
namespace declarations alone do not signal unmodeled numbering semantics.

At the public low-level Rust boundary, `CT_RPr` includes the complete language
attribute set and its retained foreign attributes, while `LayoutInput` includes
the document automatic-hyphenation boolean and the
`w:doNotUseHTMLParagraphAutoSpacing` compatibility boolean, and the laid-out
`TableCell` includes the horizontal border bands above and below its content.
Full struct literals must provide these fields. These are intentional pre-1.0 source breaks for the next stable
family. Shared `TextSegment`, `GlyphRun` and `MultilingualGlyphRun` struct
literals also provide `note_reference_source`. Exhaustive shared `FieldKind`
matches cover `Section`, `SectionPages`, `SequenceContext` and `SequenceRepeat`.
The established layout entrypoints retain their call shapes.

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
  `WD_TABLE_ALIGNMENT`, `WD_CELL_VERTICAL_ALIGNMENT` and `WD_ROW_HEIGHT_RULE`
  in `rdocx.enum.table`. All five are also top-level exports. Their checked
  integer literals cover the paragraph, run and table variants exposed by the
  facade, including `WD_ALIGN_PARAGRAPH.CENTER == 1`. Underline codes use a
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

`Table`, `Row` and `Cell` also bind the checked native formatting setters.
`Table.set_borders(style, *, size, color)` sets every table edge and
`set_border(edge, style, *, size, color)` sets one. The style is an
`ST_Border` name from `none`, `single`, `thick`, `double`, `dotted`, `dashed`,
`dotDash` and `wave`, the edge is `top`, `bottom`, `left`, `right`, `insideH`
or `insideV`, the size is in eighths of a point and the color is six
hexadecimal digits or `auto`. `border(edge)` reads a `(style, size, color)`
tuple or `None`. `set_cell_margins` takes four keyword EMU lengths and
`cell_margins` reads a `(top, right, bottom, left)` tuple of optional
`Length` values. `grid_widths` reads and replaces every grid column, and
`set_column_width(column, width)` changes one. Both keep the table width and
every covering cell width in step. `Row.height` and `Row.height_rule` follow
python-docx with `WD_ROW_HEIGHT_RULE.AT_LEAST` and `EXACTLY`. Assigning a
height keeps an exact rule, and a rule needs a height to apply to. Unlike
python-docx, a row whose `w:trHeight` has an `auto` rule or no value reads no
height, and assigning one writes a minimum.
`Row.cant_split` and `Row.is_header` are tri-state. `Cell.shading`,
`Cell.border(edge)`, `Cell.set_border`, `Cell.margins` and `Cell.set_margins`
are the same forms for one cell. These edits move no content, so they keep
live handles valid and do not advance the revision. An unknown style, edge or
rule raises `ValueError`, a value the native setter rejects raises
`RdocxError`, and either way the document is unchanged.

`Document.insert_table(index, rows, cols)` inserts a table at a direct body
index, rejects an index past the end with `IndexError` before mutation, and
returns a handle to the new table. Table handles count every table in
document order, including those inside block content controls, so the handle
is resolved from the inserted body position rather than assumed. Cells merge
through the checked table operations, never through the unchecked cell span
setter. `Table.set_cell_grid_span(row, col, span)` spans columns and absorbs
or restores untouched empty cells, and `None` or `1` removes the span.
`Table.set_cell_vertical_merge(row, col, merge)` writes `restart`, `continue`
or `None` after validating the whole merge topology. Both take the possibly
negative indexes `Table.cell` takes. A grid span that absorbs or restores
cells advances the revision once, because later cell indexes move, while a
vertical merge keeps live handles valid. `Cell.grid_span` reads the span with
the python-docx default of 1 and `Cell.vertical_merge` reads the merge state.

`Paragraph.text` is writable, in body and table-cell paragraphs, through the
native `Paragraph::set_text`, which follows the python-docx setter. The
paragraph keeps its properties and the attributes of `w:p`, while its runs,
hyperlinks, fields, pictures, content controls, tracked changes and other
inline children give way to one run without direct formatting. A tab becomes
`w:tab`, and a line feed or a carriage return becomes `w:br`. Empty text or
`None` leaves no run, as `add_paragraph("")` does. Unlike python-docx, comment
ranges, bookmarks and permission ranges found anywhere in the paragraph are
kept. Their starts move before the new run, their ends after it, and each
comment reference follows in its own run, so a range over part of the old text
covers the new text and no anchor loses its partner. Bookmarks are rebuilt from
their ID and name, and permission markers keep their source XML. Two kinds of
paragraph are rejected with `RdocxError`, since dropping part of them would
unbalance the rest of the document. The first is a paragraph whose field
characters and field code do not balance within it, such as any paragraph of a
table of contents that spans several paragraphs. The second is a paragraph
holding only one end of a tracked move range or a custom XML revision range.
When both ends of such a range are in the paragraph, both are dropped with the
tracked change they mark. A successful assignment advances the revision once
and stales every earlier handle, the assigned paragraph included, as
`Cell.text` does. A rejected one changes nothing.
`Document::set_story_text` keeps its own behavior.

`Paragraph.style` accepts a style ID the package defines or, as python-docx
does, a style name, and writes the resolved ID. Among paragraph styles only,
the ID is tried first, so a value read back always assigns the same style,
then the exact name, then the name regardless of case, so `Heading 1` finds
Word's `heading 1`. A character style therefore cannot hide a paragraph style
of the same name. As in python-docx, a value that names no style raises
`KeyError`, and one that names only a character, table or numbering style
raises `ValueError`, both before any change. Assigning a style is a value-only
mutation and keeps handles valid. The run style, table style and numbering
setters still write their value unchecked. This changes behavior on documents
that lack the style. On such a document `p.style = "Heading2"` used to store an
undefined style, which `rebuild_toc` even read as a level-two heading, and now
raises `KeyError`. A new `rdocx.Document()` defines `Heading2` and the other
common Word styles that `Document::add_common_styles` adds.

`Document.add_style(name, style_type="paragraph")` creates a paragraph,
character or table style through the native `add_style`, as python-docx's
`styles.add_style` does, and returns its `Style` snapshot. The ID keeps the
ASCII letters, digits and hyphens of the name, as Word derives one, so `Q&A`
gives `QA` and `Note (draft)` gives `Notedraft`. A name with none of them, such
as a Japanese one, gets the first unused of `a`, `a0`, `a1` and so on. The
exceptions python-docx makes for Word's lowercase built-in names `caption` and
`heading 1` to `heading 9` still give `Caption` and `Heading1` to `Heading9`.
The `style_id` keyword overrides the derived ID. `based_on` and `next_style`
accept an ID or a name, resolved as `Paragraph.style` resolves one among the
styles of the new style's type, or among paragraph styles for `next_style`. The
formatting keywords are the font name, size, bold, italic and colour, and the
spacing before and after, left, right and first-line indentation in EMU. They
are written as the `Font` and `ParagraphFormat` setters write them on content,
and a negative first-line indentation becomes a hanging one. A base or next
style that no style names raises `KeyError`. A duplicate ID, a name another
style has regardless of case, an unknown type, a base or next style of another
type, and paragraph formatting on a character style raise `ValueError`. Both
are raised before any change. `remove_style` and `set_default_style` resolve a
style of any type the same way. `remove_style` returns `False` when no style
has the ID or name, and `set_default_style` raises `KeyError` then and
otherwise makes the style the default of its type. The native refusals, such as
removing a style that content or another style names, raise `RdocxError`.
`set_style(style, **formatting)` resolves an existing style by ID or name and
passes the supplied formatting, base style and next style through native
`set_style`. Omitted properties keep their existing values. The staged native
setter compares counted style-graph defects before publishing. Existing producer
defects remain, while new defects reject the edit. A repeated style ID resolves
to its first definition, and Python handles stay valid. Unknown or wrong-type style references fail before mutation.
The update surface does not clear a theme font or theme colour inherited from
an existing Word style. Its formatting changes are additive overrides.

The Issue 168 package-write escape hatch is an external ZIP and lxml edit of
an unmodelled package part, followed by reopen and save through `rdocx`.
The typed binding does not offer arbitrary part replacement. The caller must
keep package relationships and content types coherent when editing outside
the binding. The production-chain gate changes an existing application
property through that escape hatch and proves it survives the typed round trip.

`ListLevel` is a constructible frozen value with a `format` checked against
the standard `w:numFmt` names through `ListNumberFormat::from_name`, the level
`text`, `start`, and `left_indent` and `hanging_indent` in EMU. Omitted values
take the native defaults, `%1.` or a bullet glyph and half an inch of
indentation per level with a quarter-inch hanging indent.
`Document.add_numbering_definition(levels)` creates an abstract definition from
one to nine levels and returns its ID. `add_numbering_instance(definition_id)`
creates a `w:num` for it and returns the `numId` that `Paragraph.numbering`
takes. `link_style_to_numbering(style, num_id, level)` links a paragraph style,
given by ID or name, to one level through the native method of that name, so
every paragraph of that style is numbered. Native checks raise `RdocxError`
and publish nothing. Style and numbering authoring changes no content, so the
revision and every handle stay valid.

The Python `Document` also exposes the current native comparison, main-body
comment, deterministic layout, TOC rebuild, revision, counted replacement, and
field cache update operations. `RunPosition` and
`RunRange` are constructible frozen values for zero-based half-open run ranges.
`StoryRunPosition` and `StoryRunRange` are parallel frozen values whose
`StoryItem` snapshots can identify direct body or table-cell paragraphs.
`StoryRunPosition` also accepts a body `Paragraph` handle. A handle to a
paragraph inside a block content control yields an item with the two-segment
path of `Document::paragraph_story_location`, the control's story item index
then the paragraph's position among the control's paragraphs. Comment positions
and scoped paragraph replacement accept this path. Both paragraph-handle positions and returned comment
endpoints carry checked accepted paragraph text and nonempty paragraph XML.
`Document.add_comment` accepts either range form.
`Document.add_comment_on_text` comments on the zero-based occurrence of an
exact text in the main story without run index bookkeeping. The original
direct-body constructors and call shape remain unchanged.
`Comment`, `ComparisonDiagnostic`, `BoundingBox`, `LayoutFragment`,
`LayoutPage`, `TocRebuildReport`, and `Revision` are frozen typed snapshots.
`Document.revisions` lists the revisions of every story that the accept and
reject methods resolve, each with a snake_case `kind` and the `Story` that
holds it, so its length equals the count they return. A `Revision` built
directly keeps its four-field constructor and has no story unless one is
passed. `try_replace_text` and `replace_all_regex` return their
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

`Paragraph.replace_text(old, new, *, expect=None)`,
`Cell.replace_text(old, new, *, expect=None)` and
`Document.replace_text_at(item, old, new, *, expect=None)` return a local
literal replacement count. Paragraphs keep run formatting and never join
separate paragraphs. Cells search their supported nested tables and controls,
without visiting neighboring cells. The explicit Document operation accepts
checked paragraph, table and block-control StoryItems, including related header,
footer and normal note paragraphs. StoryItem remains frozen and detached.
Paragraph handles inside block controls retain their checked two-segment paths.
Cell and cell-paragraph handles use physical table, row and cell coordinates.

Each route stages the owning native Document and reopens the complete package
before publication. `expect` mismatch raises `ReplacementCountError` with
`index=None`, `expected` and `found`, preserving the document and all handles.
Invalid XML text, stale selection, unsupported targets and preparation failures
also preserve bytes and revisions. Zero replacements, including an empty old
literal, do not advance the revision. A positive count publishes once, releases
the GIL while working and advances the revision once, staling prior handles.
The shared binding generic has checked StoryItem and physical cell closures as
its existing consumers. Global replacement and CLI signatures remain unchanged.

`Document.remove_content` returns false for an absent index and raises
`RdocxError` for an unsafe comment cut. Complete thread deletion and cell text
replacement publish one prepared and reopened candidate, then invalidate
borrowed handles once. Refused operations leave handles and package bytes
unchanged. `pop_content` refuses a commented fragment rather than detaching
anchors from their definitions. Native `Cell::try_set_text` exposes checked
errors, while its legacy setter leaves the cell unchanged on refusal.


`Document.add_picture` accepts in-memory bytes, a safe filename, optional
paired EMU dimensions, and an optional checked `StoryItem` placement. Omitted
dimensions use native 72 DPI sizing. Omitted placement appends to the body.
The returned immutable `StoryItem` names the inserted paragraph at the new
document revision. A rejected image, filename, dimension pair, or stale path
leaves package bytes and binding revisions unchanged.

`rpptx` mirrors python-pptx through an unpublished mixed-layout `rpptx-py`
crate. `Presentation` owns the Rust facade and one revision counter per
handle scope.
`Presentation(path)` opens a file and the static `Presentation.from_bytes`
opens in-memory package bytes, as the rdocx `Document` does. Lazy layouts,
slides, shapes, placeholders, text frames, paragraphs, runs, columns and cells
store only a presentation reference and `ContentPath`. The bounded
source-compatibility surface is the seven python-pptx 1.0.2 Getting Started
workflows. They change only the import namespace. Pure-Python
`Length`, `Inches`, `Pt`, `RGBColor`, and the `MSO_SHAPE`, `MSO_SHAPE_TYPE`,
`MSO_CONNECTOR_TYPE`, and `MSO_FILL_TYPE` enumerations keep native inheritance
outside the limited ABI. `MSO_SHAPE` carries the 181 python-pptx
`MSO_AUTO_SHAPE_TYPE` members whose preset rpptx can author, each with its
preset name as `xml_value`. `UP_ARROW` is absent because the generated preset
table has no `upArrow`.

Handles go stale by scope, so the handles python-pptx scripts hold across an
edit keep working. A handle belongs to the scope of the last step of its
path: slide, shape, table row or cell, paragraph, or run. An edit that
renumbers the members of one scope advances the revision of that scope and of
every narrower one. Removing, moving, or duplicating a slide, and importing
one before the end, invalidates every handle. Removing, moving, grouping, or
ungrouping shapes and `insert_picture` invalidate shape, table, paragraph, and
run handles. A row or column edit invalidates table, paragraph, and run
handles. Replacing the whole text of a shape or text frame invalidates
paragraph and run handles, and replacing a paragraph's text or a counted
replacement invalidates run handles. Appends renumber nothing, so
`add_slide`, every shape addition, `add_paragraph`, and `add_run` keep every
handle, and so do notes, comment, layout, geometry, and formatting writes.
`prs.slides`, `prs.slide_layouts`, and layout handles read the presentation
live and never go stale. A `StaleElementError` names the handle, its scope
and that scope's revisions, the call that last advanced them, such as
`SlideCollection.remove()`, and the public path that re-fetches it.

Presentation `Shape` handles expose optional `Length` values for left, top,
width, and height plus optional non-visual id and name. Those values are the
shape's own. `Shape.effective_geometry()` returns the four as rendering places
the shape: its own transform, or for a placeholder without one, the transform
it inherits from its layout placeholder, then from that placeholder's master
counterpart. It returns `None` when the resolved transform has no extent.

`Presentation.slide_width` and `slide_height` read the optional `p:sldSz` as
`Length` values. Assigning one keeps the other, and a deck without `p:sldSz`
pairs the assigned value with the bundled 16:9 size. `Slide.slide_layout`
returns the layout the slide relates to, equal to the same entry of
`slide_layouts`, and `SlideLayoutCollection.index` returns its position.
Assigning a layout of the same presentation to `slide_layout` uses the native
staged layout change and keeps every handle valid. A placeholder the new
layout does not place keeps the transform it inherited.
`Slide.hidden` reads and writes `p:sld/@show`. `Slide.background.fill` is a
live `FillFormat` over the direct background fill that never changes the slide
when read, and `follow_master_background` reports and sets whether the slide
has no `p:bg`. `SlideCollection.remove` and `SlideCollection.move(from_, to)`
use the native staged slide operations and invalidate every handle.
`SlideCollection.import_slide(slide, layout=None, index=None)` wraps the native
import. `layout` must be a layout of the destination. `index` is an insertion
position counted from the end when negative, and one outside the collection is
an `IndexError`. A slide of the same presentation is imported from a snapshot
and keeps its own layout unless `layout` is given.
`SlideCollection.duplicate(slide)` copies a slide of the same presentation,
with its speaker notes, to the position right after it through the native
staged `duplicate_slide`, invalidates every handle, and returns the new slide.
As in the facade, a slide that owns a modern
comments part is refused, even when removing its last comment left that part
empty, and the refusal leaves the package and the revision unchanged.

`Shape` geometry, `name`, and `rotation` are writable without a revision bump.
Assigning one coordinate or the rotation of a placeholder first copies the
missing parts of its inherited transform onto the shape, so the values not
assigned keep their effective values and rendering keeps drawing it. On any
other shape a missing partner coordinate becomes zero, as in python-pptx. A
negative extent is a `ValueError`. `rotation` reads clockwise degrees
normalized below 360 and writes them with round-half-even into the
60000-per-degree angle.
`shape_type` reports an `MSO_SHAPE_TYPE` member or `None`. `fill` and `line`
return live `FillFormat` and `LineFormat` views for ordinary shapes, pictures,
and connectors, and raise `ValueError` for other kinds. `FillFormat` offers
`type`, `solid()`, `background()`, and `fore_color`. `ColorFormat.rgb` reads an
sRGB colour as `RGBColor`, or `None` for any other colour, and writing it keeps
the transforms of an existing sRGB colour. Reading `LineFormat.color` changes
nothing, and assigning its `rgb` makes the line fill solid. `LineFormat.width`
reads zero without a width, writes `None` as zero, and rejects values above
the `ST_LineWidth` maximum. `LineFormat.dash_style` reads the `a:prstDash`
value as an `MSO_LINE_DASH_STYLE` member, or `None` without one or with a
custom dash, and assigning `None` removes either dash. Its members keep the
python-pptx values and XML values, and `DOT`, `SYSTEM_DASH_DOT` and
`SYSTEM_DASH_DOT_DOT` read the three presets python-pptx cannot, all three of
which PowerPoint for Mac's scripting interface writes. `head_end` and
`tail_end` are live `LineEndFormat` views whose `type`, `width` and `length`
read and write `a:headEnd` or `a:tailEnd` as `MSO_ARROWHEAD_STYLE`,
`MSO_ARROWHEAD_WIDTH` and `MSO_ARROWHEAD_LENGTH` members, numbered as the
Office enumerations of the same names. `NONE` writes `type="none"`, as
PowerPoint for Mac's scripting interface does when an arrow is removed. `None`
removes one attribute, and an end left without attributes is removed. An
`a:ln` left empty is kept, as python-pptx keeps it. A removal never creates an
`a:ln`. `adjustments` is a live `AdjustmentCollection` of
the effective preset adjustments, normalized so that 1.0 is 100000, and
assignment truncates as python-pptx does. `auto_shape_type` reads the
`MSO_SHAPE` member of the preset through `MSO_SHAPE.from_xml`, which picks the
first member in definition order when two share a preset. A picture reports
its mask or `None`, and a shape without an auto-shape preset raises
`ValueError`. Custom geometry raises too, as python-pptx does, although issue
#217 asked for `None`. A shape with `prst="upArrow"` also raises `ValueError`,
where python-pptx returns `UP_ARROW`. The generated preset table has no
`upArrow`. rpptx accepts an `MSO_SHAPE` member or preset name to replace the
geometry and reset adjustments. A text box or unsupported shape kind raises
`ValueError`. An unknown preset raises `RpptxError`.

`theme_effect_index` reads the `a:effectRef` index of an ordinary shape or
connector style, or `None` without one. Writing 0 removes the theme effect,
including a connector's shadow, while retaining its theme line style. The
setter accepts an `int` only. A shape without a typed style refuses the write.
`xml` returns the shape element as bytes. A picture's `image` is a frozen
`Image` snapshot, and `replace_image` changes only that picture through the
native staged replacement. `crop_left`, `crop_top`, `crop_right`, and
`crop_bottom` read and write the picture's `a:srcRect` insets as python-pptx
floats, where 0.25 is a quarter of the image.
A missing edge reads 0.0, a write rounds half to even as python-pptx does and
changes nothing when the value is unchanged, and a value that is not finite or
outside the `ST_Percentage` range raises `ValueError`. Other shape kinds raise
`ValueError`, and crop writes do not advance the revision.
`Shape.click_action.hyperlink.address` reads and writes the click hyperlink
of the shape's non-visual properties. It shares relationship reuse and pruning
with run hyperlinks, and a write does not advance the revision. `None` or an
empty string clears it.
`Shape.click_action.target_slide` reads the slide a named, first, last, next,
or previous slide jump opens, or `None`, and assigning a `Slide` of the same
presentation goes through the native `set_shape_target_slide`, while `None`
removes the click action. A slide of another presentation raises `ValueError`.
A group accepts a click action, where python-pptx raises `TypeError`, because
PowerPoint honours it. Two current `Slide` handles compare equal when they name
the same slide, so `target_slide == prs.slides[2]` holds, and like python-pptx a
`Slide` is not hashable. No click action write advances the revision.

`Shape.alt_text` and `alt_title` read and write `p:cNvPr/@descr` and `@title`,
where `None` or an empty string removes the attribute, and `decorative` writes
the `adec:decorative` extension PowerPoint writes in `p:cNvPr/a:extLst`, so
screen readers skip the shape. Google Slides imports only `descr`. `flip_h`
and `flip_v` read and write `a:xfrm/@flipH` and `@flipV`, copying a
placeholder's inherited transform first as a coordinate write does, and a
graphic frame, which PowerPoint does not flip, raises `RpptxError`.
`Shape.insert_picture(image_file)` fills a picture or content placeholder as
python-pptx does: the `p:sp` becomes a `p:pic` with the same id, name, and
`p:ph`, so it keeps its frame, and the image is cropped evenly on its longer
sides to the frame's aspect ratio. Another kind of placeholder or shape
raises. A connector's `begin_connect(shape, cxn_pt_idx)` and `end_connect`
record `a:stCxn` or `a:endCxn` and move that end to the connection site, which
is the preset geometry's own `a:cxnLst` site with its adjustments, flips, and
rotation, or one of python-pptx's four edge midpoints for a shape without a
preset or whose preset defines none. `begin_x`, `begin_y`, `end_x`, and
`end_y` read and write the endpoints in the parent's coordinates, as in
python-pptx, and an assignment releases that end's glue. A new slide names its placeholders as PowerPoint and
python-pptx do, such as `Title 1` and `Content Placeholder 2`.

`shadow` returns a live `ShadowFormat` for ordinary shapes, pictures,
connectors, and groups, and raises `NotImplementedError` for a graphic frame
as python-pptx does. `inherit` is python-pptx's: it reads whether the shape
has no `a:effectLst`, assigning `True` removes the list and every effect in
it, and assigning `False` adds an empty one. The outer shadow properties are
`visible`, `color`, `alpha`, `blur_radius`, `distance`, `direction`, `align`,
and `rotate_with_shape`. They read `None` without an `a:outerShdw` of the
shape's own, a theme shadow included, and read the schema default for an
omitted attribute. Assigning one to a shape without an outer shadow first
adds the shadow PowerPoint for Mac writes for its Offset Diagonal Bottom
Right preset, measured through `msoShadow21`: preset black at 40% opacity,
a 4 pt blur, 3 pt away at 45 degrees, aligned top left and not rotating
with the shape. `visible = False` removes the outer shadow and keeps the
list, so the theme shadow stays off, and changes nothing on a shape without
an outer shadow of its own. `color` is a `ColorFormat` whose `rgb` keeps the
opacity across a change of colour kind and replaces an `a:scrgbClr` or
`a:hslClr` whole. `alpha` runs from 0.0 transparent to 1.0 opaque, and 1.0
removes `a:alpha`. Beside those two unmodelled colours `alpha` reads `None`
and assigning it is a `ValueError`. `direction` uses the `rotation` degrees
and rounding. `align` is an `ST_RectAlignment` token string, and lengths are
EMU.

`ShapeCollection.add_shape` accepts a DrawingML preset name or an `MSO_SHAPE`
member. `add_connector` follows the python-pptx signature, `add_group_shape`
appends an empty group, and `add_picture` accepts a path, bytes, or a binary
file-like object, which is rewound first when it can seek. `remove` deletes one
shape of a slide with the relationships and parts only it used. `move(from_,
to)` changes the z-order like `SlideCollection.move`, so the shape ends up at
index `to` and draws above the shapes before it. python-pptx has no z-order
API. Both invalidate shape handles.

A group's `shapes` collection has the same `add_textbox`, `add_shape`,
`add_connector`, `add_group_shape`, `add_table`, and `add_picture` through the
native `ShapesMut`, so groups nest to any depth. A member takes a `p:cNvPr` id
unused across the slide, and the group, then every group enclosing it, is refit
to the union of its members as python-pptx does, so members of a new group keep
their slide coordinates. An addition to a group renumbers nothing, so the group
handle, its collection, and its earlier members stay valid.
A group added inside a group has the zero `a:xfrm` python-pptx writes, so its
`left`, `top`, `width`, and `height` read zero until its first member arrives,
while a group added to a slide's own shapes reads `None`.
`add_group_shape(shapes)` moves existing members of the collection into the new
group, as python-pptx does, and `group(shapes)` is the same call without the
empty form. The native `ShapesMut::group` keeps each member's slide position,
its order, and its id: the group's `a:chOff` and `a:chExt` equal its `a:off`
and `a:ext`, the union of the members' boxes, and the group takes the z-order
place of the topmost member. A placeholder, which PowerPoint does not group,
and a shape of another collection raise. `Shape.ungroup()` replaces a group by
its members and returns them, with the group's scale, flips, and rotation
moved into each member's own transform in PowerPoint's flip-then-rotate order.
A table or chart cannot rotate or flip, so ungrouping a rotated, flipped, or
scaled group that holds one raises, as does a group that slide animations
target. `ShapeCollection.align(shapes, alignment, *, relative_to="selection")`
and `distribute(shapes, direction, *, relative_to="selection")` run the
native `ShapesMut::align` and `distribute` on the drawn boxes, flips and
rotation included, against the shapes' bounding box or the slide. python-pptx
has neither.
Adding to the collection of a shape that is not a group raises `ValueError`
before any image is read, and so does adding to a group inside an
`mc:AlternateContent` fallback, which stays read-only. `remove` and `move` on a
nested collection raise `ValueError` too.

A table `Cell` follows python-pptx for `merge(other_cell)`, `split()`,
`is_merge_origin`, `is_spanned`, `span_height`, and `span_width`. Merge and
split run the native staged table operations, which keep the rectangular grid,
so they do not advance the revision. A cell of another table raises
`ValueError`, and a merge range that overlaps a merge or a split of a cell that
is not a merge origin raises `RpptxError` and leaves the table unchanged.
`Cell.fill` is a live `FillFormat` over the direct cell fill. `margin_left`,
`margin_right`, `margin_top`, and `margin_bottom` read the `a:tcPr` margins as
`Length` and follow the text formatting rules below. An absent margin reads
`None` where python-pptx reports its 91440 and 45720 EMU defaults, and a value
outside the 32-bit coordinate range raises `ValueError`. python-pptx has no
cell border API, so `border_left`, `border_right`, `border_top`, and
`border_bottom` are live `LineFormat` views of `a:lnL`, `a:lnR`, `a:lnT`, and
`a:lnB`, which reading never creates. `Table.rows` is a lazy `RowCollection`
of `Row` handles, like `columns`. `Row.height` reads the stored height as
`Length`, and assigning it keeps the frame height equal to the sum of the rows,
as a column width keeps the frame width. A height that is not positive raises
`RpptxError`. None of these writes advances the revision.

python-pptx has no public API to add or remove table rows and columns, and its
users call `table._tbl.add_tr(height)`, which appends a row without cell
formatting and leaves the frame height unchanged.
`RowCollection.add_row(index=None)` inserts a row before `index`, or appends
one, and returns its `Row`. `ColumnCollection.add_column(index=None)` does the
same for a grid column.
Both run the native `insert_row` and `insert_column`, so the new row or column
copies the size and cell formatting, without the text, of the row above or the
column to the left, or of the first one at index 0. The frame grows by the new
row's height or column's width, and text later written into a new cell takes
the copied formatting. `RowCollection.remove(row)` and
`ColumnCollection.remove(column)` remove one and shrink the frame by its size,
so a frame PowerPoint measured taller than its stored rows keeps that excess.
A negative index counts from the end, as in `list.insert`, and an index outside
`-len..=len` raises `IndexError`. A row or column of another table raises
`ValueError`, and removing the only row or column raises `RpptxError`. Each of
these edits invalidates row, column, cell, paragraph, and run handles, because
they name indices, and the returned handle is captured after the change.

The Presentation Python `Table` handle has no built-in table style selector.
For a style GUID not exposed by the binding, the supported fallback is to save
the deck, replace only the matching table's `a:tblPr/a:tableStyleId` in the
slide XML inside the ZIP package, and reopen it. The writer must preserve all
other decompressed parts and validate internal relationship targets after the
edit. A style GUID supplied this way is a package-level edit, not a Python
`Table` API operation.

`TextFrame.autofit`
reports `none`, `normal`, or `shape` when the body carries an explicit choice.
`Run.font` reads the run's direct Latin name, size, and sRGB colour, while the
`Run.text` setter replaces only that run's text and preserves its typed and
unmodelled properties. `Run.hyperlink` returns a live `Hyperlink` whose
`address` reads the target of the run's `a:hlinkClick`, or `None`. Assigning
an address goes through the native `set_run_hyperlink`, which reuses the
slide's relationship to the same address and removes the old relationship
once nothing on the slide names it, so retargeting does not grow the part.
`None` or an empty string removes the hyperlink, as in python-pptx, an address
with a control character raises `RpptxError`, and the write does not advance
the revision. `Hyperlink.target_slide` reads and writes a run's jump to a slide
through the native `run_target_slide` and `set_run_target_slide`, as
`click_action.target_slide` does for a shape, which python-pptx offers only for
shapes.

Text formatting follows python-pptx names and value types. Every property
reads the direct value only, `None` when the element or attribute is absent,
and assigning `None` removes the direct value. Setters validate before they
write, raise `ValueError` for an out-of-range value, and leave the package
unchanged when the value is rejected or equals the stored one, so clearing an
absent value inserts no empty `a:pPr` or `a:rPr`. They change properties in
place and do not advance the revision. `rpptx` and `rpptx.enum.text` export
`MSO_AUTO_SIZE`, `MSO_ANCHOR`, `PP_ALIGN`, and `MSO_UNDERLINE` with python-pptx
1.0.2 member values, and `rpptx` and `rpptx.dml.color` export `RGBColor`.

- `TextFrame.margin_left`, `margin_right`, `margin_top`, and `margin_bottom`
  read the body insets as `Length`, converting a universal measure such as
  `0.1in` to EMU. python-pptx reports the implied 91440 and 45720 EMU defaults
  instead of `None`, which a placeholder does not have because it inherits
  its insets. `vertical_anchor` takes `MSO_ANCHOR`, whose `JUSTIFY` and
  `DISTRIBUTE` extension members name the `just` and `dist` anchors.
  `word_wrap` maps `square` to `True` and `none` to `False`. `auto_size` takes
  `MSO_AUTO_SIZE`, and choosing `TEXT_TO_FIT_SHAPE` again keeps a stored font
  scale. `autofit` keeps its string values.
- `Paragraph.alignment` takes `PP_ALIGN`. `line_spacing` reads a float number
  of lines from `a:spcPct` and a `Length` from `a:spcPts`. Assigning a
  `Length` writes exact points and any other number writes lines, as in
  python-pptx. `space_before` and `space_after` write points for an integer or
  `Length` and lines for a float, and read lines back as a float where
  python-pptx reports `None`. `left_indent`, `right_indent`, and
  `first_line_indent` use the python-docx names for `marL`, `marR`, and
  `indent`, where a negative first-line indent hangs.
- `Paragraph.bullet` reads the direct bullet choice as its character, `False`
  for `a:buNone`, `True` for an automatic number or picture bullet, or `None`.
  Assigning a character keeps the bullet colour, size, and font and replaces
  a preserved picture bullet. `False` writes `a:buNone`, `None` removes the
  direct bullet, and `True` raises because automatic numbering is not
  writable yet.
- `Paragraph.add_run(text="")` appends a run and returns it. An append
  renumbers nothing, so the paragraph and its other runs stay valid, as after
  `add_paragraph`.
- `Font` reads and writes the same properties for a run and for a
  paragraph's default run properties: `name` (`a:latin` typeface, keeping its
  other attributes), `size` (1 to 4000 points, as python-pptx validates),
  `bold`, `italic`, `underline`, `strike`, `all_caps`, `small_caps`,
  `highlight_color`, and `color`.
  `underline` reads `True` for `sng`, `False` for `none`, and an
  `MSO_UNDERLINE` member otherwise. `strike` reads `True` for a single or
  double strike, and assigning `True` keeps a double strike. `all_caps` and
  `small_caps` share the one `cap` attribute, so assigning `True` to one
  clears the other, and `small_caps` reads `False` for an explicit `all` or
  `none`. `highlight_color` reads the sRGB `a:highlight` as `RGBColor`, or
  `None`, and takes an `RGBColor`, a hex string with or without `#`, or a
  `(red, green, blue)` tuple. Another type raises `TypeError` naming the
  property.
  `color` still reads the direct sRGB solid fill as an `RRGGBB` string for
  compatibility, and the setter takes an `RGBColor`, any triple of 0 to 255
  integers, or a six-digit hexadecimal string. It changes an existing sRGB
  value in place, keeping transforms such as `a:alpha`, replaces any other
  colour, and `None` removes the direct fill.

`Presentation.try_replace_text(placeholder, replacement, *, expect=None)`
runs the native staged literal replacement over slides and speaker notes with
the GIL released and returns its count. When `expect` is given, the
replacement runs on a clone, and a count that differs raises
`ReplacementCountError`, an `RpptxError` subclass that carries `expected` and
`found` and words its message like `rpptx replace --expect`. The presentation
and its revision then stay unchanged. Otherwise the replacement is kept, and
the revision advances once when the count is nonzero. Without `expect`, zero
matches return zero rather than raise, as the rdocx `try_replace_text` does.
A nonzero count invalidates run handles only.
Only the CLI refuses zero matches, because it would write an unchanged copy.
`Presentation.replace_text` is an alias with the same count and optional
`expect` contract.

`Slide.add_comment` accepts an optional `shape_id` to anchor a modern comment
to a drawing element. `text_start` and `text_length` together select a range
in an ordinary shape's UTF-16 text. Invalid targets, duplicate text contexts
on one slide, and invalid ranges leave the presentation unchanged. Without
these arguments, the comment uses the
existing unknown-anchor form.

`Slide.try_replace_text(placeholder, replacement, *, expect=None, notes=True)`
and `TextFrame.try_replace_text(placeholder, replacement, *, expect=None)`
keep that contract over a narrower scope. The slide form covers the slide's
shapes, groups and table cells, plus its speaker notes unless `notes` is
false, and leaves every other slide untouched. The text frame form covers that
frame only. The count, the error and its message, the unchanged presentation
and revision after a mismatch, and the single revision advance are those of
the presentation form, so a replacement that changes something invalidates
held run handles, while the slide or frame it was called on stays valid. Neither
form clones the presentation. The slide form works on a copy of that slide
and its notes and checks that they serialize. The frame form works on a copy
of the text body when `expect` is given, in place otherwise, and serializes
nothing, because replacing run text cannot make a body fail to serialize.

`Presentation.validate()` runs the native `validate` with the GIL released and
returns a tuple of frozen `ValidationIssue` snapshots in native order. Each
carries a `kind` that names the native variant in snake_case, such as
`duplicate_shape_id`, and a `message` equal to the line `rpptx validate`
prints for that issue. A clean presentation returns an empty tuple.

The presentation binding exposes `to_pdf`, `render_slide_to_png`,
`render_all_slides`, `to_notes_pdf`, and `render_all_notes` through the native
deterministic facade. Every render call releases the GIL.
`Presentation.text_layout(*, width_factor=1.0)` returns the native
deterministic text layout as a tuple of frozen `TextFrameLayout` snapshots in
slide and draw order. Each carries its zero-based slide index, optional shape
id and name, the effective autofit mode as `none`, `normal`, or `shape`,
`frame` and `usable` `BoundingBox` values, `font_scale`, `height`, `overflow`,
and a tuple of frozen `TextLineLayout` snapshots with paragraph index, text,
bounds, baseline, and font size. Every value is a float in points, as in the
rdocx `BoundingBox` and `LayoutFragment`, unlike the EMU `Length` values of
`Shape` and `Font`. The call releases the GIL, and a width factor that is not
finite and positive raises `RpptxError`, as an invalid raster DPI does.

A `Slide` exposes optional speaker-note text as a readable and writable
property. Assigning it on a slide without notes creates the notes slide, and
the notes master when the deck has none, through
`Presentation::set_notes_text`. A successful notes assignment publishes the
native staged mutation and keeps every handle valid, as python-pptx does. A
rejected assignment leaves package bytes and revisions
unchanged. A `Slide` also exposes an ordered tuple of frozen `Comment`
snapshots. Each comment contains an ordered tuple of frozen `CommentReply`
snapshots, and the presentation exposes an ordered tuple of frozen
`CommentAuthor` snapshots.
Author, comment, and reply additions accept native GUID and RFC 3339 strings.
Comment and reply moves retain native final-position semantics.
`Slide.resolve_comment(comment_id)` marks a thread resolved and treats a reply
id as unknown. `Slide.remove_comment(comment_id)` removes a thread with its
replies, or one reply. Both use the native staged operations of
`rpptx comment resolve` and `remove`. A collaboration operation renumbers no
handle path, so every handle stays valid. Constructor or
native validation failure publishes no candidate and leaves existing handles
valid.

## Native Word facade stability
Native Rust exposes `BibliographySourceKind`, `BibliographyContributorRole`,
`BibliographySourceField`, `BibliographyStyle`, source/person/contributor/property
records, source/style inspection records, `CitationOptions`,
`CitationSourceOptions`, `BibliographyOptions` and `BibliographyUpdateReport`.
`Document::bibliography_sources`, `add_bibliography_source`,
`replace_bibliography_source` and `remove_bibliography_source` inspect and mutate
qualified source records. `bibliography_style`, `set_bibliography_options`,
`insert_citation` and `insert_bibliography` provide checked native authoring.
`update_bibliography` and `update_bibliography_with_default_locale` explicitly
materialize admitted caches through one staged transaction.

Metadata models all seventeen source kinds, sixteen roles and twelve style
identities. Formatting coverage is separate. APA's 210 measured ordinary dense
locale configurations do not promise arbitrary sparse, plural, corporate or
Unicode input. The eleven other bibliography branches currently require
numeric1033, Book, Author-only contributors, a small property set, one person
and nonempty ASCII Last/First/Title/Year/City/Publisher values. Citation locale,
kind and modifier admissions have their own contract. A recognized unfinished
standard branch returns an explicit error and preserves the whole document.
Noncatalogue paths retain caches and return diagnostics in the update report.
Locked and protected owner boundaries remain distinct. Source LCID and field
selectors precede a caller-supplied actual default locale, with no ambient or
guessed en-US fallback. Non-ASCII collection sort keys refuse.

These additive pre-1.0 APIs are native Rust only. Python, WASM and CLI expose no
dedicated bibliography source or update API. Their existing document and field
compatibility checks remain required over the shared Rust changes. The remaining
full catalogue belongs to F-X192, rather than an implied wrapper capability.


Native Rust adds `IndexEntry`, `AuthorityEntry`, `IndexOptions`,
`TableOfFiguresOptions`, `TableOfAuthoritiesOptions` and
`GeneratedTablesReport`. `Document::insert_index_entry` and
`insert_authority_entry` author checked complex source markers.
`insert_index`, `insert_table_of_figures` and `insert_table_of_authorities`
insert dynamic generated owners at checked story boundaries.
`rebuild_generated_tables` atomically rebuilds supported caches.
These are additive pre-1.0 APIs with no Python, WASM or CLI additions.

Authority category `None` inserts separate numbered fields for every populated
category in native order and returns the first checked location. It never
encodes All as zero or as an omitted category. INDEX hyperlink authoring retains
its formatting request in the paragraph mark's Hyperlink character style and
in the generated label caches. Entry hyperlinks point to checked internal
bookmarks. Requested leaders and existing producer styles remain intact.

The pre-1.0 `rdocx-layout::TableRow` projection carries a public
`cant_split: bool` alongside its header and height facts. Direct and
style-resolved `w:cantSplit` values therefore reach pagination. External Rust
struct literals for `TableRow` must name the new field. The authored `rdocx`
row facade and Python, WASM, and CLI authoring surfaces do not change.

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
one exact style, `numId`, and level tuple. Python binds definition and instance
creation and the style link, as described for the Python `Document`. WASM and
CLI bindings gain no numbering graph authoring surface.
`ListNumberFormat::from_name` reads a `w:numFmt` name, and a name outside the
standard set is `Other`. `Document::numbering_level` retains namespace
declarations without counting them in `has_unmodeled_properties`. Real extra
attributes and children remain reported. Modeled producer metadata and raw
property or typed-leaf preservation overlays keep their existing reader
behavior. The richer numbering authoring completeness checks retain their
separate admission contract.

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
`Document::package_class` reads the exact main-part override. `to_bytes_as` and
`save_as_package_class` select an output class on a staged copy without
removing executable or opaque parts. `Document::save` and `to_bytes_for_path`
write the class that a `.docx`, `.docm`, `.dotx`, or `.dotm` extension names,
compared without regard to case, so a template saved as `.docx` declares a
document. `to_bytes`, `save_encrypted`, the Flat OPC saves, and a save to any
other extension retain the opened class. Path saves stage a sibling file and
atomically replace the destination whether the package class changes or stays
the same. When the class changes, a macro-free extension fails before anything
is written if the main part carries a `vbaProject` relationship, whatever the
source class, because the project would remain in a file that claims to carry
none. A macro-enabled package without one converts. `from_flat_opc_bytes`, its
limits overload, `open_flat_opc`, `to_flat_opc_bytes`, and `save_flat_opc`
provide bounded strict Flat OPC interchange through the same `Document` and
`OpcPackage` owners. These are additive pre-1.0 native APIs. The Python
`Document.save` and `Presentation.save` methods and the package outputs of both
CLIs follow the same extension rule, because they go through `save` or
`to_bytes_for_path`. WASM saves take no path and retain the opened class. No
binding gains a class selector or a Flat OPC entry point.

Native Rust also exposes the additive pre-1.0 `WordCreationProfile` enum and
`Document::new_with_profile`. `Minimal(WordPackageClass)` preserves the compact
source-built graph. `WordCompatible(WordPackageClass)` owns the standard blank
support parts, and `Document::new()` selects its DOCX form. Macro-capable
profiles select package identity without manufacturing executable content.
Python, WASM, and CLI construction continues through `Document::new()` and
therefore receives the compatible DOCX default without a new selector surface.

The minimal profile defines `Normal` and `Heading1`. The Word-compatible
profile includes the common Word styles at construction.
`Document::add_common_styles` adds the Word built-in styles most documents use
and python-docx's default template defines: `heading 2` to `heading 9`,
`Title`, `Subtitle`, `No Spacing`, `Quote`, `List Paragraph`, `caption` and
`Table Grid`. Each has the ID, name, UI priority, visibility flags and
paragraph and run formatting Word writes for it under the Office theme, with
theme colours written as literal values as in the default `Heading1`. None
names a font. `Table Grid` has no `Normal Table` base, since a new document
defines none, so it carries that base's cell margins itself, 108 twips left and
right. Without them Word gives its cells no side padding and the text touches
the grid. A style whose ID, or whose name regardless of case, the document
already defines is skipped, and the call returns how many it added. None is a
default style, so content that names none of them lays out as before. A
document without a `Normal` paragraph style is rejected unchanged.
`Document::new()` and a new Python `Document()` both expose these styles, so
python-docx code that names them finds them. A Python document opened from a
file or bytes keeps exactly the styles it has. WASM and CLI construction uses
the same Word-compatible native default.

Native Rust re-exports `StyleType`, `TableStyleRegion` and
`ConditionalTableStyle`, and `CT_PPr`, `CT_RPr` and `HalfPoint`, the property
and unit types `StyleBuilder` takes, so a caller of the facade alone can give a
style its formatting. `StyleBuilder` authors paragraph, character, and table
styles with inheritance, reciprocal links, next styles, UI flags, base
properties including a table style's own row and cell properties, and
conditional table regions carrying all five property layers. `add_style` is
fallible in the pre-1.0 API. `set_style`, `remove_style`,
`set_default_style`, and `validate_style_graph` use the same `Result` boundary.
For common formatting, the existing builder also offers fluent `alignment`,
`space_before`, `space_after`, `indent_left`, `font`, `size`, `bold`, and `color`
methods. Lengths use the paragraph facade's twip conversion and font sizes use
`HalfPoint::from_pt`. An explicit font fills all four script slots and clears
their theme references. An explicit colour clears the theme colour, tint, and
shade. Callers can mix these methods with `paragraph_properties(CT_PPr)` and
`run_properties(CT_RPr)`, with later calls taking precedence on overlapping
builder fields. This is an additive native Rust API in unreleased 0.14.0.
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
WASM and CLI retain style package and render behavior without new style
mutation entry points. Python exposes style creation, checked formatting
updates, removal and default selection, as described for the Python `Document`.

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

Python exposes the core properties as `Document.core_properties`, a
`CoreProperties` handle with the python-docx attribute names `author`,
`category`, `comments`, `content_status`, `created`, `identifier`, `keywords`,
`language`, `last_modified_by`, `last_printed`, `modified`, `revision`,
`subject`, `title` and `version`. `author` maps to `dc:creator` and `comments`
to `dc:description`. Text properties read as an empty string when absent and
accept at most 255 characters, as in python-docx. `revision` reads as an
integer, zero when absent or unreadable, and accepts only a positive integer.
The three dates read as timezone-aware UTC `datetime` values, parsed from
W3CDTF as python-docx parses them, and accept a `datetime`, a naive one being
taken as UTC. A date that cannot be read, has an offset of a day or more, or
leaves the `datetime` range in UTC reads as `None` rather than raising. Assigning `None` or empty text removes a property. Each
assignment replaces the native model through `Document::set_core_properties`,
which creates `docProps/core.xml` with its package relationship and content
type when the document has none. It changes no content, so the revision and
every handle stay valid. A value of the wrong type raises `TypeError` and an
out-of-range one `ValueError`, both before any change.

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
`StoryItemSnapshot::is_direct_child` tells a direct child of any story owner
from an item nested in another item of that owner, such as an inline content
control or a field inside a paragraph.
`Document::set_story_text` resolves a checked operation-scoped
location against a staged package and publishes only a serialized and reopened
candidate. These additions are native Rust APIs on the pre-1.0 `rdocx` crate.
Python exposes story and story item snapshots, including each item's `xml` as
bytes. Cached complex-field text agrees between owned bulk snapshots and direct
native reads when the field spans sibling runs or nests inside another field.
Snapshot reads preserve source locations, XML, direct-body coordinates and the
document revision, leaving held live handles valid. Nested instruction caches
remain distinct from their enclosing field's visible result. Story links retain
physical source ordering and relationship scope through the same bounded excerpt
projection. `Document.set_story_text` takes a `StoryItem` snapshot and resolves its
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
value. `pop_content` also accepts a `StoryItem` naming a direct child of any
story, and `insert_content` accepts a `StoryItem` for the boundary before that
item or a `Story` for the end of that story, so a paragraph authored in the
body can move into a header or footer. `try_replace_text` and `replace_all_regex` return exact native counts.
Successful structural mutations stale handles once, successful replacements
stale them only when their count is nonzero, and every rejected operation
leaves package bytes and handle revisions unchanged. The native and Python
changes are additive pre-1.0 surfaces scheduled for `rdocx` 0.14.0. WASM and
CLI gain no corresponding surface. Python resolves clone and move source and
destination locations from one owned story inventory. A rejected clone names
`source` when it is not a Paragraph or Table handle and names `destination` as
a direct body index when it is not an integer.

Native Rust exposes owned `DocumentFragment` and non-exhaustive
`FragmentConflictPolicy` values. `DocumentFragment::from_range` captures a
nonempty half-open block selection in any supported story owner. Existing
two-segment block-control paragraph paths use the enclosing control content.
Inline items are rejected because the fragment API has no inline boundary.
Final body section properties require an explicit selection ending at the
main-body boundary and a main-body destination.

`Document::import_fragment` inserts at a checked compatible block boundary and
uses that owner's physical part for relationship references. Caller policy
chooses equivalent reuse independently for styles, numbering and related leaf
parts. The import closes note, comment, binding-store and reachable OPC
companions and remaps conflicting identities in one package-authoritative
transaction. The candidate serializes and reopens before publication, and a
failed import leaves the destination unchanged. These are additive pre-1.0
APIs in the published `rdocx` crate. Python, WASM and CLI gain no fragment-import
surface here.

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
`Document::set_hyperlink_url` and `remove_hyperlink` edit the link at one
position of a story's `story_links` list in any story, through the relationship
set of that story's part. A relationship that only this link references is
retargeted in place. A shared one stays with its other references and the link
gets a new relationship. An anchor link becomes external and loses its anchor.
Removal keeps the link's runs in place and clears the built-in Hyperlink and
FollowedHyperlink character styles, as Word's Remove Hyperlink does. Both remove
a hyperlink relationship that nothing references any more. Only links that
`story_links` lists are in reach. An empty `w:hyperlink`, a link inside
`w:fldSimple`, a link in a text box inside `mc:AlternateContent`, and a
HYPERLINK field are not. Python exposes
`Document.set_hyperlink_url` and `Document.remove_hyperlink` for a `Hyperlink`
record. A snapshot from `Document.hyperlinks` resolves at its recorded position
only while its story keeps the same link paths and texts and the link keeps
its fields, and a record built from the public fields resolves only when
exactly one link matches. Neither call moves content, so live handles stay
valid.
`Document::set_picture_size` resizes every main-document `pic:pic` drawing
whose blip names one image relationship. It writes `wp:extent` and the `a:ext`
of `pic:spPr/a:xfrm`, scales `wp:effectExtent` with the extent on each axis,
and leaves anchor positions alone. It rejects a zero size. A VML `w:pict`
picture on the relationship is not resized, the relationship of an SVG blip
extension is not matched, and a picture inside a group is not resized. Python
exposes `Document.set_picture_size` with EMU sizes, and it keeps live handles
valid too.

Native Rust also exposes concrete borrowed `SectionRef` and `Section` handles.
Each handle reports its zero-based document ordinal, schema-final ownership,
and configured geometry while retaining access to the complete section
properties. Both handles read page size, orientation, margins, gutter,
equal-width columns, page-number start, header and footer distance, title-page
state, and break type. The mutable handle adds checked setters for every value,
normalizes page dimensions when setting orientation, and rejects invalid or
out-of-range inputs before changing any field. The margin, gutter, and header
and footer distance setters fill every other `w:pgMar` value the section lacks
with the default that layout already assumes for it, so the written element
carries all seven attributes `CT_PageMar` requires. `Document` adds
`section_count`, `sections`, total `section` and `section_mut` lookup, and
fallible staged `insert_section` and `remove_section` operations. Its older
final-section geometry convenience setters remain infallible and unchecked.

Native Rust also exposes non-exhaustive `HeaderFooterKind`, the existing
`HdrFtrType`, and owned `SectionStory`. `Document::section_story` resolves one
effective default, first, or even header or footer and reports its source
section and inherited state. `create_section_story`, `link_section_story`,
`inherit_section_story`, `unlink_section_story`, `replace_section_story`, and
`remove_section_story` are staged fallible operations. Rich content remains
addressed by the returned `StoryId` through the common story API. The facade
also exposes `even_and_odd_headers` and `set_even_and_odd_headers`, while first
story creation enables section `titlePg`. These are additive pre-1.0 native
Rust APIs. Python exposes immutable inspection snapshots, the default header
and footer text setters `Document.set_header` and `Document.set_footer`, and
`Document.create_section_story`, `link_section_story`, and
`unlink_section_story`. These take a section index, a `header` or `footer`
kind, and a `default`, `first`, or `even` variant, the names
`HeaderFooterVariant` reports, and return the resulting `Story`. An unknown
name raises `ValueError` and a section index out of range raises `IndexError`,
both before any change. The native operations publish a reopened package, so
a successful call advances the revision once, even when the variant already
had its own story. Rich story content is authored with the typed body
API and moved with `pop_content` and `insert_content`, which accept story
coordinates. WASM and CLI gain no corresponding binding surface.

Native Rust `Document::create_footnote(&ContentLocation, &str)` appends a normal
footnote and its reference to one direct body paragraph in a staged operation.
It returns the stable internal note ID. `footnote_story(i32)` resolves that ID
to its current checked `StoryId`. `move_footnote_before(i32, i32)` reorders the
note elements without changing IDs, and `remove_footnote(i32)` removes the note
and all matching body references together. All four methods are fallible.
Rich paragraphs, tables, fields, links, pictures, content controls, and
comments use the common story APIs on the resolved footnote story. This is an
additive pre-1.0 native API. Python, WASM, and CLI gain no matching authoring
entry point.

Native Rust `Document::create_endnote(&ContentLocation, &str)` stages a normal
endnote with its reference in a direct body paragraph and returns its stable
internal ID. `endnote_story(i32)` resolves a current checked story identity.
`move_endnote_before(i32, i32)` reorders exact note elements without changing
IDs. `remove_endnote(i32)` removes a normal endnote and every matching body
reference together. All four methods are fallible. Endnotes allocate IDs
independently from footnotes and use the common rich story operations,
including part-scoped pictures, links, and comment anchors. This is additive
pre-1.0 native API. Python, WASM, and CLI gain no corresponding authoring
entry point.

`CT_SectPr` adds typed page-number start and raw child-position state, while
`PageFrame` adds `displayed_page_number` beside its physical `page_number`.
These model and handle additions are additive APIs on the published pre-1.0
Rust crates, though exhaustive struct literals can require new fields. The
published `CT_SectPr.header_refs` and `footer_refs` types remain
`Vec<HdrFtrRef>` with the complete native vector surface. WASM and CLI gain no
section mutation entry point and retain their existing package and render
behavior.

Python `Section` values stay frozen snapshots. `Document.update_section(index,
**values)` edits one section through the checked native setters and returns
its new snapshot. The keywords are the snapshot's own field names, from
`orientation` and `page_width` to `break_type`, with EMU lengths and the
`ST_PageOrientation` and `ST_SectionType` spellings. The native setters write
page size, the four margins, equal-width columns and the header and footer
distances as pairs or quartets, so a value given alone keeps its partners as
the section already has them. A missing partner takes the native layout
default, including Letter page size, one-inch margins, half-inch header and
footer distances, and 720 twip column spacing. Column partners come from the
equal-width view the snapshot reports. A one-sided column edit on explicit
unequal-width tracks raises `ValueError` rather than rewriting those tracks.
Page size applies before orientation, which then normalizes
the dimensions as the native setter does.
The whole call is atomic: a name or partner problem raises before any change,
and a value a native setter rejects restores the section. Section edits move
no content, so they keep live handles valid and do not advance the revision.
`Document.insert_section(index)` and `remove_section(index)` call the staged
native operations, reject an out-of-range index with `IndexError`, and
advance the revision once, because they add or merge body content.

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
tuple instead of rebuilding it per returned record. Hyperlink discovery also
inventories every exact link namespace scope in one source pass, then extracts
display text from bounded namespace-complete fragments rather than reparsing
the story prefix for each record.

`Document::rebuild_toc()` is an additive pre-1.0 native Rust operation. It
updates only supported existing main-story TOC fields with deterministic
bundled-font page targets and returns `TocRebuildReport` with entry and newly
allocated bookmark counts plus exact retained-field diagnostics in physical
source order. `diagnostic_count()` is derived from the owned diagnostic
collection. A document without a TOC is unchanged and returns empty counts and
diagnostics. The cached entry paragraphs are replaced, so comment and bookmark
markers on them, including those in a table of contents content control, are
dropped with them, and a comment anchored only there stays unanchored in the
comments part. `rdocx-cli toc rebuild` publishes the validated result to an
explicit output and reports the counts through a schema-1 main-story record.
Python exposes the same operation and returns diagnostics as an immutable tuple
with a derived `diagnostic_count` property. WASM does not expose this operation.

`Document::insert_toc(index, max_level)` writes a refreshable table of contents at
a direct body index: a title paragraph and one entry per `Heading1` to
`HeadingN` paragraph, each linked to a new `_TocN` bookmark on its heading. It
writes a dynamic `TOC` field around the entry cache, so `rebuild_toc` updates
it after heading changes. Native insertion is staged and returns an error
without changing the document if bookmark allocation fails. Python
`Document.insert_toc(index, max_level=3)` binds it. An index past the body end
raises `IndexError` and a level outside 1 to 9 raises `ValueError`, both before
any change. Native insertion errors raise `RdocxError`. Success
advances the revision once.

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
The additive pre-1.0 Rust surface also exposes `create_building_block`,
`create_building_block_from_fragment`, `update_building_block`,
`update_building_block_from_fragment`, `remove_building_block`,
`building_block_fragment`, `insert_building_block` and
`bind_building_block_placeholder`. Typed creation accepts dependency-free
content. Fragment creation and insertion use `FragmentConflictPolicy` and
the source package dependency closure. Mutation checks the complete
`BuildingBlockInfo` snapshot, rejecting stale ordinals and changed values.
Glossary transfers omit review anchors and definition import before dependency
capture, including selected note and textbox review content. Reverse capture
omits before main-owned fragment construction. Dependency-free typed creation
retains comment-marker refusal. Direct replacement omits newly supplied review
markers while metadata-only updates preserve the original body. Actual local
comment relationships isolate their numeric ids from main comments. Unproved
local ownership and glossary-local note projection refuse atomically.

Placeholder binding keeps the existing control discriminator and updates
selection properties only on existing document-part control variants.
Python, WASM, and CLI bindings remain unchanged.

The additive pre-1.0 native whole-story text surface exposes `try_set_header`,
`try_set_footer`, `try_set_first_page_header` and `try_set_first_page_footer`,
each returning `Result<()>`. Existing infallible signatures remain compatibility
wrappers with their established panic convention. Complete header/footer
output proof refuses XML-invalid supplied text before publication, retaining
the OPC part and code-point diagnostic through the fallible route and exact
original bytes through either route. Raw and image setters stage
all dependency allocation and remapping before reconciliation and publication.
Python's existing `set_header(text)` and `set_footer(text)` return `None`,
propagate checked native errors, and bump the document revision once only after
success. Refusal preserves package bytes and live handles. No additional
Python, WASM or CLI setter surface is introduced.

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
text. Heading and navigation labels use only bounded projected runs, those of
content controls and tracked insertions included. A content control, a tracked
insertion or move in, a smart tag and inline custom XML are flattened: what
they hold is exported in place and the wrapper is diagnosed. Deleted and
moved-away text is left out and diagnosed.
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
serialization error, and creates no partial output. Python exposes the fallible
method as `Document.try_replace_text`. The WASM binding keeps its infallible
`replacePlaceholder`, which calls `replace_text`.

Literal replacement reaches the body, tables, content controls at every level,
headers, footers, footnotes, endnotes, the text boxes of the body, headers, and
footers, and the labels of the charts in the body. Regex replacement reaches
the same stories except chart labels. Every normal note of the
relationship-resolved footnotes and endnotes parts is searched, its tables and
block controls included. The separator, continuation separator, and
continuation notice entries are not, and neither is an untyped entry at id 0 or
below, as the typed notes reader and the story walkers read them. A text box
inside a note is not searched. A notes part is rewritten only when a note
changed, and then only the changed children of that note are serialized again.
A match inside a tracked insertion or move destination is replaced inside it.
So is a match inside a smart tag, an inline custom XML element, or the cached
result of a simple field, whose instruction is never matched. Deleted text is
not matched, except in a text box inside a deleted run, which the text box pass
rewrites like any other text box. The Rust facade, Python `try_replace_text`
and `replace_all_regex`, the WASM replacement method, and
`rdocx replace --expect` share these counts.

Literal replacement reaches the body, tables, content controls at every level,
headers, footers, footnotes, endnotes, text boxes, and chart labels. Regex
replacement reaches the same stories except chart labels. Every normal note of
the relationship-resolved footnotes and endnotes parts is searched, its tables
and block controls included. The separator, continuation
separator, and continuation notice entries are not, and neither is an untyped
entry at id 0 or below, as the typed notes reader and the story walkers read
them. A notes part is rewritten only when a note changed, and then only the
changed children of that note are serialized again. A match inside a tracked
insertion or move destination is replaced inside it, and deleted text is never
matched. The Rust facade, Python `try_replace_text` and `replace_all_regex`,
the WASM replacement method, and `rdocx replace --expect` share these counts.

Every Word replacement, literal, regular
expression, or template, edits every copy of a text box that Word writes in
`mc:AlternateContent`. The count is that of the first `mc:Choice` that holds a
text box, the DrawingML copy that layout draws and the story walkers read. The
VML `mc:Fallback` and any later Choice are edited whatever they hold, a
Fallback that a story edit left behind its Choice included, and never counted.
A replacement can therefore change a Fallback and report 0. When no Choice
holds a text box, a Fallback text box is replaced and counted as any other
text box. Replacement handles `mc:AlternateContent` at any depth outside
another text box, while the story walkers read it only as a child of a run.
A story edit of such a text box changes the Choice only, see
`03-architecture.md`. `redact_text` removes the literal from every copy and
still counts each copy.

`Document::try_replace_all_expected` takes ordered `(placeholder, replacement,
expected)` pairs and returns the count of each pair. Each pair runs over the
whole staged document after the pairs before it, so a later pair sees what an
earlier one wrote. When a pair gives an expected count and finds another, the
document stays unchanged and the inner result is a `ReplacementCountMismatch`
naming the pair index, placeholder, expected count, and found count. The outer
error is a staging failure. The legacy `replace_all` keeps its unordered map
and panicking preflight.

Python `Document.try_replace_text(placeholder, replacement, *, expect=None)`
and `Document.replace_all(pairs)` share one contract with the rpptx binding.
`replace_all` takes `(old, new)` or `(old, new, expected)` tuples and returns a
tuple of counts. A count that differs from `expect` or from a pair's expected
count raises `ReplacementCountError`, a subclass of `RdocxError` with
`expected`, `found`, and `index` attributes, and leaves the package bytes and
live handles unchanged. `index` names the failing pair of a batch and is
`None` for `try_replace_text`, whose message is the one `rdocx replace
--expect` prints. The error pickles. Zero matches without an expected count
return zero rather than raising. A replacement advances the revision once when
it replaced something. WASM gains no corresponding method.

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
additive on the native pre-1.0 facades. `rdocx` re-exports `PdfConformance`, so
a caller can name the profile without depending on `oxml-pdf`. Python
`Document.to_pdfa_deterministic(profile="pdfa-2b")` accepts `pdfa-2b` or
`pdfa-3b`, raises `ValueError` for any other name, releases the GIL, and maps a
conformance failure to `LayoutError`. WASM and CLI method names and dependency
selections remain unchanged.

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
setters remain available. These are additive pre-1.0 native APIs. WASM and CLI
do not gain new table-property methods, and Python binds them as described
under the Python API shape. Every binding's owned `rdocx::Document` remains
package-preserving when native code uses the new operations.

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
publication. These are additive pre-1.0 native APIs. WASM and CLI gain no row
or cell methods. Python binds the checked row and cell setters and the span
and vertical merge operations, as described under the Python API shape.

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
The additive native `add_comment_on_text` anchors a comment on the zero-based,
non-overlapping, case-sensitive occurrence of a literal text in main-story
paragraphs, through tables and block content controls. It splits the runs at
both ends of the match, anchors the runs between the splits like
`add_comment`, and refuses a missing occurrence or a match that cannot be
anchored exactly without changing the document. Tabs and breaks have no width
in the literal text, and a match whose range would also show text that the
literal text leaves out, such as a field result, is not exact. Python exposes
it with keyword `author`, `text`, `occurrence`, `initials` and `date`
arguments.
Native `move_comment(id, StoryRunRange)` and `move_comment_to_text(id, anchor,
occurrence)` move an existing root thread without changing its numeric id or
metadata. Python exposes `move_comment(id, range)` and
`move_comment_to_text(id, anchor, *, occurrence=0)`, advancing the document
revision exactly once after successful package publication. Refusals retain
bytes and revisions. `rdocx comment move <file> <id> --text <anchor>
--occurrence <n>` uses the existing required output, JSON and atomic save
conventions. Unknown ids, replies, stale destinations and unsupported source
ownership produce checked errors.
The Python `add_comment` and `reply_to` methods expose the same value as the
optional `date` keyword, and `rdocx comment add` and `rdocx comment reply` as
the optional `--date` flag. Omission writes no date and remains deterministic.
Returned ids keep naming the same comment or reply after rdocx save and reopen,
although third-party editors may renumber them. `RunPosition` and `RunRange`
define top-level paragraph run
boundaries with an inclusive start and exclusive end. Run boundaries count the
accepted-view runs that `Paragraph.runs` and `rdocx-cli text --json` list,
including the runs inside inline content controls and tracked insertions, and
a range that cannot be anchored exactly is an error rather than a shifted
range. `Document::split_run`
splits one direct-body run at a Unicode scalar offset of its literal text, so
such a boundary can fall inside what was one run. Its first argument is the
same direct body child index as `RunPosition` and `find_content_index`. An
index that names a table, a block content control, or preserved XML is an
error naming that kind. The second part keeps the run
properties and the enclosing hyperlink. Tabs, breaks, fields, drawings,
references, and preserved raw children have zero width. Zero and the literal
text length return the existing boundary without mutation. Interior success
returns the new continuation index. Python exposes the same method and also
accepts a `Paragraph` handle in place of the index, which reaches paragraphs
inside block content controls. A table cell paragraph handle is refused. The
binding revision advances only when a continuation is created.
`Paragraph::remove_run` removes one run at the index `Paragraph::run` counts,
and Python exposes it as `Run.remove()`. The comment, bookmark, and permission
markers around the run stay in place. A run inside a hyperlink, an inline
content control, or a tracked insertion is removed inside it. A hyperlink or
a tracked insertion left with nothing in it is removed too. An emptied content
control stays with its properties, as Word keeps it to show its placeholder,
and a removed hyperlink leaves its relationship in place. A run holding a
comment, footnote, or endnote reference, part of a complex field whose other
parts are in other runs, or part of a tracked move destination is refused
without change. A field whose parts are all in the
paragraph is one run and is removed whole. A removal advances the binding
revision. `CommentRef` exposes
comment metadata, text, parent identity, resolved state and checked accepted-view
anchor access without permitting part-local mutation. Frozen Python `Comment`
snapshots retain the seven original keyword arguments and add optional
`anchor_text` and `anchor` values, both defaulting to `None` for manually
constructed records. Equality retains the original seven metadata fields, since
derived revision-bound locations are not comment metadata identity. Document
listing materializes real `StoryRunRange`
endpoints with checked story identity, paragraph path and revision. Paragraphs
directly inside block controls retain their two-segment path and actual containing
body index. Recursive facade handles use the actual direct paragraph identity,
so nested control or table descendants have no two-segment location. Non-body
owners never acquire a body index. Snapshots remain readable
after mutation, while reuse of their stale range refuses. Listing propagates
extraction errors instead of converting unsupported sources to orphans.
CLI `comment list --json` retains metadata and adds nullable `anchor_text` and
`anchor`, with explicit start/end story kind, part, owner index, paragraph kind,
index path, run index and containing body index. Its scope is `all_stories`.
Empty point text differs from orphan null text, and replies have no invented
parent range. `rdocx-cli comment` lists, adds, replies to, resolves, and
removes comments. Add ranges use explicit zero-based, half-open body paragraph
and run coordinates, and the run coordinates count the runs that `text --json`
lists. `comment add --anchor TEXT`, with an optional zero-based
`--occurrence`, replaces those coordinates with `add_comment_on_text` and keeps
its refusals. Every mutation publishes a complete validated document
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
direct range, current text, and marker issue. The range counts paragraphs
recursively through tables and block content controls. The direct range uses
the `RunPosition` body index that `add_bookmark` takes and is present only when
both markers sit in direct body paragraphs. Its run indexes stay the
accepted-view boundaries of the range, which are the run indexes
`add_bookmark` takes. Insertion validates the Word name and both
boundaries, rejects duplicate or producer-reserved names, and returns the
allocated nonnegative id. The shared recursive `Field` model retains the
complete `REF` and `PAGEREF` instruction, target argument, cached display,
dirty state, source form, and producer XML. Python `Document.bookmarks`
returns a tuple of frozen `Bookmark` snapshots with the same fields, both
ranges as `RunRange` values or `None`. `Document.add_bookmark(name, range)`
returns the id. Its markers sit between runs, so it keeps live handles valid
and does not advance the revision. A rejected name or range raises
`RdocxError` and leaves the document unchanged. WASM and CLI consumers keep
their existing surface and preserve the typed content when they save the owned
document.

The additive native pre-1.0 story range API exposes `StoryRangeKind`,
`StoryRangeRef` and `Document::story_ranges` with immutable story-qualified
`StoryRunRange` endpoints. `add_story_bookmark`,
`add_story_permission_range`, `add_story_proofing_range`,
`move_story_range` and `remove_story_range` validate one physical owner,
accepted run boundaries and pair order before committing. Permission starts
retain an editor or group, while proofing pairs use `spell` or `gram` and no
numeric id. Removing a comment pair retains its reference and definition as
a point comment. Moving the pair relocates the reference with it. Existing
Python, WASM and CLI APIs remain source compatible and preserve these markers.

Native Word callers author captions with `CaptionOptions` and `CaptionTarget`,
sequence fields with `SequenceOptions`, and REF fields with
`CrossReferenceOptions` and `CrossReferenceNumber`. `insert_sequence` and
`insert_cross_reference` take checked accepted `StoryRunPosition` boundaries.
Caption targets expose the actual allocated whole-caption, label-and-number
and number bookmark names. Invalid instructions, ambiguous physical ownership
and rejected ranges publish no partial change. These additive pre-1.0 native
APIs add no Python, WASM or CLI methods.

Native Word callers evaluate fields with `Document::evaluate_fields` and an
explicit `FieldEvaluationContext`. `FieldDateTime` supplies deterministic civil
time. Caller maps supply merge values and included text, including
`source#bookmark` keys for bookmark-scoped includes. Each `FieldEvaluation`
records a snapshot-local visible document-order index, original instruction,
stored display, and a `FieldOutcome` that is resolved text, pagination deferral, a
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
methods and continue to preserve the same package content. SEQ and REF include
qualified physical related and selected text-box owners. Other fields preserve
their established discovery boundary. Hidden ordinary fields keep their stored
cache and raw attributes. Public evaluation indices do not identify physical
layout source fields. Python exposes
`Document.update_layout_backed_fields() -> LayoutBackedFieldUpdateReport` and
the count-only `Document.update_page_fields() -> int` wrapper. The frozen owned
report carries separate PAGE, NUMPAGES, PAGEREF, SECTION and SECTIONPAGES counts
plus an immutable diagnostic tuple. Its total is the sum of all five counts. Both operations release the GIL while native deterministic
layout runs and advance the document revision only when a cache is written, so
handles taken earlier then raise `StaleElementError`. The new `field_source`
member of `oxml-layout`'s `TextSegment`, `GlyphRun`, and
`MultilingualGlyphRun` is a pre-1.0 source break for native struct literals.

Native immutable layout snapshots expose page_sections, field_placements and
bookmark_page_section while preserving bookmark_page compatibility. WordStory
has a physical text-box variant with actual OPC part and logical owner path.
Hidden ordered cached-field segment, section inventory and physical text-box owner
accessors support checked staging. The hidden Field.cached_display_owner_is_locked
accessor reports effective cached-owner ancestry and conservatively protects an
owner outside the cached tree. LayoutInput adds actual story part names,
authoritative optional rich story bodies and projected footnote compatibility
booleans. CT_Shape adds optional text_body and source_text_box_owner projections.
Format-neutral TextSegment, GlyphRun and MultilingualGlyphRun add optional
note_reference_source for structural ownership independently of exact text spans.
Hidden physical revision/control metadata accessors and the existing CT_R
selected-MC raw-slot predicate support layout-only binding without changing
producer serialization. CT_P.source_runs borrows initialized physical revision
projections while the existing runs reader keeps its prior scope. Field readers
expose effective instruction text, typed cached result runs and owned comment
range markers. Cache setters validate their typed result on a staged clone and
preserve source form, instructions, lock controls and unrelated XML. Absent projections retain legacy callers. These pre-1.0 public struct and enum
additions require exhaustive constructor updates and package consumer checks.
Python gains the two report count getters. WASM and CLI gain no update methods.

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
captured raw WordprocessingML subtree. The additive
`Document::story_revisions` returns owned `StoryRevision` snapshots for every
story that revision resolution reaches: the main document, headers, footers,
comments, normal footnotes, endnotes, and the text boxes inside them. Each
snapshot adds the `StoryId` of its Word story, with table cells folded into the
story that holds the table. It scans the parts as resolution stages them, so
the list has one entry per resolved revision. A compared comment package
change uses one private package revision, reported under the main story,
because its added or removed comment may have no surviving comment owner.
The list length equals the count from `accept_all` and `reject_all`.
`rdocx-cli revision list` and Python `Document.revisions` expose this all-story
listing. WASM load and save paths preserve the revision XML without a revision
inspection method.

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
methods are additive on the pre-1.0 native Rust facade. Python `Run` binds
`add_tab()` and `add_field(instruction, cached_result="")`. Both append inside
the run through the same accepted run path as the text setter, so no run index
moves. A tab keeps live handles valid. A field becomes a story item of its
own, which moves the index path of every later `StoryItem`, so `add_field`
advances the revision once. An instruction without a field name raises
`RdocxError` and leaves the run and the revision unchanged. The field
serializes as a simple field after the run's earlier content, and
`update_layout_backed_fields` fills a `PAGE`, `NUMPAGES`, or `PAGEREF` cache.
WASM and CLI gain no implicit surface.

Native Rust also exposes `Run::add_field_value(Field)`. It accepts a checked
`Field::from_raw` or `Field::from_instruction` with explicit `FieldForm::Simple`
or `FieldForm::Complex` and ordered `Vec<CT_R>` cached content. The shared
`FieldInstruction::new` constructor validates names and switches, quotes text
operands and accepts recursive typed operands. A simple instruction cannot hold
nested fields. Known flag switches reject explicit operands. Otherwise unknown
quoted or nested switch operands retain their operand position across reopen.
This is a pre-1.0 semantic projection extension for ambiguous unknown switches,
while original producer bytes remain unchanged. `Field::form`, `locked` and
`set_locked` expose representation and three-state field locks, and immutable
`FieldRef::locked` reads the lock. Existing conservative unmodeled semantic
attribute reporting still identifies producer lock attributes. These additions
are native Rust only. Existing `Run::add_field` retains its simple-field API,
including its default PAGE and NUMPAGES cache behavior. Ordered cache controls
and properties survive attachment and reopen. Invalid attachment leaves the
run unchanged.

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
Each method returns the number of revisions resolved, including a compared
comment package change. Shared
ids select every matching placement, author matching is case-sensitive, and
missing dates do not match a date range. Invalid bounds and malformed selected
changes return an error before mutation. Resolution covers the main document,
headers, footers, comments, normal footnotes, endnotes, and nested text boxes.
`Document::revisions` remains main-story-only, while
`Document::story_revisions` lists exactly the revisions these methods resolve.
These eight methods are additive on `rdocx::Document`.
`rdocx-cli revision accept|reject` exposes the all-story resolution boundary
with mutually exclusive id, exact-author, or paired date selectors. An omitted
selector resolves all modeled revisions. Python and WASM continue to preserve
the resulting document when they save it.

Native callers generate tracked changes with `Document::compare`, supplying an
edited document, author, and RFC 3339 timestamp. The additive
`ComparisonDiagnostic` value reports stable locations and messages for
differences that stay out of the revisions, and the redline keeps the
original for each. A message starts with a stable prefix,
`formatting differs` for unsupported formatting and
`content-control <name> differs` for a content control's metadata, where
`<name>` is `tag`, `alias`, `lock`, `placeholder`, or `docPartGallery`.
Comparison rejects existing modeled revisions and unsupported structural shell
differences, a content control's type or data binding included, and it commits
only after accepting and rejecting staged copies reproduce their respective
package-wide modeled baselines. Those baselines read each paragraph as one
sequence in document order: compared units, preserved raw children such as
bookmarks, comment range markers, inline content controls, and hyperlink edges.
A granular revision that splits a run therefore still reproduces a bookmark or
comment range beside the edit, while a result that moves any of them relative
to the text or to each other is refused. Ignored whitespace, fields, and comment
references stay in the sequence where they touch one of those markers, so
moving a marker across them is refused too, and `ignore_comments` leaves
comment range markers out. `Document::compare` keeps its
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
When edited image bytes replace an existing image relationship payload,
comparison carries the edited bytes in a distinct media part and owner
relationship. The drawing run receives a tracked replacement. Acceptance
resolves the edited relationship and bytes, while rejection retains the
original relationship and bytes. Unrelated package parts remain unchanged.
When a matched paragraph changes a comment range, hyperlink, inline control,
or preserved child boundary, comparison tracks a complete paragraph deletion
and insertion if its bookmarks remain in place. The rebuilt TOC entry
transition to a hyperlink and PAGEREF field uses this path. Granular matching
cuts at bookmark and comment range boundaries so inserted text stays on its
edited side of a marker.
When comment owners or metadata change, comparison carries the edited comments
and their related package graph in the redline. A related private custom XML
part retains the original comment graph for rejection. Acceptance keeps the
edited graph, rejection restores the original graph, and either resolution
removes the private part. The package change appears as one selectable revision
in `story_revisions` and in Python and CLI revision listings. Compatible
comment text edits continue as ordinary comment-story revisions. This changes
no public Rust, Python, or CLI signature and adds no semver break.
Detached inline and anchor wrappers retain only the inherited namespace
bindings they use and that are not already carried by the story root. Dirty
typed inputs recover matching package drawing payloads before serialization.
Complex fields map every physical source run to one modeled comparison owner,
and sibling fields from one physical run share that owner. Text read out of a
field's physical run is compared as its own runs, and the comparison source
writes that span as one physical run per modeled run.
It emits same-story moves and supported run, paragraph, table, and section
property revisions. Diagnostic locations retain the actual story identity and
stable owner path. `rdocx-cli compare` takes an explicit author, RFC 3339
timestamp, and output, and exposes every `ComparisonOptions` field as a flag.
Its `--granularity` defaults to the source-compatible whole-run `run`, like
the native default, and `word` or `character` marks only the changed words or
characters. `--ignore-story` is
repeatable and takes the Python `Story.kind` names, where `body` selects the
main story. An unknown granularity or story name is a usage error, and a
duplicated story keeps the native rejection. Python
`Document.compare` takes the `ComparisonOptions` fields as keyword-only
arguments. `granularity` is `"run"`, `"word"`, or `"character"` and defaults to
the native `"run"`. `ignore_formatting`, `ignore_whitespace`, `ignore_fields`,
and `ignore_comments` default to false. `ignored_stories` takes `Story.kind`
names, where `body` selects the main story and `table_cell` is not a comparison
category. An unknown granularity or story name raises `RdocxError` before the
document changes, and a duplicated story keeps the native rejection.
`ignore_comments` also leaves comment anchors to the original, while ignoring
the `comment` story excludes only the comments part. Python and WASM preserve
comparison output when they save their owned document.

Native Word rendering exposes `rdocx::RevisionView` and the concrete
`rdocx::RenderOptions`, whose default selects the accepted view. Additive
option-taking counterparts cover PDF bytes and files, single-page and all-page
raster output, page layout, deterministic rendering, and caller-supplied font
paths. The existing methods keep their accepted default. Python `to_pdf`,
`render_page_to_png`, `render_all_pages`, and `render_pages` take a keyword-only
`revision_view` of `"accepted"` or `"tracked"`, with `"accepted"` as the
default. Any other value raises `ValueError`. CLI `convert` for PDF and image
formats and `render` take `--revision-view accepted` or
`--revision-view tracked`, with `accepted` as the default. An unknown value is
a usage error.
HTML and Markdown conversion refuse tracked view before creating output. WASM
retains its existing rendering behavior.
Native selected-image rendering adds zero-based page-list entry points that
share `rdocx::RasterFormat`, `rdocx::RasterOptions` and
`rdocx::RasterOutput` with `oxml-pdf`. The existing PNG methods remain
source-compatible opaque defaults. Python exposes the same image controls as
keyword-only `render_pages` arguments, keeps zero-based page indices, releases
the GIL for rendering, returns `list[bytes]` for PNG or JPEG, and returns one
`bytes` value for TIFF.

Python `Document.to_pdf(*, fonts=None, font_dir=None, revision_view="accepted")`
uses `Document::to_pdf_with_options` when no fonts are supplied. With `fonts`,
a sequence of `(family, bytes)` pairs, or `font_dir`, a directory whose `.ttf`,
`.otf`, and `.ttc` files
`Document::load_fonts_from_dir` labels by file name, it calls
`Document::to_pdf_with_fonts_and_options` with the given fonts first, as
`rdocx convert --font-dir` does. That call lays out with the caller fonts only,
so a family they do not provide, even through the automatic label aliases and
metric-compatible names, raises `LayoutError`. The native loader reads a
missing directory as an empty one, so the binding raises `FileNotFoundError`
for a missing `font_dir` and `NotADirectoryError` for a file, before any
layout. SVG and raster output take no caller fonts.

Native Word SVG adds `SvgDiagnostic`, `SvgRenderResult`, and four additive
`Document` methods. `render_page_to_svg` and
`render_page_to_svg_with_options` reuse normal layout. Their deterministic
counterparts reuse bundled-font-only layout. Every method takes a zero-based
page index and returns `None` beyond the laid-out document. The result contains
self-contained searchable SVG plus layout-first, path-specific lowering
diagnostics. Python `Document.render_page_to_svg(page_index)` binds the normal
layout method, releases the GIL, and returns a frozen `SvgRenderResult` with the
SVG text and a tuple of frozen `SvgDiagnostic` values, or `None` beyond the
last page. WASM, CLI, Presentation, and public `oxml-pdf` APIs do not gain SVG
methods or values.

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
`add_comment`, `reply_to_comment`, `resolve_comment`, `remove_comment`,
`move_comment`, `move_reply`, `sections`, `set_sections`,
`notes_header_footer_mut`, and `handout_header_footer_mut`. Python exposes the
comment snapshots, additions, moves, resolution, and removal described with the
presentation binding, and `rpptx comment` exposes the comment operations
described under CLIs. WASM consumers gain no collaboration or navigation
methods and continue to preserve these package parts through their existing
`Presentation` owner.

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

## Native PowerPoint text layout

The published pre-1.0 `rpptx` facade exposes the concrete `TextFrameLayout`
value and one render-feature method:

```rust
Presentation::text_layout_deterministic(&self, width_factor: f64)
    -> Result<Vec<TextFrameLayout>>;
```

`TextFrameLayout` carries the zero-based slide index, the optional shape id and
name, the effective `AutofitMode`, and an `rpptx_render::ShapeTextLayout`. The
published pre-1.0 `rpptx-render` crate adds the concrete `ShapeTextLayout` and
`TextLineLayout` values and `layout_shape_text`, which shares its stacking path
with slide lowering. `08-rendering-spec.md` owns the coordinate, overflow, and
width factor semantics.

Python gains `Presentation.text_layout` with frozen `BoundingBox`,
`TextFrameLayout`, and `TextLineLayout` snapshots, described under the Python
API shape above. WASM and CLI consumers gain no text layout method. This is an
additive semver change for `rpptx` and `rpptx-render`. It adds no production
dependency, feature flag, trait, dynamic dispatch, generic parameter, or
builder.

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

Current source prepares the exact seven-package Word Rust family, `rdocx`
Python distribution, `rdocx-wasm` and inherited support carriers at0.16.0.
The exact fifteen-package shared OOXML and PowerPoint family plus `rpptx-py`
and its Python distribution are prepared at0.14.0. The separate unpublished
`rpptx-wasm` npm carrier remains0.12.1. Word's current dependency pins require
shared0.14.0, so the shared family precedes Word publication. Preparation is
local metadata and artifact verification, without registry publication or tag
authority. The previous unified v0.15.0 and rpptx-v0.13.1 releases remain
immutable at reviewed main9d019472f7e6b95ac4ba0770c4dee35dcbc28f0e.

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

The historical S73 Rust package trains remain separate. The exact 15-package
shared OOXML and PowerPoint workspace family was published at 0.12.1 from immutable annotated tag
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

`wheels.yml` handles the unified `v*` Word and `rpptx-v*` shared/PowerPoint
namespaces. A tag selects one exact Rust allowlist, CLI and Python distribution.
Historical `py-*` tags remain immutable and cannot start new publication.
Manual dispatch builds both Python projects across six declared targets plus
one source distribution each. It cannot publish to registries or create a
GitHub release and does not build the tag-only CLI archives.

Each freshly built native wheel runs clean installed priority runtime checks.
The supported floor remains Python3.9. Exact `mypy==2.3.0 --strict` and
recursive package `stubtest` run under Python3.12, including the native
extension and handwritten stubs. The current local preparation builds real
wheels and source distributions and checks their complete metadata and README
payloads. Compatible oracle exclusions on hosted runners retain their pinned
native and CI coverage. The musllinux cells use a clean Python3.9 Alpine
runtime. Build jobs have read-only repository permissions, while selected
tag-only publication uses OIDC through the `pypi` environment.

After final integrated full verification, clean sprint review and sprint push,
the SHA-bound build-only rehearsal must produce twelve `cp39-abi3` wheels and
two source distributions at that reviewed source. Downloaded provenance,
metadata and clean Python3.9/3.12 installed checks are required before close.
Local worker artifacts do not substitute for this final hosted evidence.

After sprint close, each family requires its own immediate approval at the
reviewed main merge SHA. The tagged workflow verifies six CLI archives, six
wheels, one source distribution and a complete `SHA256SUMS`, then publishes
only the selected family. Registry state, ownership, attestation, exact release
body and release-bound contribution notifications are verified afterward.
Neither local preparation nor build-only dispatch earns publication or hosted
artifact availability. Historical separately tagged Python releases
`rdocx0.13.2` and `rpptx0.11.0` at2b009243ed39ab66470d7484d490985368e865a8
remain available, and the later unified family releases remain immutable.

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
native Rust APIs. Python and CLI path saves write the class that the output
extension names, while byte saves and WASM preserve the source main content
type, and none of them gains a package-class selector in this milestone.
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
`text`, `convert`, `diff`, `replace`, `validate`, `render`, `thumbnail`,
`outline`, and `comment`. It uses clap derive and `serde_json` for `--json`.

`inspect` reports the file, slide and layout counts, slide size, core metadata,
and each slide's identity, hidden state, and shape count. Its JSON form uses the
shared schema-1 envelope. Beside each slide's shape count, `shape_details`
lists every immediate shape in z-order with its index, non-visual id and name,
kind, placeholder type and index, direct position and size in EMU, rotation in
degrees, direct autofit mode, paragraphs, table size, and children. A
placeholder that inherits its transform reports null geometry, and a
placeholder without an explicit type reports a null type. `text` emits slide
text in presentation order. `text --json` emits schema-1 slides with a one-based
slide number, the slide id, paragraphs, and speaker notes, which are null when
the slide has no notes part. Each paragraph carries a typed zero-based path of
shape, table row, table cell, and paragraph positions, the owning shape id, its
level, its visible text, and its regular runs. Run indexes match
`TextParagraphRef::run`, so fields and line breaks appear only in the paragraph
text, where a line break is U+000B. Run formatting is null without direct run
properties. Otherwise it contains nullable direct bold, italic, underline token,
Latin font, point size, and sRGB colour fields.
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
omitted, and embedded paragraph breaks become spaces. `outline --json` reports
the same title, or null for an untitled slide, the same items with their
levels, and the speaker notes. Notes are the plain text of the notes body, with
paragraphs and line breaks both written as newlines. `text --notes` and
`outline --notes` print one `Notes:` line for each non-empty notes line after
the slide's plain output. JSON output always carries the notes.

The structured output reads public facade values only. `ShapeRef` gains
`rotation` and `placeholder_type`, and `PhType::as_str` and
`TextUnderline::as_str` become public. These are additive changes to the
pre-1.0 `rpptx`, `rpptx-oxml`, and `oxml-drawing` crates.

`comment` lists, adds, replies to, resolves, and removes modern PowerPoint
comments. Legacy comment parts stay preserved and unlisted. `list` shows a
comment or reply without a status, or with the `active` status, as open, and a
`resolved` or `closed` one as such. Its JSON carries the `resolved` flag
beside the raw `status` token. `add` takes a one-based `--slide`. `add` and
`reply` require an RFC 3339 `--date`, reuse the first author with the given
name, and otherwise add an author whose `userId` is that name and whose
`providerId` is `None`. New author, comment, and reply ids are the first
unused sequential GUIDs, so the output depends on neither a clock nor a random
source. `resolve` accepts only a thread id. `remove` accepts a thread id,
which removes its replies, or a reply id. Every mutation requires an explicit
output, refuses an existing one, and publishes through the shared staged
output set. Its schema-1 record states the action, comment id, one-based
slide, and output path. The commands use the additive
`Presentation::resolve_comment` and `Presentation::remove_comment` facade
methods, which rest on the new `Comment::remove_reply`.

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
document order through the facade plain-text representation. It gives each
paragraph the same accepted-view text as `text --json`, but it leaves out
paragraphs inside block-level and cell-level content controls and inside
nested tables, which `text --json` reports. `text --json`
emits schema-1 accepted-view paragraphs with a zero-based direct body index,
typed zero-based nested path, direct style and numbering, text, and ordered
runs. Run formatting is null when no direct run properties exist. Otherwise it
contains nullable direct bold, italic, strike, underline, font, point size,
colour, highlight, language, and character style fields. Both views then read
every other story through `Document::story_item_snapshots`, leaving out the
main body and its table cells, which they already cover. Only the direct
paragraphs and block content controls of a story are read, so the text of an
inline control or a field is not repeated after its paragraph. Plain `text`
prints each package part of those stories under a line such as
`--- header (/word/header1.xml) ---`, one item text per line, and skips a part
without any text. `text --json` states `all-supported-stories` scope and adds
a `stories` array with the Python `Story.kind` name, part name, owner index,
and items of each story. Each item carries its story `index_path`, its kind,
and its accepted-view text.
`convert` to Markdown or HTML appends the same parts after the body under a
bold label, and leaves comments out as review annotations. A story part that
cannot be read leaves the body of these views intact. They print one stderr
warning that names the part when it can, list no other story, keep a zero exit
status, and `text --json` states `main` scope. `validate` checks
that every XML part the main document relates to is well formed, and that the
paragraph, character, and table style ids named in the main document, headers,
footers, notes, and comments are defined. Both are errors. A style id inside a
tracked property change or inside `mc:Fallback`, and an empty id, are not
checked. Parts are scanned before the document
opens, so a malformed styles, numbering, or settings part is named in the
report beside the open failure it causes. `layout --json` uses
bundled deterministic fonts and reports every direct body item. Its point-space
fragments carry one-based physical and displayed page numbers, and preserved
unlaid items retain an empty fragment list. `diff` compares accepted-view
paragraph text story by story: the body paragraphs, the body's table cells
through the `text --json` traversal, and every other story through
`Document::story_item_snapshots`, reading only the direct paragraphs and block
content controls that `StoryItemSnapshot::is_direct_child` marks. Headers and
footers pair by type and the first section that references them, and every
other part pairs by kind. Each story is one sequence matched by a Myers
linear-space shortest edit script, in O((N+M)D) time and O(N+M) memory, and
removed and added paragraphs between two matches pair in order as changed
paragraphs. The text form keeps the `[i]` body lines and adds
bracketed story locations, a story that cannot be read is reported as not
compared, `--json` emits a schema-1 record with each unreadable side in
`not_compared`, and `--exit-code` exits with 1 for a difference and 2 for an
error or incomplete comparison. `replace --expect N` checks the
run-aware replacement count before staged publication. A mismatch creates no
output and leaves an existing destination untouched. Both the selected page
and all-page `render` paths use bundled deterministic fonts. The compiled
surface also includes nested comment thread commands with optional RFC 3339
comment dates, all-story revision inspection, all-story filtered revision
resolution, comparison with explicit granularity, ignore options and per-story
revision counts, and TOC rebuild. Every new mutation requires an explicit output and publishes through
the shared staged output set. Their schema-1 records state `main` or
`all-supported-stories` scope, and the comparison record also states the
options that ran. Revision selectors are mutually exclusive, and
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

Generated-table rebuild errors include checked target mutations that would
require unsafe canonical serialization beneath conflicting retained namespace
bindings. The error leaves the package and existing caches unchanged. Namespace
aliases are accepted when the source projection and mutation are safe.
