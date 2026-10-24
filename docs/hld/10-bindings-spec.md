# 10, Bindings spec

Owners: `oxml-py-support`, `rdocx-py`, `rpptx-py`, `rdocx-wasm`, `rpptx-wasm`,
`oxml-cli-support`, `rdocx-cli`, `rpptx-cli`.

## Native numbered outline

The additive Rust `Document::numbered_document_outline(&self) ->
Result<Vec<NumberedOutlineNode>>` retains accepted heading level, title-only
text, optional visible marker, display text, physical WordSourcePath and child
nodes. Display text separates a nonempty marker and title with one readable
space. Existing `headings` and `document_outline` retain their signatures and
legacy title-only scope. No new Python or WASM navigation surface is implied.

Doc-hidden layout hooks expose private production paragraph facts and the
validated generated-result exclusions used by the facade. They add no public
struct fields. The iterator `paragraph_navigation()` and direct
`paragraph_navigation_source(SourceNodeId)` read the same private facts. The
latter returns `Option<(&WordSourcePath, Option<u32>, &str, Option<&str>,
Option<&ResolvedNumbering>, Option<usize>)>`. Neither creates a second heading
tree or counter. The source stamp binds exclusions to the current parsed main
input and physical part, while every sequence counter is regenerated.

Ambiguous generated-owner discovery travels as a private source-stamped failure,
not an empty successful navigation inventory. The narrow doc-hidden
`WordSequenceSnapshot::set_navigation_failure(&mut self, &LayoutInput, String)`
uses existing types. `WordLayoutResult::navigation_failure() -> Option<&str>`
lets the new outline refuse without making ordinary cached painting fatal.
Ordinary REF/SEQ counter semantics remain unchanged.

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
// ContentPath retains public path/revision fields and private captured authority.
// RevisionCounter owns prepared scope generations and first retirement causes.
// Use existing constructors or counter capture methods, not struct literals.

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

Captured paths carry owner authority and the first causal retirement, not
only a numeric revision. Word captures a finite prepared inventory of physical
paragraph, table, cell, section and hyperlink owners. Regular runs retain
conservative document-wide structural invalidation, and frozen Story and
StoryItem captures retain snapshot-inventory authority. Presentation scopes
bind physical slides, shapes, grids and text
owners. Unrelated edits and supported owner relocation retain handles, while a
removed and later reused numeric shape id cannot resurrect an old handle.
Stale errors name the actual public operation and recovery path. Foreign
presentation identity is checked before stale lookup or mutation.

Native staged edits finish validation, checked scope preparation and every
returned Python object allocation before committing and publishing the
candidate. `capture_prepared` binds returned handles to the successor state.
Run tab and break edits prepare the changed paragraph read fact before
publishing native content and keep existing handle scopes live.
Failures and actual no-ops publish nothing. Manual unbound native snapshots
cannot acquire authority from a matching numeric position. Private authority
changes pre-1.0 struct-literal source compatibility, while existing constructors
remain available. No stable-id-on-disk or generic handle registry is implied.

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
an intentional pre-1.0 struct-literal break. Python roots expose concrete
`CategoryChartData`, `XyChartData` and seven-kind `XL_CHART_TYPE` values. Word
`add_chart(chart_type, chart_data, width, height, after=None)` returns a prepared
owned Paragraph and accepts only a checked same-document direct body owner.
PowerPoint shapes add charts through the existing owned frame handle. Its live
Chart reads category and series snapshots. Its chart_type returns int or None
from qualified source metadata without requiring dense caches or a workbook.
The four qualified imported subtypes return 58, 52, 4 and 69. Unsupported valid
types return None, while malformed source and invalid owners raise errors.
Type reads preserve bytes and live authority. Chart replaces data through the native
unique-ID graph transaction, including nested groups. Returned allocations
precede publication. Invalid or retired owners, invalid data and exact no-op
leave bytes and captured authority unchanged. The Python data objects do not
add WASM or CLI chart entrypoints.

Word `chart_legend(paragraph)` returns a dictionary with `visible`, `position`
and `overlay`. `set_chart_legend(paragraph, *, visible=None, position=None,
overlay=None)` edits a live same-document direct-body Paragraph with exactly
one supported qualified chart drawing. Opaque ancestry cannot grant authority.
PowerPoint's existing live Chart exposes writable `has_legend`,
`legend_position` and `legend_overlay`, including grouped chart owners. Positions
are `r`, `l`, `t`, `b` and `tr`. Boolean inputs require actual bool values.
Missing settings read as right and false. Hiding discards the entire legend
subtree and settings. Position or overlay alone never makes a hidden legend
visible or stores hidden settings. Showing creates default or explicitly
combined settings. Repeated effective values preserve exact bytes and held
owner authority. These additive APIs reuse native `LegendPosition` and
`LegendSettings` re-exports without new WASM or CLI surfaces.

Legend mutation prepares binding publication before committing the native
candidate. Shared chart parts are isolated without copying or rewriting the
workbook. Only the selected qualified Word relationship attribute is retargeted,
including when an opaque marker shares its old relationship. Unknown, stale,
foreign, related-story and ambiguous Word owners refuse atomically.

A Chart view captures its frame XML-child authority. Checked Shape XML
replacement retires old Chart views while retaining the proven Shape owner.
Fresh chart lookup resolves the current owned graph. Data-only replacement
retains the chart and frame, isolating shared chart/workbook targets through
the native copy-on-write transaction.

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
doc.add_picture("img.png", width=Inches(2), height=Inches(1))
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

Word paragraph creation accepts empty text and an optional resolved paragraph
style. `add_heading` resolves Title or Heading1 through Heading9, and
`add_page_break` stages its paragraph and page-break run together. Table style
creation follows the existing checked table-style resolver. Validation and
returned-owner allocation finish on a candidate before publication. Missing
list styles name explicit numbering and style providers without changing the
default template. Run bold, italic and underline aliases share `run.font`'s
owner validation and tri-state behavior.

Document construction/open accept binary seekable streams as well as paths.
Save serializes fully before writing to a binary stream, detects short writes
and preserves caller stream ownership. A final case-insensitive .pdf path
extension raises ValueError naming to_pdf before any write. Existing document
variants, other path extensions and binary streams retain their routes. A
stream's display name does not turn it into a filesystem path. I/O errors
propagate. Picture streams
reuse native byte admission. Supplying only one picture dimension refuses with
the explicit width-and-height alternative. Direct-body paragraph insertion
resolves the physical body index. Hyperlink anchor=Paragraph refuses with
explicit add_bookmark/add_internal_hyperlink guidance. Cell and related-story insertion refuse with
the supported story or document authoring path.

`styles` is an immutable tuple subtype containing frozen Style snapshots.
Integer indexing, slicing, iteration and equality retain tuple behavior, with
slices returning ordinary tuples. String indexing raises a first-line
`set_style` pointer. Exact builtin tuple type identity is outside this contract.
Frozen Section left_margin read and assignment, and header/footer/first-page properties name the live
`document.section(snapshot.ordinal)` alternative. left_margin names the matching
live left_margin property, while frozen margin_left remains the stored field. `add_section` names native
insert/update operations without promising python-docx section cloning.
Table column compatibility names refuse with `insert_column`,
`set_column_width`, `grid_widths` and `cell.width` alternatives. Paragraph
`_element` and Presentation `_sldIdLst` name `remove_content` and
`slides.remove` respectively, without exposing lxml mutation. Paragraph._p,
Run._r and Run._element validate current ownership before refusing with their
xml and replace_xml(xml) alternatives. Retired owners retain stale errors.

The `rdocx.enum.text` module defines bounded `WD_BREAK_TYPE`/`WD_BREAK` and
`WD_COLOR_INDEX`. The concrete `rdocx.enum.section` module defines
`WD_SECTION_START`/`WD_SECTION` and `WD_ORIENTATION`/`WD_ORIENT` with pinned
python-docx numeric values and identity-preserving aliases. Its inline annotations
provide the same direct import path at runtime and under strict typing. These
constants do not add XML conversion methods or section-cloning semantics. The
table module's `WD_ALIGN_VERTICAL` aliases `WD_CELL_VERTICAL_ALIGNMENT` by identity.
`Font.highlight_color` maps the supported color-index values to native Word
highlight names and clears with None. Other break forms retain named
line/page/column alternatives.

## Python Word formatting and checked list publication

Numbering restart admission and decimal continuation use the concrete instance's
replacement level before its abstract definition. An effective bullet cannot be
restarted as a number or reused by the decimal helper. An effective decimal over
an abstract bullet remains eligible. The existing abstract-level reader projection
retains its separate inspection contract.

ListLevel accepts optional nonempty XML-valid font text, carried through native
marker properties to ascii and hAnsi font slots. Numbered-item restart=True
creates a fresh instance starting at1. Later compatible decimal items continue
the most recent physical explicit level restart, including after reopen, while
unrelated plain or bullet items do not reset it. Without an active explicit
restart, historical default selection remains. restart_numbering accepts a live
same-document direct-body numbered non-bullet paragraph, retains its level and
other overrides, and returns the new instance id. Its non-bool start accepts the
full unsigned32 domain. Unprojectable imported override source refuses atomically.


Native `Run::set_width_scale_value` accepts `Option<u32>` percent and
`Run::set_position_value` accepts `Option<i32>` signed half-points. Some writes
an explicit direct property, including zero. None removes only that property
and leaves an absent `w:rPr` absent. Existing width-scale and position setters
and builders retain their signatures and delegate to these nullable forms.
The Python Font surface retains its existing properties.

`Font` exposes optional superscript, subscript, small_caps, all_caps,
double_strike, hidden and rtl. Superscript and subscript share one direct
`w:vertAlign` source. True selects that mode, false clears only its own active
mode, and None removes the direct mode. An absent mode reads None, while a
direct baseline reads false for both. Clears on absent source insert no nodes.
Language, language_east_asia, language_bidi, name_east_asia and
name_complex_script edit their selected slots. Authored language accepts a
nonempty ASCII alphanumeric sequence of subtags separated by single hyphens,
beginning with a letter. Case is preserved and None clears the property. This
is a finite lexical rule, without registry or complete BCP47 validation.
Imported values and the separate East Asia and bidi setters keep their existing
contracts. General font name retains its
established four-slot assignment. `character_spacing` accepts signed integer
EMU, explicit Length or None and returns Length or None. `space_after` retains
nonnegative EMU, Length or None. Both reject bool and unsupported types before
publication. A nonzero bare integer that truncates to zero twips refuses with
Pt, Twips or Emu guidance. An explicit Length retains intentional truncation,
and explicit zero remains valid. The full signed32 twip quotient is checked
before narrowing. None clears only the selected property.

`ParagraphFormat.line_spacing` distinguishes genuine Length absolute spacing
from non-bool integer or float line multiples. Integer2 and float2.0 both
write480with auto rule. Pt(18) writes360with exact rule. None clears both direct
value and rule. Finite signed values and zero retain the native domain, with
signed32 range checked on the truncating EMU/635 quotient for absolute lengths
or the finite multiple-times240 product for multiples. Unsupported types,
nonfinite products and saturation refuse before mutation. Native off-grid
multiple truncation remains distinct from python-docx's twip rounding.

`ParagraphFormat.tab_stops` is a live parent-bound collection with length,
indexing, iteration, indexed deletion, add_tab_stop and clear_all. Returned
TabStop values are immutable position/alignment/leader snapshots. Alignment
assignment refuses with the delete-then-readd collection route, preserving
position and leader. Delete the occupied entry before add_tab_stop. New Python
positions use the existing signed EMU conversion, truncating toward zero at
635 EMU per twip with signed32-bit bounds. A repeated stored position raises
ValueError before publication. Distinct positions insert before the first
greater position without sorting imported occurrences or rejecting imported
duplicates. The native equal-position insertion contract is unchanged.
Alignment includes List, Start and End.
An absent leader reads SPACES, while an unsupported owned token raises an
explicit error and remains preserved. Collection edits keep the paragraph
handle live. Paragraph formatting also exposes border read/set/clear, shading,
outline_level and bidi, with checked input and retained foreign payload.

`Document.default_font_name` and `default_font_size` edit docDefaults without
restyling Normal. Size uses Length or integer EMU and returns Length or None.
`add_style` and `set_style` accept alignment, line_spacing, underline,
keep_with_next, keep_together, page_break_before, borders, shading and tab_stops
in addition to their existing arguments. Omitted values retain source, and
explicit empty border/tab collections clear active owned entries. The selected
Word border or tab container can remain empty to suppress earlier duplicate
owners while retaining foreign attributes and opaque payload. Border clearing
leaves no active owned edge, and tab clearing leaves no active owned tab.
These retained containers do not represent active formatting. Invalid input or
style graph changes fail before source publication.

`add_numbering_instance(definition_id, *, start=None, level=0)` creates an
instance of one existing definition. Start is a checked u32, including zero,
and level is 0 through 8. A nonzero level without start is rejected.
`add_bullet_list_item(text, level=0)` and
`add_numbered_list_item(text, level=0)` use checked native factories and return
a fresh Paragraph. `ListLevel.checklist(checked=False)` creates a real bullet
level using the unchecked or checked glyph, not an ordinary text prefix.
Structural factories stage the native candidate, validate state, allocate the
returned Python handle against prepared owner state, then commit and publish.
Failure, allocation refusal or no-op leaves source and revision state intact.
Earlier handles retire under the existing causal document owner policy, while
ordinary formatting edits retain live parents and fonts.

The low-level `CT_BorderEdge.source_xml` field retains the original physical
edge payload used by source-aware paragraph and style merges. Full public
struct literals must supply it or use default update syntax. This is an
intentional pre-1.0 Rust source break. List/Start/End tab enum variants also
require exhaustive matches to cover their exact tokens. Start and End retain
the existing left/right rendering mapping. List retains its established
fallback and has no additional paint-parity claim.

## Python colour inputs and live font formats

Python numeric setters reject bool before conversion. Word font.size accepts
its existing EMU/Length representation and refuses nonzero values whose
unchanged native conversion stores zero half-points. The error names EMU,
0.5 point storage and rdocx.shared.Pt. Explicit zero and None retain their
existing behavior. Native constructors and imported zero source are unchanged.

Presentation before/after spacing rejects nonzero EMU values that truncate to
zero centipoints, while explicit zero, None and existing float-line roles
remain valid. Character spacing applies the same refusal symmetrically to
small signed values, preserving meaningful negative spacing. Line spacing
must produce positive stored points or percentage units, with positive line
multiples and rpptx.util.Pt named by its refusal. Role-specific positivity does
not change valid zero before/after spacing. Every refusal precedes owner or
source publication, including warmed wrappers and pre-existing nonzero values.


Both families accept `RGBColor`, legacy integer channel triples and exactly
six ASCII hexadecimal digits with an optional single leading hash at every
existing scalar RGB setter. Hex values normalize to uppercase. The inventory
covers Word style creation/update, run fonts, run and cell shading, and table
and cell borders. PowerPoint covers run and paragraph-default fonts, including
table cells, plus supported fills, lines, cell borders and outer shadows.
Named Word highlight values retain their separate vocabulary.

One shared concrete extractor validates before mutable source access. Wrong
types raise `TypeError` naming the full parameter, accepted forms and actual
Python type. Malformed, non-ASCII or surrogate-bearing hex strings and channel
range or overflow failures raise parameter-qualified `ValueError`. Arbitrary
string coercion is not accepted. Failed assignments preserve source and
revisions. `RGBColor.from_string` accepts the same optional hash syntax, and
`str(RGBColor(...))` remains uppercase hexadecimal.

`Font.color` always returns a live `ColorFormat`, including when no direct
colour exists. It re-resolves the concrete run owner and follows structural
and cell-text invalidation. Getter reads create no properties or revisions.
Direct scalar reads use `font.color.rgb`. The getter deliberately replaces
the earlier Word `RGBColor | None` and PowerPoint `str | None` return types.
Direct `font.color = value` remains a shortcut, and `None` removes its direct
colour. The shared signatures and handwritten stubs are:

```python
Font.color -> ColorFormat
Font.color = RGBColor | tuple[int, int, int] | str | None
ColorFormat.rgb -> RGBColor | None
ColorFormat.rgb = RGBColor | tuple[int, int, int] | str | None
ColorFormat.type -> MSO_COLOR_TYPE | None
ColorFormat.theme_color -> str | None
ColorFormat.theme_color = str | MSO_THEME_COLOR | None
```

Word ColorFormat.type returns the concrete rdocx.enum.dml.MSO_COLOR_TYPE values
RGB=1, THEME=2 and AUTO=101, or None. This replaces the former rgb/theme/auto
strings, so callers migrate comparisons to enum members despite patch release
numbering. Theme source takes precedence over cached RGB or auto. Reads retain
source tint, shade and opaque metadata. Word theme_color remains str or None.

PowerPoint uses its own enum-valued getters: ColorFormat.type returns
MSO_COLOR_TYPE or None, and
theme_color returns MSO_THEME_COLOR. A direct nonscheme color returns
NOT_THEME_COLOR, an absent direct color raises AttributeError, and an
unmapped scheme token refuses with preserved-source inspection guidance.
Existing string setters remain supported. The concrete PowerPoint contract
and finite language mapping are defined under Native PowerPoint notes and
handout export.

`None` in these colour-format setter signatures clears only font sources.
Existing fill, line and shadow colour setters reject it. Word `type` reports
its RGB, THEME or AUTO enum member. PowerPoint models RGB, SCHEME, PRESET and SYSTEM members. Unset or unmodeled direct colours report `None`. Scheme references
are not flattened to guessed RGB. Word `rgb` can expose an existing cached
literal beside its theme reference, and reads `None` for `auto`. Editing a
Word theme reference keeps that cached literal, tint, shade and opaque source.
Clearing the reference removes tint and shade but keeps the cached literal.
An explicit literal assignment retains the existing native clearing policy.
The native Word tint/shade calculation remains unchanged.

Both top-level packages export `MSO_THEME_COLOR` from their existing enum
files. Supported values 1 through 16 cover dark/light, six accents, hyperlink,
followed hyperlink, text and background slots. Word tokens use `dark1`,
`light1`, `dark2`, `light2`, `hyperlink`, `followedHyperlink`, `text1`,
`background1`, `text2` and `background2`. PowerPoint uses `dk1`, `lt1`, `dk2`,
`lt2`, `hlink`, `folHlink`, `tx1`, `bg1`, `tx2` and `bg2`, and also accepts
`phClr`. Both use `accent1` through `accent6`. `NOT_THEME_COLOR` and `MIXED`
are exported sentinels that cannot author a theme reference.

Word shading snapshots return `RGBColor | "auto" | None`, and border snapshots
return `(style, size, RGBColor | "auto" | None)`. Their existing `auto`,
omission and clearing behavior remains surface-specific. Python examples,
runtime assertions and typing smoke use the live getter and these scalar
snapshot types.

The S33 formatting inventory is intentionally bounded to font name, size,
colour, bold, italic, underline and strike, plus paragraph alignment, spacing,
indentation, keep-with-next, keep-together, page-break-before and widow
control. Assigning `None` clears direct tri-state formatting. `Paragraph` also
exposes its style ID and its list numbering as a `(num_id, level)` pair, `Run`
its character `style_id`, and `Font` its named Word highlight and separate
shading fill. Highlight reads and writes `ST_HighlightColor` names through
`w:highlight`. Shading accepts the shared RGB forms or `auto` through
`w:shd`, and reads `RGBColor`, `auto` or `None`. These setters also clear with `None`. The S33 table
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

Word `Table.indent` is an optional Length with a setter accepting the existing
integer EMU/Length convention or None. Positive, zero and negative values are
explicit. The checked native set_indent_value validates signed-i32 twip bounds
before creating properties and truncates EMU toward zero at635 EMU per twip.
set_indent_checked delegates to that signed optional setter. None removes only
direct tblInd and clearing an absent value creates no tblPr. Overflow and wrong
Python types preserve source and handles. Existing alignment, cell margins,
compatibility compensation and floating-table placement remain unchanged.

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
`set_style`, `add_style` and `insert_document` keep unmodelled style-root children, including `w:latentStyles`, in schema order ahead of `w:style`. Empty style records and foreign attributes on those children round-trip through those edits.
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

`Document.remove_content_range(start, end)` returns how many direct body
children it removed. The range is half-open. An empty range returns 0. A
reversed or out-of-bounds range raises `RdocxError` and publishes nothing.
The call does not recurse into tables.

A plain direct body `Paragraph.replace_xml` resolves its captured generation
through a derived current-inventory index and prepares only that paragraph's
content, properties and exact run scopes. An empty plain paragraph may acquire
direct runs with prepared newborn authority and exact route buckets. Subsequent
equal-path edits resolve those appended owners through the same route index.
Other topology changes retain the staged path. A last direct plain paragraph
can be removed locally when the suffix contains only sections or appended runs
whose parents precede the removed owner. Their indices are shifted locally.
Non-contiguous descendants of the removed owner retain the staged path.
A shared removal epoch retires all held physical owner paths, while earlier
current inventory entries remain available to fresh handles. Removing another
body child keeps conservative address reconciliation. Generic paragraph XML
refuses adding, removing or changing physical section properties, including
self-closing forms. Unchanged imported section XML, including opaque foreign
attributes, is preserved. A temporary model failure on an admitted Empty
section retains exact protected raw bytes and namespace meaning. Retained
nested namespace owners require staged identity proof, except a plain paragraph
removal when every owner is a `w:sdt` whose declarations replay on the body.
When that replay fails, or a shadowed serializer prefix is retained, `insert_paragraph`, `Paragraph.insert_paragraph_before`,
`add_paragraph`, `add_heading`, `add_page_break`, `insert_table` and `add_table`
refuse before any change, with the error a save would give. A retained shadowed serializer prefix refuses that local publication
and leaves the package unchanged. PowerPoint `ShapeCollection.add_textbox`, `Shape.alt_text`, `Shape.left`,
`FillFormat.solid` on a shape, and a stable `Shape.element.replace_xml` do the
same for one slide child when the edit does not change relationships, connectors,
or owner identity, and the package has no retained signature. Each local publication validates that child and advances the
revision once. An edit that must read the rest of the package uses the existing
staged transaction.


`Document.add_picture` accepts in-memory bytes, a safe filename, optional
paired EMU dimensions, and an optional checked `StoryItem` placement. Omitted
dimensions use native 72 DPI sizing. Omitted placement appends to the body.
The returned immutable `StoryItem` names the inserted paragraph at the new
document revision. A rejected image, filename, dimension pair, or stale path
leaves package bytes and binding revisions unchanged.

`rpptx` mirrors python-pptx through an unpublished mixed-layout `rpptx-py`
crate. `Presentation` owns the Rust facade and one revision counter.
`Presentation(path)` opens a file and the static `Presentation.from_bytes`
opens in-memory package bytes, as the rdocx `Document` does. Lazy layouts,
slides, shapes, placeholders, text frames, paragraphs, runs, columns and cells
store only a presentation reference and `ContentPath`. The bounded
source-compatibility surface is the seven python-pptx 1.0.2 Getting Started
workflows. They change the import namespace. Captured collections and items
resolve current owners after supported insertion or relocation. Retirement
occurs at the affected structural or text owner, with an explicit stale error
when that owner is no longer valid. Pure-Python
`Length`, `Inches`, `Pt`, `RGBColor`, and the `MSO_SHAPE`, `MSO_SHAPE_TYPE`,
`MSO_CONNECTOR_TYPE`, and `MSO_FILL_TYPE` enumerations keep native inheritance
outside the limited ABI. `MSO_SHAPE` carries the 182 python-pptx
`MSO_AUTO_SHAPE_TYPE` members, each with its preset name as `xml_value`.
`UP_ARROW` is value 35 and its preset name is `upArrow`.

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
staged layout change and advances the revision once. A placeholder the new
layout does not place keeps the transform it inherited.
`Slide.hidden` reads and writes `p:sld/@show`. `Slide.background.fill` is a
live `FillFormat` over the direct background fill that never changes the slide
when read, and `follow_master_background` reports and sets whether the slide
has no `p:bg`. `SlideCollection.remove` and `SlideCollection.move(from_, to)`
use the native staged slide operations and advance the revision once.
`SlideCollection.import_slide(slide, layout=None, index=None, *,
keep_source_formatting=False)` returns one live imported Slide. Existing
positional arguments and the destination-formatting default remain supported.
`layout` must belong to the destination. Python insertion positions are
zero-based, count from the end when negative, and refuse out-of-range positions
with `IndexError`. In destination-formatting mode, a same-presentation import
reads a source snapshot and keeps its own layout unless an explicit layout is
given. Source-formatting mode uses
the checked source layout, master, theme and required resource closure.
Combining source formatting with an explicit destination layout is refused
before publication.
`import_slide_with_diagnostics` accepts the same arguments and returns
`(Slide, tuple[str, ...])`, exposing import diagnostics without changing the
ordinary single-slide return type.

`SlideCollection.import_slides(source, indexes=None, *, layout=None,
index=None, keep_source_formatting=False, keep_sections=False, slide_map=None)`
returns `(tuple[Slide, ...], tuple[str, ...])`. Omitted indexes select every
source slide in source order. Explicit zero-based indexes retain caller order
and reject duplicates and out-of-range selections before publication. Valid
negative source indexes count from the end, without modulo wrapping. An empty
selection changes nothing. Duplicate selection is a native operation refusal
reported as `RpptxError`. Out-of-range source indexes raise `IndexError` before
the native transfer. `keep_sections` preserves selected source section
blocks without reordering the selection. Explicit `slide_map` pairs identify
source-to-destination jump targets through checked existing slide indexes.
Both map endpoints accept valid negative indexes counted from their respective
source or pre-insertion destination collection ends, without modulo wrapping.
The destination layout and insertion rules match the single-slide method.
All returned Slide objects and diagnostics are allocated before the one
success-only publication. Foreign owners, bad mappings and refused native
transfers leave the destination and held-owner state unchanged.

The CLI command is `rpptx slide import DEST SOURCE [--slides 1,3-5] [--at N]
[--keep-source-formatting] -o OUT`. CLI source ranges and insertion positions
are one-based. Omitted selection imports all slides, omitted insertion appends,
and `--at 1` inserts before the first destination slide. Selection retains
caller order and rejects duplicates and out-of-range values. Structured or
human diagnostics use the existing CLI channel. Atomic output publication
never overwrites either input.
`SlideCollection.duplicate(slide)` copies a slide of the same presentation,
with its speaker notes, to the position right after it through the native
staged `duplicate_slide`, advances the revision once, and returns the new slide
captured at that revision. As in the facade, a slide that owns a modern
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
sRGB colour as `RGBColor`, or `None` for any other colour. RGB replacement
keeps transforms and opaque children of an existing sRGB colour. Theme
replacement keeps those of an existing scheme colour. Replacing another
colour kind discards that kind's transforms and raw children, except that an
outer shadow keeps its modeled alpha. Brightness edits keep all non-luminance
transforms and exact opaque children, including foreign metadata. Reading `LineFormat.color` changes
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
#217 asked for `None`. A shape with `prst="upArrow"` reads as `UP_ARROW`.
rpptx accepts an `MSO_SHAPE` member or preset name to replace the geometry and
reset adjustments. `Shape.remove_style` removes `p:style` from a shape or
connector and returns `None`. Absence does not publish. A picture, graphic
frame, group, or alternate content raises `RpptxError` and does not publish.
A text box or unsupported shape kind raises `ValueError` for an auto-shape
assignment. An unknown preset raises `RpptxError`.

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
shape of a slide with the relationships and parts only it used and advances the
revision once. `move(from_, to)` changes the z-order like
`SlideCollection.move`, so the shape ends up at index `to` and draws above the
shapes before it, and advances the revision once. python-pptx has no z-order
API.

A group's `shapes` collection has the same `add_textbox`, `add_shape`,
`add_connector`, `add_group_shape`, `add_table`, and `add_picture` through the
native `ShapesMut`, so groups nest to any depth. A member takes a `p:cNvPr` id
unused across the slide, and the group, then every group enclosing it, is refit
to the union of its members as python-pptx does, so members of a new group keep
their slide coordinates. An addition stages a cloned presentation and advances
the revision once after native validation, checked owner preparation and
allocation of the returned Python shape. Existing physical owners retain their
authority, including the group, its collection and its existing members. The
returned member is captured against the prepared successor before publication.
The same group collection can make the next addition without being re-fetched.
An error leaves native state and owner generations unchanged.
A group added inside a group has the zero `a:xfrm` python-pptx writes, so its
`left`, `top`, `width`, and `height` read zero until its first member arrives,
while a group added to a slide's own shapes reads `None`. `add_group_shape` has
no python-pptx `shapes` argument for moving existing shapes into the new group.
Adding to the collection of a shape that is not a group raises `ValueError`
before any image is read, and so does adding to a group inside an
`mc:AlternateContent` fallback, which stays read-only. `remove` and `move` on a
nested collection raise `ValueError` too.

A table `Cell` follows python-pptx for `merge(other_cell)`, `split()`,
`is_merge_origin`, `is_spanned`, `span_height`, and `span_width`. Merge and
split run the native staged table operations, which keep the rectangular grid.
They advance only the cell-text generation, leaving Table, Row, Column and
Cell grid handles live while retiring previously captured nested text handles. A cell of another table raises
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

`Cell.text_frame` returns the same TextFrame surface used by shapes, including
paragraphs, runs, fonts, spacing, alignment, bodyPr anchor and wrapping. The
getter validates the complete explicit Row/Cell path and never creates missing
text bodies. A merge continuation retains its own grid-cell address. The
native immutable accessor returns TextFrameRef for the cell borrow lifetime,
and TableCellMut::into_text_frame consumes its handle for safe mutable access.
Missing-body reads remain explicit, while successful text authoring may create
the body. Counted replacement preflights a missing body so zero matches, count
mismatch and invalid empty placeholders preserve source and generations.

Cell frames and descendants capture the owning table-grid and cell-text
scopes. Text replacement, merge/split and paragraph/run topology edits retire
affected descendants while retaining the cell/table and unrelated owners where
the operation preserves them. Formatting and run-value edits preserve text
identity. Failed or no-match edits do not retire scopes. `Cell.vertical_anchor`
reads and writes the direct `a:tcPr` anchor, independently of
`Cell.text_frame.vertical_anchor` on `a:bodyPr`. Both retain physical text
handles. Cell hyperlink reads use the owning body. Assigning `address` uses
the same slide-scoped external relationship as a text-box run. Assigning
`target_slide` on a cell refuses before relationship mutation.

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
these edits advances the revision once, because row, column, and cell handles
name indices, and the returned handle is captured after the change.

The Presentation Python `Table` handle has no built-in table style selector.
For a style GUID not exposed by the binding, the supported fallback is to save
the deck, replace only the matching table's `a:tblPr/a:tableStyleId` in the
slide XML inside the ZIP package, and reopen it. The writer must preserve all
other decompressed parts and validate internal relationship targets after the
edit. A style GUID supplied this way is a package-level edit, not a Python
`Table` API operation.

`TextFrame.clear` leaves one paragraph, keeps its properties, and drops its
runs. Held runs of that frame become stale. Other shapes stay live. Assigning
empty text writes a paragraph with no runs, so the next `add_run` is the only
run. `TextFrame.autofit`
reports `none`, `normal`, or `shape` when the body carries an explicit choice.
`Run.font` reads the run's direct Latin name and size and exposes its live
colour format, while the
`Run.text` setter replaces only that run's text and preserves its typed and
unmodelled properties. `Run.hyperlink` returns a live `Hyperlink` whose
`address` reads the target of the run's `a:hlinkClick`, or `None`. Assigning
an address goes through the native `set_run_hyperlink`, which reuses the
slide's relationship to the same address and removes the old relationship
once nothing on the slide names it, so retargeting does not grow the part.
`None` or an empty string removes the hyperlink, as in python-pptx, an address
with a control character raises `RpptxError`, and the write does not advance
the revision.

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
  direct bullet, and `True` raises because a scheme is required.
  `auto_number` selects the concrete scheme and optional start instead.
  `bullet_char` is a checked alias with the same str, bool or None contract,
  including clear/no-op behavior and atomic invalid-value refusal.
- `Paragraph.add_run(text="")` stages an append and allocates the new returned
  run against prepared successor scopes before publication. The owning
  paragraph and text frame remain live, while prior affected run descendants
  retire with the exact public operation label.
- `Font` reads and writes the same properties for a run and for a
  paragraph's default run properties: `name` (`a:latin` typeface, keeping its
  other attributes), `size` (1 to 4000 points, as python-pptx validates),
  `bold`, `italic`, `underline`, `strike`, `all_caps`, and `color`.
  `underline` reads `True` for `sng`, `False` for `none`, and an
  `MSO_UNDERLINE` member otherwise. `strike` reads `True` for a single or
  double strike, and assigning `True` keeps a double strike. `all_caps` does
  not name `cap="small"`, which is preserved until `all_caps` is assigned.
  `color` is the shared live `ColorFormat`, whose `rgb` reads `RGBColor`
  rather than an `RRGGBB` string. Direct assignment accepts the shared RGB
  forms, and `None` removes the direct fill. Run and paragraph-default font
  sources reuse the existing font property update path, including cell fonts.
  Clearing an absent direct value remains a byte-stable no-op.

PowerPoint `ColorFormat.brightness -> float | None` reports `None` without a
modeled explicit colour, the first luminance offset when present, otherwise
the first luminance modulation minus one, or zero without either transform.
Its setter takes a finite number from -1 to 1 and requires an explicit modeled
colour. Booleans and nonnumbers raise `TypeError`, while nonfinite or
out-of-range values raise `ValueError`, before mutation. Positive brightness
writes `lumMod = 1 - b` and `lumOff = b`, negative brightness writes only
`lumMod = 1 - abs(b)`, and zero removes both. Values use round-half-even at
100000 units. Other transforms and raw children remain intact. This finite
policy deliberately refuses boolean and NaN inputs accepted by the pinned
python-pptx oracle. Existing fill foreground prerequisites, line
solidification and shadow opacity behavior remain unchanged.

`Presentation.try_replace_text(placeholder, replacement, *, expect=None)`
runs the native staged literal replacement over slides and speaker notes with
the GIL released and returns its count. When `expect` is given, the
replacement runs on a clone, and a count that differs raises
`ReplacementCountError`, an `RpptxError` subclass that carries `expected` and
`found` and words its message like `rpptx replace --expect`. The presentation
and its revision then stay unchanged. Otherwise the replacement is kept, and
the revision advances once when the count is nonzero. Without `expect`, zero
matches return zero rather than raise, as the rdocx `try_replace_text` does.
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
frame only. Both binding routes stage a cloned presentation before native
replacement and finish checked owner preparation before publishing it. The
count, error and message, unchanged presentation and revision after an
`expect` mismatch, and single revision advance for a nonzero count retain the
presentation form's contract. Zero matches publish nothing.

The slide route uses `Slide.try_replace_text` as the retirement cause. Only
physical owners whose accepted native text changed receive a new descendant
text generation. The frame route uses `TextFrame.try_replace_text` and retires
the selected owner's descendant text generation when its count is nonzero.
Affected paragraph, run, font and hyperlink descendants must be re-fetched.
The owning slide, shape, cell and frame handles remain valid, as do unrelated
physical owners and their descendants. Neither route performs global
presentation invalidation or publishes an in-place edit before validation.
A cell with no text body returns zero, or raises the same count-mismatch error
for a nonzero `expect`, without creating a body or retiring an owner.

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
native staged mutation through its checked owner publication. Affected notes
owners retire while unrelated slide and shape owners remain live. A rejected assignment leaves package bytes and revisions
unchanged. A `Slide` also exposes an ordered tuple of frozen `Comment`
snapshots. Each comment contains an ordered tuple of frozen `CommentReply`
snapshots, and the presentation exposes an ordered tuple of frozen
`CommentAuthor` snapshots.
Author, comment, and reply additions accept native GUID and RFC 3339 strings.
Comment and reply moves retain native final-position semantics.
`Slide.resolve_comment(comment_id)` marks a thread resolved and treats a reply
id as unknown. `Slide.remove_comment(comment_id)` removes a thread with its
replies, or one reply. Both use the native staged operations of
`rpptx comment resolve` and `remove`. A successful
collaboration operation publishes its affected owner scopes once. Constructor
or native validation failure publishes no candidate and leaves existing
handles valid.

### Checked presentation surfaces

`Row.cells` returns a tuple of checked live cell views.

Held `Table.rows` and `Table.columns` collections survive supported row and
column additions and removals under the same physical table. Earlier indexed
rows, columns and cells retire with the first topology-edit cause. Newly
returned indexed members bind the prepared successor. Invalid edits and
no-ops leave both collection and member authority unchanged. Parent removal
retires the table collections. `Theme.name` reads and
updates the owned theme root. An absent `CoreProperties.author` reads as None,
while an explicitly empty creator remains an empty string. `Paragraph.text`
retains its paragraph and properties, replacing its run children and retiring
the affected captured runs.

`Shape` exposes title, alternative text, decorative state and direct flips.
`alt_title`, `alt_text` and `decorative` use the actual shape's cNvPr owner.
`ShapeCollection.group`, `Shape.ungroup`, `align` and `distribute` stage selected
physical owners and preallocate returned handles before publication.
The `relative_to` compatibility argument uses the existing arrangement
reference provider. Conflicting explicit reference arguments refuse before
publication. Cardinality refusal states the actual required minimum, selected
count and reference context while retaining the public exception class and
existing operation cause. Connector
endpoint methods validate actual connection sites and current target owners.
`Shape.insert_picture(image_file)` fills a picture placeholder under its
existing identity with checked crop-to-fill. Invalid image or unrepresentable
source metadata leaves bytes and every handle valid. Generated placeholder
names follow native type/orientation/collision rules, without renaming imports.

`Run.hyperlink.target_slide` gets, sets and clears same-presentation slide
jumps under checked relationship ownership. External addresses and action-only
clicks retain their distinct semantics. A live `Font.highlight` colour view
shares the RGB/theme/brightness contract and refuses destructive unknown
transform edits. `Font.highlight_color` names that same checked colour view.
`Font.small_caps` and `Font.all_caps` select mutually exclusive
direct modes. These getter reads create no owner nodes or scope generations.

## Native Word facade stability

Native Rust exposes `Document::paragraph_visibility(&self) -> Result<Vec<bool>>`.
Each flag matches one `Document::paragraphs()` item in encounter order. The
inventory includes direct body paragraphs and admitted block-control
paragraphs, entering control-owned tables. Direct body tables and direct
nested cell tables remain excluded. Qualified physical owner paths and typed
content slots must match before any flags are returned.

The query clones one unpublished candidate, flushes its current model to the
resolved main part once, and classifies that source without changing the live
document or package. It uses retained namespace bindings, including aliases,
defaults and shadows. Unsupported or ambiguous source construction, malformed
XML and owner mismatches return Error. It does not save, reopen or serialize
each paragraph separately.

Accepted non-whitespace text, drawings, pictures, objects, equations, symbols,
tabs, breaks, note references and supported special characters are visible.
A soft hyphen alone, deleted text, comment references, properties, range
markers and producer hints are not. Complete simple and complex fields use
consistent accepted typed results, including nested nontext caches, before
their current effective text cache. Loose markers and instructions alone do
not count. Current empty PAGE/NUMPAGES serialization supplies its existing
visible placeholder. Parser-owned results retain qualified source identity
and one authoritative nested child graph. Public cache reads retain their
consistency filter, while child edits and refresh use that graph even when a
parent's effective cache is temporarily stale. Unsupported replacement of a
retained child owner is refused.

`Document::audit_accessibility` uses the same classifier with only direct
body paragraphs. Every nonparagraph body item resets its empty streak.
Pictures and other displayed content interrupt the streak. If visibility
cannot be determined, the infallible audit retains metadata, image-alt and
heading findings, adds one Warning with the actual error detail, and omits
consecutive-empty findings. It does not invent empty flags on failure.

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
`number_context`, `number_suffixes_without_text`, `num_id`,
`definition_id`, `has_start_override`, and the hidden
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
key. Python and CLI expose checked licensed font operations through these
native providers. WASM gains no new font-authoring entrypoint.

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
bytes. That snapshot carries the namespace declarations `Paragraph.xml` already
writes, so the item parses alone. The native exact-slice API stays the separate
source excerpt. Cached complex-field text agrees between owned bulk snapshots and direct
native reads when the field spans sibling runs or nests inside another field.
Snapshot reads preserve source locations, XML, direct-body coordinates and the
document revision, leaving held live handles valid. Valid empty comment owners
remain visible as empty comments without an anchor. Reading them, listing story
items and performing no-op field or TOC updates preserve package bytes and live
paragraph handles. The native operation-local source inventory changes no
public snapshot identity or binding signature. Nested instruction caches
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
record. A captured snapshot from `Document.hyperlinks` resolves its surviving
physical owner and current ordinal. Removal retires that authority even when
another link has identical public fields. A record built from public fields
resolves only when exactly one link matches. These operations preserve proven
surviving content owners and retain conservative regular-run invalidation.
`Document::set_picture_size` resizes every main-document `pic:pic` drawing
whose blip names one image relationship. It writes `wp:extent` and the `a:ext`
of `pic:spPr/a:xfrm`, scales `wp:effectExtent` with the extent on each axis,
and leaves anchor positions alone. It rejects a zero size. A VML `w:pict`
picture on the relationship is not resized, the relationship of an SVG blip
extension is not matched, and a picture inside a group is not resized. Python
exposes this compatibility operation as `Document.set_picture_sizes_for_relationship`
with EMU sizes. The distinct `Document.set_picture_size` takes a frozen physical
Picture snapshot and changes one occurrence, invalidating affected handles after
publication.

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
`orientation` and `page_width` to `break_type`, with explicit Length values and
the `ST_PageOrientation` and `ST_SectionType` spellings. Nonzero public geometry
requires a genuine Length, including Twips or Emu. Bare zero remains subject
to native field rules, while None or omission leaves that argument unchanged.
This applies to page dimensions, four margins, gutter, column spacing and
header/footer distances. Nonzero bare integers refuse with explicit-unit
guidance. Existing live Section setters retain their integer EMU contract and
wrap only their already validated internal forwarding values as Length.
Snapshot lengths remain integer EMU and can be explicitly wrapped with Emu
when passed back to update_section. The native setters write
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
native Rust, with Python's existing `Paragraph.add_equation(source, format=...)`
route reusing it. The finite Issue365 matrix admits named Greek variants,
relations, operators, arrows and degree, upright sin/log/ln/exp/operatorname,
plain mathrm and literal text. Function runs use normal=true, plain mathematics
uses style=Plain and normal text uses both. Literal text retains boundary and
interior whitespace. Supported MathML function application and text-boundary
spacing normalize into these existing carriers. Unknown marker contexts and
unsupported properties retain explicit diagnostics.

frac, dfrac and tfrac normalize to stacked Bar fractions and canonical frac
output. Forced TeX display/text style is not preserved. binom uses existing
parentheses around a NoBar fraction. Thin, semicolon and quad spaces retain
U+2009, U+2004 and U+2001 respectively, with canonical MathML widths .167em,
.278em and1em. Native semicolon paint uses an explicit Unicode approximation.
These bounded pre-1.0 behavior expansions add no public signature, model field,
variant or binding method and do not claim unrestricted TeX support.

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
fetch external resources. Python `Document.from_html` and `open_html` return
`(Document, tuple[HtmlDiagnostic, ...])` with immutable location, property and
message records. `from_markdown` and `open_markdown` additionally accept an
optional loaded Document or template path. Their explicit image map contains
source, bytes and filename triples. Inputs are bounded UTF-8. WASM gains no
import entry point.

Python `Document.insert_html(html, *, after, images=())` accepts a required
live paragraph of this document and returns detached diagnostics. It inserts
after the checked physical direct child in a supported body, cell, header or
footer. Refusal preserves the document and handles, while changed publication
invalidates previous handles. CLI `rdocx convert` admits HTML and Markdown
input with `--to docx` or a DOCX output path, and optional `--template`. It
rejects rendering options, protects input and template paths and stages output
atomically. Local image paths resolve only inside the input directory with
bounded reads. Traversal, escaping symlinks and unsupported schemes refuse,
while missing images retain the importer diagnostic policy.

Native Rust callers can also insert rich HTML into an existing Word story with
`Document::insert_html_fragment(&ContentLocation, &str,
&[HtmlImageResource])`. The additive `HtmlImageResource` value contains an
exact source key, caller-owned bytes, and a filename. The operation accepts
body, cell, header, and footer destinations supported by the story API. Its
`HtmlFragmentInsertResult` returns the refreshed `StoryId`, the half-open direct
child range, and ordered `HtmlDiagnostic` values. Links and images are scoped
to the physical destination part. Unsupported or unresolved markup is
diagnosed, while invalid locations, malformed resources, bound failures, and
package failures leave the document unchanged. Python exposes this transaction through its checked paragraph-relative
`insert_html` method.

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
and CLI surfaces gain no presentation HTML import entry point and retain their
existing import methods and errors.

Presentation HTML text collapses ASCII whitespace across inline boundaries.
The retained space keeps the formatting and hyperlink ownership of the first
source text in that whitespace sequence. Leading and trailing paragraph spaces
are trimmed, and explicit breaks reset pending whitespace. Nonbreaking spaces
and other non-ASCII spacing characters remain literal text, including supported
positioned NBSP-only content and table cells. Unpositioned content still receives
the explicit unsupported implicit-layout diagnostic.


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

Every document-wide Word replacement, literal, regular
expression, or template, edits every copy of a text box that Word writes in
`mc:AlternateContent`. The count is that of the first `mc:Choice` that holds a
text box, the DrawingML copy that layout draws and the story walkers read. The
VML `mc:Fallback` and any later Choice are edited whatever they hold, a
Fallback that a story edit left behind its Choice included, and never counted.
A replacement can therefore change a Fallback and report 0. When no Choice
holds a text box, a Fallback text box is replaced and counted as any other
text box. Replacement handles `mc:AlternateContent` at any depth outside
another text box, while the story walkers read it only as a child of a run.
A story-item paragraph edit of such a text box uses the paired transaction below, so both admitted copies change or neither does. The handle that published the edit retires. `redact_text` removes the literal from every copy and
still counts each copy.

A positive checked scoped replacement in a qualified alternate text box maps
its selected paragraph or descendant cell to every admitted counterpart.
Physical signatures, topology and independent occurrence counts must agree
before all original source ranges are spliced. Each representation retains
its run formatting and opaque source, and unselected owners remain exact.
Asymmetric, ambiguous or unsupported mapping refuses the whole package.
Expected-count mismatch and zero-match behavior precede alternate-copy
admission. Invalid replacement text and stale locations fail without
publication. Foreign namespace lookalikes cannot establish alternate ownership.

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
DER or the admitted single unencrypted PEM envelopes on native targets. It flushes typed state into a staged document, asks the
shared package layer to sign and verify the complete graph, and commits only
the verified candidate. The additive APIs distinguish cryptographic validity
and complete declared coverage from certificate-chain trust. Python and CLI adapters explicitly enable signing and owned reports. WASM
does not select the feature.

The native `Presentation` facade exposes the parallel default-off security
surface. `open_encrypted`, `from_encrypted_bytes`, and the bounded bytes
variant use the shared package reader. `save_encrypted` and
`to_encrypted_bytes` use the shared fixed Agile profile, with sibling-file
atomic publication for the path method. `verify_signatures` stages current
typed presentation state before returning shared reports. Native `sign`
accepts PKCS#8 private-key DER and X.509 certificate DER or the admitted PEM
envelopes, then commits only a
candidate whose signature is cryptographically valid and has complete declared
coverage. Relevant mutation retains signature infrastructure for inspection
but makes verification report it as invalid. Python and CLI adapters explicitly select digital signatures. WASM and
native defaults remain unchanged, and the tool adapters do not select Agile
encryption.

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

An empty, namespace-qualified cell `w:tcMar` models the same typed margin
container whether its source is self-closing or expanded. Namespace
declarations alone do not make it unmodeled. Its four absent edge values
inherit existing table margins rather than inventing explicit zeros. Genuine
attributes on self-closing containers and foreign or rebound namespace
lookalikes retain exact opaque XML and remain unmodeled. Typed serialization
keeps the existing cell-property schema order and unsupported sibling slots.

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
anchor access without permitting part-local mutation. `resolved` is `Option<bool>` in Rust and `bool` or `None` in Python. A missing `commentsExtended` entry is `None`. A library-written comment that still has an extension is `False`. The CLI treats only `Some(true)` as resolved. Frozen Python `Comment`
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

Template occurrence staging uses the existing accepted-run mapper through the
doc-hidden `replace_prepared_occurrences(&mut CT_P,
&[(usize, Range<usize>, String, Option<CT_RPr>)]) -> Result<()>` provider.
Prepared original stretch/ranges are validated together before publication.
A run-property override affects only inserted text, preserving source prefix
and suffix properties. This helper adds no binding surface or public carrier.
The same placeholder carrier exposes doc-hidden
`replace_original_literals_in_xml_part`. It rewrites original literal stretches
outside checked physical text-box ranges and does not admit an unknown
structural owner.

Native Word callers render templates with
`Document::render_template(&serde_json::Value)`. Scalar tags use
`{{ path.to.value }}` syntax and may cross ordinary Word run boundaries.
Dedicated marker paragraphs and rows use `{% for item in path %}` with
`{% endfor %}`, or `{% if path %}` with `{% endif %}`. Blocks nest within one
container. Loops require arrays and introduce lexical variables. Conditions
treat false, null, zero, empty strings, empty arrays, and empty objects as
false. Other JSON values are true. Grammar 1 pairs structural controls inside
one physical owner. That owner may be the main body, a table cell, a header, a
footer, one normal note, or one checked text box. Markers do not pair across
those boundaries. Chart labels stay scalar. Missing paths, malformed markers,
invalid scalar leaves, invalid numbering references, and crossed container
boundaries fail without mutation. One row loop may own several adjacent
template rows. Each iteration retains table banding, grid and merge properties,
and preserved row and cell XML. Repeated list items retain one source numbering
identity and level, so their sequence continues across iterations.
`Document::render_template_versioned(&Value, u32)` admits only grammar 1.
`Document::render_template` keeps the legacy grammar.
`Document::render_template_with_formatter` selects grammar 1 and invokes the
existing `MailMergeFormattedText` carrier once per scalar occurrence, in final
physical order, with the original paragraph location. Callback text and run
properties are validated before commit. Chart text and unqualified raw text
boxes have no original paragraph location, so formatter mode refuses those
active scalars before any callback. Built-in chart and raw text-box
substitution remains available without the formatter. The existing Python
`render_template` method takes optional `grammar_version`. `None` keeps the
legacy grammar and `1` selects grammar 1. Other versions and malformed values
refuse without publishing. Python has no formatter callback. WASM and CLI gain
no template surface.

Native Word callers can also inspect content controls through
`Document::content_controls` and the tag or alias lookup methods.
`ContentControlRef` exposes immutable metadata and display text. Direct setters
update every matching tag or alias, while `bind_content_controls` applies a
string map with tag precedence and alias fallback. Bound values update their
custom XML datastore and displayed text atomically through the
package-preserving facade. Python retains callable frozen snapshots and the
existing all-match tag/alias setters. Its bind_content_controls(dict[str,str])
wrapper uses the same tag precedence, alias fallback and ignored unmatched keys.
WASM has no corresponding new form surface.

`Document::insert_content_control` takes a `ContentLocation` and a
`ContentControlSpec`. Python
`Document.insert_content_control(destination, placement, kind, tag=None, alias=None, lock=None, appearance=None, placeholder=None)`
uses the same fields and does not expose `ContentControlSpec`. Placement names
are `block`, `row`, `cell`, `paragraph`, and `inline`. Kind names match the
frozen snapshot. Creation admits `rich_text` and `plain_text`. Other known
kinds return `RdocxError` and name F-286. An unknown kind is `ValueError`.
Lock names are the stored values `unlocked`, `sdtLocked`, `contentLocked`, and
`sdtContentLocked`. Appearance names are `boundingBox`, `tags`, and `hidden`.
Placeholder is `bool` or `None`, and any other type is `TypeError`.
`set_content_control_tag`, `set_content_control_alias`,
`set_content_control_lock`, `set_content_control_appearance`, and
`set_content_control_placeholder` each take a location and an optional value.
An omitted value or `None` clears that field and does not replace content.
`remove_content_control` and `unwrap_content_control` use the same location.
Python `move_content_control(source, destination)` and
`clone_content_control(source, destination)` take two `ContentLocation`
values. Move keeps the id and stays inside one story. Clone allocates a new
id and drops `w14:paraId` and `w14:textId`. A destination whose parent cannot
hold the control content refuses before publication. A nested content
control is transparent under that grammar. Decoded references at that
boundary use the same rule. Python returns `None`.
Rust returns `Result<()>`. A refusal is `Error::Other` and does not publish.
`ContentControl` stays a frozen snapshot. WASM and CLI gain no method.

`Document::set_content_control_type(location, payload)` and
`Document::repeat_content_control(location)` use the same location.
`ContentControlTypePayload` is rich text, plain text, date, checkbox, combo
box, drop-down list, picture, group, repeating section, or repeating item.
Python constructs that payload with `rich_text`, `plain_text`, `date`,
`checkbox`, `combo_box`, `drop_down_list`, `picture`, `group`,
`repeating_section`, and `repeating_section_item`, then passes it to
`Document.set_content_control_type`. `insert_content_control` does not grow a
payload argument. Citation, equation, bibliography, and document-part kinds
stay refused. Combo and drop-down values still go through the existing tag,
alias, and bind setters. Picture and group still refuse a text value. The last
repeating item refuses `remove_content_control` and publishes nothing. WASM
and CLI gain no method. Repetition intentionally clones the first item, unlike
Word's addressed-item default. Last-item refusal intentionally preserves the
section instead of Word's whole-section deletion. Native qualification uses
the documented equivalent operations and records both default differences.

`Document::add_custom_xml_store(item_xml, schema_uris)` returns the braced item
id. Python `Document.add_custom_xml_store(item_xml, schema_uris=())` takes
`bytes` and a sequence of strings.
DOCTYPE/DTD declarations refuse before allocation, including in the prolog.
An XML declaration must be first and follow XML 1.0 declaration syntax.
Input bytes must be UTF-8. Declarations accept UTF-8 and, for wholly ASCII
inputs, US-ASCII. Unsupported or contradictory declared encodings refuse.
Legal processing instructions and comments remain supported. Invalid document
events refuse atomically. Valid item bytes, including supported unbound-prefix
Names, remain verbatim.
`Document::set_content_control_binding(location, store_item_id, xpath, prefix_mappings)`
and Python
`Document.set_content_control_binding(location, store_item_id, xpath, prefix_mappings=None)`
use the three `CT_DataBinding` strings. There is no new binding object. WASM
and CLI gain no method.

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
Word and character revisions group changed stretches as deletions followed by
insertions. Whitespace between changes folds, while a non-whitespace anchor
folds only when its joined word count is strictly shorter than both changed
sides before and after it. Character units count joined words rather than
letters. Structural, opaque and formatting boundaries, fields, drawings and
ignored units remain barriers. Default and explicit run granularity retain their
existing output. Inside supported inline controls, retained raw children keep
their physical text ownership on both sides through accept, reject and reopen.
No new comparison option or public signature is introduced.
The native facade stages the main story from package-authoritative XML and
preserves exact unchanged drawing wrappers even when sibling text in the same
paragraph, table, cell, or control changes. Accepting and rejecting the result
retain the drawing payload, relationship graph, and media bytes.
When edited image bytes replace an existing image relationship payload,
comparison carries the edited bytes in a distinct media part and owner
relationship. The drawing run receives a tracked replacement. Acceptance
resolves the edited relationship and bytes, while rejection retains the
original relationship and bytes. Unrelated package parts remain unchanged.
Chart-only and chart-plus-text comparison use the same tracked drawing
replacement boundary. Complete chart graph identity includes workbook-only edits,
retained sidecars and shared topology. Edited chart resources receive collision-safe
relationships and part names, with one shared physical destination per edited
source resource across owner stories. Saved and reopened accept-all and reject-all
select their corresponding chart caches and embedded workbooks through the actual
drawing relationships. Part-name stability is not part of the comparison contract.
No Rust, Python, CLI or WASM signature changes.
When a matched paragraph changes a comment range, hyperlink, inline control,
or preserved child boundary, comparison tracks a complete paragraph deletion
and insertion if its bookmarks remain in place. The rebuilt TOC entry
transition to a hyperlink and PAGEREF field uses this path. Granular matching
cuts at bookmark and comment range boundaries so inserted text stays on its
edited side of a marker.
Whole inserted, deleted and moved hyperlink paragraphs keep the hyperlink shell
outside revisions and track its content inside. Resolved hyperlink targets and
display text survive save and reopen. Empty whole-paragraph hyperlinks and simple
fields outside a complex field carried whole remain explicit atomic refusals.
Nested captured Word hyperlinks are also refused atomically, keeping wrapper
call depth bounded to one paragraph and one direct hyperlink owner.
Edits before retained tables carry the changed paragraph's own mark, and a
removed mark before a retained table requires no content to remain. Opaque
children and meaningful raw gaps prevent that removal. Empty paragraphs can gain
properties or text, and multiple trailing deletions retain both property states.
Foreign same-name XML stays opaque. This changes no public method signature.
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

Native Word rendering exposes `rdocx::RevisionView`, `rdocx::RevisionStyle`
and the concrete `rdocx::RenderOptions`, whose default selects accepted view
and plain style. `RevisionStyle::Legal` requires `RevisionView::Tracked`.
The public pre-1.0 source changes add `revision_style` to `RenderOptions` and
`LayoutInput`, replace `ParagraphBlock::has_visible_revision` with the optional
colour `change_bar`, and change `NoteLayout::revision_ranges` from ranges to
`(Range<usize>, Color)` pairs. Rust struct literal consumers supply the new
field and direct layout consumers use the colour-bearing fields. Additive
option-taking counterparts cover PDF bytes and files, single-page and all-page
raster output, page layout, deterministic rendering, and caller-supplied font
paths. The existing methods keep their accepted default. Python `to_pdf`,
`render_page_to_png`, `render_all_pages`, and `render_pages` take a keyword-only
`revision_view` of `"accepted"` or `"tracked"`, with `"accepted"` as the
default. Any other value raises `ValueError`. CLI `convert` for PDF and image
formats and `render` take `--revision-view accepted` or
`--revision-view tracked`, with `accepted` as the default. An unknown value is
a usage error. Python `to_pdf`, `render_page_to_png`, `render_all_pages` and
`render_pages`, including their diagnostic counterparts, accept keyword-only
`revision_style="plain"` or `"legal"`. SVG selected through `render_pages`
uses those revision options. The dedicated `render_page_to_svg` method accepts
neither revision operand and uses the accepted view with plain style.
CLI `convert` and `render` accept
`--revision-style plain|legal`. Unknown styles and legal style with accepted
view fail before output or cached layout is returned. Existing positional
Python arguments and caller font inputs retain their meaning.
HTML and Markdown conversion refuse tracked view before creating output. WASM
retains its existing rendering behavior.
Native selected-image rendering adds zero-based page-list entry points that
share `rdocx::RasterFormat`, `rdocx::RasterOptions` and
`rdocx::RasterOutput` with `oxml-pdf`. The existing PNG methods remain
source-compatible opaque defaults. Python exposes the same image controls as
keyword-only `render_pages` arguments, keeps zero-based page indices, releases
the GIL for rendering, returns `list[bytes]` for PNG or JPEG, and returns one
`bytes` value for TIFF.

Python Word `to_pdf`, `render_page_to_png`, `render_page_to_svg`,
`render_all_pages`, `render_pages`, `layout` and `layout_page` accept
keyword-only `fonts=None`, `font_dir=None` and `font_dirs=None`. Fonts are
ordered `(family, bytes)` pairs. Singular and plural directory inputs are
mutually exclusive. A scalar string or bytes value is not a plural sequence.
Directories load in caller order, supported files in lexical path order,
then explicit bytes load last. Later same-rank additional faces win. Embedded family
identity and filename aliases remain available. Caller faces override embedded
and fallback faces, and empty loaded inputs use the ordinary omission route.

Python copies bytes and path inputs before releasing the GIL. Metadata,
strict file loading and native layout or rendering run detached. Missing and
non-directory paths raise their existing I/O exceptions with the path. Other
strict loader refusals retain the public layout/domain error and exact failing
path. Invalid inputs do not produce output or mutate source. Normal Word
caller rendering keeps normal fallback, while deterministic entrypoints keep
bundled fallback. Existing caller-only native layout methods remain isolated.

Configured `render_pages` constructs one owned layout. Omitted page selection
uses that bundle's actual page count. Explicit selections retain order and
refuse duplicates, empty selections and out-of-range indices. PNG and JPEG return lists of bytes,
TIFF returns one multi-page byte stream. DPI, quality, transparency and revision
selection preserve the existing raster contracts. Measurement results remain
immutable, with no binding-private renderer or public invocation counter.

Diagnostic companions add `_with_diagnostics` to existing PDF, PNG and raster
byte-render methods. Native results pair the unchanged payload with an owned
`Vec<Diagnostic>`. Python results pair that payload with `tuple[str, ...]`.
Python signatures retain the original parameters and the caller-font keywords
`fonts`, `font_dir` and `font_dirs`. Native companions take the corresponding
render options or selection and one explicit font-files slice. The facade
prepares one configured layout and encodes or paints that same bundle once.
Existing byte-only methods retain their signatures. No rerender or mutable
last-render registry supplies diagnostics for earlier bytes.

Native and Python Word companions cover PDF, individual PNG, all-page PNG
and selected raster pages. Native Word also covers archival PDF. Native and
Python PPT companions cover PDF, individual slide PNG and all-slide PNG.
Native PPT additionally covers archival PDF, notes PDF, all-notes PNG,
handout PDF and all-handouts PNG. Python PPT adds notes PDF and all-notes
PNG companions. Python archival and handout companions are not exposed.
Word omitted page selection means all actual pages. Explicit empty selection
retains `RasterError::EmptyPageSelection` and Python `LayoutError`, with source
unchanged. PNG/JPEG keep list payloads and TIFF keeps its byte stream.
Existing SVG results retain their named diagnostic result.

Native PPT `TextFrameLayout` owns the diagnostics from the same measured
shape. The frozen Python frame exposes those messages as an immutable tuple.
A subsequent shape measurement cannot replace an earlier frame's facts.
Adding the native public field requires external pre-1.0 struct literals to
include it. CLI render and conversion publish the successful output's ordered
facts to stderr once. Fit reports consume those same output-owned facts.

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

`FontData` carries selected variation coordinates and synthetic style flags
in addition to source bytes and face index. External pre-1.0 struct literals
must include those fields. Checked PDF output refuses failed selected-font
preparation before allocating partial output, and native Result consumers
propagate the existing error. Valid output and legacy vector signatures remain.

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


### Repeating control mutation context

The existing native and Python type setter, move/clone and unwrap calls retain
repeating-type placement and ownership. A repeated item requires section
containment and block, row or cell level. A section's children remain items.
Empty wrapper contents do not establish physical placement. Retiring a section
marker or shell cannot orphan retained items, and retiring a direct item
cannot leave non-item section content. Compatible enclosing ownership and
item-only unwrap replacement remain supported. Errors publish nothing. The
bindings invoke the native policy and add no parallel implementation. These
checks do not claim full validation of malformed imported control contents.


## Native PowerPoint collaboration and navigation

Native parsing and supported authoring admit recognized PresentationML nesting
through32 and reject33 before descent or publication. Group and compatibility
owners share a counter, while timing wrappers do not increase node depth.
Python retains the existing `XmlError` mapping for malformed producing parts,
and CLI text, inspect and validate use the same checked native opening route.
No Python signature or stub changes are required. The additive OXML constant
and validator retain existing public constructors and raw/equality access.


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

Python Slide.has_notes_slide is a pure presence predicate, including existing
empty notes. Slide.notes_slide lazily creates native notes/master relationships
only when absent and prepares its returned NotesSlide owner before publication.
NotesSlide.notes_text_frame returns None for a missing body, or an owner-aware
NotesTextFrame exposing scalar text. One exact direct effective Body shape is
required. Ambiguous, grouped or alternate bodies refuse with the existing
aggregate Slide.notes_text read alternative. Every operation validates the
physical slide owner, follows surviving moves and rejects retired slides.
Existing notes are not rewritten by wrapper access.

Notes text writes stage the existing native provider and publish checked owner
survival. Outside-body placeholders, relationships and opaque source remain
owned by their original parts. Within the selected body the existing scalar
setter retains body-level state and first-paragraph/template formatting rather
than arbitrary rich-run distributions. Actual serialized source equality is a
no-op, while equal aggregate text alone does not bypass source comparison.

The enum module provides identity alias MSO_THEME_COLOR_INDEX and
MSO_COLOR_TYPE with pinned RGB1, SCHEME2, HSL101, PRESET102, SCRGB103 and SYSTEM104
values. PowerPoint ColorFormat.type returns the corresponding enum for its
modeled RGB, scheme, preset or system choice, or None when unset. The approved
getter migration replaces the earlier string results. ColorFormat.theme_color
returns MSO_THEME_COLOR for the sixteen supported scheme slots and
NOT_THEME_COLOR for a direct nonscheme color. Reading theme_color with no
direct color raises AttributeError. Unknown scheme tokens, including phClr,
refuse enum projection with an owning element.xml inspection alternative.
Reads preserve transforms and opaque source. Existing string and enum setter
inputs and source-specific None clearing remain supported.

Slide.slide_id reads the producer-assigned registration through its current
physical owner. Reordering preserves that ID, and retired slides retain causal
stale errors. Shape._element and TextFrame._txBody validate their existing
owners, then refuse the unsupported lxml graph with shape.element or
text_frame.element and its xml/replace_xml operations named. They do not
construct alternate XML owners.

The concrete rpptx.enum.lang.MSO_LANGUAGE_ID defines the 216 literal Office
numeric/tag declarations pinned to python-pptx 1.0.2. Font.language_id projects
the same native language slot as Font.language. Missing language returns NONE,
and None or NONE clears it without creating absent properties. MIXED, bool,
wrong types and unmapped codes refuse before publication. Known tags use the
pinned declaration-order mapping. Unknown imported tags refuse enum projection
while Font.language still exposes the preserved string. The existing string
setter retains its tag admission. Finite enum assignment also supports the
pinned es-ES_tradnl tag without broadening that string admission. There is no
locale inference or normalization.


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

Python exposes notes and handout PDF through the native exporters. WASM and
CLI add no new notes or handout methods and preserve the underlying parts. No
new public surface is added to `rpptx-layout`, `rpptx-render`, or the OXML
crates. The additive facade API is reviewed through the pre-1.0 release gate.

PowerPoint Python `to_pdf`, `render_slide_to_png`, `render_all_slides`,
`text_layout`, `to_notes_pdf` and `render_all_notes` accept the same three
keyword-only caller-font inputs and ordered strict loader contract. Native
caller-font counterparts retain existing arguments and result forms. Common
preparation supplies selected caller instances to slide, notes, handout,
timeline and fit measurement paths. Deterministic fallback remains bundled.
Neither WASM wrapper adds a host-directory input.

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

## Explicit PowerPoint text fitting

The render-feature facade provides `refresh_autofit(&mut self, font_files)`,
`fit_report(&self, font_files)` and `fit_text_at(&mut self, location,
font_family, max_size, font_files)`. Locations come from the checked existing
physical text owner. The input font slice follows the shared caller-font
selection rules. Refresh and fit prepare one font manager and publish one
staged package only after every relevant context succeeds. Reports do not
mutate or imply refresh.

Python `Presentation.refresh_autofit(*, fonts=None, font_dir=None,
font_dirs=None)` returns a frozen `AutofitRefreshResult`. Its `updates` and
`diagnostics` are tuples and `changed` is a boolean. Checked
`TextFrame.fit_text(font_family=None, max_size=18, *, fonts=None,
font_dir=None, font_dirs=None)` returns one frozen `AutofitUpdate`. The maximum
is an integer from one through4000, with boolean and noninteger refusal.
Caller bytes are copied before detached work and directory validation follows
the existing shared font rules.

Updates expose physical owner part/kind, optional slide index and shape id,
optional cell coordinates, mode/status, optional saved integer scale and
reduction, optional integer point size, old/new four-integer EMU bounds,
changed and owned diagnostics. Empty frames leave fitting values absent.
Native reports additionally retain typed locations and accepted Q/S versus
current persisted-trial measurements. Snapshot values do not select a future
mutation. Existing formatting-only operations keep live logical handles.
These concrete additive APIs do not add a public text-segment field or a new
layout trait.

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
Python distribution, `rdocx-wasm` and inherited support carriers at0.16.3.
The exact fifteen-package shared OOXML and PowerPoint family plus `rpptx-py`
and its Python distribution are prepared at0.14.3. The separate unpublished
`rpptx-wasm` npm carrier remains0.12.1. Word's current dependency pins require
shared0.14.3, so the shared family precedes Word publication. Preparation is
local metadata and artifact verification, without registry publication or tag
authority. The previous unified v0.16.2 and rpptx-v0.14.2 releases remain
immutable at reviewed main57f4d7438483e109cb8e45097f3498fe1c223b23.

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
one source distribution each, and both CLIs across six platforms each. It
cannot publish to registries or create a GitHub release. The build-only run
has27 successful build/aggregate jobs and four publication jobs skipped,
with26 unique payloads: twelve wheels, two source distributions and twelve
CLI archives. Aggregate copies do not add unique artifacts.

Each freshly built native wheel runs clean installed priority runtime checks.
The supported floor remains Python3.9. Exact `mypy==2.3.0 --strict` and
recursive package `stubtest` run under Python3.12, including the native
extension and handwritten stubs. The current local preparation builds real
wheels and source distributions and checks their complete metadata and README
payloads. Compatible oracle exclusions on hosted runners retain their pinned
native and CI coverage. The musllinux cells use a clean Python3.9 Alpine
runtime. Build jobs have read-only repository permissions, while selected
tag-only publication uses OIDC through the `pypi` environment.

After local scoped checks and independent clean review, release preparation
may use the canonical sprint source availability checkpoint for its named
hosted gates. Its plan stays approved and feature stays in-progress until
actual SHA-bound build-only jobs, downloaded metadata and clean
Python3.9/3.12 installed checks pass. Local artifacts do not substitute for
hosted evidence. Final integrated full verification and sprint review remain
required before close, and material source changes invalidate affected
hosted evidence.

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
The real constructor and byte round-trip operate within the pinned default
WASM stack. Their complete package-graph and populated text-style controls use
the ordinary facade without a test-only stack allowance. Private parser
storage changes preserve native signatures and retained XML.

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

`--font-dir` configures Word layout, conversion and image rendering, and
PowerPoint conversion, rendering, thumbnails and read-only fit reports. The shared strict loader
retains lexical file order and embedded identity with filename aliases.
Deterministic CLI fallback remains bundled. PDF, PNG, JPEG and TIFF consume the
selected layout instances. Unsupported output formats refuse a supplied font
directory before publication. Missing or invalid inputs preserve staging and
existing destination limits.

Both CLIs use clap derive, native family providers and the shared schema-1
JSON envelope. Word exposes inspect, text, layout, convert, diff, replace, validate,
render, comment, revision, compare, toc, fields, images, meta and fill. PowerPoint exposes inspect, text,
convert, diff, replace, validate, render, thumbnail, outline, comment,
footer, slide, notes, fit and meta.

Ordered `replace --map PATH` accepts a JSON array of placeholder/value/optional
expect objects, retaining duplicate pairs and parser last-key-wins behavior.
Unknown keys, invalid types and command-line pair/expect conflicts refuse.
Null or absent expectations allow zero, and an empty map succeeds unchanged.
Pairs cascade on one native candidate, with one final publication and ordered
counts. Native presentation `try_replace_all_expected` returns the same
family-local index/placeholder/expected/found mismatch facts as Word. Python
Presentation.replace_all exposes pairs with optional counts and ordered tuple
results. New counted wrappers allocate results before publication, preserve
held handles on failure or actual unchanged bytes, and retain causal owner
survival on changed results. Single CLI replacement retains family unmatched
behavior. Reports contain action or operation, mode, ordered index/placeholder/
value/count/nullable expected records, total_count and output.

Word `--regex` consumes `try_replace_all_regex_expected` and Python
replace_all_regex_expected, with the existing native dialect and $1 captures.
Patterns and replacement text are admitted before staging. Required raw-owner,
header/footer, note rewrite and final reopen errors propagate. Legacy regex
methods retain their existing behavior. Logical counts and alternate physical
copies use the established document-wide matcher rather than scoped picture
correspondence. Chart labels remain outside regex scope.

Word fields update consumes `update_all_fields(context)`, returning the
FieldsUpdateReport through native/Python. The report retains ordinary_updated,
layout, ordered evaluations with original instructions and cached results,
combined diagnostics, total_updated and saved_cache_scope. Python allocates
the complete owned report before publication. Evaluation outcome dictionaries
are owned report data and do not mutate the document.
One outer candidate performs ordinary and deterministic page-cache work.
`--now YYYY-MM-DDTHH:MM:SS` is a checked Gregorian civil value, without offset,
fraction or leap second. Omission keeps now=None and unresolved date caches.
Existing header/footer dynamic caches and unsupported-switch diagnostics stay
visible. The input path supplies field filename context.

Word images extract uses physical picture occurrences and actual internal
media identity, extracting each unique source part once in first-occurrence
order. Flat generated image-000001 names derive extensions from bounded content
types, with bin for unknown types. External links retain URL metadata and a
null extracted file, without fetching. Missing required media refuses.
Complete bytes and destinations are admitted before nonreplacing staged
publication. Inspect adds pictures and public word/Unicode-scalar character/
deterministic page counts, with accepted-main-body scope, spaces included and
provider diagnostics. Existing inspect keys remain compatible.

Word meta get/set uses typed current core/custom providers. Repeated
`--custom NAME=TYPE:VALUE` admits str, int32, finite float64, true/false bool,
checked provider filetime and empty without payload. Existing PID/fmtid/string
subtype/order and opaque values survive updates. Remove-custom is explicit,
and missing removal is unchanged. Tag fill splits NAME=VALUE once, uses tags
only and refuses duplicates, missing tags or unsupported bound owners as one
transaction. Native/Python fill_content_control_tags returns per-tag counts.
The mutually exclusive `fill --values VALUES.json` route accepts a nonempty
flat UTF-8 JSON object, at most1MiB and4096 assignments. Duplicate decoded keys,
empty keys, unsupported values, unknown selectors and ambiguous cross-kind
matches refuse. Keys select exactly one physical SDT through its tag or alias,
or one uniquely named legacy field. A second key selecting the same target
refuses. Strings fill SDTs and legacy text, actual bool fills checkbox controls,
and nonnegative usize integers select legacy drop-down indices. No implicit
stringification is performed. JSON assignment rows retain input order and name
the route, target kind and legacy physical part/ordinal where applicable.

One private candidate applies SDTs first, then legacy fields. Exact selected
SDT routes and metadata are revalidated before edits, after the phase and after
final reopen. Existing nested SDT retention is admitted when selected targets
remain valid. When any legacy assignment exists, the complete original legacy
identity, metadata and typed-value inventory must survive the SDT phase.
Removing an earlier legacy field inside an edited SDT therefore refuses instead
of selecting another field at a shifted ordinal. Each subsequent legacy edit
permits only planned earlier value substitutions. Final reopen precedes one
atomic output publication. Existing input/output collision protection and
existing --tag all-match semantics remain unchanged.

Core set in either family changes only supplied fields, treating empty text
as removal and preserving imported metadata.

PowerPoint slide commands use one-based CLI positions, checked decimal layout
indexes or exact unique names. Add-at composes native creation/placement on
one candidate and exposes SlideCollection.add_slide_at(layout,position) with
zero-based Python positions. Duplicate inserts after source, move uses final
positions, and hide/show preserve no-ops. Notes set requires exactly one text
or complete UTF-8 file source, preserving LF/CRLF/carriage returns, all spaces,
empty text and existing notes formatting/opaque payload through native/Python.
New editing commands require a distinct output and refuse existing files and
input aliases, symlinks or hard links before native editing/publication.

PowerPoint fit consumes the F-X207 native fit_report and exposes frozen Python
FitReport/FitFrameReport/FitTrial snapshots, without a second solver or mutable
owner selector. CLI JSON preserves physical owner part/child path/cell,
identity/name, contexts, empty/status facts, current measured/stored values,
needed total scale/reduction, quantized/persisted trials and diagnostics.
Trial dimensions/heights are points, stored bounds EMU and integer scales or
reductions hundred-thousandths. No fitting permitted trial reports null needed
scale and unfit-at-floor. Unsupported mandatory owners refuse. Fit does not
save, exiting0 without overflow,1 for overflow and2 for checked error.

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
report beside the open failure it causes. Empty-paragraph advisories count
body paragraphs and paragraphs nested in block content controls, including
control-owned tables. Direct body tables and direct nested cell tables remain
excluded, while tables wrapped by a cell control remain included. `Document::paragraph_visibility` supplies the final accepted text and
nontext classification, including visible content inside supported inline
wrappers. Complete fields require accepted displayed results, not loose field
markers or instructions. A consistent typed deleted-only result remains empty
even when its public text-extraction cache is nonempty. Properties, bookmarks,
range markers, deleted content, producer hints and soft hyphens alone do not
count. The read-only current-source scan follows native typed content-control
admission and only the first expanded content slot. Duplicate slots, invalid
opaque controls and self-closing raw slots add neither visibility nor owners.

A visibility failure returns Error when no structural error has already been
collected. Otherwise validation retains those errors, adds one indeterminate
paragraph-visibility Warning with the actual detail, and omits the empty-count
advisory. It does not substitute plain text or all-false flags. Title, author
and heading advisories retain their existing behavior. `layout --json` uses
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

## PowerPoint footer, field, theme and transition authoring

Presentation.set_header_footer accepts slide_number, footer, date, hide_on_title
and date_format, with keyword-only date_cache. None disables a field, an empty
string is an explicit cache or fixed value, and auto selects an automatic date.
Live Slide.header_footer changes individual fields and exposes presence, mode,
cache and legal policy provenance. Shape.placeholder_format exposes immutable
source idx and placeholder type through the live shape owner. Paragraph.add_field
uses the existing staged owner publication and presentation-wide field identity
inventory. Presentation.slide_number_start preserves optional source presence.

Presentation.slide_masters exposes linked per-master Theme proxies. Colors and
major/minor Latin, East Asian and complex-script fonts are read-only.
Slide.transition exposes the eight common effects, effect-specific direction,
optional duration, click and after policy. Seconds are finite and nonnegative,
with checked truncation to milliseconds. Apply-to-all stages all slides before
publication. Removing the transition does not remove timing. Retained sound,
extensions and compatibility branches survive supported edits.

The footer CLI uses explicit date caches and safe output publication without
overwriting its input. Inspect reports each master's theme, footer policy and
displayed numbering separately from physical slide identity.

## Live Word sections and rich related owners

`Document.sections` remains a tuple of immutable snapshots. `Document.section`
returns a checked `SectionHandle` with geometry, first-page flags, note policy
and six `HeaderFooter` accessors. Absent story reads create no package part.
Mutation uses the effective inherited owner or creates one explicit absent
slot. Unlinking an absent slot creates an explicit empty owner, while linking
removes the explicit reference. Structural publication refreshes the mutating
owner and makes previously fetched element handles stale.

`HeaderFooter.paragraphs` and `tables` return ordinary rich element handles.
`add_paragraph` resolves explicit paragraph styles by actual ID or name. Its
default is the actual Header or Footer paragraph style when present, otherwise
no direct style. `add_table` and nested `Cell.add_table` use checked physical
story boundaries. `FootnoteEndnote` exposes paragraphs, tables, rich additions
and deletion. Note creation accepts a same-document paragraph end or accepted
run boundary and rejects invalid note hosts before publication.

Page colour and per-edge page-border setters preserve opaque source. Real text
watermarks retain styled text in both DrawingML and VML fallback branches.
Image watermarks retain actual image bytes with checked bounds, rotation and
opacity. Replacement and removal affect only the unchanged API-owned header
representation. Legacy VML setters remain separately named compatibility
methods. Length operands reject bool, floats and overflow rather than coercing
Python values. `Twips` converts a checked signed twip count to EMU.

SectionHandle.set_text_watermark admits only opacity 0.5. Other valid-range values raise the named unsupported RdocxError before publishing any story, linkage or source change. Legacy native VML methods and image watermark opacity are unchanged.

## Extended PowerPoint properties, fills and export

Font.baseline reads float or None and accepts fractions from-1 to1. Font.spacing
reads Length or None. Font.language, east_asian_name and complex_script_name
read direct source values. Paragraph.auto_number accepts any of the41 admitted
DrawingML schemes with start_at from1 to32767. bullet_color, bullet_size and
bullet_font remain independently editable. A fractional bullet size from 0.25
to 4.0 writes `buSzPct` with ties-to-even rounding. A `Length` writes
`buSzPts`, and `Pt(17.85)` is 1785 centipoints. `None` removes the size. The
getter returns a float for a percentage and a `Length` for points. A bool is
refused. Table first_row, last_row, first_col, last_col,
horz_banding, vert_banding and style_id preserve grid and cell ownership.

FillFormat.gradient_stops exposes checked live stops with runtime identity
independent of XML IDs. Equal-valued stops remain distinct, reordering retains
identity, and deletion refuses when fewer than two would remain. Shape and
slide picture fills use the actual owning slide's relationship allocation.
FillFormat.transparency reads float or None without creating a fill and writes
only a solid fill. Transparency, gradient-stop position, percentage bullet
size, and baseline round ties to even into percentage thousandths. `0.8`
writes alpha `20000`. EMU, twips, time, and adjustment values still truncate.
Rejected writes are atomic.

Presentation.core_properties, app_properties and custom_properties are checked
live views. Sections and Slide.media return immutable snapshots.
Shapes.add_movie and native media operations preserve poster and relationship
ownership. render_slides accepts keyword-only dpi, format, quality, transparent,
slides, fonts, font_dir and font_dirs. PNG and JPEG return lists of bytes, TIFF
returns one multi-page byte stream. to_pdfa accepts pdfa-2b or pdfa-3b, and
to_handout_pdf accepts1,2,3,4,6 or9 slides per page. Both accept the common
keyword-only caller-font inputs. from_odp(file) and to_odp return the
presentation or bytes with ordered OdpDiagnostics. save_reports exposes the
ordered report payload.

## Python Word document operations

Document.charts is a read-only tuple of owned Chart snapshots. chart_type is an
optional integer, categories a tuple of strings and series a tuple of existing
CategorySeriesData values. Reading preserves package bytes and held owners.
Paragraph.add_equation retains the required format keyword and adds display=False.
True selects existing display OfficeMath, exposed by is_display, while finite
conversion diagnostics and inline defaults remain unchanged.
Document.remove_footnote(id) and remove_endnote(id) use current checked native
removal, independent of retired note handles. Non-bool signed32 ids are admitted
before publication. Missing or reserved ids retain native refusal behavior.
Selected note removal cannot revive retired aliases or remove unrelated notes.


Both families' PDF diagnostic companions return the encoded payload and an
owned diagnostic tuple from the same successful layout. Explicit Word symbols
and numbering markers, and accepted PowerPoint character markers, preserve
unknown source codes and name their family and original scalar. Unknown symbol
meaning does not imply missing glyph coverage. Caller font labels alone do not
authenticate an original symbol face. Genuine physical requested-face coverage
and known semantic mappings retain precedence.

Document exposes HTML and Markdown, CLI-equivalent text and owned TextReport, typed ODT, RTF and EPUB payloads with complete ordered diagnostic tuples, and explicit body word and character counts. Validation accepts an owned document snapshot, bytes or a path. Package decode and path errors propagate, while admitted advisory findings preserve native severity, order and validity.

Document.validate and module validate_file share the native physical-owner
validation used by the Word CLI. Identified paragraphs inside block content
controls retain inherited producer namespace bindings, including qualified
paraId and textId, through open, save and reopen. This does not relax genuine
owner mismatch or malformed-source refusal. Imported content-control type markers keep
their namespace identity across actual modeled mutation, serialization and
reopen, including producer aliases and locally shadowed canonical prefixes.
Move and clone refuse an incompatible repeating type before publication with
unchanged package bytes. Compatible block edits retain the qualified markers,
opaque metadata and the existing move or clone id policy.

Template input accepts a string-keyed dict of bounded JSON values. Bool precedes integer conversion. Nonfinite values, cyclic containers, out-of-range integers and arbitrary objects refuse before publication. FragmentConflictPolicy exposes independent style, numbering and related-part reuse choices. ContentLocation and DocumentFragment retain physical source boundaries. Checked insertion uses the native fragment path and excludes final section properties by default.

Python Run.add_break refuses Page in a relationship-resolved physical header,
including table-cell descendants and nonstandard part names. Body Page and
header Line remain admitted. Footer and Column behavior is unchanged. Imported
breaks remain source-owned through unrelated edits.

CustomFileTime validates explicit RFC3339 text. CustomProperties returns immutable typed snapshots and stages checked set/delete operations. Its mapping
assignment, deletion and membership reuse checked set, delete and get. Names
are case-sensitive str without coercion. Empty assignment raises ValueError,
empty membership is false, and missing or empty deletion raises KeyError.
Existing delete returns bool and get returns an immutable CustomProperty.
`__iter__` yields each `CustomProperty` in native part order. `__len__` is the
count. `items` yields each name and property in that order. `__getitem__(name)`
returns the property and raises `KeyError` when the name is absent. An absent
part has length 0. `keys` is not added.
Membership and failed operations create no properties part.

Python `part_names` returns a tuple of package paths with a leading slash, in lexicographic order, including a part with no relationship. `part_xml` returns the UTF-8 text of one staged part. Rust returns `Result`. A missing, binary, or non-UTF-8 part raises `RdocxError` and names the part and content type. The read sees unsaved staged parts, returns no mutable handle, and has no WASM surface.

Python `DefaultParagraphFormat` exposes `space_before`, `space_after` and `line_spacing` on `Length`. An omitted keyword keeps the stored value. A keyword set to `None` clears only that property. Rust `default_paragraph_properties` and `set_default_paragraph_properties` expose the stored `CT_PPr`. Clearing the whole native object preserves `w:rPrDefault` and unrelated opaque source. A float, int, or bool assigned to `line_spacing` refuses and does not publish. Assigning a `Length` writes `w:line` in twips and `w:lineRule="exact"`. The twip conversion divides EMU by 635 toward zero, so `Pt(12)` is 240 twips. Assigning `None` clears both `w:line` and `w:lineRule`. The getter returns a `Length` for `exact` and `atLeast`, the stored integer divided by 240 as a float for `auto` and for a missing rule, and `None` when `w:line` is absent or the rule is any other string. Omitting the keyword keeps `atLeast`, a missing rule, and an unknown rule.

AppProperties and presence-sensitive track_revisions are live document-owned handles. Equations return immutable source snapshots, ordered per-child conversion results and native loss diagnostics. Tag and alias control updates stage all selected qualified values once. Runtime constructors, explicit exports and strict stubs describe the same public classes and defaults.

Python `Document.legacy_form_fields()` returns a tuple of frozen owned
LegacyFormField snapshots. Each exposes source_part, ordinal, optional name,
enabled, calculate_on_exit, exact text/checkbox/dropdown type, typed str/bool/int
value and immutable choices. Snapshots grant no mutation authority and retain
historical values. The setter resolves current unique names or a complete
source_part/ordinal pair, with optional matching name. Partial, stale, ambiguous,
negative or bool identities refuse. Text requires str, checkbox actual bool,
and drop-down a nonnegative int excluding bool. A non-integer dropdown
refusal names the actual Python type, integer index and legacy_form_fields()
snapshot choices. It does not select SDT display text or values. Native bounds and malformed
source checks remain authoritative across supported related legacy stories.

Bind dictionaries validate all keys and values before staging. Both new Python
mutations allocate returned Python values before prepared owner publication.
Exact source no-ops preserve bytes and authority. Failed staging preserves live
source and held owners. SDT scope remains main body with nested controls, while
legacy inventory follows native supported related stories. Individual tag and
alias setters raise KeyError when their selector is missing. Selected native
InvalidValue errors map locally to ValueError in those individual setters.
Other error families and the generic bulk error mapping remain unchanged.
Native missing selectors still return zero. Bind ignores unknown keys and
retains tag-first alias fallback. Matched counts include exact no-ops. Strict
fill keeps missing and duplicate rejection with ordered assignment counts.

Date SDTs accept valid Gregorian YYYY-MM-DD values in years0001 through9999.
The explicit display masks are yyyy-MM-dd and dd/MM/yyyy. Calendar absence,
absent val and gregorian are admitted. Missing or unsupported masks, other
calendars and ambiguous metadata refuse atomically. Selected fullDate uses
canonical UTC midnight. Bound text uses the formatted display, date uses the
ISO date and dateTime uses fullDate. An absent mapping element or val uses
text semantics without adding metadata. Unselected date metadata is unchanged.

Dropdown selection matches an exact declared value or display to one unique
mapping. An absent displayText falls back to a present value. Empty attributes
remain distinct from absence. Missing values, unknown items and conflicting
mappings refuse before publication, with legal choices in the error. ComboBox
retains free text. Successful fills clear only qualified Word placeholder
metadata. Dropdown lastValue is preserved as producer-owned cache, so display
and bound scalar parity do not promise every native cache field. These scalar
form operations do not add form creation, live snapshot proxies or a second
XML provider.

Form publication preserves an owner only when independent exact physical-route
maps identify one old and one candidate owner, the existing matcher agrees,
and native owning XML is byte-identical. Paragraph/table sources share the
same batched native inventory as their addresses. Section source uses the
existing checked section XML provider. Every ancestor must also be certified.
Cells bind their exact table, row, cell and physical story, so a changed table
retires all its cells and descendants. Changed or ambiguous paragraphs retire.
Equal text or equal inventory counts alone grant no authority. Run/Hyperlink
and existing global run/snapshot epochs retire on actual mutation. Exact
no-ops and refused operations keep all current tokens. Other conservative
publishers retain their existing behavior, and retired pre-existing tokens
cannot be revived by a same-route or equal-text form result.

## Live PowerPoint masters and layouts

Presentation.slide_masters and slide_master expose live physical part owners.
SlideMaster exposes slide_layouts, shapes, placeholders, background, theme and
text_styles. Slide placeholder iteration yields every placeholder in ascending
idx and does not stop at a missing idx. Layout and master placeholder
iteration follows document order. `placeholders[idx]` remains an idx lookup on
every owner. `shapes` stays in tree order. Notes slides do not expose a
placeholder collection. SlideLayout exposes writable name, shapes, placeholders, background,
optional show_master_shapes, follow_master_background, slide_master and actual
used_by_slides. Duplicate and remove operate on the live layout collection.
Used or unsafe layout removal refuses without publishing revisions.

Theme.colors is a live twelve-slot scalar RGB mapping. Named assignment edits
the resolved theme resource, so shared theme owners see the same change. Major
and minor font sets expose writable Latin, East Asian and complex-script faces.
Value edits preserve owner inventories. Qualified source patches retain
unselected opaque bytes and refuse unsafe selected rewrites.

MasterTextStyles.title, body and other expose nine direct paragraph-formatting
levels through TextStyleList. Indexing is zero-based and level(n) is one-based.
The font property selects the first level. Levels reuse checked Paragraph/Font
formatting, including font colour, bullet choice/colour and indents. They do not
represent a text-bearing frame and do not synthesize content for text methods.

Presentation.apply_theme accepts local path or bytes, optional source_master,
import_masters and explicit layout_map pairs. Imported source closure is staged
and reopened before one success-only causal publication. Property changes retain
existing physical owners. Layout removal retires the removed owner and
descendants while surviving owners resolve their current positions through the
existing provider. Errors name the operation and retain source and revisions.

The native CT_Picture.hidden query is an additive pre-1.0 API. It reads the
qualified drawing owner's explicit true or 1 flag without changing retained
source XML or public struct fields. Picture rendering honors that flag when
Office saves master visibility into layout-owned pictures. Foreign or rebound
lookalikes do not supply the flag. This adds no Python hidden property.

Rendering honors show_master_shapes on the physical owners. Ordinary master
shapes require both effective slide and layout flags. Ordinary layout shapes
require the slide flag. Direct slide shapes and occupied placeholders retain
visibility and inheritance independently. Missing flags mean true.

## Word tables, physical pictures and internal links

Table exposes signed removable indent, autofit, actual table style ID or display-name resolution, six look flags and checked empty-row, logical-column and complete-cell-split operations. Undefined, ambiguous and wrong-type styles refuse without publication. Existing nested Cell.add_table uses the checked physical owner provider. Structural edits invalidate held rows and cells, while a no-op split preserves their revision.

PictureOptions exposes checked metadata, crop and the supported inline and floating placement vocabulary. Document.add_picture, Run.add_picture and Cell.add_picture import path or bytes and return frozen factory-created Picture snapshots. Document.pictures returns physical source order across admitted stories. Snapshots carry document origin, revision, physical owner and unchanged source identity. Document.set_picture_size accepts one snapshot and resizes that occurrence only. The former relationship-wide Python operation is explicitly named set_picture_sizes_for_relationship. Native relationship-wide sizing and replace_image retain their separate compatibility contracts. Foreign, stale, malformed and unsupported inputs refuse before publication.

Paragraph.add_internal_hyperlink and first-character # URLs use the same checked bookmark provider. Ordinary URL fragments remain external links. Tooltips retain native hyperlink formatting. Document.add_heading_hyperlink stages a suitable existing or unique valid supported heading bookmark and its source link together. Supported related-story source paragraphs use their actual owner. Bookmark grammar, owner origin and revision are validated before mutation. No external relationship is allocated for internal links.

Scoped alternate text-box replacement maps each selected logical paragraph or descendant cell to every admitted physical Choice and Fallback counterpart. Independent occurrence counts, formatting and source ranges decide publication. Unselected owners and opaque XML remain exact. Ambiguous or asymmetric mapping refuses the whole operation, while logical zero-match and count behavior retain the existing provider contract.

## PowerPoint Element XML surface

Python save refuses a case-insensitive pdf destination before writing and names
to_pdf. Authored changed slide dimensions must lie in the inclusive1to56 inch
range, 914400through51206400 EMU. Exact retained imported values remain no-op,
and changing one axis does not coerce the other. Changed horizontal or vertical
flip flags on table and chart frames refuse before geometry materialization.
Ordinary supported shape flips keep their established behavior.

Table style selection accepts None, unchanged source, an existing built-in
style or a uniquely defined qualified style in the actual deck style part.
Braced GUID identity is case-insensitive without rewriting the selected spelling.
Foreign lookalikes, absent styles and ambiguous or invalid style relationships
refuse. Invalid imported values remain source-owned until explicitly changed.

The finite XML diagnostic carrier exposes invalid_xml_child with actual part,
path and message. Refused edits preserve package bytes, revision and held
physical owners. The four source frontiers and inherited compatibility scope
are the native admission contract, not a Python XML validator or a renderer
whitelist.

The exported Element exposes immutable xml:str and replace_xml(xml:str)->None.
Paragraph, regular Run, Table, Cell, Shape, TextFrame, Slide, SlideLayout and
SlideMaster expose element. Legacy Shape.xml remains bytes. Reads capture the
actual source owner and close required inherited namespace bindings. Public
replacement validates the live owner before native resolution and publication.
A Word paragraph, table, or cell root attribute binds when its prefix is declared on that same root, in either attribute order. An unbound prefix refuses and leaves the target bytes unchanged. A checked paragraph write emits canonical `w14:paraId` without a local declaration, so the part root binds `w14` before the section-profile parse and before that source is parsed again.

A checked replacement preserves its proven selected logical owner and retires
unsafe former descendants, even when counts or values match. Table XML keeps
the selected table owner but retires its old grid, cells and text bodies.
Shape replacement retires replaced text-body and property views. Body XML
keeps the selected body owner and retires paragraph, run and font descendants.
Removed and recreated bodies cannot resurrect old handles. No-op, refusal and
allocation failure preserve document bytes and live tokens. Ordinary supported
setters retain their established lifetime behavior, while notes edits preserve
unrelated shape and text owners.

Binding publication chooses the local or staged route before changing live
native state. Exact owned XML is a no-op. Local slide shape replacement needs
a concrete scalar change and unchanged physical topology. Graphic frames also
retain identical graphicData. Descendant scopes and grid retirement are prepared
from the shared selected-owner address projection. Template and unproved edits
stage native mutation, post-edit inspection and all owner preparation before
publishing together. Scalar setters use a proven stored-state comparison or
stage. Non-blip shape fill prepares its revision before native mutation.

## Licensed embedded fonts in Python and CLI

Both Python packages expose FontEmbeddingLicense with required keyword-only
authorized and identity inputs. Embedding requires true authorization and a
nonempty exact identity. Font inputs accept bounded bytes, paths and binary
readers. Binary readers are read in bounded chunks through actual empty-bytes
EOF, with cumulative bounds checked before copying. Invalid chunk types, later
read exceptions and excess input refuse before native publication.

Word exposes immutable FontDefinition/EmbeddedFont snapshots through fonts,
set_font, remove_font, remove_embedded_font and embed_font. The four kind names
are regular, bold, italic and bold_italic. Omitted font_key derives a valid
uppercase OOXML GUID, while explicit keys retain checked native semantics.
subsetted is explicit face metadata. embed_true_type_fonts, embed_system_fonts
and save_subset_fonts are bool-or-None properties, where None removes only the
owned setting. PowerPoint exposes immutable EmbeddedTypeface/EmbeddedFontFace
snapshots through embedded_fonts, equivalent embedding/removal and tri-state
embed_true_type_fonts/save_subset_fonts. Authored PowerPoint EmbeddedFont values
retain explicit kind and license. Snapshot bytes and Python result allocations
precede native publication, and exact no-ops preserve existing owner handles.

Word authored font input requires a bounded sfnt directory with the exact name
tag before key derivation or native publication. Licensed valid faces keep
existing input and snapshot behavior. A permissive external font parser's
ability to load nameless bytes does not widen that authoring contract.

PowerPoint CLI replacement maps require every pair without explicit expect to
match at least once. Explicit expect0 remains valid. An absent selected font or
face fails without publishing a pretend removal. Non-force font commands retain
input/alias identity checks and no-overwrite staging while their existing-output
messages do not recommend an unavailable force option.

Both CLIs use font list, font embed, font remove and font settings. Embedding
requires --license-authorized and --license-identity, accepts --name, --font
and --kind regular/bold/italic/bold-italic, and writes a new --output. Word also
accepts --font-key and --subsetted. Removal without --kind removes a record,
while an explicit kind removes only that face. Settings read without mutation
options, and explicit true/false/none values require a new output. JSON uses
the existing schema1 envelope and stable native record/face order, with actual
part/format/length facts. Input/font aliases and existing output collisions
refuse. Word font errors use truthful nonreplacement diagnostics without an
unavailable force option. Identity/file-kind preflight grants no replacement
authority, and final staging remains nonreplacing. Real force-capable commands
retain their existing diagnostics. Complete native serialization precedes
staged output publication.

The four face kinds describe authored package entries and the native-backed
API. They are not a promise that every external application selects every
embedded face. PowerPoint has independently qualified actual regular and bold
use. Native Word's documented one-font-per-name behavior remains explicit,
including synthetic bold and changed or empty saved parts. Word16.114
build26101123 uses a genuine prepared Bold face in the qualified bold-only
control, but uses Regular outlines with synthetic slant in the paired Italic
control and empties that Italic part on native save. Broader native multi-face
support remains unqualified under the accepted documented-limit outcome.
A proper bold-slot control uses the existing
bold kind, explicit font key and ordinary bold run without a regular face.
It needs no family alias, new API or rewrite of the run's document semantics.
API inputs, exact license authorization, immutable snapshots and own-renderer
face selection retain their contracts.

## Presentation semantic text export

Python `Presentation.to_html` and `to_markdown` return strings with keyword
options `include_notes=False`, `skip_hidden=False` and `inline_images=True`.
Detached image mode refuses with directions to an export result or save method.
`export_html` and `export_markdown` return a frozen `PresentationTextExport`
with `text: str`, `assets: tuple[tuple[str, bytes], ...]` and
`diagnostics: tuple[str, ...]`. Detached results own their data and remain
independent of subsequent presentation edits.

`save_html` and `save_markdown` accept a path and the same keyword options,
with `inline_images=False`, and return a tuple of diagnostics after staged
publication. Invalid option types refuse before output. All exports preserve
current owner handles and presentation revisions, including on complete-body
image validation failure. Package exports and stubs
include the result type and every method.

`rpptx convert --to html` and `--to md` use semantic source order. `--notes`
includes notes, `--skip-hidden` omits hidden slides and
`--embed-images=false` writes adjacent resources. Images embed by default.
These formats reject all rendering options before any font discovery or
rendering. Semantic flags refuse on PDF and image formats. The existing
`--force` policy applies to the whole document/asset set and never permits
replacing the input. Diagnostics are separate from exported text. WASM gains
no semantic export entry point.

## Word checked XML owner authority

Word retains a finite current physical-owner inventory prepared before handle
capture and candidate publication. Paragraph, table, cell and section handles
bind actual owner and ancestor scopes. Supported append and owner-local
replacement preserve only source-proven surviving owners. Unproven structural
publishers retire current owners conservatively. Regular runs retain the
document epoch, so Document.add_paragraph retires an older run with that first
cause. Frozen Story and StoryItem captures also carry snapshot-inventory
authority. Paragraph XML retires content and property children, Run XML
retires Font and ColorFormat views, table XML retires grid descendants, cell
XML retires content descendants, and section XML retires furniture-derived
handles without retiring independent aliases of a shared physical header.
Ordinary Paragraph.text preserves paragraph properties. Ordinary Run.text
preserves the same run and rPr while retiring removed drawing and text-box
descendants. No-op and refusal publish nothing. Candidate inventory, revision
successors and returned Python allocations all complete before the single
prepared publication. The five Word element objects return namespace-complete
str and replace_xml(str) returns None through the concrete checked native
operation. Checked paragraph replacement retains ancestor-declared namespace
bindings needed by preserved inline and anchored drawing payloads when the
typed paragraph serializer drops the original declaration. It accounts for
bindings retained on paragraph and run roots and for local namespace shadows.
The opaque payload interior remains exact. Namespace-complete paragraph XML
keeps its existing public fragment contract. Checked replacement separately
reads actual physical ancestor MC authority and admits the finite35 expanded
direct-child names from the selected Word vocabulary, with pPr first and
single. Unknown newly authored Word children and non-Ignorable foreign children
refuse before mutation. Qualified OMML and valid markers remain admitted.
Unchanged opaque direct-child exceptions retain unique ordered physical
correspondence and namespace meaning, with exact snapshot no-op first. A
raw-subtree multiset cannot lend source authority across owners or structural
routes. Nested typed owners keep their existing providers. An unchanged
unsupported MC attribute retains its own raw value and resolved meaning,
without freezing supported text and metadata edits on its owner. This does
not establish a recursive schema validator or broaden ordinary import.

Current paragraph, table and cell capture uses derived exact-route indexes.
Body and cell-paragraph aliases are prepared from the native batch location
projections when the owner inventory is prepared, including before candidate
publication. The additive doc-hidden native methods paragraph_story_locations and
cell_paragraph_story_locations return the existing ContentLocation vocabulary. Unsupported entries retain the single lookup and
its refusal behavior. No Python signature changes. Read text, collection counts
and nested-table enumeration use current admitted inventory facts after scope
validation. Full or property publications replace those facts. Plain paragraph
XML prepares its accepted text before live mutation and tail removal drops only
the removed facts. No-op, refusal and numeric-slot reuse preserve their existing
causal lifetime semantics. Source facts do not bypass native mutation admission.
Ordinary parsing and physical picture snapshots keep their source authority,
and invalid replacement or save candidates fail atomically.

Complete Word owner snapshots use the native hidden
`Document.story_owner_snapshots` provider across the existing crate boundary.
Its four owned families retain full story items, physical cells, paragraph XML
and hyperlink information with the same validation as the separate providers.
One operation shares source, owner and item qualification without caching
across revisions. Each cell tuple contains the checked enclosing table, physical row and
cell coordinates and source-qualified cell identity. Nested tables and controls
participate in the same correspondence traversal. Python exposes no additional
method. Registry construction and candidate revision reconciliation perform
bounded work per current physical owner. Indexed matching preserves the existing
pairwise body-slot, projection, parent and run-path rules and first-match order,
including hyperlink index shifts after removal. A scalar cell lookup retains
its separate checked compatibility contract.

Section margin aliases use the canonical EMU length interface and truncating
Twips conversion. Header/footer paragraph append preserves the selected held
Section and independent physical header aliases, while retiring unsafe
furniture descendants. Note and equation append preserve a surviving held
Paragraph and retire affected content descendants. These operations do not
relax the conservative regular-Run document epoch.

Cell.merge admits both checked owners in the same physical table and prepares
the surviving returned cell before publication. A no-op preserves revisions.
Removing another cell with opaque cell children, cell-property children or
border children refuses atomically rather than losing those bytes.
The style ID `LightShading-Accent1` and name `Light Shading Accent 1` resolve
to the same actual canonical style definition. Lazy admission preserves
custom definitions and leaves the default document factory unchanged.

## Signing publication

Python signing uses a private signing publication policy. Exact unchanged physical
source and ancestor certificates preserve supported paragraph, table, cell and
section owners. Runs require their certified parent and identical accepted path.
Hyperlinks require the same unique physical paragraph source and local index.
Unproved owners retire. Local revision preparation preserves document authority.
Story snapshot authority survives only when the ordered story identities and
complete item snapshots are equal across staging. Report and registry allocation
finish before token commit and infallible native replacement. Read-only signature
verification serializes current Word content on a private clone and publishes no
state. Certificate identity does not establish certificate-chain trust.

Both Python owners expose sign(private_key: bytes, certificate: bytes) and
verify_signatures(). Signing returns an owned report dictionary. Verification
returns a tuple of those dictionaries, empty for unsigned packages. Fields are
signature_part, signer, cryptographically_valid, coverage_complete,
covered_parts, covered_relationships and issues. The optional signer contains
subject, issuer and serial_number. Relationship objects contain source_part
and relationship_id. Issues contain the exact snake_case native kind and a
details dictionary with every named string payload. Only the fieldless
signature_value_mismatch has empty details. Private TypedDict descriptors in
the existing stubs describe these values without adding runtime exports.
Malformed and unsupported signatures retain package exceptions. Native
re-signing and external-relationship refusals remain in force.

Both CLIs expose sign INPUT --key KEY --cert CERT -o OUTPUT and
verify-signatures INPUT, each with optional --json. Key and certificate reads
are bounded at1MiB before allocation. Signing stages a verified complete
candidate and refuses existing or aliased destinations. JSON retains schema1
and carries signature for signing, or status and signatures for verification.
Verification reports valid, invalid, partial or unsigned. Invalid outranks
partial. Exit0 requires at least one report and all reports valid and complete.
Unsigned, invalid, incomplete and runtime errors exit1. Usage errors exit2.
Both signing and verification propagate report output errors, including broken
pipes, as exit1. Signing publishes before reporting, so an output error can leave
a valid new signed package. Input and existing destinations remain unchanged.

## Native ZIP error compatibility

OpcError::Zip exposes zip::result::ZipError from the version9 dependency, replacing its zip8 type identity. Native Word and PowerPoint error families continue to carry that OPC error. This is an accepted pre-1.0 source-compatibility boundary for downstream construction, matching and From conversions. No new wrapper error or method is introduced. Python retains the existing package/error mappings.
