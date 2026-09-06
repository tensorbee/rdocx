# 03, Architecture

## Three families, one workspace

```
crates/
  # format-neutral infrastructure
  oxml-core          units, XML helpers and strict lexical validation, entity
                     decoding, raw-XML capture, core / app / custom properties
  oxml-opc           ZIP and OPC package, relationships, content types
  oxml-media         image and media sniffing, dimensions and DPI, MIME, naming
  oxml-drawing       DrawingML: colour, transforms, geometry, fills, lines,
                     effects, theme, text body
  oxml-layout        output types, font manager, bundled fonts, line breaking
  oxml-pdf           PDF writer and tiny-skia rasteriser
  oxml-sml           minimal SpreadsheetML writer, chart workbooks only
  oxml-cli-support   range parsing, JSON envelope, output-path defaulting
  oxml-chart         ChartML model and renderer
  oxml-py-support    content paths, revision checks, stale-domain errors,
                     Length conversion helpers

  # WordprocessingML
  rdocx-opc          deprecated shim over oxml-opc
  rdocx-oxml         WordprocessingML and OfficeMath types, re-exports oxml-core
  rdocx-layout       flow engine, paginator, blocks, tables, style resolver
  rdocx-pdf          deprecated shim over oxml-pdf
  rdocx-html         outbound HTML and Markdown emitter
  rdocx              the python-docx-shaped facade and inbound HTML importer
  rdocx-cli  rdocx-wasm  rdocx-py

  # PresentationML
  rpptx-oxml         PresentationML types
  rpptx-layout       inheritance resolver, chart routing and flattener
  rpptx-render       resolved slides to page frames
  rpptx-chart        deprecated shim over oxml-chart
  rpptx              the python-pptx-shaped facade, plus assets/default.pptx
  rpptx-cli  rpptx-wasm  rpptx-py
```

## The dependency rule

The graph is acyclic and layered. **Nothing in `oxml-*` may depend on
`rdocx-*` or `rpptx-*`.** There is no exception, and
`no_shared_crate_depends_on_a_format_crate` in `oxml-drawing` keeps it that
way.

```
oxml-core ──┬─→ oxml-drawing ──→ rpptx-oxml ──→ rpptx-layout ──→ rpptx-render
            │         │                                              │
            │         ←────────────────── rdocx-oxml ──→ rdocx-layout │
            ├─→ oxml-opc                                    │        │
            ├─→ oxml-media                                  ↓        ↓
            └─→ oxml-layout ──→ oxml-pdf ←──────────── rdocx-pdf   rpptx
                                                            ↓        ↓
                                                          rdocx   rpptx-cli
```

**The theme adapter.** `rdocx_oxml::theme::Theme` is a thin adapter over
`oxml_drawing::CT_OfficeStyleSheet` (`impl From<&CT_OfficeStyleSheet> for
Theme`), so that `rdocx-layout`'s existing `LayoutInput.theme` field does not
churn. The impl lives in `rdocx-oxml`, which owns `Theme`, so the edge runs
`rdocx-oxml → oxml-drawing` like every other cross-family edge.

It used to sit in `oxml-drawing` and point the other way, as the one documented
exception. That single edge made the two publication trains mutually dependent,
because `rdocx-layout` already depends on `oxml-layout`. Once both trains
carried breaking changes neither could publish first, so the adapter moved to
the side that owns its target type.

## Why these seams

**`oxml-opc` does not depend on `oxml-core`.** It has its own small local-name
handling. Staying independent means it is publishable first and consumable
alone. `rdocx-wasm` consumes the complete `rdocx` facade rather than using this
lower-level seam as a second document model. Its default-off
`agile-encryption` edge stays inside the package boundary and adds CFB plus
cryptographic primitives only when a named native consumer enables it.
The default-off `digital-signatures` edge follows the same boundary. It keeps
exclusive XML canonicalization, OPC relationship transforms, RSA-SHA256
creation and verification, and X.509 parsing in `oxml-opc`. `rdocx` and
`rpptx` forward the native package reports and stage typed facade state before
requesting package verification or signing. Both facades also forward the
default-off `agile-encryption` feature for native encrypted opens and writes.
Ordinary, Python, WASM, and CLI graphs do not include either security feature
or its cryptographic dependencies.

**`oxml-media` has no dependencies at all.** It owns byte sniffing, image header
probing, intrinsic EMU sizing through its local `NativeSize` value, safe MIME
grammar, bounded media-container signature checks, and collision-safe naming.
It remains a leaf that anything can take cheaply without importing
`oxml-core`. The format facades depend on it directly for package media names,
sniffed package metadata, and byte-first MIME inputs. The container checks do
not claim a codec decoder.

**Inbound HTML belongs to the `rdocx` facade.** The private importer uses
`scraper` for HTML5 document and fragment tree repair, then projects supported
content directly into the one owned WordprocessingML document model. The edge
does not enter `rdocx-html`, which remains an outbound emitter. This avoids a
dependency cycle and avoids a second public intermediate document model.

Bounded MHTML import and export use that same seam. The private MIME reader and
writer live in the existing `rdocx` HTML owner, project through the existing
HTML import path, and use `rdocx-html` only through its existing outbound
facade. MIME resource indexing, transfer decoding, content-derived boundaries,
and loss diagnostics add no crate, dependency edge, feature flag, or second
document model.

**ODT conversion belongs to the `rdocx` facade.** The private importer validates
the complete bounded ZIP index, parses ODF XML by expanded namespace, and
projects supported content directly into the one owned WordprocessingML
document model. The private writer walks that same tree and its package media,
materializes effective formatting, and writes deterministic ODF 1.3 content,
manifest, and image entries. ODT is not OPC, so neither direction enters
`oxml-opc` or retains an ODT object model. The facade uses the existing
workspace `zip`, `quick-xml`, and `oxml-media` dependencies for conversion.

**ODP conversion belongs to the `rpptx` facade.** One private module validates
the complete bounded non-OPC ZIP, parses ODF XML by expanded namespace, and
projects the supported subset into a fresh `Presentation`. The writer walks
that same owner and emits deterministic ODF 1.3. It does not retain a second
presentation model or shell out to LibreOffice. Unsupported safe content is
represented only by stable diagnostics.

**Inbound presentation HTML belongs to the `rpptx` facade.** Its private
module uses `scraper` for bounded HTML5 tree repair and supported selector
matching, then projects explicit absolute CSS boxes directly into a fresh
owned `Presentation`. It reuses the existing shape, text, table, image,
hyperlink, package, and validation owners. It retains no browser layout model,
performs no resource fetch, and adds no `oxml-*`, binding, or CLI edge.

**Inbound PDF belongs to the `rpptx` facade.** Its private module uses `lopdf`
for strict bounded syntax, page-tree, resource, stream, and content decoding.
It normalizes the supported operator subset into existing `oxml-layout` values,
then projects through the ordinary presentation owner. Preserved mode uses the
existing `oxml-pdf` rasterizer. Editable mode creates ordinary text, picture,
custom-geometry, and URI-link shapes. No PDF model or renderer crosses the
facade boundary, and no `oxml-*`, binding, or CLI dependency edge is added.

**Outbound EPUB belongs to the `rdocx` facade.** The private writer projects
the owned Word document through the established outbound HTML semantics, then
packages deterministic EPUB 3 metadata, navigation, reflowable XHTML, shared
CSS, and body-referenced relationship images. It builds bounded render-only
style and numbering projections instead of cloning preservation trees.
The render projection copies only supported paragraphs, runs, hyperlinks,
tables, rows, and cells. Content-control subtrees and table grids are diagnosed
from the source and never cloned into the render document. Fields are reduced
to bounded cached display values, and drawings are rebuilt without preserved
raw drawing subtrees. Paragraph, run, table, row, and cell properties are
rebuilt from the bounded values consumed by XHTML rather than cloning typed
revision trees.
Top-level outline entries split the spine, while pre-heading content becomes
front matter. EPUB is not OPC, so the writer uses the existing workspace `zip`
dependency directly and does not add a second publication object model.

**`oxml-layout` is where the format boundary genuinely falls.** Its
output, font, and line modules hold page frames, positioned elements, glyph
runs, colours, fonts, and owned line parameters, none of which name a document
format.

Presentation timeline execution keeps the same layered boundary.
`rpptx-oxml` projects typed timing, transitions, target presence, and
non-visual names without executing them. `rpptx-layout` owns pure slide-local
evaluation, stable source identity, group-space geometry, resolved frame
state, and synchronized media playback state. `rpptx-render` lowers that state
through the existing page path and composes page-frame groups for transitions
and bounded morph. The `rpptx`
facade assembles incoming and optional outgoing slides and returns the page,
state, and diagnostics together. Its media-aware entry point adds ordered
playback states and applies explicit poster fallback policy inside that same
assembly. Static resolver and renderer entry points do not pass through the
timeline path, and shared output backends remain unaware of PresentationML
timing or media playback.

Presentation media editing follows that boundary. `rpptx-oxml` projects the
schema-owned picture and timing XML while retaining raw serialization sources.
`rpptx` owns relationship graphs, package payloads, and atomic mutation.
`oxml-media` owns format-neutral signature classification and naming.
`rpptx-layout` evaluates checked trim, trigger, command, loop, volume, and
position state. The facade admits only poster images to renderer media, freezes
an unresolved poster as a deterministic labelled group when policy permits,
and never offers audio or video payload bytes to a renderer or codec decoder.

Presentation executable-content editing uses the same package seam.
`oxml-opc` owns the Transitional and Strict OLE, control, ActiveX, VBA, and
signature relationship constants. `rpptx` owns producing-scope discovery,
relationship-graph validation, opaque payload hashing and extraction, and
transactional replacement or removal. OLE owner XML remains in the slide,
layout, or master model, while ActiveX and VBA ownership is resolved through
their package relationships. No OLE, ActiveX, or VBA decoder enters an
`oxml-*`, layout, renderer, binding, WASM, or CLI crate.

Word executable-content editing follows that boundary independently.
`rdocx` owns schema-positioned story-owner discovery, relationship-graph and
content-type validation, exact payload hashing and extraction, and staged
replacement or removal. The private `embedded` module scans supported Word
story parts without promoting executable XML or bytes into the typed document
model. It commits only a serialized, reopened, and re-inventoried candidate.
The facade uses the existing `oxml-opc` relationship vocabulary and a direct
`sha2` dependency. It adds no cross-family edge, decoder, binding, WASM, or CLI
surface.

Deterministic animation export also belongs to the `rpptx` facade. It validates
and samples explicit segments, prepares the package, resolver, font, chart,
picture, and media state once, then evaluates and lowers one timeline sample at
a time. The completed opaque raster immediately enters either the facade-owned
GIF encoder or the Motion JPEG AVI writer before the next sample is resolved.
`gif`, `jpeg-encoder`, and `tiny-skia` are optional dependencies of the existing
`rpptx` render feature. No codec edge enters `rpptx-layout`, `rpptx-render`, or
an `oxml-*` crate, and audio and video package payloads are still never decoded.

The same boundary owns multilingual text mechanics. Rich segments retain
logical text and source ranges while carrying script, language, direction,
bidi level, glyph clusters, and two-dimensional advances and offsets. Script
and coverage segmentation, ICU break opportunities, conditional hyphenation,
HarfRust shaping, and line-local UAX 9 visual ordering therefore remain shared
by document formats. Logical order is the extraction contract. Visual order is
applied only to completed lines for painting.

Word and Presentation both project complex text into those shared rich values.
`rdocx-layout` selects the effective direct, bidirectional, or East Asian
`w:lang` value for each logical run, retains its exact Word source interval,
and leaves script, coverage, cluster, and line segmentation to `oxml-layout`.
It also projects `w:bidi` as the paragraph base direction and `w:rtl` as an
exact run-span override. Shared layout resolves the complete logical paragraph,
applies line-local whitespace reset and visual ordering after fitting, and
keeps the original text and source intervals authoritative.
Paragraphs containing Arabic, Devanagari, Thai, or CJK use the rich path across
PDF, raster, and SVG. Paragraphs on the established Latin path retain their
legacy shaped values and byte identity.

Language-aware automatic hyphenation keeps that boundary intact. `rdocx-oxml`
owns the Word `w:autoHyphenation`, `w:lang`, and paragraph suppression values.
`rdocx-layout` resolves them and projects only a boolean plus a BCP 47 language
tag into `oxml-layout`. The shared line breaker maps the `en`, `fr`, `de`, and
`es` primary subtags to embedded `hypher` 0.1.7 Liang patterns. The dependency
therefore stays inside the format-neutral crate and creates no reverse edge to
a Word crate.

Completed layout results share each immutable page frame and each font byte
buffer through `Arc`. Word and Presentation producers establish that ownership
at their format boundary. PDF, raster, SVG, facade page access, diagnostics,
and provenance consumers borrow the shared values without gaining a
format-specific dependency.

`SourceNodeId` and `SourceSpan` are the format-neutral provenance carriers at
this boundary. A text segment and its positioned glyph run can hold one
result-local node id plus an exclusive Unicode-scalar range. Shared line
breaking preserves and subdivides that range without learning what the node
means. Consumers must resolve an id through the format-specific result that
created it and must not compare ids from different results.

`DocumentStructure`, `StructureNode`, `StructureRole`, and `StructureId` are
the format-neutral accessibility carriers at the same boundary. A
`PositionedElement::MarkedContent` container assigns one structure node to its
exact positioned children, or assigns no node when those children are an
artifact. Word layout builds this tree from source headings, lists, tables,
and drawing descriptions before pagination. Presentation layout leaves the
optional tree absent. Shared walkers and raster output recurse through the
container without treating it as drawing geometry.

An otherwise empty Word paragraph crosses the same boundary as one empty,
zero-width text segment. The segment resolves the paragraph mark's default
font and metrics without shaping a glyph. Attributed results attach the
paragraph node with scalar range `0..0`, while ordinary results retain the
same structure with no source id. The shared PDF and raster backends treat an
empty run with no glyph ids as non-drawing content.

One construct is an exception and is called out rather than glossed. A text
segment carries an optional `NoteRef`, a footnote or endnote reference, and
notes are a WordprocessingML idea with no PresentationML counterpart. It sits
here because a note reference has to survive line breaking, which is the shared
code, and the alternative is a parallel segment type for one field. The pair
`NoteStream` and `NoteRef` replaced an untyped `footnote_id` that had the same
problem less visibly.
`rdocx-layout` keeps its Word-specific input and converts paragraph alignment,
tabs, leaders, underlines, spacing, wrapping, and twips in `convert.rs`. The
converter preserves Word's automatic line height and emits one shaped text
segment for each formatting and provenance span. Shared line breaking alone
discovers UAX 14 opportunities, reshapes each exact text slice, and subdivides
source spans. That seam is the reason the PDF backend transfers for free.
Its reusable normal engine retains one exact private identity for every
non-body layout input, section properties, and the document-wide wrapping
state that can affect cached paragraph work. The native Word facade can move a
compatible engine between two documents without exposing mutable cache state.
The facade owns a separate deterministic-base engine for layouts where caller
fonts override the bundled inventory. Missing families resolve from bundled
faces without consulting system fonts. A caller font label that differs from
its embedded family is a label-derived alias for that exact caller face. Native
callers can also install byte-free family aliases. Resolution prefers an exact
embedded family, then an explicit caller alias, then a label-derived alias,
before existing mapped and generic fallbacks. Alias state retains a
deterministic prefix of at most 256 mappings and 64 KiB of complete mapping and
lookup identity. Checked transfer includes that exact bounded alias identity
and the exact caller-font bytes in the complete retained-work context and keeps
the engine private.

**`oxml-pdf` consumes `LayoutResult` and shared image metadata.** It depends on
`oxml-layout` for the rendering contract and on `oxml-media` for byte sniffing
and header probing. It has no format-specific workspace dependency. A slide is
a page with a fixed size, so the same crate serves both formats without knowing
either exists. The `rdocx` facade renders through this crate directly, while
`rdocx-pdf` remains an exact deprecated re-export shim.
The same backend owns raster encoding through `RasterFormat`, `RasterOptions`,
`RasterOutput` and `render_pages`. PNG, JPEG and TIFF behavior is shared across
Word, Presentation, Python and CLI consumers, so format semantics and page
selection validation live at the format-neutral backend rather than in each
facade.

**Native Word SVG belongs to the `rdocx` facade.** Its private renderer consumes
the same immutable `LayoutResult` used by PDF and raster without moving the
surface into the already published `oxml-pdf` family. It preserves searchable
text, embeds used fonts and images, and recursively lowers the complete page
element tree. Stable path diagnostics make every unsupported or approximate
lowering visible while preserving supported siblings. The production edge adds
only `base64` for self-contained data URLs. Exact resvg 0.48.1 is development
test infrastructure and does not enter the runtime graph or package archive.

When the optional document structure is present, `oxml-pdf` alone owns PDF
marked-content operators, page-local MCIDs, structure elements, list bodies,
the parent tree, conditional PDF/UA metadata, and catalog accessibility
entries. The writer rejects malformed public structure graphs and withholds a
PDF/UA claim when shown text uses the `.notdef` glyph. This backend work does
not introduce a Word dependency. A result without structure uses the existing
untagged writer path.

**`rpptx-layout` is separate from `rpptx-render`.** The inheritance resolver
produces a `ResolvedSlide` in which every theme reference, colour transform and
inherited property is already collapsed to a concrete value. The renderer
consumes that and nothing else. Freezing this contract is what lets the resolver
and the renderer be built and tested independently.

**`oxml-chart` depends on `oxml-layout` for backend-neutral geometry.** Its
typed ChartML caches lower directly to `PathElement` and `Group` values. The
edge stays inside the format-neutral family, and no PDF or raster backend
becomes a chart dependency. `rpptx-chart` is an exact deprecated re-export shim
over this shared owner.

**`rpptx-layout` depends on `oxml-chart` for native chart projection.** Package
assembly parses scoped ChartML targets, then the resolver freezes a completed
backend-neutral group or a visible fallback in `ResolvedContent`. The
PresentationML resolver depends inward on the shared chart engine.
`rpptx-render` and the format-neutral backends consume only the frozen group
and do not parse ChartML.

**`rdocx-layout` depends on `oxml-chart` for native chart projection.** The
Word facade resolves document-scoped chart and theme relationships into layout
input. The layout engine freezes each inline or anchored chart as a
backend-neutral group before pagination. `oxml-layout` carries that group
through line breaking and page placement without gaining a ChartML or document
family dependency.

The `rdocx` crate uses `rpptx` only as a development dependency for the exact
cross-family chart golden. The production dependency tree has no Word to
PowerPoint edge. The all-target tree admits this test-only edge and retains the
rule that no `oxml-*` crate depends on either facade family.

## What stays put

`rdocx-oxml` remains a real crate holding the WordprocessingML grammar for
text, properties, tables, styles, numbering, borders, headers and footers,
footnotes, comments, settings, placeholder replacement, and `drawing.rs`. The
`wp:` inline and anchor code in the latter is Word-only and has no pptx value,
so it is not migrated.

That grammar also owns the bounded reader projections for document and section
completeness, table, row, and cell formatting, numbering metadata, tracked
insertions, complex-field display, and drawing relationship safety. Expanded
names and schema positions decide typed meaning. Ancestor namespace bindings
travel with retained table and document content, while malformed row revisions
and every unmodelled property remain in their original schema slots. The
`rdocx` facade exposes borrowed reader facts and computes effective paragraph
and run properties from the final direct, style, and numbering identities. It
does not maintain a second reader model.

The low-level text reader decodes visible `w:t` and `w:delText` content
fallibly and rejects malformed encoded values instead of publishing partial
text. The numbering grammar retains producer-defined `w:numFmt` tokens in
`ST_NumberFormat::Other(String)`. Its writer emits those tokens unchanged,
while render and export consumers decline to invent a marker for an unknown
format.

The same grammar owns the bounded `w:ffData` projection on complex legacy form
fields and the `w:glossaryDocument` root model. Typed form values and glossary
properties are namespace aware, while retained XML remains the serialization
source for every unsupported attribute and subtree. The `rdocx` facade owns
relationship resolution, story-part identity, staged validation, and package
commit for form-value and existing building-block replacement.

Strict XML 1.0 lexical policy is shared by these glossary and facade scanners
through `oxml_core::xml::validate_strict_xml_1_0`. The shared pass owns UTF-8,
declaration grammar, literal characters, names, namespace bindings, expanded
attribute uniqueness, references, comments, and processing instruction
targets. Each format owner keeps its document roots, schema positions,
declaration placement, doctype policy, semantic whitespace, diagnostic labels,
and public error variants.
The validator and its concrete error enum are additive pre-1.0 Rust APIs.

The same grammar crate owns Transitional OfficeMath. One concrete recursive
tree covers inline and display equations, math runs, fractions, scripts,
radicals, matrices, limits, n-ary operators, delimiters, accents, and the
bounded properties needed by authoring and later layout. Paragraph equations
use the existing raw-boundary sidecar rather than a second document model.
Expanded names decide typed meaning, fixed `m:` names are written, unsafe
prefix collisions fail closed, and unsupported descendants retain their owner
and schema slot. The settings owner projects the single document-wide
`m:mathPr` subtree through the same source-preserving replacement path.

`rdocx-layout` consumes that typed OfficeMath tree at the paragraph's retained
raw boundaries. Its private recursive math module measures the supported
expressions with the shared font manager and lowers them to backend-neutral
text, line, path, and group elements. `LayoutInput::math_properties` carries an
optional concrete copy of the document-wide defaults into that projection.
The shared `oxml-layout` group variants carry an optional baseline through line
breaking, while `None` retains the established top-aligned drawing behavior.
The PDF and raster backends continue to consume only `LayoutResult` and gain no
Word grammar dependency.

**MathML and LaTeX conversion belongs to the `rdocx` facade.** One private
module projects both formats directly into the `rdocx-oxml` `MathArgument`
tree. The MathML side reuses `quick-xml` with expanded names. The LaTeX side is
a bounded local recursive-descent parser. Pandoc is an exact-version test
oracle only and is absent from the production dependency graph.

The settings model owns the separate `w:settings` root and read-only
projections for `w:documentProtection` and valid `w:docVars` entries. It reports
the four supported editing modes, the recorded enforcement and formatting
flags, password-verification metadata, and ordered document-variable names and
values. Prefix aliases are accepted on read. Parsed producer bytes remain the
sole serialization source, so root attributes, schema order, unmodelled
children, and unsupported or malformed protection and variable elements
survive unchanged. Invalid elements are preserved but are not reported through
the typed projections.

The comments model owns typed comment entries and the three body anchor forms.
Comment bodies retain ordered paragraphs, producer attributes, and unmodelled
children. Paragraph and run models retain each anchor at its insertion boundary
without moving neighbouring raw XML. Parsing accepts in-scope aliases for the
WordprocessingML namespace, while serialization uses the fixed `w:` prefix.
The comments-extended model owns paragraph-id parent linkage and resolved state,
with unmodelled attributes and root children retained at their original
boundaries. The `rdocx` facade owns the relationship-resolved pair of comment
parts and coordinates them with the anchors in the main document.

The Word text model projects bookmark starts and ends at accepted-view and
tracked-view run boundaries while retaining every marker as ordered raw XML.
The projection combines direct paragraph markers with markers owned by typed
hyperlinks, accepted insertion and move-destination revisions, and inline
content controls in exact document order. Each nested content control retains
its own namespace scope for this projection. Opaque wrappers remain excluded,
and direct markers are not duplicated. Complex-field collapse remaps both run
views. Direct-run and marker mutation rebuild the same read projection in
memory. Simple and complex fields share one recursive
`Field` grammar with a normalized name, text or
nested arguments, switches, cached result, and optional dirty state. Its private
source records the original field form, run partition, and producer XML.
Complex fields expose ordered cached-display segments with each segment's
direct run properties. Tracked insertion projection retains inline paragraph
structure and nested revision boundaries, with a fixed depth ceiling checked
before recursive projection.
Unchanged fields therefore write their original bytes. Cache and dirty updates
rewrite only the typed values while preserving run formatting and unmodelled
neighbours. Markers are recognized only as direct run children through their
in-scope WordprocessingML namespace bindings. Malformed sequences remain opaque
raw XML, while unsupported valid fields retain their cached display. Dirty
complex hyperlinks are not reported as `Document::links()` until the update
policy defines how to handle them. The `rdocx` facade correlates bookmark ids
and owns mutation across typed body paragraphs, including supported table and
content-control traversal.
`rdocx-layout` resolves bookmark text and maps page targets, while the shared
`oxml-layout` boundary exposes only format-neutral `Target` and `TargetPage`
field kinds.

The `rdocx` facade owns pure field evaluation over that recursive grammar. It
walks every typed paragraph in main text, tables, content controls, distinct
header and footer parts, footnotes, and endnotes. Package-backed inputs come
from unique bookmarks, styles, core and custom properties, and settings
document variables. Date-time, filename, merge, and included-text values come
only from an explicit caller context. Evaluation reports resolved text,
pagination deferral, a structured TOC or TC request, a mail-merge control
decision, a validated barcode request, or a stable cached-display fallback
without mutating the package. Formula parsing is bounded to 4,096 bytes, 512
tokens, and 32 parenthesis levels. It supports numeric operands, arithmetic
precedence, comparison operators, postfix percentages, nested textual field
operands, and the existing numeric picture formatting. Unformatted
non-integral results round to 15 significant decimal digits before stable
shortest-decimal display. TOC and TC requests retain the validated heading,
style, outline, entry, bookmark, and page-number selections required by the
native rebuild boundary. TOC page-number selection also retains an optional
`\s` sequence identifier and the corresponding `\d` separator. A bare TOC
selects the built-in heading styles at levels one through nine. An explicit
`\t`, `\f`, or `\u` source does not inherit that bare-field default unless
`\o` is also present. Custom style pairs trim list-separator whitespace, and
the `\p` page-number separator is exactly one character. Every positional and switch
operand uses the shared recursive field grammar, including quoted escapes and
nested fields. Barcode requests carry validated data, symbology, dimensions,
correction, colour, and typed symbology-specific display options without
generating renderer content. `CASE` retains its public spelling while using
the same payload and case-style rules as `ITF14`. Sequence counters and
mail-merge record state remain
isolated by story. Mail-merge record and output sequence numbers come only from
the explicit caller context. Raw text boxes and other untyped XML remain
outside this evaluation boundary.

The facade also owns explicit field cache updates across that same typed story
scope. It evaluates the complete field set before changing cloned document and
package-backed parts. Resolved values replace the stored display and clear the
field-local dirty flag. Pagination deferrals, structured non-text outcomes,
and stored-display fallbacks keep their cache and become dirty so Word may
retry them. Only validated staged XML
is committed, then both layout caches are invalidated once. Existing save and
byte methods remain leave alone operations that preserve cache content and
dirty spelling. Update-aware save methods opt into the same atomic operation
before writing. The settings-level `w:updateFields` value remains untouched.

The native facade also rebuilds supported existing main-story table of
contents fields. It reparses each owned instruction through the same recursive
field grammar, discovers selected headings, custom paragraph styles, direct
outline levels, and TC entries in document order, and excludes the old owned
result range from source discovery. A complete-paragraph source reuses one
valid whole-paragraph bookmark when available. If multiple valid ranges cover
the whole paragraph, the earliest start in document marker order wins. A
source restricted to the surviving fragment of a TOC begin or end paragraph
never reuses a whole-paragraph range. Otherwise the staged candidate allocates
a unique hidden bookmark id and `_Toc` name. Level-specific
TOC paragraphs contain the selected title, optional internal hyperlink,
configured separator, and a PAGEREF field unless that level omits page
numbers. Sequence-selected entries prefix the final page value with the
nearest preceding matching SEQ value and the configured separator. TOC and
SEQ identifiers share the evaluator's ASCII case-insensitive namespace.
Source field discovery and heading title extraction use the accepted revision
projection. Direct and content-control runs retain their stored order,
insertion and move-to runs are included at their typed boundaries, and deleted
or moved-from runs are excluded. Direct outline values use checked one-based
conversion, so values outside the supported TOC level domain are ignored.
When a content control and accepted revision share a direct-run boundary, their
raw sidecar positions retain exact serialized order. Simple TOC fields inside
accepted insertions and move-to revisions stay unchanged and contribute one
diagnostic each.

Rebuild changes no instruction or source paragraph content. Byte-position XML
edits insert only newly owned bookmark markers and replace only the span after
the TOC separator run through the run containing its matching end marker.
Ownership scanning requires expanded-name Word ancestors, a recognized typed
paragraph path, and field markers as direct children of typed Word runs.
Word-shaped field content retained below an opaque descendant is ignored.
Same-namespace inline wrappers qualify only when their parent-child grammar
and required metadata match a typed owner path. Block-level document, body,
table, row, cell, and content-control ownership follows the same parent-child
grammar as the typed parser. A body or cell content control types only
paragraphs and tables, a table control types only rows, and a row control types
only cells. The byte-position rebuild scanner retains that placement on each
content-control owner, so an invalid child cannot enter its paragraph index or
shift a later bookmark insertion. Before entering a content container, the
scanner also requires the complete control shell to parse successfully in that
same placement with its inherited namespace bindings. A rejected control and
all of its fields and markers remain opaque. The public standalone
content-control parser
keeps its context-free union of paragraph, table, row, cell, and run children.
Parent document, table, row, cell, and paragraph parsers use placement-specific
internal entry points.
Other Word children remain preserved raw XML and contribute no
headings, fields, bookmarks, or other typed sources. A content control owned by
a paragraph types only inline children. Paragraph, table, row, and cell
children below that control remain opaque and contribute no runs, markers, or
TOC sources. Accepted
revision ownership stops at the parser's
32-wrapper nesting ceiling. Wrapped instructions are validated with a balanced
synthetic end marker inside their complete owner chain, and result replacement
restores the matching wrapper closures. At the matching end paragraph,
replacement retains only the exact paragraph opening, paragraph properties,
and accepted wrapper prefix needed to contain the end-marker run. Direct
self-closing paragraph properties retain their exact prefix spelling and
attributes. A content-control prefix retains its properties, end properties,
identity, binding, type, and raw property slots through its content opening.
Cached result runs before that marker are removed. Whole-paragraph bookmarks
and every other valid bookmark range with exactly one marker inside a replaced
result are narrowed to the surviving source fragment, including partial
same-paragraph and cross-paragraph ranges and entries that need no generated
target. The exact end-marker run boundary stays separate from wrapper-prefix
reconstruction. Deterministic pagination uses the staged direct marker, then
the final XML moves each generated or repaired start immediately after the end
run inside its accepted hyperlink, revision, or content-control owner. The
bookmark therefore covers surviving post-end text without changing the
wrapper structure or its metadata. Save and reopen reconstruct the same
accepted-view marker and run boundaries used by the public bookmark facade,
PAGEREF layout targets, and later TOC ownership scans, so the rebuilt document
can be rebuilt again.
Whole-paragraph bookmark reuse
compares marker bounds with every ordered direct paragraph child, including
content controls and revision wrappers, rather than only direct-run indexes.
Bookmark range validation also compares each marker's raw position at a shared
run boundary, so an end before its start is rejected.
Nested TOC ownership is rejected before any edit is applied. Page placeholders
are chosen outside the complete source and generated byte sets, then
substituted only at their unique owned result offsets. Bookmark ids and names
are allocated lazily, including the final representable id when it is the next
free value.
Unsupported valid TOCs stay unchanged and increment the report diagnostic
count. Malformed ownership, ambiguous bookmarks, missing selected bookmarks,
layout failure, or serialization failure rejects the whole staged operation.
The candidate is reopened before deterministic bundled-font pagination and
again after final displayed page substitution. The live document receives the
validated package once and both completed layout caches are invalidated once.
Main-story pagination visits paragraphs and tables inside valid body, table,
row, and cell content controls in the same order used by TOC source and
bookmark discovery. A generated PAGEREF target that does not reach pagination
is diagnosed and is not exposed as a resolved target-page value. Missing or
invalid target-page output rejects the staged rebuild before the live package
changes.

The `rdocx` facade also owns flat native mail merge over
`BTreeMap<String, String>` records. Separate mode stages, serializes, and
reopens one complete package clone per record. Its private evaluator policy
maps only an absent `MERGEFIELD` value to empty text, while ordinary field
evaluation retains its cached-display fallback. Section mode first rejects a
record-varying merge dependency in a referenced header, footer, footnote, or
endnote. It then concatenates candidate body entries in record order, moves
each non-final body section properties value to a next-page section-ending
paragraph, and retains the final body-level section properties value. A
namespace-aware serialized-body pass remaps bookmark, content-control, and
drawing identities together with bookmark field and hyperlink references,
including values held in preserved raw XML. The operation does not evaluate
structured template tags, and ordinary field traversal keeps its existing
typed story scope.

The same facade owns additive native rich mail merge over `MailMergeData` and
owned text, image, and DOCX fragment values. Whole-paragraph and whole-row
`TableStart:<name>` and `TableEnd:<name>` merge fields form bounded nested
regions. Scalar lookup walks lexical records from inner to outer scope, while
each region resolves first from the current record and then from the named
top-level source. Text replacement retains the existing merge-field switch
grammar, then an optional formatter receives the local source record number,
region path, field name, and global emitted-field sequence number. Images and
fragment relationship closures are imported into a staged candidate with
collision-free package, style, numbering, bookmark, content-control, and
drawing identities. Consumed markers and field shells are removed. Every
candidate is serialized and reopened before publication, and rich section mode
reuses the flat section assembler. The flat methods and non-body story scope
remain unchanged.

The `rdocx` facade owns structured template evaluation over
`serde_json::Value`. The focused `template` module recognizes scalar tags
across ordinary run boundaries and pairs nested `for` and `if` controls with a
container-aware stack parser. Top-level marker paragraphs clone body entries,
including section-ending paragraphs and their section properties. Marker rows
clone every row in a multi-row template group inside their owning table. The
owning table is retained, and each row and cell is deep-cloned with its merge,
banding, content-control, and ordered raw XML state. Numbered paragraphs in a
loop retain their source `numId` and level, which keeps one continuous list
without allocating definitions. Numbering references are validated before
evaluation. Loop variables form lexical scopes, and dotted lookup searches the
innermost scope before the root value. Structural controls are limited to the
main body and its tables. Headers, footers, text boxes, and chart labels retain
scalar-only replacement through the existing Word placeholder mapper. A
successful render commits the staged document and package together and
invalidates both layout caches once.

The content-control model owns one recursive `CT_Sdt` grammar at block, row,
cell, paragraph, and run placement boundaries. It reports tag, alias, numeric
id, bounded control type, and custom XML binding metadata from `CT_SdtPr`.
Unmodelled attributes, properties, and content children remain in ordered raw
slots. Empty or malformed controls remain opaque. Prefix-tolerant readers and
fixed-prefix writers follow the same boundary rules as the surrounding
WordprocessingML model.

The `rdocx` facade owns content-control value mutation because one operation
can cross the typed document and package parts. Immutable summaries expose the
control metadata and display text. Lookup and mutation select tags before
aliases, so map application updates each control at most once. Display changes
preserve the control shell, direct run formatting, and nested control
boundaries. The facade stages every selected display and custom XML change on
cloned state, validates the resulting XML, then commits once and invalidates
layout once. Any rejected control or binding leaves both document and package
state unchanged.

The revision model belongs to `rdocx-oxml`. Insertions, deletions, moves,
property changes, deleted text, and contextual markers are typed read-only
projections over captured WordprocessingML subtrees. The captured raw subtree
is the sole serialization source until an explicit accept or reject operation
replaces it. Invalid revision metadata remains preserved but is not reported.
Prior run, paragraph, table, and section properties are projected with the
namespace context of the revision element, including nested properties.
The ordered preservation sidecars added to the public low-level Word model and
the `RunContent::DeletedText` variant form the breaking pre-1.0 0.8.0 boundary
for the next published stable family. The higher-level `rdocx::Document`
revision API remains additive.

The `rdocx` facade owns revision resolution because one operation can replace
content wrappers, property owners, paragraph boundaries, and table rows across
the main document, deduplicated headers and footers, comments, normal
footnotes, endnotes, and text boxes nested in those stories.
Accepting keeps insertions and move destinations, while rejecting keeps
deletions and move sources and converts deleted text to ordinary text.
Property rejection restores exactly one namespace-correct prior property
value. Contextual markers act on their owning run, paragraph mark, numbering
property, or row. Resolution stages every affected package part, resolves
selected descendants before their enclosing subtree, reparses the complete
candidate package, and commits once only after validation succeeds.

The `rdocx` facade also owns deterministic comparison of those same stories. A
source index assigns stable public story categories and private owner paths to
modeled paragraphs, tables, fields, and nested text boxes. A concrete
hierarchical longest-common-subsequence alignment covers paragraphs, runs,
table rows, nested tables, lists, and modeled content inside existing
content-control shells. The default path retains whole-run behavior. An
additive options value selects Unicode-scalar character units or maximal
Unicode word, whitespace, and punctuation-or-symbol units, and applies
left-biased ignores for formatting, textual whitespace, fields, comments, and
selected story categories. Selected categories leave the original story bytes
untouched and are excluded before shell checks and revision-id allocation.
Non-text content remains atomic. Unmatched identical owners become move pairs
only within one story.
Changed field results remain inside their field owner, while instruction or
form changes replace that complete owner. Supported run, paragraph, table, and
section properties emit property revisions that retain the original property
sidecars. Unsupported formatting differences retain the original bytes and
produce stable `ComparisonDiagnostic` values at the actual story path. Inputs
with existing modeled revisions or differing story and control shells are
rejected unless their story category is ignored. Attributed text alignment
retains owner, formatting, content position, and raw-child boundaries, then
coalesces adjacent equal-owner edits into minimal revision wrappers.
Comparison patches only owned source spans, preserves every unowned byte,
stages the complete package, proves that acceptance matches the edited policy
projection and rejection matches the original, then commits once.

`rdocx-layout` owns the renderer-only revision projection. The
`LayoutInput::revision_view` selector chooses an accepted or tracked view. The
engine merges ordinary runs and typed revision runs at their preserved
boundaries without mutating the package. Accepted layout keeps insertions and
move destinations and omits deletions and move sources. Tracked layout keeps
both sides, applies neutral decorations, and carries changed-paragraph state
through pagination. The `rdocx` facade owns the concrete `RenderOptions` value
that passes this selection into layout. Default accepted renders reuse the
normal and deterministic caches, while tracked renders remain uncached.

The same Word projection owns glyph provenance. `rdocx-layout` allocates one
deterministic `WordSourcePath` for each modeled paragraph in the document,
arbitrarily nested tables, headers, footers, footnotes, and endnotes. A
`WordLayoutResult` resolves shared node ids through that result-local table and
records the selected revision view. Ordinary text ranges address the selected
projection. Parsed complex-field caches advance projection offsets, while new
simple fields do not. Generated markers, evaluated fields, note labels, and
non-bijective display transformations remain unattributed.

The `rdocx` facade caches that complete `WordLayoutResult` for accepted normal
and deterministic layout. It also retains one synchronized normal-font
`rdocx-layout::Engine` across document edits. Mutation clears completed result
caches but preserves the engine's bounded paragraph and shaping work.
`Document::layout` and `Document::layout_with_options` return a shared `Arc`
for accepted normal-font layout. A tracked result stays uncached, while its
revision-view cache identity prevents reuse of accepted paragraph blocks.
`Document::layout_with_fonts` and `Document::layout_with_fonts_and_options`
construct isolated engines and return owned uncached bundles because arbitrary
caller font sets have no stable cache key. Deterministic layout also retains
its separate result cache and never enters the normal engine. PDF, raster, and
cloned page access borrow the backend-neutral `layout` field from the same
bundle, so external renderers receive the exact font bytes and source table
used for each glyph run.

`rdocx-layout` keeps the flow model: the engine, the paginator, blocks, tables
and the style resolver. Slides do not paginate, so none of it transfers. The
normal Word engine caches only ordinary body paragraphs that are independent
of traversal state. Its exact identity includes the paragraph, width, revision
view, styles, theme, and active embedded fonts. Numbering, drawings, fields,
hyperlinks, relationships, generated markers, and other contextual content
bypass reuse. A successful whole-document layout publishes staged entries,
including their diagnostics and exact font-resolution trace. Failure publishes
nothing. Cached scalar ranges use a placeholder source node and are rebound to
the current result-local node before pagination.

The same engine owns bounded loaded-face, coverage, shaping, and paragraph
state. Paragraph entries count all retained owned buffers, including reflow tab
stops, against both entry and byte ceilings. Font traces are bounded and shrunk
before retention. Active fonts are canonicalized by first resolution order so
a warm layout returns the same font table and ids as a cold layout, including a
legitimate working set larger than the memo ceilings.

The flow engine resolves Word relationship IDs to content-addressed `MediaId`
values before pagination, and page output carries the resolved bytes and MIME
type rather than a relationship-scoped placeholder. One `MediaRegistry` per
layout compares complete bytes, assigns deterministic alternate IDs when two
compact keys collide, and is shared by the lower-level layout and pagination
entry points.

Header VML watermarks keep the same ownership split. `rdocx-oxml` retains the
complete `w:pict` source and projects only supported `v:shape` text paths and
images for layout. The `rdocx` facade owns atomic text and image authoring,
header-local relationships, section inheritance, and package-visible first and
even variants. `rdocx-layout` consumes that projection, resolves header images
through the shared `MediaRegistry`, and lowers each selected watermark to a
backend-neutral group before pagination. Unsupported VML stays opaque and no
backend parses WordprocessingML.

Footnotes and endnotes are laid out into a `NoteRegistry` before pagination, and
the paginator reserves, splits and draws them. Note placement is part of
pagination rather than a pass that runs after it, because a page's body height
depends on the note area it owes, and a note that does not fit continues on the
following page. The registry pre-shapes each note's marker, so the paginator
places notes without needing a mutable font manager.

Each note is laid out once per distinct section content width rather than once
per document, and is looked up by the width of the section drawing it. A note is
broken to the measure of the section carrying its reference, since that is the
measure it is drawn at, and reserve and render therefore still read the same
lines. A document whose sections share a page size registers one width and lays
each note out once, which is the common case. Endnotes are measured against the
final section, because they are emitted after the last body page and drawn
against that section's geometry wherever their references sit.

The paginator also reflows a paragraph around any floating drawing that wraps,
because whether a drawing overlaps a line is only known once the paragraph has a
position on a page. The inputs to line breaking are therefore kept alive past
layout, but only for a document that actually holds a drawing whose wrap is not
`none`, since those inputs hold the same shaped glyphs the laid-out lines do.

Text also flows around a wrapping drawing anchored to a **later** paragraph,
which Word documents do routinely. A drawing framed by the page or a margin has
a position without its own paragraph being placed, so one pass is enough. A
drawing framed by its own paragraph does not, so a section holding one
paginates **twice**: the first pass records where each such drawing landed and
on which page, and the second offers those rectangles to the text above them.
The first pass is identical to a single-pass run, and a section holding no such
drawing paginates once, which is every sample and every corpus document today.

Two passes, and not a fixed point. The second pass reflows earlier text, which
can move the drawing's own paragraph, so the rectangle it flowed around may be
slightly stale. Iterating is not guaranteed to terminate, since growing a
paragraph can push a drawing to the next page, which shrinks the paragraph,
which pulls the drawing back. Two passes give one answer, always.

The two note streams are placed differently and are keyed apart. A footnote
sits at the foot of the page carrying its reference and takes height from that
page. An endnote costs its page nothing and is emitted after the last body
page, where endnotes flow from the top of their own pages without a separator
rule. A reference therefore carries a `NoteRef`, its stream and its number,
because the streams number independently and a document may hold a footnote and
an endnote sharing a number.

## Versioning

The 15 shared and PowerPoint publication candidates use the explicit common
incubating version 0.11.0 in their manifests and workspace pins. The latest
published coherent family is 0.11.0 from immutable annotated tag
`rpptx-v0.11.0` at reviewed SHA
`0b6bd622f8a14189d7d1281d011f81319ef8ad2a`. All 15 registry entries and their
sole owner are verified, while the `rpptx-wasm` preparation member remains
unpublished at 0.11.0. The earlier 0.10.0 family remains available. The family
includes `oxml-chart` as the format-neutral owner while
retaining `rpptx-chart` as a source-compatible deprecated shim. The released
`rdocx-*` crates use the separate workspace version. The stable workspace and
its nine internal pins, eleven inherited lockfile packages, two Python project
versions, and unpublished `rdocx-wasm` package are prepared at 0.13.1. The latest
published exact seven-package crates.io family is 0.12.0 from immutable
annotated `v0.12.0` tag at reviewed SHA
`19adaacfcf82e3918bba4f8c3648747f1969b746`. Those published archives retain
their shared 0.9.0 registry requirements, while current source prepares shared
0.11.0. The immutable v0.13.0 tag at reviewed SHA
`05332b17f481741e7d5ab4e39699c6d1536475af` published five low-level stable
packages, then stopped because packaged `rdocx` required the four Word main
content-type constants added after shared 0.10.0. `rdocx`, `rdocx-cli`, and the
GitHub release are absent. The immutable v0.11.0 attempt at reviewed SHA
`25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4` published `rdocx-opc` and
`rdocx-oxml`, then stopped before the other five packages and GitHub release
when `rdocx-layout` proved it needed `TextSegment.direction` from a newer
shared registry family. The complete 0.11.1 recovery is published and verified.
The separately approved cleanup yanked exactly the incomplete
`rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0` entries. Complete coherent stable
releases remain live and unyanked. The v0.11.0 tag remains immutable, and no
v0.11.0 GitHub release exists. The last published complete stable family remains
0.12.0 until the separately approved 0.13.1 recovery. Earlier immutable
registry releases remain available. Version
preparation and manifest eligibility do not authorize any later publication.
`oxml-cli-support` is the
format-neutral owner of range parsing,
JSON envelope, and output-path contracts. It has no dependency on either
document family, while CLI binaries depend inward on it.

The immutable `rpptx-v0.1.2` release contains the earlier 12-package family.
`oxml-cli-support` and `rpptx-cli` remain unpublished at 0.1.2. The original
14-package family is published at the immutable 0.1.3 and 0.2.0 boundaries,
and the earlier 15-package family remains available at 0.4.0. No existing tag
or registry version was moved or overwritten.

The `rpptx` facade owns formatting-preserving presentation text replacement.
`Presentation::replace_text` applies literal, non-recursive replacement across
contiguous regular runs in ordinary shapes, nested groups, and table cells.
Fields, breaks, and selected alternate-content fallbacks remain traversal
boundaries so the facade preserves their unmodelled or separately typed XML.

The facade also owns modern PresentationML package identity. The exact main
part content type distinguishes PPTX, PPTM, POTX, POTM, PPSX, and PPSM. Normal
serialization preserves that source class. An explicit output conversion
changes only a staged content-type override, retains opaque executable parts
and relationships, and invalidates retained package signature evidence when
the signed table changes. Binary `.ppt` never enters this OPC path.

The `rdocx` facade owns the corresponding Word package identity and its private
Flat OPC boundary. `oxml-opc` supplies the four exact Word main content types
and the existing in-memory package. The facade validates DOCX, DOCM, DOTX, or
DOTM before building a `Document`, while `flat_opc.rs` projects XML package
parts directly into that same package owner. No second public package model or
format dependency edge is introduced.

`rpptx-*` crates carry their own `keywords` and `categories`, because the
workspace values say `["docx", "word"]` which would be wrong on a presentation
crate. Once publication is approved, the rpptx family uses its own pre-1.0
version train so breaking releases do not drag the released rdocx family with
them. The families fold into a lockstep train once rpptx stabilises.

## Crate-level conventions

- **quick-xml pull parsing only.** No serde, no derive, no macros, no codegen.
  Every element's parser and serialiser is hand-written. This is a deliberate
  existing choice and the new crates follow it.
- **Spec names.** Types are `CT_*` and `ST_*` after the schema, under a
  crate-level `#![allow(non_camel_case_types)]`.
- **Root parts** get `from_xml(&[u8]) -> Result<Self>` and
  `to_xml(&self) -> Result<Vec<u8>>`. **Nested elements** get
  `from_xml(reader: &mut Reader<&[u8]>)` and
  `to_xml<W: Write>(&self, writer: &mut Writer<W>)`.
- **Prefix-tolerant on read, fixed prefix on write.** `matches_local_name`
  strips any prefix and compares the local part.
- **An unmodelled enumerated value reads as an absent attribute.** A value
  parser rejects a string it does not list, and the property parsers treat that
  rejection as "not specified" rather than propagating it. An absent attribute
  means the element's default, which is usually inheritance from the style
  chain, so the surrounding properties survive and the document opens. The
  parsers stay fallible, so a caller that wants strictness keeps it: the
  tolerance belongs to the reader, not to the type. A value carried this way is
  lost on save, which is the accepted cost of opening the document at all.
- **Unmodelled subtrees are preserved verbatim** via `capture_element` into
  `raw_xml` fields. This matters far more for PresentationML than for
  WordprocessingML, and it is the scope control for an otherwise unbounded
  format: parse only what you render, preserve the rest.
- **`thiserror`, no `anyhow`.** One error enum per crate plus a `Result` alias.
- Edition 2024, MSRV 1.93.

## Facade conventions

Both facades use the same borrow-handle idiom rdocx already has: a mutable
`Foo<'a>` wrapping `&'a mut CT_Foo` and a read-only `FooRef<'a>`, with
consuming builders for formatting so calls chain, `&mut self` methods for adding
content that return a nested handle, and index-based `Option`-returning
accessors that never panic.

The `rdocx` facade also provides direct immutable paragraph lookup. Mutable
and read-only paragraph handles each provide total run count and lookup, while
only the mutable handle provides mutable run lookup. These accessors let the
Python binding re-resolve lazy index paths without allocating paragraph
snapshots, clearing layout caches for reads, or reaching through private OOXML
fields.

Native paragraph handles expose borrowed equation iteration and indexed
lookup. Mutable handles also expose indexed mutable lookup and one
`add_equation(OfficeMath)` operation. `ParagraphItemRef::Equation` retains
equation order among runs, controls, revisions, and unsupported XML.
`Document::math_properties` and `Document::set_math_properties` use the
relationship-resolved settings part. These are additive pre-1.0 Rust APIs.
Python, WASM, and CLI surfaces do not gain OfficeMath entry points implicitly.

The native facade re-exports four free equation conversion functions because
the normalized tree is owned by `rdocx-oxml`. The functions return the same
`MathArgument` tree or canonical text together with concrete ordered
`MathConversionDiagnostic` values. The generic result container is used for
both `MathArgument` and `String`. No wrapper, trait, feature flag, or binding
surface is introduced.

`Document::text` traverses body paragraphs and table cells in document order.
The WASM binding uses that additive facade accessor for its existing `getText`
method and otherwise owns one complete `Document`. It never reaches into
`rdocx-oxml` or maintains a second package representation.

`Document::render_page_to_svg`, its option-taking counterpart, and their two
deterministic variants expose one zero-based page as self-contained searchable
SVG. Out-of-range pages return `None`. `SvgRenderResult` carries the SVG and
ordered `SvgDiagnostic` values, with layout diagnostics before recursive
lowering diagnostics. These additive methods are native Rust only. Python,
WASM, CLI, Presentation, and the public `oxml-pdf` surface remain unchanged.

`Document::from_html` and `Document::open_html` are additive native facade
constructors. They return the converted document with stable path-aware
diagnostics for parser repairs, unsupported CSS, dropped resources, and safely
skipped visible constructs. Input, DOM, projection, text, table, and diagnostic
limits fail closed before a partial document is published. The importer saves
and reopens its candidate through the typed Word package model before returning
it.

`Document::from_mhtml_bytes`, `Document::open_mhtml`,
`Document::to_mhtml_bytes`, and `Document::save_mhtml` are additive native
facade methods. Concrete read and write results carry the converted document or
bytes plus stable diagnostics. Read and write failures use one contextual
`Error::Mhtml` variant. The paths publish only after bounded conversion,
MHTML reparse, and DOCX save and reopen checks succeed.

`Document::from_odt_bytes`, `Document::from_odt_bytes_with_limits`, and
`Document::open_odt` are additive native facade constructors. They return a
fresh converted document with ordered path-aware diagnostics for safe lossy
skips. Archive, XML, retained-text, projection, table, and diagnostic limits
fail closed before a partial document is published. The importer saves and
reopens its candidate through the typed Word package model before returning it.

The same direct lookup rule covers document tables and paragraphs nested in
table cells. `Document::table` and `Document::table_mut` are total, and cell
handles provide paragraph counts plus immutable and mutable lookup. Run and
paragraph formatting expose direct `Option<bool>` values and clear-capable
setters, preserving the distinction between inherited, explicitly false, and
explicitly true formatting without bypassing the facade.
The binding-only underline variants travel through a bounded integer-code
accessor so the published exhaustive Rust `UnderlineStyle` enum stays stable.

Low-level content-control traversal is recursive and ordered. Body, table,
row, cell, and paragraph accessors expose each wrapped ordinary paragraph,
table, row, cell, and run once while retaining the surrounding `CT_Sdt` for
metadata lookup. The facade consumes this single WordprocessingML ownership
tree and does not maintain a second content-control representation.

`Document::body_items` exposes the direct body ownership vector without
flattening it. Its borrowed items distinguish paragraphs, tables, body-level
content controls, and preserved unsupported XML in exact source order. The
recursive `paragraphs()` and `tables()` accessors keep their existing behavior.
Self-closing Word paragraphs and tables normalize to typed empty values, while
self-closing final section properties remain outside the item vector. Empty
foreign and unsupported children remain captured raw rather than being lost.

The same ordered compatibility view extends through `CellRef::items`,
`ParagraphRef::items`, `HyperlinkRef::items`, and `RunRef::items`. Their
non-exhaustive borrowed item enums retain each typed child and unsupported raw
subtree at its direct source boundary, including borrowed drawing and field
facts. Existing flattened run, paragraph, table, and body accessors keep their
established semantics. `Document::body_content` reports unsupported modeled
content through `UnsupportedXmlRef` name, namespace, and child-content facts,
while exposing raw bytes only when the facade owns an actual preserved raw
subtree. Save replays the namespace scope required by retained raw content and
fails closed when a safe prefix binding cannot be maintained.

The `rdocx` facade also owns RTF import and export. The private RTF reader
parses the Word-written subset for text, run and paragraph formatting, tables,
lists, and PNG or JPEG pictures directly into the same `Document` ownership
tree that DOCX opens use. Its scanner and destination stack stay inside
`rdocx`, while media bytes flow through `oxml-media` for sniffing and intrinsic
size. The private RTF writer walks the same typed document and package media
state without flushing or rewriting the DOCX package. It allocates font,
colour, list, and image references deterministically, emits RTF header tables
before body content, writes formatting resets at paragraph, run, cell, and row
boundaries, and serializes non-ASCII text as signed UTF-16 `\uN` code units
with surrogate pairs where needed. Safe lossy skips return stable
`RtfDiagnostic` records, malformed reader state fails through the facade error
enum, and writer output is bounded before retained byte growth or picture hex
expansion. RTF path saves serialize first, stage a same-directory temporary
file, sync that file, and publish through the shared portable replacement
helper.

Revision traversal follows that ownership tree through the main body, tables,
cells, and content controls. `Document::revisions` reports every valid modeled
revision once in document order as a borrowed `RevisionRef`. The facade does
not copy or reparse the raw subtree, and revisions outside the main document
part remain outside this traversal.

Revision mutation uses explicit all, exact-author, inclusive RFC 3339 instant,
and id selectors. One id operation resolves every modeled element carrying the
id, while undated revisions do not match date ranges. Invalid bounds or a
malformed selected revision leave the typed document, package bytes, and both
layout caches unchanged. A successful operation invalidates layout once.

Word comment mutation uses `RunPosition` and half-open `RunRange` values whose
body indexes select top-level paragraphs and whose run indexes select insertion
boundaries. `Document` validates both endpoints before mutation, allocates
collision-free comment and paragraph ids, updates the comment parts and all
three anchors together, then invalidates layout once. `CommentRef` is a
read-only view over the typed comment and its comments-extended thread entry.
Replies follow paragraph-id parent linkage, resolution applies to the thread
root, and removal deletes the selected comment plus descendant replies without
deleting unrelated runs or producer XML.

Word bookmark mutation input reuses the same top-level `RunPosition` and
half-open `RunRange` boundary as comments. `Document::bookmarks` returns
immutable correlated summaries in typed main-story paragraph order through
tables and block content controls. A reported body index is that recursive
paragraph ordinal, and its run index is the accepted-view boundary used to
extract bookmark text. Marker encounter order resolves direction when start
and end share one accepted boundary, so end before start remains reversed and
start before end is a valid empty range. Isolated projection refresh after a
run, comment, or bookmark edit carries the original Word namespace aliases.
The facade reports malformed, unmatched, reversed, or duplicate markers
without hiding their preserved XML. `Document::add_bookmark`
validates both mutation endpoints and the Word name, rejects producer-reserved
and duplicate names, allocates the first free nonnegative id, stages both
marker insertions, commits once, and invalidates layout once.

The `rpptx` facade provides the same total lookup boundary for slides, nested
shape trees, placeholders, text frames, paragraphs, regular runs, tables and
cells. Consuming mutable accessors transfer a facade borrow into its nested
handle, which lets the Python binding re-resolve a path without exposing
PresentationML internals or storing a Rust borrow in a pyclass.

SmartArt follows the same ownership split. `rpptx-oxml::diagram` owns the five
namespace-aware diagram part models and raw-preserving relationship payload.
`oxml-opc` owns the standard diagram relationship constants and the Microsoft
cached-drawing relationship constant. The `rpptx` facade alone resolves those
parts in the producing slide, layout, or master scope, stages checked node-text
edits, and copies the bounded owned graph for duplication or explicitly
layout-bound cross-presentation transfer. No diagram model depends back on the
facade, layout resolver, or renderer.

For rendering, the same facade projects the six exact pinned authentic layout
resources into transient ordinary PresentationML groups in the producing
slide, layout, or master clone. A bounded private evaluator consumes the typed,
doc-hidden layout and colour render data from `rpptx-oxml`, validates the
authoritative data and presentation graphs, and fails closed for every other
instruction program. The unchanged layout resolver and renderers then apply
their shared geometry, text, paint, effect, clipping, timeline, and media
paths. No cached drawing becomes authoritative and no persistent second
diagram model is created.

Notes-page and audience-handout export follow the same facade-owned assembly
boundary. `rpptx` resolves the notes slide, notes master, handout master, and
their themes from the OPC relationship graph, remaps notes-master and
notes-slide relationship scopes into one collision-free transient package
scope, and composes ordinary `PageFrame` values. The existing layout and render
crates consume those frames without a notes-specific parser, renderer, public
type, or dependency.

The facade also owns package-to-render-input assembly. Its deterministic render
entry points resolve the current package once and return either the shared
render input and layout or a complete PDF. The corpus example and
`rpptx-wasm` call that boundary, so neither binding nor development tooling
maintains a second PresentationML package interpretation path.
The CLI image paths reuse that resolved layout once per command and pass the
selected slide order to the shared raster backend. Separate PNG and JPEG files
are encoded one selected slide at a time, while TIFF remains one multi-page
stream.

Every consuming formatting builder on `Paragraph`, `Run`, `Table`, `Row`, and
`Cell` has a non-consuming `set_*` twin because a `mut self -> Self` builder
cannot back a Python property setter. The 61 consuming builders delegate to
their setter twins, so Rust callers retain chaining while borrowed handles and
Python properties use in-place mutation.
