# 03, Architecture

## Numbering-aware navigation sources

Navigation reads the production paragraph cascade and its resolved numbering
from one immutable Word layout result. Accepted body, cell and nested-control
headings retain physical WordSourcePath identity, including unplaced eligible
headings. Result-local IDs are not interchangeable across snapshots.

The facade's existing generated-table scanner supplies accepted-run exclusion
facts for owned TOC, figure, authority and bibliography result caches. Ordinary
CITATION is not an outer generated owner. Private facts travel on the existing
WordSequenceSnapshot with exact current document and physical main-part stamps.
Fresh sequence evaluation regenerates values and IDs, transferring only
validated exclusions. Stale facts refuse instead of changing another source.
Genuine text beside a same-paragraph result remains eligible. This bridge adds
no public input fields, second counter, XML parser or navigation registry.
The scanner's disjoint physical order bounds each paragraph's owner window.
Per-run end-position search narrows that window to one candidate for the existing
ownership predicate, including many disjoint owners inside one paragraph.

An ambiguous generated-owner scan becomes a private stamped failure fact on
WordSequenceSnapshot through the narrow doc-hidden
`set_navigation_failure(&mut self, &LayoutInput, String) -> Result<()>` bridge.
Fresh sequence regeneration preserves only the exact-current-document and
physical-part-qualified failure. Ordinary rendering keeps existing stored cache
paint and counter behavior and reports an existing layout Diagnostic. The new
numbered outline refuses through `WordLayoutResult::navigation_failure() ->
Option<&str>`, and newly supported STYLEREF retains cache with an ordered
diagnostic. Explicit ambiguous rebuilds retain their existing atomic refusal.
This is not an empty successful source discovery or a new ordinary REF/SEQ
policy. Snapshot equality includes the private status and invalidates caches.

TOC selectors reuse the existing source-qualified navigation tuple. Explicit
custom-style selection remains independent. Heading-style selection without
u keeps its established rule, while u requires the qualified shared heading
eligibility, including effective inherited or direct outline9 suppression.
An unavailable owner is an explicit refusal, not a fabricated suppression fact.
No additional cascade, flag, counter or per-field source scan is introduced.

## Owned publication and causal paths

`oxml-py-support` owns captured path authority, owner scopes and prepared
revision publication. It does not own native presentation objects or keep a
second handle registry. Word bindings prepare a finite inventory of current
physical paragraph, table, cell, section and captured hyperlink owners.
Supported publication retains only proven surviving owners. Regular runs keep
the conservative document epoch, while frozen Story and StoryItem captures
also retain snapshot-inventory authority.
Word owner inventories share source, owner and item qualification across full
item, cell, paragraph XML and hyperlink projections in one operation. They
materialize each source once and pair qualified physical owners with one traversal of each parsed outer table and its nested models.
Physical coordinates retain row and cell vectors rather than logical grid
columns. Candidate revision matching indexes current owner routes and retains
all admissible parent matches until descendant matching completes. Body-slot
and projection fallback classes remain distinct because their pairwise match
relation is not transitive. Earliest matching old ownership determines retained
scopes. No historical owner registry is introduced. The binding keeps the
inventory by body block and story part, each piece with the value it was read
from: a block, read from a main part whose body holds only that block, or a
related part's bytes. The next inventory reuses a piece only where the
document still holds an equal value. For a block that is the equality of
`BodyContent`, which keeps its owners only while the manual equality of
fields, tab stops and bookmark markers ignores nothing that changes them.
Blocks keep their pieces over the common prefix and suffix of the body, and
the blocks between are read again. Parts are compared one by one by their
bytes. Block pieces are placed by offsets of body story items, table cell and
text box stories, and tables. Another part root or namespace fact of the
opened package reads every block again. A part read twice, and form and
signing publications, use the whole-document inventory without reading a
piece. So does a block that does not read as one direct body item, once read.
The body story of a piece has no fingerprint, since that covers the whole
body, and handles refresh a story before native use.
Namespace scans retain a frame per element and copy inherited maps only when
the element declares a namespace. Attribute validation and scope restoration
retain their checked source semantics.

Presentation bindings capture concrete slide, shape, table-grid and text-owner
scopes, then locate retained physical owners in their current source graph.
Manual native snapshot constructors remain unbound and visible snapshot
values retain their existing equality semantics.

Table rows and columns collections bind their surviving physical table owner.
Grid topology edits retire captured indexed rows, columns and cells while
preserving those collections and the table and shape ancestors. Removing the
parent owner retires its collections too.

Bindings stage a native candidate and validate all conversions, source graph
constraints and returned Python allocations before committing prepared scopes.
The final native assignment cannot fail after publication. A rejected edit or
an actual no-op leaves source, scope generations and live handles unchanged.
First retirement records the exact public operation label for later stale
errors. The native facades retain ownership of relationships, geometry and
package authoring, while bindings reuse that publication protocol.

A plain direct body paragraph replacement or removal edits that one child.
A package that retains a shadowed serializer prefix does not take that local
path. The staged transaction then fails closed and leaves the bytes unchanged.
The prefix is recorded when the package is opened or saved, so the call does
not read the other body children. Admission binds `w`, `r`, `mc` and `wp`,
the prefixes `w:document` always declares. Stored `xmlns` records are
attribute names and do not replace those prefixes. A paragraph that is not
plain, a relationship change, or a removal whose media ownership is not
proved keeps the staged package transaction. `Document::remove_content_range`
also keeps that transaction. It removes a half-open range of direct body
children and does not recurse into tables. Modelled body children and
preserved raw children are stored in boxes, so a direct removal shifts
pointers.

Python Word owner lookup uses derived indexes over the current inventory.
Plain paragraph XML prepares only the selected content/property scopes and
its exact run owners before native mutation. An empty plain paragraph can
acquire direct runs locally, with newborn scopes, storage and route buckets
prepared before admission. New run owners append without shifting earlier
physical owners. Other run-topology changes retain staged reconciliation.
A plain tail removal can keep earlier entries when the suffix contains sections
or appended runs whose parents precede the removed owner. It updates those
suffix indices. Non-contiguous descendants of the removed owner retain the
conservative path. This is not a bound independent of prior local edits.
Every captured physical Word owner includes a removal epoch, so that shortcut
retires held handles without allocating successors for every paragraph. Fresh
handles capture the new epoch. Other removals retain conservative reconciliation.

Word read projections are prepared with that same immutable current-owner
inventory. The facade supplies batch body and cell-paragraph locations matching
its existing single lookups, and the inventory by block and part gives them
block by block. Missing or unsupported projections retain those lookups.
Prepared handles and the edited paragraph's publication route bind through
the same projections before those lookups. Route hashes index candidate
owners, with exact route comparison and first-match selection after hashing.
Hash equality alone grants no authority.
Accepted paragraph text and table row/cell paragraph counts are derived read
facts, never serialization sources. Full replacement and property relocation
prepare fresh facts. Local plain XML prepares its new accepted text before
mutation, and installs it only on successful publication. Tail removal retires
captured epochs and removes its projections without scanning unrelated owners.
Shared paragraph XML admission compares physical section properties for both
Start and Empty events, using a namespace-complete temporary semantic parse.
It refuses changed or introduced section boundaries on local and staged paths.
An admitted opaque Empty section that the paired parser cannot model remains
protected by exact bytes and expanded namespace meaning. Only an unchanged
instance may remain. Submitted and stored XML are not rewritten by comparison. Retained nested
modeled namespace owners require staged identity reconciliation, even for
benign aliases, except for a plain paragraph removal when every retained owner
is a `w:sdt` and their declarations replay on the body as a save serializes it.
The removal adds or removes no `w:sdt` candidate of that replay, so it stays
local. When the replay fails, or a serializer prefix is shadowed, the seven
binding body insertions that 10-bindings-spec names refuse with the save's
error before they change anything. The Rust API's insertions are not guarded.
A failed proof refuses before source publication.

Run tab and break edits prepare accepted paragraph text from the native candidate,
validate its current run-to-paragraph owner and publish source and fact together.
They preserve existing run, font and paragraph scopes. Cell edits resolve the
same physical paragraph location as immutable reads, including controlled
paragraphs, rather than using the direct-only mutable cell ordinal accessor.
Reads validate current scopes before using facts. Mutations retain native
fingerprint refresh and admission. There is no lazy first-handle package scan
or second owner registry.

PowerPoint edits one child for `add_textbox` on a connector-free slide, for
alt text, for non-connector geometry, for a non-blip shape fill, and for a
shape XML replacement whose relationship facts and owner signature stay the
same. A connector, a blip fill, a relationship change, a group refit, a
retained package signature, or a whole-part replacement keeps the staged
package transaction. These local edits do not read unrelated slides.
Shape-id allocation reads typed ids and preserved raw blobs. The whole tree
is serialized only when a preserved prefix other than `p` is not declared in
that blob.

Python PowerPoint also indexes current owner generations, routes and shape-id
multiplicity. Those indexes are derived, not a second ownership registry.
Appending a simple root textbox prepares three newborn owners and extends
the current inventory without cloning existing entries. Exact owned XML is a
no-op. Slide shape XML can use local publication only when a typed scalar
change witnesses actual mutation and its physical owner topology is unchanged.
Graphic frames also require unchanged graphicData. The shared address projector
enumerates the selected shape's descendants before native mutation, including
table grids and regular runs. Fields and breaks are not regular run owners.
Unproved XML, template owners and topology changes use a native candidate, with
all post-edit reads and owner preparation completed before publication.
Scalar property edits likewise require a pre-native change/no-op proof or use
a candidate. Non-blip shape fill prepares its property revision before native
mutation. Replacing the inventory rebuilds its indexes, while local publications
update only affected entries. No fallible discovery or staged fallback follows
a successful live native edit.

Word binding admission refuses PDF save paths, duplicate stored tab positions,
page breaks in actual physical headers and malformed authored language values
before publication or output. Header ancestry follows the relationship-resolved
native story graph, including cell descendants. Native font admission requires
the exact name directory entry before key derivation or package allocation.
Word font CLI mutation preflight composes identity and file-kind checks with
nonreplacement availability and staged publication. These checks preserve
source bytes and captured authority on refusal.

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
                     Length conversion helpers, shared PyO3 RGB input validation

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

**Python RGB input belongs to `oxml-py-support`.** One concrete
`rgb_channels(value: &Bound<PyAny>, parameter: &str) -> PyResult<(u8, u8, u8)>`
serves both binding crates. This unpublished support crate takes the existing
workspace PyO3 dependency without enabling `extension-module`. That feature
stays in `rdocx-py` and `rpptx-py`. The facades and both WASM graphs do not
consume the Python support crate. Each binding adapts the channels to its
existing native colour representation and owns its surface-specific `None`
and Word `auto` policy. Live colour formats keep document and content-path
ownership in their binding, not in a second document or colour model.

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
Native defaults and WASM do not select either security feature. Both Python
bindings and both CLIs explicitly select digital signatures. The two Python
consumers share concrete owned report conversion in oxml-py-support. The two
CLIs select the default-off oxml-cli-support report feature even when their
wrapper default features are disabled. Agile encryption remains unselected
by Python, WASM and CLI consumers.

**`oxml-media` has no dependencies at all.** It owns byte sniffing, image header
probing, intrinsic EMU sizing through its local `NativeSize` value, safe MIME
grammar, bounded media-container signature checks, and collision-safe naming.
It remains a leaf that anything can take cheaply without importing
`oxml-core`. The format facades depend on it directly for package media names,
sniffed package metadata, and byte-first MIME inputs. The container checks do
not claim a codec decoder.

Inbound CommonMark also belongs to the `rdocx` facade. The pinned
`pulldown-cmark` parser produces bounded events and an escaped HTML projection
for the existing importer. Ordered diagnostics retain original source byte
spans. Outbound Markdown derives from saved Word styles, numbering and content,
without storing an input Markdown copy.

**Inbound HTML belongs to the `rdocx` facade.** The private importer uses
`scraper` for HTML5 document and fragment tree repair, then projects supported
content directly into the one owned WordprocessingML document model. The edge
does not enter `rdocx-html`, which remains an outbound emitter. This avoids a
dependency cycle and avoids a second public intermediate document model.
Semantic `mark` highlighting carries its final cascaded background through
nested inline formatting. A child's own background takes precedence on that
child, while ordinary CSS backgrounds remain scoped to their own elements.
Nested marks choose their own highlight, and unmarked siblings inherit none.
`Document::insert_html_fragment` reuses that importer against a staged document
and emits concrete body content at a checked `ContentLocation`. Body and
main-document cell destinations mutate their typed owners so retained sibling
XML keeps its exact boundary. Other supported stories receive namespace-closed
fixed-prefix fragments. Numbering, media, hyperlinks, and drawing identifiers
are allocated in the destination story scope before the staged package reopens.
No resource callback, network access, or second fragment model is introduced.

**Word package identifiers belong to the `rdocx` facade.** One private
`DocumentIdentifiers` value owns category-specific occupied sets for each
document. It scans normalized package parts, relationship owners, typed
WordprocessingML, preserved XML, and content types. Part-name and resolved
relationship-target occupancy uses ASCII-case-insensitive package identity
while package maps retain producer spelling. Facade operations reserve
complete identifier bundles on staged state, while serialization derives
authored identifiers from final recursive document order before publishing the
candidate. Relationship identifiers remain scoped to their owning part, and
producer drawing-definition uniqueness is also validated within each physical
XML part. The union of every valid producer drawing value remains occupied for
package-global authored allocation. Only identities captured at package open
have preserved provenance. Current graph edges added later join authored
semantic canonicalization even when no modeled `r:id` refers to them. Internal
and external edges retain their targets and modes while receiving deterministic
type-and-target order. Relationship attributes in retained body XML and raw
drawing payloads are fixed occupants.
Their values are found by expanded office relationship namespace, including
bindings inherited by the document body, and are never remapped without a safe
raw rewrite.

Story-scoped image and hyperlink operations resolve the normalized OPC owner
from `StoryId`. Cells and text boxes inherit the containing part, while
headers, footers, notes, and comments use their related parts. Main-part
occurrence provenance reconciles against live expanded-name references during
serialization. Repeated same-owner references share one relationship, removed
references retire only occurrence provenance, and each live authored picture
receives a distinct globally canonical drawing identity. Producer drawings may
reuse one preserved identity in different physical parts without being
renumbered.

Bookmark, comment, drawing, abstract-numbering, and numbering-instance values
remain separate namespaces. Rich-merge content-control `w:id` and non-visual
drawing `cNvPr` values remain separate merge-local scopes rather than package
identifier-owner categories.

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

Word picture replacement stays at the same facade and package boundary.
`rdocx` resolves one exact body or story relationship and `oxml-media` sniffs
the replacement format. An unshared target with a compatible extension may be
reused. A shared target or format change allocates a new media part and changes
only the selected relationship target. The drawing XML and relationship
identifier remain stable. The live document receives only a serialized and
reopened candidate.

Deterministic animation export also belongs to the `rpptx` facade. It validates
and samples explicit segments, prepares the package, resolver, font, chart,
picture, and media state once, then evaluates and lowers one timeline sample at
a time. The completed opaque raster immediately enters either the facade-owned
GIF encoder or the Motion JPEG AVI writer before the next sample is resolved.
`jpeg-encoder` is an optional dependency of the existing `rpptx` render
feature. The facade also uses `gif` and `tiny-skia` unconditionally for bounded
image validation when rendering is disabled. No codec edge enters
`rpptx-layout`, `rpptx-render`, or an `oxml-*` crate, and audio and video package
payloads are still never decoded.

The same boundary owns multilingual text mechanics. Rich segments retain
logical text and source ranges while carrying script, language, direction,
bidi level, glyph clusters, and two-dimensional advances and offsets. Script
and coverage segmentation, ICU break opportunities, conditional hyphenation,
HarfRust shaping, and line-local UAX 9 visual ordering therefore remain shared
by document formats. Logical order is the extraction contract. Visual order is
applied only to completed lines for painting.

The shared PDF backend owns the final extraction projection. It groups only
adjacent rich runs with one semantic owner, one transformed baseline, and
contiguous source spans or logical indices. The first painted run carries the
complete logical line through `ActualText`, and later painted runs carry empty
replacement text so extraction sees the line once. The backend retains paint
traversal and exact glyph positions, including through nested group transforms.

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
callers can also install byte-free family aliases. Within the caller tier, resolution prefers an exact
sfnt family, then an explicit caller alias, then a label-derived alias,
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
Bibliography source authoring and formatting stay in the native
`rdocx::bibliography` module. It owns qualified custom XML source inspection,
checked source CRUD, document style/options and concrete measured citation and
bibliography consumers. `rdocx::field` owns physical field discovery, recursive
instruction spans, result replacement and structured paragraph/table cache
boundaries. The facade stages a complete document and publishes only after
source graph and output preparation succeed. An unfinished standard formatter
aborts that transaction. Noncatalogue paths retain the owner with diagnostics.
Normal save preserves caches and does not invoke this materializer.

Source kind, contributor role and style metadata do not imply formatter parity.
APA has 210 earned ordinary dense locale configurations with separately measured
sparse, source-script and rich operations. Other bibliography styles retain
lean numeric1033 Book admissions. Citation consumers remain separate. Actual
source LCID, field locale and explicit application default select existing
consumers without an en-US fallback. Multi-source non-ASCII sort keys refuse.
The remaining catalogue and Word comparator are F-X192 backlog work, with no
new collation dependency in this boundary.

The bibliography table-width caller shares `rdocx_layout::engine::resolve_font_family`
with the existing layout callers. Its public inputs must already contain the
effective cascaded run properties. Font-slot and theme resolution retain the
same concrete algorithm and deterministic font manager. Rich bibliography
runs carry the measured font, language, script, italic and direction facts.
They do not introduce a general Unicode font-coverage classifier.


Generated-table source discovery and rebuilding stay in `rdocx::field`.
Concrete XE and TA markers, INDEX, caption-selected TOC and numbered TOA options
use the existing checked story inventory and shared dynamic owned-span scanner.
The shared instruction grammar in `rdocx-oxml::text` owns bare and quoted
operands, including INDEX LCIDs and TA/TOA categories. It distinguishes INDEX
run-in and TOA heading flags from operands without changing other opcode rules.
Source discovery excludes every generated cache and correlates accepted marker
and caption paragraphs with their exact retained OPC owner. Source SEQ values
come from the existing physical sequence snapshot rather than another counter
interpreter. Existing `rebuild_toc` keeps its compatibility contract.

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

`CT_PPr` types the complete authored paragraph-property set. Text frame
placement, line-number suppression, document-grid right indentation, contextual
spacing, mirrored indentation, frame overlap suppression, text flow, vertical
character alignment, text box tight wrap, and the web-settings division all
have typed members beside the borders, shading, tabs, indentation, spacing, and
paragraph-mark state that were already modeled. Every remaining `w:pPr` child
keeps its original schema slot. A newly typed toggle retains its source element
as an attribute carrier, so an attribute the model does not own survives being
modeled. `CT_FramePr` and `CT_BorderEdge` retain their unmodelled attributes in
source order and replay them ahead of the modeled ones. `CT_PPr` stores its
frame and its borders behind a `Box`, which keeps the value small enough for
recursive callers that carry paragraph properties on the stack.

`CT_RPr` types the complete `EG_RPrBase` sequence. Outline, shadow, emboss,
imprint, proofing suppression, grid snapping, web-only hiding, kerning, the
animated text effect, the character border, fitted text, the complex-script
toggle, the emphasis mark, East Asian run layout, numbered-paragraph vanish,
and the Office Math flag all have typed members beside the fonts, sizes,
colour, shading, language, and direction state that were already modeled.
`w:rFonts` owns all four explicit script slots, all four theme slots, and
`w:hint`, and `w:color` owns its theme colour, tint, and shade, so nothing on
either element is dropped on save. `CT_Shd` owns the same six theme attributes
for run, paragraph, and table shading. Like `CT_FramePr` and `CT_BorderEdge`,
each of these retains its unmodelled attributes in source order and replays
them ahead of the modeled ones. `CT_RPr` stores its character border, its
fitted-text and East Asian layout members, its shading, and its property change
behind a `Box`, and `CT_PPr` boxes its shading for the same reason, so the
grown model still fits the stack budget the F-084 precedent set.

Ordered run content types `w:sym` and, behind one `SpecialCharacter` value,
`w:cr`, `w:noBreakHyphen`, `w:softHyphen`, and `w:ptab`. One variant for the
four keeps the number of cases every consumer must match on flat.
`w:lastRenderedPageBreak` is a producer hint, so it stays in positioned raw
capture with a read-only classification and no authoring surface. A symbol
whose code point is not four hex digits, a positional tab missing one of its
three required attributes, and an empty special character carrying an attribute
all stay raw rather than losing bytes to a partial projection.

The low-level text reader decodes visible `w:t` and `w:delText` content
fallibly and rejects malformed encoded values instead of publishing partial
text. The numbering grammar types the complete standard `w:numFmt` token set
and retains producer-defined tokens in `ST_NumberFormat::Other(String)`.
`CT_Lvl` owns typed restart, legal-numbering, template, tentative, style-link,
indentation, marker-property, and sequence-ranked raw state. `CT_NumLvl` owns
typed start and replacement-level overrides. Their writers emit standard and
producer-defined values unchanged while render and export consumers decline to
invent a marker when they do not implement a format's visible semantics.

The `Document` facade projects abstract definitions and numbering instances as
owned concrete values. Create, update, and remove operations validate level
ranges, placeholders, identifiers, override ownership, style references, and
the complete candidate graph before committing the staged package. Standalone
style-link mutation is forbidden. `link_style_to_numbering` and
`unlink_style_from_numbering` stage the complete document, validate both
graphs, write the style paragraph properties and effective numbering-level
style link together, flush the candidate package, and publish once. A missing
or conflicting style, instance, level, or reciprocal edge leaves the original
document unchanged.

The layout counter owner keys visible state by concrete `numId` and level. A
reused instance continues its sequence across ordinary body paragraphs, table
cells, and section boundaries unless its definition requests a section
restart. A distinct instance starts independently even when it shares an
abstract definition. Base-level starts and restart controls, instance start
overrides, replacement-level formatting, legal numbering, and `numId` zero
suppression enter the
same result-local projection consumed by marker layout, TOC rebuild, and REF
evaluation.

The same grammar owns the bounded `w:ffData` projection on complex legacy form
fields and the `w:glossaryDocument` root model. Typed form values and glossary
properties are namespace aware, while retained XML remains the serialization
source for every unsupported attribute and subtree. The `rdocx` facade owns
relationship resolution, story-part identity, staged validation, and package
commit for form values and the glossary lifecycle. Structural glossary edits
retain untouched entry spans, container attributes and unsupported siblings.
Glossary capture and insertion share the document-fragment dependency
transaction, using the physical glossary part as relationship owner.

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

`SUPPORTED_SETTINGS` closes that model. It names, in `xsd:sequence` order, the
thirty-one top-level `w:settings` children the typed model owns, and the nested
supported names are `compat/compatSetting`, every `CompatibilityOption` local
name under `compat`, `docVars/docVar` and the bounded `mailMerge` members. One
`SETTINGS_ORDER` table drives both schema-position insertion and that closed
set, so a new member is added in one place rather than three. Every other
child, including `attachedTemplate`, `rsids`, the drawing-grid family and every
`w14:` or `w15:` extension, stays preservation-only and is never a diagnostic.
`CT_Settings::diagnostics` reports each supported child the model could not own,
as `Duplicated` when more than one occurrence exists and as `Malformed` when a
single occurrence did not parse. A package authored entirely through the public
surface reports nothing, which is what makes "no unmodeled supported children"
decidable rather than open-ended. Removal refuses a duplicated or malformed
member instead of reinterpreting one of several conflicting occurrences.

Document-protection authoring records caller-supplied metadata verbatim.
Deriving a hash from a password, choosing a salt and picking a spin count are
explicit non-goals, so nothing in this path introduces fresh cryptographic
randomness.

The web settings model owns the separate `w:webSettings` root with the same
source-preserving architecture. `w:frameset` and `w:divs` stay
preservation-only, because the first carries relationship targets and the
second is a nested tree no story asked for. `CT_WebSettings::div_ids` projects
the retained `w:divs` subtree read-only, so a caller can tell whether a
`w:divId` resolves to a division the part declares without opening HTML
division authoring.

The document `w:defaultTabStop` reaches layout. `LayoutInput::default_tab_stop`
carries it into `oxml-layout::LineBreakParams::default_tab_interval_pt`, whose
`36.0` default is the exact literal it replaced, so a document that says
nothing keeps Word's half-inch implicit tab grid.

Tab stops resolve as Word resolves them. Positions are measured from the zero
indent, not from where a line starts. A paragraph's `w:tabs` add to the stops
of its styles and list level, a stop replacing an inherited one at its
position and a `clear` stop removing it. A hanging indent is a left stop at
the left indent. Bar stops stop no tab. Default stops start after the last
explicit stop. A right, centre or decimal tab takes its width from the text
after it, a decimal tab aligning the first full stop or else the end of the
first number. `LayoutInput::clamp_tabs_past_margin`, set for a
`w:compatibilityMode` of 15 or later, moves a stop past the right margin to
the end of its line as Word 2013 does. Earlier modes keep the stop.
The page-field tab adjustment crosses the shared layout boundary as
`TabAlignedField` on `GlyphRun`. `LineBreakParams` and `LayoutInput` also expose
`clamp_tabs_past_margin`. These additions are source compatible for callers
using the existing defaults and methods. Existing public struct literals must
supply the new fields when updating to these pre-1 crate versions.

Two consecutive paragraphs of the body, or of one table cell, are separated by
the larger of the first one's space after and the second one's space before,
which is what Word does. The paginator reduces a paragraph's space before by
the space after already above the cursor, because body paragraph blocks are
cached and must not carry a value that depends on their neighbour. Table
layout reduces the space before of a cell paragraph that follows another, so a
row height stays a plain sum. Spacing that meets across a table edge or between
two rows is added. `w:compat/w:doNotUseHTMLParagraphAutoSpacing` restores the
sum everywhere and reaches layout as
`LayoutInput::do_not_use_html_paragraph_auto_spacing`. Space before stays
suppressed at the top of every page.

The comments model owns typed comment entries and the three body anchor forms.
Comment bodies retain ordered paragraphs, producer attributes, and unmodelled
children. Paragraph and run models retain each anchor at its insertion boundary
without moving neighbouring raw XML. Parsing accepts in-scope aliases for the
WordprocessingML namespace, while serialization uses the fixed `w:` prefix.
The comments-extended model owns paragraph-id parent linkage and resolved state,
with unmodelled attributes and root children retained at their original
boundaries. The `rdocx` facade owns the relationship-resolved pair of comment
parts and coordinates them with the anchors in the main document.
Authored and parsed comment ids are stable facade identities across rdocx save
and reopen. Allocation takes the lowest unused nonnegative id and never
reorders existing comments to canonicalize ids. Dates are optional RFC 3339
values validated before staging. An absent date remains the deterministic
default. Third-party editors may independently renumber comment identities.

The Word text model projects bookmark starts and ends at accepted-view and
tracked-view run boundaries while retaining every marker as ordered raw XML.
The projection combines direct paragraph markers with markers owned by typed
hyperlinks, accepted insertion and move-destination revisions, and inline
content controls in exact document order. Each nested content control retains
its own namespace scope for this projection. Opaque wrappers remain excluded,
and direct markers are not duplicated. Complex-field collapse remaps both run
views. Fields that share a physical run form one span, so a run where one field
ends and the next begins keeps both. Visible content that shares a physical run
with a complex field, such as the text of `Page {PAGE} of the report`, is read
as sibling runs in source order around the field, so such a paragraph reports
more runs than it has physical runs. The span writes its original bytes while
those runs and its fields are unchanged. Otherwise each field writes its own
part of the span, so one field of a shared run can be updated alone. Direct-run and marker mutation rebuild the same read projection in
memory. Simple and complex fields share one recursive
`Field` grammar with a normalized name, text or
nested arguments, switches, cached result, and optional dirty and lock states. Its private
source records the original field form, run partition, and producer XML.
Complex fields expose ordered cached-display segments with each segment's
direct run properties. Native callers construct either representation through
checked raw instructions or typed arguments and switches, using `CT_R` values
for ordered cached content. Typed operands cannot inject instruction tokens.
Unknown switch operands are quoted on write and recognized as quoted or nested
operands on read. Known flag switches keep their positional-operand grammar.
This enriches the pre-1.0 projection of ambiguous unknown producer switches
without changing preserved producer bytes. Cached page and column breaks remain
typed XML controls rather than literal XML-invalid display characters. A nested
field replacement invalidates the unchanged-source shortcut even when its
display text stays equal. Lock edits rewrite only the physical simple owner or
complex begin marker, retaining opaque subtrees and namespace bindings.
Nested simple-cache projection reuses the producer's effective Word alias without
replacing an occupied prefix. Cache writes retain aliased owner bindings and source text aliases. Only
source-less generated result elements receive a local canonical binding. Foreign cached subtrees keep
their original bytes and never become typed pagination identities.
 Tracked insertion projection retains inline paragraph
structure and nested revision boundaries, with a fixed depth ceiling checked
before recursive projection.
The mutable native run facade owns one logical ordered sequence of text, tabs,
typed breaks, inline pictures, fields, and Unicode symbols. A field remains a
paragraph child in physical WordprocessingML. Serialization therefore divides
the surrounding logical run into schema-valid physical runs, copies the direct
run properties to each authored segment and field result, and preserves the
logical child order on reopen.
Unchanged fields therefore write their original bytes. Cache and dirty updates
rewrite only the typed values while preserving run formatting and unmodelled
neighbours. Markers are recognized only as direct run children through their
in-scope WordprocessingML namespace bindings. Malformed sequences remain opaque
raw XML, while unsupported valid fields retain their cached display. Dirty
complex hyperlinks are not reported as `Document::links()` until the update
policy defines how to handle them. The `rdocx` facade correlates bookmark ids
and owns mutation across typed body paragraphs, including supported table and
content-control traversal. Its story-qualified range inventory also pairs
bookmarks, comment endpoints, permission ranges and proofing ranges within one
physical owner. Checked add, move and remove operations stage a candidate,
validate accepted-view order and reopen the result before publication.
Proofing pairs carry their spelling or grammar type without an invented id.
Removing a comment range leaves its reference run and definition as a point
comment. Moving a comment range moves its reference run with the pair.
`rdocx-layout` resolves bookmark text and maps page targets, while the shared
`oxml-layout` boundary exposes only format-neutral `Target` and `TargetPage`
field kinds.

The `rdocx` facade owns pure field evaluation over that recursive grammar. It
registers physical typed field owners in main text, tables, content controls,
distinct header and footer parts, notes, original annotations and selected
text boxes. SEQ and REF use this expanded source scope. Other field kinds keep
their established discovery scope, including direct related-story paragraphs.
Hidden ordinary owners retain cache and dirty attributes. Public result indices
number only visible evaluation records and are distinct from physical source
IDs. Accepted revision selection uses the same initialized borrowed physical
projection as layout, retaining deleted source IDs without evaluating them.
Package-backed inputs come
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
the same payload and case-style rules as `ITF14`. Mail-merge record state remains isolated by story. Its record and output
sequence numbers come only from the explicit caller context. SEQ uses one
font-independent accepted-source snapshot owned by `rdocx-layout`. The facade
and renderer consume its indexed counter events, field values and physical
source identities. Selected text boxes participate in main-document order.
Related-story increments produce the measured main-document-only error.
Source-qualified note repeats use their document context, while shared
furniture repeats resolve against actual page event boundaries. Pure evaluation
retains caches when physical context is unavailable. Tracked views retain SEQ
caches with diagnostics rather than joining accepted indices to tracked runs.
Opaque fallback XML does not manufacture sequence sources. Producer instructions with unclosed
quoting retain their cache and contribute no sequence event. The shared XML
instruction quote validator honors quoted escapes before snapshot mutation.

Snapshot equality includes retained field XML, effective instructions, inherited
locks and selected owner projections before layout cache reuse. Rich selected
text boxes bind exact physical source paths before shaping and restore local
indices through the same injective map for the existing outer binding. Generated
cache descendants cannot increment counters. Instruction descendants emit
zero-width structural events at their owning accepted run. REF number and
position projections also use the same source-only registration when no SEQ
exists, so selected owners cannot alias local body indices.

Ordinary REF target text and numbering use uniquely paired physical bookmark
owners in the main story, headers, footers, notes and selected text boxes.
The layout result exposes the existing resolved numbering for a uniquely paired
bookmark, including nested table and control owners, without facade source-path
reconstruction. Duplicate physical names and malformed ranges have no projection.
Relative position uses the same registered physical owner, accepted paragraph
path and actual run boundary in the shared source-only snapshot. Header and
footer aliases share a physical part identity, while notes and selected boxes
retain their actual owner identity. A field inside its target, an ambiguous range or an unavailable
cross-story comparison retains the complete cache with a diagnostic, including
combined numbering and position requests.

The native facade authors checked captions, sequence fields and REF fields at
accepted story boundaries. Caption insertion allocates separate whole-caption,
label-and-number and number bookmark ranges atomically. REF numbering shares
one resolved-numbering formatter with layout, including current-level, relative,
full-context, text omission and current-leaf delimiter behavior.

REF note and annotation copies freeze original source ownership before staged
allocation. Typed cache runs and paragraph-sibling comment endpoints preserve
rich payload and schema order. Unique cached copy owners may refresh from the
original source without accumulating note, comment or companion graph entries.
Ambiguous ownership or unsupported companion edges retain source caches and
ordered diagnostics. Optional comment metadata parts remain absent when absent.
Unmodelled XML and existing relationship scope remain authoritative.

The facade also owns explicit field cache updates across that same typed story
scope. It evaluates the complete field set before changing cloned document and
package-backed parts. Resolved values replace the stored display and clear the
field-local dirty flag. Pagination deferrals, structured non-text outcomes,
and stored-display fallbacks keep their cache and become dirty so Word may
retry them. Only validated staged XML
is committed, then both layout caches are invalidated once. Existing save and
byte methods remain leave alone operations that preserve cache content and
dirty spelling. Update-aware save methods opt into the same atomic operation
before writing. The settings-level `w:updateFields` value remains untouched by
these updates. Callers read, set, or remove it explicitly through
`Document::update_fields_on_open` and `set_update_fields_on_open`. A duplicate
or malformed producer value stays unmodelled and byte-preserved, and an
explicit mutation rejects that ambiguous ownership before publication. The two
methods are additive pre-1.0 `rdocx` API.

`Document::update_layout_backed_fields` materializes PAGE, NUMPAGES, SECTION,
SECTIONPAGES and resolved PAGEREF through one staged deterministic layout.
Its immutable snapshot retains physical pages, displayed numbering, every section
occupying a continuous page, bookmark target ownership and recursive field
placements. Field indices follow preorder through instructions and cached results.
Cached children have their own display placement. Instruction operands inherit
only their owning field's placement. Locks propagate to descendants.

Complete story-range discovery materializes each physical source once per operation.
Owner, control and namespace scopes are discovered in batches, then paragraph
excerpts reuse checked source offsets. Cache projection shares those owned
source/span facts. The evaluator reuses one paragraph listing for range pairing
and bookmark-name counts, with independent failure handling for each. This
operation-local inventory adds no global cache or revision invalidation state.

Physical source identity covers modeled paragraphs, tables, controls, headers,
footers, notes and selected typed text boxes. Actual OPC names remain distinct
from logical source paths. Rich related-story and text-box bodies replace their
paragraph-only projections when present. Text-box source owner ordinals bind
from actual namespace-qualified anchor occurrences before shaping, including
identical anchors and earlier unregistered owners. Original DrawingML and opaque
fallbacks remain serialization authorities. Cache staging validates the exact
field source inside the matched physical owner before publishing atomically.
Physical binding traverses all revision and raw owner slots before accepted
selection. Opaque occurrences consume identity without becoming editable stories.
Accepted revision projection metadata survives without serialization and reparsing.
Authoritative rich story bodies and actual part names participate in reusable
context and header/footer cache identity, including note-only invalidation.


The report separates five counts and retains ordered diagnostics. The count-only
`update_page_fields` wrapper uses the same operation. PAGE follows its field's
section format and PAGEREF follows the target's format. A unique placed target
permits unlocked PAGEREF updates in unused source stories. Other unplaced fields,
missing or ambiguous targets, unsupported switches and locked fields retain their
caches. Saved header and footer PAGE and NUMPAGES caches remain unchanged under
the pinned Word policy, while dynamic rendering uses per-page values. Note PAGE
and SECTION follow the unique body reference owner even when document-end note
text is printed on a later physical page. No second layout resolves these facts.

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
Unsupported valid TOCs stay unchanged and append their exact retained-display
message to the report in physical source order. Simple TOCs contribute the
same ordered report without mutation. The compatibility diagnostic count is
derived from that collection. Malformed ownership, ambiguous bookmarks,
missing selected bookmarks,
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
including values held in preserved raw XML. The same pass removes `w14:paraId`
and `w14:textId` from every copied paragraph and table row, since a copy must
not share them and Word assigns new ones, and keeps revision-save identities.
Rich merge region copies and imported fragments go through it too. The
operation does not evaluate structured template tags, and ordinary field
traversal keeps its existing typed story scope.

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
banding, content-control, and ordered raw XML state. Every paragraph and table
row a loop renders drops its `w14:paraId` and `w14:textId`, as other copies do,
and keeps its revision-save identities. Numbered paragraphs in a
loop retain their source `numId` and level, which keeps one continuous list
without allocating definitions. Numbering references are validated before
evaluation. Loop variables form lexical scopes, and dotted lookup searches the
innermost scope before the root value. Both grammars defer explicit item paths
and unknown implicit item properties while preflighting an empty loop. Known
outer or root properties still validate their complete path and intermediate
container kinds. Structural paths and scalar leaves use the same deferred
scope policy. Grammar 1 is selected by `render_template_versioned` and by the native
formatter. `render_template` keeps the legacy property-key grammar. Structural
controls pair inside one physical owner. That owner may be the main body, a
table cell, a header, a footer, one normal footnote or endnote, or one checked
text box. Markers do not pair across owners, notes, cells, field boundaries or
alternate text-box representations. A shared header or footer is evaluated once
per physical part. Note separators and continuation separators are excluded.
Qualified equivalent text-box representations form one logical program and are
evaluated once. Chart labels stay scalar. Template preparation may ask the
existing story-item scan for owner-scoped field intervals. Ordinary callers do
not. Structural programs inside or crossing those intervals refuse before
callbacks, and the paragraph-local field projection stays unchanged. Repeated
ranges reuse the staged fragment transaction so copied bookmarks, controls,
drawings, notes, comments and relationship operands stay unique. Original
callback locations remain the provenance. Fresh physical ranges are the edit
authority. A successful render commits the staged document and package together
and invalidates both layout caches once.

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
deletions and move sources and converts deleted text and deleted field codes
(`w:delText` and `w:delInstrText`) back to ordinary text and field codes.
Property rejection restores exactly one namespace-correct prior property
value. A `w:pPrChange` holds only the base paragraph properties, so rejecting
it keeps the paragraph mark `w:rPr` and the `w:sectPr`, which resolve their
own markers, `w:rPrChange`, and `w:sectPrChange`. Rejecting a mark
`w:rPrChange` replaces the whole mark formatting, including unmodelled
children, and keeps only the mark markers. Redlines from rdocx
0.14 and earlier carry prior mark formatting as a `w:rPr` in the prior
properties. That child replaces current mark formatting and retains current
paragraph-mark markers that stay. A mark formatting change that stays has no
owner in this form and is refused. When only the edited paragraph had mark
formatting, the prior holds no `w:rPr` and reads like Word's form, so rejecting
it keeps the edited mark formatting. Contextual markers act on their owning
run, paragraph mark, numbering property, or row. A removed paragraph mark
merges with the next paragraph across a table that immediately follows it when
every row of that table is removed by the same resolution. Before a retained
table, removal drops an emptied paragraph, permitting proofing markers but
refusing retained content, opaque children or meaningful raw gaps atomically.
A hyperlink whose selected revisions remove every child is dropped only when
its resolved inner bytes contain no opaque content. Previously empty hyperlinks
remain. A removed paragraph mark inside removed content, such as a cell paragraph of a removed row, needs
no merge partner. The merged paragraph keeps the resolved properties of the
last paragraph, whose mark remains. Resolution stages every affected package part, resolves
selected descendants before their enclosing subtree, reparses the complete
candidate package, and commits once only after validation succeeds.

The `rdocx` facade also owns deterministic comparison of those same stories. A
source index assigns stable public story categories and private owner paths to
modeled paragraphs, tables, fields, and nested text boxes. A concrete
hierarchical longest-common-subsequence alignment covers paragraphs, runs,
table rows, nested tables, lists, and modeled content inside existing
content-control shells. Changed gap owners pair by kind and in order, so a
paragraph beside an edited table cannot force whole-table replacement. The
default path retains whole-run behavior. An
additive options value selects Unicode-scalar character units or maximal
Unicode word, whitespace, and punctuation-or-symbol units, and applies
left-biased ignores for formatting, textual whitespace, fields, comments, and
selected story categories. Selected categories leave the original story bytes
untouched and are excluded before shell checks and revision-id allocation.
A linear post-pass groups word and character changes into consecutive deletions
followed by insertions within each structural and formatting owner segment.
Whitespace anchors between changes fold. Other anchors fold only when their
joined word count is strictly smaller than both deleted and inserted word counts
on each side. Character fragments carry word state across units, so an unchanged
one-word anchor between one-word edits stays even when those words are long.
Fields, drawings, ignored units, opaque children, bookmark and comment markers,
hyperlinks, controls and formatting transitions stop grouping. Formatting-only
changes and replacements using ignored formatting also stop it. Different runs
retain their own formatting and coalescing ownership. Run granularity retains
its existing alignment and revision output. Supported inline controls retain raw
child run cuts before extraction, align Word and Character units within those
shell segments and copy each shell before its first inserted unit.
Non-text content remains atomic. Unmatched identical owners become move pairs
only within one story. A changed run of consecutive paragraphs that holds a
pair whose hyperlinks, bookmarks, comment ranges, preserved raw children, or
inline controls differ, that gains or loses a modeled field, or that holds a
hyperlink or simple field, is replaced whole: all its original paragraphs are
deleted, then all its edited paragraphs are inserted. The run grows over its
neighbouring paragraphs until each side holds every complex field it begins or
ends, so a table of contents, whose end Word writes in a paragraph of its own,
is deleted and inserted whole, which Word needs to accept or reject it. A
whole deleted, inserted, or moved paragraph carries its bookmarks and comment
ranges inside its revision wrappers. Hyperlink shells remain outside the
wrappers, which track their own children, so whole hyperlink paragraphs can be
inserted, deleted or moved. Simple fields still require a complex field carried
whole around them. Empty whole-paragraph hyperlinks remain refused. Captured
nested Word hyperlinks remain unsupported and are refused before deeper wrapping.
Wrapper calls are bounded to the paragraph and one direct hyperlink owner. Comparison
also refuses when the two sides of such a run share a bookmark or comment range,
or a moved paragraph holds one, since the redline would hold it twice. Expanded
Word names distinguish typed owners from foreign lookalikes, whose raw bytes
remain opaque. Whole-owner wrappers retain comments and processing instructions
in their original positions. Deleted content writes a field code as
`w:delInstrText`, since Word refuses to open a deletion that holds
`w:instrText`. A carried hyperlink whose target only the edited side has gets
its relationship in the redline, and a paragraph mark lists its revision
markers in schema order before its formatting.
When a matched paragraph changes its comment range, hyperlink, inline control,
or preserved child boundary without moving a bookmark, comparison writes a
paired paragraph deletion and insertion. This also covers a TOC entry whose
cached text becomes a hyperlink and PAGEREF field. Empty self-closing paragraphs
are opened before granular properties or text are inserted. A paragraph before
a retained table carries its own mark revision. Trailing deletion chains mark
the kept predecessor once and retain accepted and rejected paragraph properties.
Changed field results remain inside their field owner, while instruction or
form changes replace that complete owner. Supported run, paragraph, table, and
section properties emit property revisions that retain the original property
sidecars. A changed paragraph records its prior base properties in
`w:pPrChange` and its prior mark formatting in a `w:rPrChange` in the mark
`w:rPr`, as Word's own Compare writes them. Unsupported formatting
differences retain the original bytes and produce stable
`ComparisonDiagnostic` values at the actual story path. Inputs
with existing modeled revisions or differing story shells are rejected unless
their story category is ignored. The root and owner start tags of a comment
or note story compare as namespace-resolved trees, so a part written again
with other declarations, attribute order, or empty-element forms keeps its
shell. Edited comment additions, removals, replies, resolved state, and dates
use a related custom XML part when a normal comment story revision cannot
express the changed owner shell. The redline holds the edited comment parts
and their relationship targets. The custom part holds the original comment
package state and one selectable revision identity. Accepting keeps the edited
comments. Rejecting restores the original comments and removes edited-only
related assets. A comment asset whose producer path already holds a different
asset gets a distinct package path without changing the unrelated owner.
Unrelated comment root extensions still refuse a changed
shell. Compatible comment text edits continue through WordprocessingML
revisions. A content control's shell is its type and data binding, and a
difference there is rejected too. Its `w:id` is producer
identity and ignored. Its tag, alias, lock, placeholder, and document-part
gallery are metadata, so controls that differ only by those align, compare,
keep the original `w:sdtPr`, and report one `content-control <name> differs`
diagnostic per property. Attributed text alignment retains owner,
formatting, content position, and raw-child boundaries, then coalesces
adjacent equal-owner edits into minimal revision wrappers. That alignment
runs separately between consecutive hyperlink, inline-control, bookmark, and
comment-range boundaries,
so no text matches across a shell and words inserted or deleted beside a
shell move it. Text inserted between two boundaries with no original run
between them, such as before a hyperlink that opens its paragraph, has no
original bytes to go between and refuses the pair. When every run of a
paragraph matches, the runs stay whole and only the differing inline
controls are compared.
When a main story or a content control gains a trailing run of paragraphs,
comparison marks the original final paragraph boundary once, marks each
intermediate inserted paragraph boundary once, and leaves the final inserted
paragraph mark as the story or control terminator. A self-closing original
final paragraph expands around its marker without creating a raw sibling. This
ownership lets acceptance retain every appended paragraph and rejection
reconstruct the original without an empty terminal residue.
When an inserted or deleted table stands between that boundary and a final
inserted or deleted paragraph, the paragraph mark before the table carries the
boundary, because resolving that change removes every row of the table.
An inserted or deleted row carries its marker in `w:trPr`, and every cell
paragraph mark and every cell run carries one too, nested tables included, as
Word's own Compare writes it. Without the cell marks Word merges the adjacent
paragraph into the first cell.
A paragraph-mark marker is the first child of `w:rPr`, and a new `w:rPr` goes
before `w:sectPr` and `w:pPrChange`, as the schema orders them.
Removing the marked paragraph mark merges that paragraph into the final one,
which keeps its own properties. When the two resolutions need different
properties there, the final paragraph carries the accepted paragraph
properties with a `w:pPrChange` that holds the rejected ones. When mark
formatting differs, its mark `w:rPr` ends with a `w:rPrChange` that holds the
rejected mark formatting, as Word's own Compare writes them. An inserted or
moved-in paragraph therefore
records the original properties of the paragraph before it, and a deleted or
moved-out one the edited properties of that paragraph.
Comparison patches only owned source spans, preserves every unowned byte,
stages the complete package, proves that acceptance matches the edited policy
projection and rejection matches the original, then commits once.
Alignment leaves paragraph properties out so that a paragraph whose
properties changed still matches, but the accept and reject projection adds
the modeled properties of every paragraph of each story, including those in
tables and content controls, unless formatting is ignored. A resolution that
loses paragraph or mark formatting, apart from section breaks, therefore fails
the postcondition.

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

Caller-width content measurement stays on this ownership boundary. The
`rdocx` facade resolves one checked paragraph or table location and builds the
ordinary `LayoutInput`. A fresh deterministic `rdocx-layout::Engine` then uses
the production paragraph or recursive table path and returns point height plus
ordered layout diagnostics without pagination. The query never enters or
invalidates either facade cache. Related header and footer content uses the
same part-scoped media view as whole-document layout.

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

Configured picture and text-box authoring follows the same story ownership.
The facade stages the selected story, allocates package-global drawing ids, and
creates media relationships only on that story's OPC owner. Pictures carry
typed crop, size, inline or floating position, wrap, distance, and z-order
choices. Text boxes emit one DrawingML WPS primary branch plus a self-contained
VML fallback inside `mc:AlternateContent`. Section-aware text watermarks target
one effective header variant without changing the caller-owned even-and-odd
header setting. Every operation serializes and reopens the staged package
before publishing it.

Footnotes, endnotes, and authored separator and continuation records are laid
out into a `NoteRegistry` before pagination, and the paginator reserves, splits
and draws them. Note placement is part of
pagination rather than a pass that runs after it, because a page's body height
depends on the note area it owes, and a note that does not fit continues on the
following page. The registry pre-shapes each note's marker, so the paginator
places notes without needing a mutable font manager.

Each note is laid out once per distinct section content width rather than once
per document, and is looked up by the width of the section drawing it. A note is
broken to the measure of the section carrying its reference, since that is the
measure it is drawn at, and reserve and render therefore still read the same
lines. A document whose sections share a page size registers one width and lays
each note out once, which is the common case. Section-end endnotes use their
section's width, while document-end endnotes use the final section's width.

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
sits at the foot of the page carrying its reference or immediately beneath
body text when the effective policy selects that position. The paginator
reserves its full note area before admitting body lines. Endnotes flow from the
last page of their section or from the final body page, then continue on later
pages as needed. The body and note marker use one label assigned at the first
body reference under document and section policy. Page restarts use the
reference's physical page after a bounded pagination pass. A custom mark does
not consume the numeric stream. A reference carries a `NoteRef`, its stream
and its package ID, because footnotes and endnotes can share an ID.

The native document facade stages normal endnote creation, exact element
reordering, and removal with matching body references. Endnotes use their own
relationship-resolved part and ID allocation. Rich edits use the same checked
story operations as footnotes, then reopen the candidate package before
publication.

## Versioning

The 15 shared and PowerPoint publication candidates use the explicit common
incubating version 0.12.1 in their manifests and workspace pins. They are
published from immutable annotated tag `rpptx-v0.12.1` at reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b`. Every registry entry and its
sole owner are verified. The failed immutable
`rpptx-v0.12.0` tag targets reviewed SHA
`54f4567b54b4028cd5126bcf66054f3e0588a4a9`. Its Windows CLI archive used CRLF
for reviewed text, so aggregate asset validation stopped before any registry
publication or GitHub release. The `rpptx-wasm` preparation member remains
unpublished at 0.12.1. Earlier coherent families remain available. The family
includes `oxml-chart` as the format-neutral owner while
retaining `rpptx-chart` as a source-compatible deprecated shim. The released
`rdocx-*` crates use the separate workspace version. The stable workspace, its
eight stable-version internal pins, ten inherited lockfile packages, the
`rdocx` Python project, and unpublished `rdocx-wasm` package are at 0.14.0.
The exact seven-package stable crates.io family is published from immutable
annotated tag `v0.14.0` at reviewed SHA
`58ca5a279277f7cd8de0b8f250fb4650de14371b` and depends on shared 0.12.1.
PyPI `rdocx 0.14.0` and `rpptx 0.12.1` are published from the corresponding
Python tags at the same reviewed SHA. Each has six `cp39-abi3` wheels and one
source distribution with complete crate-local README metadata. Every selected
registry owner is verified as `mantissaman`, while binding and WASM crates
remain unpublished on crates.io. The unpublished 0.13.2 crates.io train is
superseded by 0.14.0 rather than backfilled. Earlier coherent Rust and Python
releases remain available. The immutable v0.13.0 tag at reviewed SHA
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
v0.11.0 GitHub release exists. Earlier immutable registry releases, including
the complete 0.12.0 family, remain available. Version preparation and manifest
eligibility do not authorize any later publication.
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
contiguous regular runs in ordinary shapes, nested groups, table cells, and
speaker notes. `Presentation::try_replace_text` stages that complete mutation,
serializes the candidate, and publishes it only after validation. Fields,
breaks, and selected alternate-content fallbacks remain traversal boundaries
so the facade preserves their unmodelled or separately typed XML.

The facade also owns modern PresentationML package identity. The exact main
part content type distinguishes PPTX, PPTM, POTX, POTM, PPSX, and PPSM. Byte
serialization preserves that source class, while a path save writes the class
that its PowerPoint extension names. An explicit output conversion
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

Destructive Word facade edits stage the document and reconcile qualified raw
comment ownership before publication. One inventory covers body, related
stories, selected and opaque source owners, and linked comment companions.
`try_remove_content` distinguishes an absent index from an unsafe cut, while
its legacy boolean twin returns false for either. Complete owned threads may
be removed. Note deletion, section story pruning and revision resolution use
the same candidate boundary. Unchanged or increased source marker counts leave unrelated
producer annotation metadata untouched rather than adding global save-time
validation. Decreases of demonstrably undefined markers retain their legacy
lifecycle because no owned definition can be orphaned, with absence proved from
actual qualified raw definitions, malformed or ambiguous owned linkage refused,
and CLI validation remaining strict. Partial ranges and commented fragment detaches refuse without
changing the live package. Cell text edits retain checked paragraph anchors,
and a document-unaware cell removes only a proven plain empty paragraph when
another required paragraph remains.

Both facades use the same borrow-handle idiom rdocx already has: a mutable
`Foo<'a>` wrapping `&'a mut CT_Foo` and a read-only `FooRef<'a>`, with
consuming builders for formatting so calls chain, `&mut self` methods for adding
content that return a nested handle, and index-based `Option`-returning
accessors that never panic.

`Paragraph` follows that idiom for the whole paragraph-property surface. Each
property has a consuming builder, an in-place setter, and a `_value` form that
takes an `Option` and removes the direct value, and every index accessor for a
tab stop or a border edge returns an `Option`. `Paragraph::mark` returns a
nested `ParagraphMark` handle over the paragraph mark's own run properties
rather than adding a prefixed method per mark property. `ParagraphFrame` is the
checked mirror of `w:framePr` and converts through the pinned truncating twip
constructors. Relative frame alignment, the frame height rule, and any producer
attribute are not mirrored and survive a facade write unchanged.

The native Word facade creates a complete Word-compatible DOCX graph by
default. `WordCreationProfile` separates package completeness from the four
`WordPackageClass` identities, so native callers can select DOCX, DOCM, DOTX,
or DOTM while retaining an explicit compact package option. The compatible
profile owns its main document, styles, settings, theme, font table, core
properties, and application properties without loading a template.
Fresh Word-compatible documents initialize the common Word paragraph and table
styles in the native facade. Minimal-profile documents retain the compact
Normal and Heading1 set.

`Document` keeps relationship-resolved typed theme and font-table state beside
their resolved part names and dirty flags. The native facade re-exports the
shared concrete DrawingML theme and owns concrete font-definition, embedded
face, face-kind, and license-authorization values. Theme, font-table, embedded
part, relationship, content-type, and layout-input changes publish from one
staged clone. Caller font bytes and exact license identity are explicit input.
The facade never searches the host for bytes to embed.

Core, application, and custom property models are relationship-resolved
package state owned by `Document`. The native facade exposes their existing
concrete `oxml-core` types. Property creation, replacement, selective custom
property removal, and whole-part removal all run on a staged document clone so
the part, package relationship, content type, and typed model publish together.
An empty custom-properties part is pruned only when the current facade created
it. Python lists those custom properties on the existing class in part order.
Iteration yields each property, length is the count, items yield the name and
the property, and name lookup raises `KeyError` when the name is absent. An
absent part has length 0. There is no `keys` method and no second collection
type. Settings mutations use the same staged boundary and keep the existing
relationship-resolved target. Web settings mutations reuse that boundary and
allocate a collision-safe part only on the first authored value, and the part,
its relationship and its content-type override are pruned together when the
last value is removed from a part this facade created.

Facade serialization compares each relationship-resolved typed root with the
typed state parsed from its current package part. An equal root keeps the exact
source bytes, including producer formatting and namespace choices. A changed
root serializes only its owned part. Relationship collections apply the same
rule to their source `.rels` bytes.

Container-neutral Word story editing also belongs to the `rdocx` facade.
Concrete `StoryKind`, `StoryId`, `ContentLocation`, `StoryItemKind`, and
`StoryItemRef` values address the body, cells, headers, footers, ordinary
footnotes and endnotes, comments, and nested text boxes without constructing a
second document tree. The facade resolves package owners and stable source
order. The existing `rdocx-oxml` grammar remains the authority for admitting
content controls, revisions, and fields as typed content. Content rejected by
that grammar remains one opaque preserved boundary and cannot expose nested
owners or editable text.

Footnote creation and removal stage the note part and body references as one
package candidate. Reorder moves the exact note element bytes and retains its
internal ID. Common story content, relationship, and comment operations edit
rich note bodies through the resolved note owner. Each successful mutation
reopens the candidate before it replaces live state.

Word writes a text box twice in a run's
`mc:AlternateContent`, as DrawingML in `mc:Choice` and as VML in
`mc:Fallback`. That text box is one text-box story, read from the first
Choice that holds a text box, the one layout draws. Any other Choice and the
Fallback stay opaque, and the Choice drawing adds no drawing item to the story
that holds it. A paragraph replacement in a qualified alternate text box updates every admitted copy of that paragraph in one transaction, or changes none. Matching topology and accepted text are required. An asymmetric pair refuses before publication. A text box outside alternate content changes only its own paragraph. The Python handle that published the edit retires, so a later read of that handle fails. Document-wide text replacement still edits every copy, see `10-bindings-spec.md`.

`StoryItemRef::links` returns modeled hyperlinks in item source order. Display
text comes from the existing story text projection, while relationship targets
resolve only through the item's checked physical `StoryId`. Nested cells and
text boxes remain separate owners and are not folded into the enclosing item.
`Document::story_links` merges item-owned links by their physical XML position
and returns each existing `LinkInfo` with its checked `ContentLocation`. This
keeps nested content-control ownership without reordering interleaved links.
Owned story-item and story-link snapshots build the package source and owner
inventory once per accessor. Namespace scopes for all selected owners, items,
and hyperlink starts are collected in bounded source passes. Item and hyperlink
text is then projected from namespace-complete fragments without restarting at
the physical story root for each returned value. Complex-field text and link
excerpts share one namespace envelope around the exact sibling-run bytes.
The envelope inherits the context before the first run. Each original element
retains its local declarations, which never become sibling context. Single-root
excerpts likewise keep their own declarations authoritative.
Enclosing fields establish typed discovery admission but carry no local offsets
into an isolated excerpt. Checked inverse translation restores hyperlink source
positions before relationship lookup or smallest-owner deduplication.

The Python projection materializes each story item as a frozen value with its
exact XML bytes and the binding revision that produced it. Story mutation
resolves the story kind, part name, owner index, item kind, and item path
against the live native document, then rejects a stale revision before calling
the staged native operation. Default header and footer setters and story or
paragraph hyperlink creation use the same native ownership and relationship
paths. Cloning omits comment anchors from the copy, while story text replacement
removes only hyperlinks that had visible content before the replacement and
became empty because of it. Body text lookup returns direct paragraph matches
before enclosing content controls and exposes every matching body coordinate.
Paragraph text and run handles use one accepted-view walk. Direct runs, inline
content-control runs, insertion runs, and move-destination runs retain recursive
source paths in exact order. Deletion and move-source text stays excluded.
Literal and regex replacement read the same runs in the same order. Body,
table-cell, nested-control, header, footer, and text-box traversal also uses
the accepted view for readable text, replacement, export, links, and pictures.
Producer identity on paragraph, run, table-row, and content-control owners
does not create a content revision or hide an editable run. The runs of
one inline content control, insertion, or move destination form a stretch of
their own, and a match must lie within one stretch, so a match that crosses an
insertion boundary is not replaced and is not counted. A replacement inside an
insertion is written back inside the same wrapper, which keeps its id, author,
and date, so the edited text stays attributed to that tracked change. A
deletion or a move source is never matched and does not split a stretch.

The paragraph model keeps a smart tag and an inline custom XML element as raw
XML, and a simple field as a field run whose source is its raw XML. Paragraph
text and replacement parse the runs of these wrappers on demand from that
source, in document order, nested wrappers included, and the paragraph run
handles do not address them. The runs of each wrapper form a stretch of their
own under the same rule. A replacement inside one writes the wrapper again from
its start tag, its changed content, and its end tag, so its attributes and a
field instruction keep their bytes, and a wrapper without a match keeps all of
them. For a simple field, the replaced text is its cached result, which is what
a reader sees until Word updates the field. A wrapper inside a content control,
a revision, or a hyperlink is not read.

`ContentFragment` owns one paragraph, table, block content control, or removed
preserved node. Insert, remove, clone, and move resolve canonical
`StoryItemRef::location` values only when they name actual direct owner
children of the matching kind. `ContentLocation::end` is the distinct boundary
after final direct content. It works for empty and self-closing owners and
remains before body section properties. Moves stay within one unchanged story
owner. Clones allocate fresh document identities and drop the `w14:paraId` and
`w14:textId` of their paragraphs and table rows, while relationship-bearing
fragments require the unchanged owner scope. Every operation serializes and
reopens a staged candidate before publishing it.

`Document::insert_content_control` writes one rich-text or plain-text shell at
a `ContentLocation`. `ContentControlPlacement` is block, row, cell, paragraph,
or inline, and it must match the destination parent. A block story or a cell
accepts a block or paragraph control. A table accepts a row control. A row
accepts a cell control. A paragraph accepts an inline control. A mismatch
returns `Error::Other` and publishes nothing. Kind is `SdtType::RichText` or
`SdtType::PlainText`. Any other kind names F-286 and publishes nothing. The id
is the smallest unused positive `i32` across story parts. Tag and alias reject
non-XML characters. Lock, appearance, and placeholder edits rewrite only
`w:sdtPr`. `None` clears that field. Unmodelled property children and
foreign-namespace `showingPlcHdr` lookalikes stay. Child order on a created
shell is `w:sdtPr`, then `w:sdtEndPr` when present, then `w:sdtContent`. The
default inline shell is one empty run. The default block or paragraph shell is
one empty paragraph, and a plain-text paragraph in that shell contains one
empty run. The default row shell is one schema-valid row. The default cell
shell is one schema-valid cell whose block content is one empty paragraph.
Direct block children use `move_content`, `clone_content`, and
`remove_content_at`. Row, cell, paragraph, and inline controls use that same
staged transaction through location-aware wrappers. Clone drops `w14:paraId`
and `w14:textId` and allocates a new sdt id. Move keeps the id and stays inside
one story. Both refuse before publication when the destination parent cannot
hold the control content. A nested content control is transparent, so move,
clone and unwrap apply that same parent grammar to its content. Admitted
`w:customXml` wrappers also inherit that grammar. Their optional
`w:customXmlPr` stays metadata, while their actual row, cell, block or inline
payload must fit the destination. Alternating custom-XML and control wrappers
do not reset the context. Repeating conversion uses this same policy. Character
references, predefined entities and CDATA at that boundary count as their
decoded text. Remove deletes the shell and its content. Unwrap leaves the content
only when the parent can hold it, and refuses a cell that would lose its last
paragraph. Story-item classification seeds each owner with the ancestor prefix
bindings already collected for that part, including a document-level `w14`
binding, so a control whose paragraph carries `w14:paraId` stays typed. That
scan does not start a second per-offset namespace walk. Imported recognized type markers retain their
original qualified name and existing namespace bindings during modeled
serialization. This includes empty and nonempty markers and inherited or local
bindings. A foreign binding on a canonical prefix cannot change the marker's
expanded name or the meaning of retained opaque children. Newly authored or
replaced kinds use their canonical names and declared bindings. Placement
checks therefore see the same qualified type before and after staging.

`Document::set_content_control_type` replaces the type element on a shell from
`insert_content_control`. The payload is rich text, plain text, date,
checkbox, combo box, drop-down list, picture, group, repeating section, or
repeating item. Date writes `w:dateFormat`, then `w:storeMappedDataAs` with
`text`, then `w:calendar` with `gregorian`, and omits `w:lid`. The format is
`yyyy-MM-dd` or `dd/MM/yyyy`. Checkbox writes `w14:checked` `0`, then
`w14:checkedState` `2612` in `MS Gothic`, then `w14:uncheckedState` `2610` in
`MS Gothic`. List items keep caller order. An empty list, a repeated value, or
a non-XML character refuses. Picture requires an image relationship the part
already owns and a drawing in the shell that embeds it. Repeating section
content must already be one repeating item, or one block sequence that the
call wraps in one nested repeating item. Empty content refuses. Admission
uses the same decoded-text rule at each transparent nested block boundary. A
repeating item requires a containing repeating section and block, row or cell
placement. Containment is distinct from immediate parenthood. Repeating
section conversion also requires a physical level capable of holding such
items, independently of whether empty payload fits a block grammar. Qualified
custom-XML ancestors inherit that physical level. Relocation checks the outer
and retained boundary types as well as payload grammar and preserves item
containment. Moving the last direct item refuses. Type changes and unwrap
refuse orphaned items or non-item children in a section. An inner section can
retire when an outer section retains ownership. An item can unwrap into an
item-only sequence that leaves at least one direct item. Removing a section
deletes its content and does not orphan retained items. Citation,
equation, bibliography, and document-part kinds stay refused at creation and
have no payload. Tag, alias, id, lock, appearance, placeholder, and unmodelled
property children stay. `Document::repeat_content_control` clones the first
repeating item, assigns a new sdt id, drops `w14:paraId` and `w14:textId`, and
appends the clone. Zero items refuse. `remove_content_control` deletes one
item of several and refuses the last item in its section. Combo and drop-down
fill share the existing list table, so an unknown or ambiguous value refuses.
A refusal returns `Error::Other` and publishes nothing. Python exposes
`ContentControlTypePayload` and the same two methods. WASM and CLI gain no
method. Layout and render keep the existing content walk. Native Word
qualification compares documented equivalent operations. The API intentionally
clones the first item rather than Word's addressed-item default, and refuses
last-item removal rather than Word's whole-section deletion. Authentic producer
version and input identities remain mandatory for the comparison.

`Document::add_custom_xml_store` stages one custom XML item, its properties
part, both relationships, and both content-type overrides through the
bibliography store sequence. The call returns the braced item id. An occupied
conventional name is not overwritten. Item bytes must be one well-formed
XML element. DTD declarations are explicitly unsupported and refuse before
allocation in every position. XML declarations are allowed only first and
must satisfy the XML 1.0 declaration grammar. Input bytes are UTF-8. A declared
encoding must be UTF-8, or US-ASCII for an entirely ASCII input. Other encodings
refuse rather than being interpreted as UTF-8 under a contradictory label.
Processing-instruction targets
must be legal XML Names and cannot be case-insensitive `xml`. Every parser
event has an explicit admission rule. That check includes XML names, attribute values, entity
references, unique attribute names and comment well-formedness. One reference decoder feeds the XML
1.0 character rule. A reference, CDATA section or nonspace text outside the
root is a fragment. Valid item bytes stay verbatim,
including an unbound prefix. An empty schema list writes an empty `ds:schemaRefs`. Each schema URI
is stored in caller order and is not fetched. `Document::set_content_control_binding`
writes `w:dataBinding` with `w:prefixMappings` when a mapping is present, then
`w:xpath`, then `w:storeItemID`. `None` and an empty mapping omit
`w:prefixMappings`. The store must resolve to one part, and the path must
select one simple element through the existing absolute child-path subset. A
refusal returns `Error::Other` and publishes nothing.
`set_content_control_value_by_tag` and `bind_content_controls` still update the
display and the store together. A failed sibling rolls both back. Unknown
sibling bytes stay byte-identical. Bibliography keeps its own item bytes and
still returns the part name from its private caller. The helper stays in
`rdocx`. Python exposes both methods. WASM and CLI gain no method.

`DocumentFragment` owns a package-authoritative half-open block selection from
any supported story owner. Existing two-segment paragraph paths inside block
controls resolve their enclosing control content. Source and destination keep
their physical part's relationship scope. Inline items and split owners are
invalid fragment boundaries.

Import closes only dependencies reachable from selected XML and its required
note and comment companions. Style and numbering references reach a fixpoint
across those companions, including cyclic style links and numbering overrides.
Custom XML stores carry item properties and receive fresh store IDs on
collision. Notes, comment threads, bookmarks, paired endpoints and revisions
receive deterministic new identities. REF targets share the bookmark-name map.
Charts, diagrams, embeddings and opaque extensions use the existing OPC graph.
Internal cycles retain exact payload bytes and part-local relationship IDs,
with allocated part names and rewritten internal targets. External edges keep
their target, type and mode without fetching resources. Reuse applies only to
structurally equivalent styles, numbering and relationship-free leaf parts.

One staged document serializes and reopens before publication. Insertion into
a note or comment part retains companions imported into that same part.
Dangling, malformed, split, exhausted and unsafe opaque-reference graphs fail
without changing the live document. An integrity-bound package signature part
that cannot survive remapping is rejected. Existing destination signature
invalidation policy applies to changed coverage. Preserved opaque companions
gain no new rendering or decoding support.

Ordered Word section ownership also belongs to the `rdocx` facade. Concrete
`SectionRef` and `Section` handles borrow the existing paragraph-level or
schema-final `CT_SectPr` owner and carry its document-order ordinal and final
owner identity. They expose page size, orientation, four margins, gutter,
equal-width columns, page-number start, header and footer distance, title-page
state, break type, and the complete properties. Checked setters validate the
full `Length` value before publishing any narrowed twip. Orientation mutation
normalizes page dimensions. The legacy final-section `Document` convenience
setters remain unchecked and infallible. `Document` counts, iterates, and looks
up those owners without constructing a second section tree. Insertion and
removal stage the complete document and package, retain body order and
unmodelled section XML, serialize and reopen the candidate, then publish once.

Per-section header and footer ownership is a native `rdocx` facade concern.
`HeaderFooterKind` and `HdrFtrType` select the story family and default, first,
or even variant. `SectionStory` reports the resolved `StoryId`, its source
section, and whether it is inherited. Lookup follows only the same variant
through preceding sections. Create, link, inherit, unlink, replace, and remove
stage the complete document and package, serialize and reopen the candidate,
then publish once. First-page creation enables `titlePg`. Even-page selection
is controlled separately by the typed document setting. Rich edits continue
through the container-neutral story operations rather than a second header or
footer content model.
Comment authoring uses the same checked story run range in every header and
footer variant. The facade anchors selected paragraphs in the owning related
part on a staged package, retains untouched producer XML around those spans,
then publishes the part and comment definition together after validation.

Word layout retains that physical ownership. The facade loads header and
footer images under the main-part relationship that selects the story plus the
story-local image relationship. `MediaRegistry` exposes a scoped view for each
selected part while sharing one immutable media payload map. An equal local
relationship identifier in the body, a header, and a footer therefore resolves
to three independent images.

The Python facade projects sections, styles, stories, story items, effective
header and footer variants, and hyperlinks into detached frozen records.
Document accessors return tuples in native source order. Records retain the
physical story owner, item index path, inheritance source, and relationship
identifier without exposing raw XML or adding a second binding-side tree.

The Python document facade also exposes direct-body structural editing through
live Paragraph and Table handles. Native identity mapping converts their lazy
paragraph or table ordinal to the interleaved body coordinate and rejects a
nested, stale, or foreign handle. Popping returns one opaque owned
`ContentFragment`, and insertion clones that value without exposing its XML.
Insert, pop, clone, and move publish their native staged result before advancing
the binding revision once. Counted literal and regular-expression replacement
advance the revision only when the native operation reports a nonzero count.

`rdocx-oxml` authors only `w:pgNumType/@w:start` for M23. Number format,
chapter style, chapter separator, and every other unsupported attribute or
child remain in the retained source for M24. Section serialization replays raw
children at schema slots and at predecessor or successor boundaries between
repeated header and footer references. Distinguishable references follow their
value through removal and reorder. Equal public values are indistinguishable
through the preserved `Vec<HdrFtrRef>` surface and resolve by deterministic
source ordinal, with a resolvable following reference taking precedence.

`Document` also owns one relationship-resolved `CT_Styles` graph. Public style
creation, update, default selection, and removal build a complete candidate,
validate identifiers, reference types, defaults, links, next styles, and
inheritance cycles, then publish the staged graph and invalidate layout once.
The facade reuses the existing paragraph, run, table, and conditional property
types. It does not introduce a second style model.

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
Nested tables and the content controls at every level contribute their
paragraphs in place. `Document::images` and `Document::word_count` reach the
same content. `Document::headings` and `Document::links` read the body
paragraphs and those that body-level content controls wrap, and do not search
table cells. The Markdown, HTML, MHTML, EPUB, ODT and RTF exporters write
what the text readers read: what content controls wrap in place, as if the
control were not there, tracked insertions and moves in, and the runs of smart
tags and inline custom XML, while deleted and moved-away text stays out.
MHTML sizes and EPUB packages the pictures of that same view, in document
order. Each exporter that reports losses notes every control, revision and
wrapper it flattens or leaves out, and reports the losses of the content it
writes as it does outside them.
Each paragraph contributes the same accepted-view text as paragraph text, so
tracked insertions are included and tracked deletions are left out.
The OOXML model owns accepted structural visibility. A deleted or moved-away
body paragraph mark joins its paragraph to the next direct body paragraph
unless a section break or intervening block prevents that join. Text omits the
joining newline, while HTML and Markdown emit one block with the following
paragraph's block properties. Accepted layout omits an empty joining
paragraph and lays out retained joined content in the following block.
Deleted or moved-away table rows are absent from accepted text, JSON, HTML,
Markdown, MHTML picture traversal and layout. A table whose rows are all
removed contributes no accepted block or spacing. Source model rows and
editing indices remain intact, and tracked layout still shows them.
The WASM binding uses `Document::text` for its existing `getText` method and
otherwise owns one complete `Document`. It never reaches into
`rdocx-oxml` or maintains a second package representation.

`Document::render_page_to_svg`, its option-taking counterpart, and their two
deterministic variants expose one zero-based page as self-contained searchable
SVG. Out-of-range pages return `None`. `SvgRenderResult` carries the SVG and
ordered `SvgDiagnostic` values, with layout diagnostics before recursive
lowering diagnostics. Python binds the normal-layout method as
`Document.render_page_to_svg`. WASM, CLI, Presentation, and the public
`oxml-pdf` surface remain unchanged.

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
tree and does not maintain a second content-control representation. The first
supported type child keeps its unmodelled attributes and ordered child bytes.
Its typed discriminator selects a fixed-prefix type wrapper on output, and a
discriminator change replaces only that payload with the canonical empty type.

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
established semantics. `StoryItemRef::direct_body_index` returns the containing
direct main-body child for any safely anchored body item while preserving the
flat recursive `index_path`. Items from another story and final section
properties return no body coordinate. `Document::body_content` reports unsupported modeled
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
part remain outside this traversal. `Document::story_revisions` covers every
story instead. It stages a copy of the document as revision resolution does
and scans the main part and each related story part left in the staged
package with the element inventory that resolution counts, so its length
equals the accept and reject counts, text boxes included, and it fails where
staging fails. Each revision belongs to the innermost owner that
`Document::stories` reports around it, with table cells folded into their
story. Where the typed serialization that `stories` scans drops a namespace
binding the staged part keeps, owners pair across the two, and a text box
that `stories` does not report folds into the story around it. A revision
outside every owner, such as one in a footnote separator, is an error rather
than a silent omission.

Revision mutation uses explicit all, exact-author, inclusive RFC 3339 instant,
and id selectors. One id operation resolves every modeled element carrying the
id, while undated revisions do not match date ranges. Invalid bounds or a
malformed selected revision leave the typed document, package bytes, and both
layout caches unchanged. A successful operation invalidates layout once.

Word comment mutation uses `RunPosition` and half-open `RunRange` values whose
body indexes select top-level paragraphs and whose run indexes select insertion
boundaries. Run indexes count the accepted-view runs that `Paragraph::runs`
lists, including the runs inside inline content controls and tracked
insertions. `CT_P::anchor_accepted_range` writes the markers inside
`w:sdtContent` when a boundary falls between two runs of a control, and around
the control when the range covers it. A range that crosses the edge of a
control, has a boundary between two runs of a tracked insertion or move, sits
next to a tracked change inside a hyperlink, or continues into another
paragraph from inside a control is refused instead of shifted. Removal also
clears the markers and reference runs that are direct children of a control's
content. `Document` validates both endpoints before mutation, allocates
collision-free comment and paragraph ids, updates the comment parts and all
three anchors together, then invalidates layout once. `CommentRef` is a
read-only view over the typed comment and its comments-extended thread entry.
Its checked `anchor()` and `anchor_text()` accessors share the document's
qualified comment ownership proof. `Document::comment_anchor` returns an
accepted-view `StoryRunRange`, and `comment_anchor_text` returns exactly the
selected accepted display, with one newline between paragraphs. A known
orphan returns no range or text, and a reference-only point returns no paired
range with empty text. Replies do not inherit a fabricated parent range.
Unknown identities and malformed, ambiguous or unprojectable selected sources
return errors. Listing validates all owned definitions and companions, while
raw marker ids proven absent from qualified definitions do not invent records
or block established comparison reads. Strict deletion and CLI ownership
validation still reject those undefined markers, and reads preserve their XML. Source-order temporary projections reuse the existing paragraph
parser and accepted run axis without changing source XML. An interval around
only a raw wrapper can display text while having zero accepted runs. Such an
interval is explicitly unsupported when re-anchoring those coordinates cannot
represent it faithfully. Supported controls, revisions and ordinary field
caches retain accepted-view semantics. Multi-paragraph insertion qualifies both
source endpoints before publishing either, retaining stale-owner checks.
`Document::move_comment` selects a complete root range and reference through
that same qualified proof. Replies, unknown ids and incomplete or ambiguous
sources refuse atomically. Moving preserves thread identity and metadata,
transports the original reference properties and opaque attributes, and leaves
mixed-run neighbors in source order. Destination paragraph identity and run
boundaries rebase after removal. Only newly empty pure Google marker controls
are pruned. `move_comment_to_text` shares the existing recursive literal finder
and splitter without allocating a temporary thread.
Rich comment query restores main-part namespace context once through the existing strict declaration replay. Paired XML events prove that names, non-namespace attributes and payload remain identical, yielding exact physical boundary correspondence. The mapped canonical and restored story-owner inventories must be bijective in physical spans, kind and owner index before canonical paragraph paths receive actual namespace scopes. Additional or reinterpreted owners refuse instead of redirecting a query. The existing checked paragraph span traversal serves both discovery and projection, with paragraph scopes inventoried in a bounded pass. Canonical locations, fingerprints and accepted run indices remain authoritative. Imported paragraph and block-local alias fields retain rich A7B, source-byte purity and usable move/reopen endpoints. Comment destination anchoring replays retained declarations before replacing package authority, while other anchor kinds keep their existing path. Redundant declaration replay is omitted only for a unique exact owner whose complete non-namespace event payload and local declarations match, with identical resolved element and attribute namespace facts throughout that owner. Redundant child declarations on a transported reference do not require global replay to reinterpret its shape. This owner-local proof permits unrelated modeled source normalization. Needed declarations still use the unchanged strict replay. The private pre-removal endpoint probe validates the caller's exact accepted endpoints without publishing destination namespace context, while the real Comment destination always preserves that authority before replacing source. Same-owner modeled alias plus local raw-field fixtures that already fail initial save remain outside this correction. Supported controls include separate unrelated root-bound modeled aliases and normalized properties in the same package.
`StoryRunPosition` and `StoryRunRange` add checked `ContentLocation` ownership
for body and table-cell paragraphs without changing `RunPosition`. A body
location can also name a paragraph inside a block content control with a
two-segment path, the control's story item then the paragraph's position among
its paragraphs, which `Document::paragraph_story_location` returns for a
paragraph index. The staged
path validates both endpoints and edits cloned paragraphs before it creates
comment relationships, so any path, run, or package failure publishes nothing.
Replies follow paragraph-id parent linkage through each comment's last
paragraph, as `w15:commentEx` keys it, each line of a comment text is one
paragraph. Resolution applies to the thread root, and removal deletes the
selected comment plus descendant replies without
deleting unrelated runs or producer XML.
`CommentRef::resolved` is `Option<bool>`. `None` means the comment has no `commentsExtended` entry. A present entry without `w15:done` is `Some(false)`, including a comment this library just wrote. `Some(true)` means the entry is done. Python exposes the same three states as `bool` or `None`. The CLI lists a comment as resolved only when the value is `Some(true)`.
Whole-story header/footer setters stage the complete text, raw XML or image
operation and reconcile removed comment ownership. A clone of that candidate
prepares and reopens the complete output once before the original candidate
publishes, preserving live authoring identities until save. Glossary deletion and body replacement reuse that same strict graph.
The graph accepts an explicit physical main or glossary owner, so equal numeric
ids never join unrelated review parts. Imported retired section parts stay in
the package. Their qualified markers are removed only after every modeled,
inherited and proved opaque physical use disappears. Unknown incoming use
refuses. Glossary raw edits refresh the typed model before a later flush.

The additive `add_comment_with_date` and `reply_to_with_date` operations own
validated optional timestamps. The original operations delegate with no date.

Checked picture insertion uses the same story owner and package relationship
scope. `insert_picture_to_story` accepts bytes, a safe filename, paired
explicit dimensions or native 72 DPI sizing, and an optional direct item after
which to insert. It serializes, reopens, and returns the refreshed paragraph
location only after the image part, relationship, drawing identity, and story
content all validate together.

`Document::split_run` creates an exact accepted-view run boundary without
changing `RunPosition`. Its paragraph argument is the same direct body child
index as `RunPosition` and `find_content_index`, and an index that names a
table, a block content control, or preserved XML fails with an error naming
that kind. It clones the selected paragraph, resolves the selected
recursive source path, counts Unicode scalar values only in literal text,
partitions ordered zero-width children at their source boundary, repairs
hyperlink and marker coordinates, and publishes the clone only on success.
Zero and end offsets select existing boundaries and leave typed state, layout,
and binding revisions unchanged. A Python `Paragraph` handle to a paragraph
inside a block content control splits through `Document::paragraph_mut`, which
clears the cached layout even for those offsets. A structural edit makes an
earlier path-backed Python run handle stale.

Word bookmark mutation input reuses the same top-level `RunPosition` and
half-open `RunRange` boundary as comments. `Document::bookmarks` returns
immutable correlated summaries in typed main-story paragraph order through
tables and block content controls. A reported body index is that recursive
paragraph ordinal, and its run index is the accepted-view boundary used to
extract bookmark text. `BookmarkRef::direct_range` reports the same range with
the direct body child index of `RunPosition` when both markers sit in direct
body paragraphs, and `None` otherwise. `Document::add_bookmark` takes the
same accepted-view run boundaries and places its markers as comments do.
Marker encounter order resolves direction when start
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
type, or dependency. The resolved slide-to-notes edge is authoritative when a
producer omits the optional reverse notes-to-slide edge. A present reverse edge
must still be internal, singular, correctly typed, and equal to the known
owner. Notes text mutation stages one body placeholder, retains its identity
and first regular-run formatting and raw XML, and publishes only serializable
notes state.

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

## Shared Word document operations

The native Word facade owns accepted-view text presentation and advisory package validation. Python and CLI callers consume the same TextReport and ValidationReport. The native crate has no CLI dependency. Plain text presentation groups readable related stories in physical inventory order and retains main text with one ordered omission diagnostic if extraction fails. Body counts use accepted content rather than presentation labels.

Python template conversion admits explicit JSON scalar and container types before native mutation. Its existing workspace serde_json dependency edge introduces no separate converter model. Checked fragment publication and property mutation stage complete candidates, preserve the original on refusal and invalidate live handles only after changed publication. Markdown requests recursive physical-owner boundaries from the existing accepted projection. The default accepted view remains unchanged for its other consumers.

## PowerPoint checked XML ownership

The rpptx facade resolves checked XML against actual slide, layout or master
parts and concrete physical owner paths. Existing rpptx-oxml and oxml-drawing
parsers validate qualified roots and modeled child order. The facade stages
package candidates, validates relationship target roles and MIME identities,
and serializes and reparses before publication. Supported raster targets use
the shared bounded PNG, JPEG, GIF, TIFF and BMP decoder in rpptx. Existing
unsupported media facts may survive only with exact original target identity.

The Python binding stages that native candidate and prepares returned handles
and revision tokens before committing. Physical paragraph and run positions
address current owners, never prove that removed children survived. Actual
owner, text-body and XML-property child generations establish alias identity.
Shared revision primitives remain owned by oxml-py-support.

## Embedded font ownership

Word owns its relationship-resolved font table and ODTTF parts. PowerPoint
owns its source-preserving presentation font list and EOT parts. Both native
facades stage font model, package bytes, relationships, content types and flags
as one transaction. Shared source targets use copy-on-write and pruning removes
only genuinely unreachable owned parts. Descriptive records and opaque producer
source remain authoritative, and unsafe ambiguous edits refuse before publication.

Both families require explicit caller authorization and exact nonempty license
identity. Font bytes come from bounded caller inputs. Word checked listing
bounds projected owned copies while retaining the legacy snapshot layout. Its
indexed metadata reader resolves the actual part through the same checked
relationship provider. PowerPoint decodes each distinct supported part once
and shares its immutable body among aliases. Python preallocates immutable
result snapshots before publishing the native candidate, and CLIs use those
same native providers before staging outputs. No second ZIP/font/revision engine
or low-level dependency on a facade is introduced.

Native Office is an independent consumer of those package graphs, not a second
font provider. PowerPoint qualifies actual regular and bold font use through
source open, native save and normal reopen. Word's documented one-font-per-name
behavior can synthesize styles and change saved embeddings. That native
interoperability boundary does not change the four-face APIs, deterministic
renderer selection or caller precedence. Native Word multi-face work belongs
to F-X223.

## Presentation semantic text projection

The `rpptx` facade owns semantic HTML and Markdown projection in its existing
HTML carrier. Python and CLI consumers call that projection without invoking
font discovery, layout or rendering. Export options and detached results are
concrete facade values. The source package and live owners remain unchanged.

The facade directly uses the existing workspace `tiny-skia`, `gif`,
`zune-jpeg` and `tiff` dependencies for bounded image validation, including
when rendering and the default template are disabled. PNG and JPEG bodies are
fully decoded, all GIF frames and TIFF pages are checked, and uncompressed BMP
pixel extents are validated. The existing crate root owns the crate-private bounded picture-body validator.
Semantic export attaches its source-path error context at the call site. The
validator remains a concrete facade helper, with no dependency from a shared
OXML crate into a facade.

## ZIP package dependency

The three direct ZIP consumers are oxml-opc, rdocx and rpptx. They use zip9.0.3 with disabled defaults and the existing Rust Deflate backend. No extra codec or dependency alias is selected. Native OPC errors expose the upstream zip9 error identity. Existing package and ODF validation owns admission after fallible decoded-name access.

ZIP9 also selects its upstream macOS/BSD-only nix0.31.3 dependency with fs and
disabled defaults, plus cfg_aliases0.2.2 for that dependency build. They do not
add a ZIP codec and remain subject to actual target, license and MSRV checks.
