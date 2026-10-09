# 04, OPC and packaging

Owner: `oxml-opc`, with media naming in `oxml-media`.

## The package
Bibliography source transactions resolve the qualified source collection
through the document-owned custom XML graph. Source identity, GUID, relationship
owner and schema particle checks precede publication. Unmodeled members,
namespace scopes, repeated producer values, raw LCID and RefOrder occurrences
and unrelated package parts remain preserved. Duplicate or ambiguous identities,
malformed graphs, referenced deletion and stale paths refuse atomically.
No-op source replacement retains identical source bytes.

Citation and bibliography cache updates use a staged package and only publish
after every eligible owner succeeds. One recognized unfinished standard branch
aborts the whole update, including earlier candidate edits. Noncatalogue paths
retain their complete cache with report diagnostics. Locked fields and protected
producer topologies keep their own retained or refusal boundary. Simple fields
expand only when their qualified owner and producer attributes are proved.
Structured results preserve physical begin/separate/end spans, outside content,
first/interior paragraph context and IEEE label/cell/grid ownership.
Ordinary save does not refresh these caches or rewrite source metadata.


Generated-table edits stage source targets and provisional cache paragraphs
in one candidate. Only the owned separator-to-end cache span is replaced.
Supported simple generated owners expand to complex fields with the exact
producer instruction and result XML retained during expansion. Expanded names
and ancestor namespace bindings qualify every owner and generated paragraph.
Locked fields, unsupported formatting, unsupported locales and unmeasured sort
keys retain their complete owner and cache with diagnostics. Unknown source
XML and original style definitions stay untouched. Missing, ambiguous or
unplaced required targets and malformed spans abort without publishing any
candidate part.

```rust
pub struct OpcPackage {
    pub content_types: ContentTypes,
    pub package_rels:  Relationships,
    pub part_rels:     HashMap<String, Relationships>,  // key: "/ppt/slides/slide1.xml"
    pub parts:         HashMap<String, Vec<u8>>,        // same key shape
}
```

Fully in memory. Every part is decompressed at open. Part names are normalised
to a leading slash. This design is already format-neutral and is carried over
essentially unchanged from `rdocx-opc`.

Archive entry identity is checked after path normalization and before the
entry map is populated. Two ZIP entries that resolve to one normalized,
ASCII-case-insensitive part name are invalid. Duplicate default extensions and
override part names in `[Content_Types].xml` are also invalid, so map construction never hides a
package collision. Default extension and override part-name identities compare
ASCII case-insensitively in parsing and API lookup. Package part get, set,
contains, and remove operations, relationship-owner lookup, required Word and
PowerPoint part resolution, and digital-signature discovery and coverage use
the same identity. The first authored spelling remains the serialized key.
Direct mutation of the public maps remains supported, with deterministic lookup
if callers create an invalid case conflict. Serialization rejects that conflict
before writing output. Every loaded package retains unchanged producer
content-types bytes and element order regardless of enabled features.

Relationship XML rejects duplicate `Id` values during parsing and before
serialization. Package serialization validates content-type identity, part
identity, relationship-owner identity, every relationship-id scope, and output
ZIP entry identity before opening or truncating a destination path.

**No save writes a character XML 1.0 cannot carry.** Serialization refuses
`[Content_Types].xml`, every relationship part, and every UTF-8 or UTF-16
part whose content type ends in `xml` when its bytes hold such a character:
a C0 control other than tab, LF and CR, or U+FFFE or U+FFFF. The error names
the part, the code point, and its line and column. A part whose content type
does not end in `xml`, such as a VML drawing
(`application/vnd.openxmlformats-officedocument.vmlDrawing`), is not scanned.
This is the one place every writer of both formats passes, so a value stored
through an infallible setter or a public field is caught here. An entry that
already held such a character when the package was read is written back while
its bytes are unchanged, so producer content passes through verbatim. The
check is fail-closed: any edit that re-serializes such a part, or a
relationship part whose source held one, refuses the save rather than write it
again. The fallible entry points that take free text refuse such a value
earlier, naming it and the one-based character position, as python-docx and
python-pptx refuse such strings. DrawingML `a:t` text set through the text
setters is the exception, stored as `_xHHHH_` as python-pptx stores it.

**Saves are deterministic.** Both `part_rels` and `parts` are emitted in sorted
key order, so writing the same package twice produces byte-identical output.
That property is load-bearing for the round-trip corpus and must not regress.
The Word facade preserves relationship identities captured when a package is
opened. Relationships added to the current graph afterward are authored state,
including internal theme and other typed edges that have no modeled `r:id`.
They join deterministic semantic ordering by relationship type and normalized
target. Unknown and external authored edges retain their target and mode.

Modern presentation package identity is the main presentation part's exact
content type. `oxml-opc` names the ordinary presentation, macro-enabled
presentation, ordinary template, macro-enabled template, ordinary slideshow,
and macro-enabled slideshow values. `rpptx` accepts only those six values.
Changing an output class replaces only that override in a staged package.
Executable and unrelated parts remain opaque and byte-preserved.

Modern Word package identity follows the same content-type authority at the
common `Document` package-open boundary. Exactly one normalized internal
`officeDocument` relationship must resolve to an existing part with an exact
override for DOCX, DOCM, DOTX, or DOTM. Extension defaults, external targets,
unsafe targets, duplicate relationships, and unknown types fail closed. An
output class conversion changes only that override on a staged clone and uses
the existing signature invalidation path.

Fresh Word construction has two explicit completeness profiles. The minimal
profile owns only the main document and styles graph. The Word-compatible
profile also owns settings, the Office theme, the font table, core properties,
and application properties with exact content types and internal
relationships. `Document::new()` selects Word-compatible DOCX. DOCM and DOTM
declare macro-capable main-part identity without inventing a VBA project. Empty
core properties omit created and modified timestamps, so equivalent fresh
constructions remain byte-identical. Fresh application properties stamp the
crate major and minor version as an `XX.YYYY` AppVersion, the form Word opens.
Reading a package repairs a malformed AppVersion on save. An `rdocx` property
gets the current stamp, another producer loses that attribute, and a valid
AppVersion leaves its original part byte-identical.
Word-compatible construction includes the common Word style definitions in its
initial styles part. It does not add a template part or alter producer styles
when opening an existing package.

Word story discovery begins with the main body and its nested cell and text-box
owners, then follows header and footer references in document order and note
and comment relationships in package relationship order. Ordinary footnotes
and endnotes keep their source order, while conventional nonpositive separator
records are not public story owners. Each `StoryId` includes the normalized
source part, owner kind, source-order ordinal, and a structural fingerprint.
Any changed owner makes a retained identity stale before indexed content can be
resolved.

Rich footnote authoring reserves the relationship-resolved footnotes part,
adds one normal note with an available internal ID, and inserts its body
reference in one staged package. Reorder moves only the selected note element
within the part. Removal deletes that element and every matching body
reference before the candidate is reopened. Common story edits preserve
untouched note children, separator records, producer prefixes, and unrelated
part relationships. Pictures and hyperlinks use the footnotes part relationship
set, and comment anchors in note paragraphs use the comments part.

Endnote authoring uses the relationship-resolved endnotes part and its own
normal ID namespace. Creation inserts the note and body reference together.
Reorder moves an exact endnote element, while removal deletes that element and
its body references. The staged package reopens before publication. Endnote
story edits preserve separator records, unknown children, and unrelated
relationships. Pictures and hyperlinks are owned by the endnotes part.

Checked document and section note policies retain unmodelled property children
while writing position, format, start, and restart in schema order. The
document policy lives in the relationship-resolved settings part. An authored
separator, continuation separator, or continuation notice replaces only its
special note owner in the relevant note part. The corresponding settings note
property records the special ID. This namespace remains separate from normal
note IDs and unselected special records do not render. Removing only the
numbering or placement policy retains special-record selection and unmodelled
property children. Special records are never public note stories. Custom reference
marks live in the body reference and in the note owner's marker run. Package
mutation is staged through save and reopen before publication.

Revision inventory uses these same supported story owners and reports their
`StoryId` with each record. A revision reachable by resolution without a
discoverable owner is an error. CLI text extraction retains readable body text
and names any malformed related part in a warning. Validation instead fails
on malformed related XML or an undefined style reference, including in a
package produced by the Rust facade.

Story items are projections over the existing typed and retained package
sources. Body and comment items can expose owned XML serialized from their
typed owners. Other package-backed items borrow their exact subtree bytes and
can depend on namespace declarations on retained ancestors. A complex field
instead exposes one owned, namespace-complete paragraph projection because its
source can span sibling runs. Complete typed admission decides whether
controls, revisions, simple fields, and complex fields cross the preservation
boundary. Unmodelled and rejected subtrees remain opaque and byte-preserved.
Story-wide hyperlink projection retains the source byte position only while
building its result, then returns existing locations and link records in that
physical order. Relationship lookup remains scoped to the owning story part.
Private complex-field excerpts wrap exact sibling runs in one namespace context,
including inherited aliases, default bindings and authoritative foreign shadows.
Local declarations retain their original positions. The wrapper is never saved.
Text and hyperlink scans preserve the field marker interval inside that excerpt,
and hyperlink endpoints translate back only within the original source interval.
A hyperlink wrapper inside a complex field remains subject to the existing typed
source-equality admission proof rather than becoming an admitted field merely
because the private scanner can read that shape.

The main Word document reader accepts one namespace-correct `document` root
and one namespace-correct `body` child, rejects truncation, duplicate roots,
foreign lookalikes, and non-whitespace content outside that root, and retains
the first body `sectPr`. Later section owners remain opaque rather than
disappearing. Self-closing paragraphs, tables, and cells are modeled as empty
typed owners. Start-and-end forms are modeled only when their complete
attributes and content satisfy the same grammar. Otherwise their exact
namespace-complete subtree remains opaque. Header and footer references
recognize `r:id` only when the attribute is bound to the package relationships
namespace.
Body comparison excludes the source's direct final `sectPr` from its
interleaved content before appending the compared section properties. This
keeps the strict first-section reader from selecting a stale duplicate and
retains the tracked `sectPrChange` as the only final section owner. Comparison
wrappers for extracted body XML declare the standard relationships namespace,
so namespace-aware section parsing retains header and footer references.

Story text replacement runs on a staged clone. It resolves the owner,
fingerprint, one-element index path, item kind, and text-bearing capability,
patches only the selected owner, then serializes and reopens the complete
candidate before one commit. Missing or wrong owners, stale fingerprints,
invalid paths, bounds failures, kind mismatches, non-text items, XML failures,
and reopen failures leave the original document bytes and facade-owned package
state unchanged.

Scoped literal replacement uses the existing run-aware matcher through an owned
Document transaction. `try_replace_text_at` checks paragraph, table and block
control locations, including supported two-segment control paragraph paths.
A recursive facade handle maps by actual direct paragraph identity. Nested
control and table paragraphs have no two-segment location.
The hidden cell-coordinate entrance selects one physical cell or one of its
paragraphs. Whole-cell scope includes supported nested tables and controls.
Fields, drawings and opaque story items are not independently writable targets.
Matching retains existing wrapper, revision and field boundaries and never
joins paragraphs or searches another textbox story. A text box item also edits
the VML copy of the box in each later `mc:AlternateContent` branch, paired by
item slot after proving the same items and paragraph texts, and counts once.
Copies that cannot be paired, or a cell inside a text box with copies, are
refused. Count mismatch and zero matches publish nothing. Positive edits
serialize and reopen the complete candidate before one commit.

Main-source selection preserves original unselected bytes. Existing strict
namespace replay proves the current canonical physical owner inventory. Two
private existing-model projections replace the canonical and proposed retained
selected spans with the same collision-free schema-valid sentinel. Normalized
baseline equivalence and exact probe equality, with one observable sentinel in
each, prove the physical position even among identical siblings. Selected and
unrelated producer property normalization is supported. The sentinel never
enters the published candidate. Namespace, projection or reopening refusal
leaves the live model and package untouched.

Generic story content mutation uses the same package boundary. Existing-item
destinations are canonical flattened locations that must resolve to actual
direct owner children. The explicit story-end destination remains before body
section properties and expands a self-closing owner into a complete element.
Removal returns one namespace-aware owned fragment. Same-owner moves retain
its exact bytes, and clones rewrite document identities before insertion.
Relationship references can move or clone only when their original owner scope
is unchanged and complete. Content-control fragments validate the complete
serialized block grammar by expanded name, including retained root slots and
raw direct Word children. Foreign raw subtrees stay opaque. Invalid locations,
relationships, identities, XML, and reopen results discard the staged
candidate without changing the live document. Direct body inventory recognizes
section properties only among preserved nodes, so ordinary paragraphs and
tables do not repeat namespace-prefix scans. One clone still performs one
identity-freshening pass and one complete staged reopen.

HTML fragment insertion uses the same transaction. `HtmlImageResource` maps an
exact source string to caller-owned bytes and a filename. Data-URI images are
decoded within the aggregate input bound. Raster headers must pass
`oxml-media` probing, duplicate explicit sources fail closed, and unresolved
markup images produce ordered diagnostics with alternate text retained. Image
and hyperlink relationships are created under the selected story part. Every
image occurrence receives its own drawing identifier, while repeated source
strings reuse their relationship within one projection. Numbering, media,
content types, owner XML, and identifiers publish only after package reopen.

Story-scoped picture and hyperlink authoring resolves the relationship owner
from the checked `StoryId`. Body, cell, and text-box stories use the main
document relationship set. Headers, footers, notes, and comments use the
relationship set of their resolved part. Internal relationship validation
checks exact type, internal mode, normalized target, and target existence.
Image and hyperlink lookup applies the same owner boundary. New relationships,
media parts, content types, XML, and drawing identities publish only after the
staged package serializes and reopens.

Existing-picture replacement uses that owner boundary without changing the
drawing's relationship identifier. Unsupported byte signatures fail before a
candidate exists. Compatible unshared targets can be updated in place. Shared
targets and format changes use copy-on-write with the next deterministic
`/word/media/imageN.<ext>` name. The old part and its override are removed only
after the complete relationship graph proves the part unreachable. The new
extension and content type come from the sniffed bytes, and publication follows
a successful package reopen.

Hyperlink inventory follows the same rule. Each modeled story item discovers
its own hyperlink elements, then resolves each relationship identifier through
the checked story owner. Equal identifiers in the main part and a header or
footer therefore remain distinct package relationships.

Main-document authored occurrence provenance is derived from live
namespace-aware XML on every canonicalization. Removing a paragraph retires a
zero-use occurrence without deleting the relationship definition retained by
its owned fragment. Same-owner clones keep a shared relationship definition
and remap every live reference simultaneously. Each live authored picture
occurrence receives its own global `wp:docPr` identity. A producer may repeat
a drawing ID within one part. Its occurrence count is retained, and a staged
mutation rejects an increase in that count. Fresh authored IDs avoid the
package-wide occupied union. Producer-owned raw references remain fixed
occupants, and image part naming follows final serialized relationship order.

Theme and font authoring retain the relationship-resolved targets already in a
package. A missing theme or font table receives one collision-safe part,
content-type override, and internal main-document relationship on the staged
candidate. Embedded faces are obfuscated with the caller-provided OOXML font
key and live below a font-table-owned internal `font` relationship. Replacing a
face reuses its relationship and part only when both are exclusively
referenced. Shared producer relationships or parts remain intact while the
replacement receives a new package edge. Removing a face removes its owned
part only when no remaining internal package relationship targets it.

Flat OPC is a private Word facade codec over `OpcPackage`. Import first applies
the shared strict XML 1.0 lexical gate, then resolves expanded package names,
unique canonical absolute part names, exact data-kind selection, strict base64,
relationship ownership, and entry, part, and cumulative decoded-size limits.
Relationship parts populate `package_rels` and `part_rels`, never ordinary
parts. Export flushes a staged document and emits sorted fixed-prefix parts.
XML uses `pkg:xmlData` and opaque bytes use `pkg:binaryData`. Relationship-owned
alternative-format import targets remain opaque and use `pkg:binaryData` even
when their media type ends in `+xml`. Namespace bindings inherited from Flat
OPC wrapper elements are materialized only when the extracted XML payload uses
them in qualified names or markup-compatibility namespace-bearing values. A
reopened Flat package cannot claim retained cryptographic signature validity
because the original content-types byte ordering is unavailable.

ODT is a ZIP package but not an OPC package. The private `rdocx` ODT reader
therefore indexes it directly with the workspace `zip` dependency and does not
create an `OpcPackage`. The index rejects unsafe or duplicate names, non-files,
unsupported compression, encryption, and configured expansion-limit violations
before XML parsing. It requires the exact root `mimetype` and `content.xml`,
checks manifest encryption state, and reads only styles, content, manifest, and
referenced image parts. The resulting Word document is saved and reopened
through the normal OPC owner before it is published.

The private ODT writer also stays outside OPC. It writes the stored `mimetype`
entry first with no extra field, followed by deflated `content.xml`, image
entries in encounter order, and deflated `META-INF/manifest.xml`. The manifest
names exactly the root, content, and emitted images. Ordered style allocation,
fixed namespace prefixes, fixed ZIP metadata, and bounded retained output make
two writes of one document byte-identical.

MHTML is MIME rather than OPC. Its private `rdocx` reader requires one bounded
`multipart/related` entity with a unique HTML root and unique normalized
Content-ID and Content-Location identities. It accepts folded headers and the
base64, quoted-printable, 7bit, and 8bit transfer forms, but rejects ambiguous
roots, duplicate identities, trailing delimiter material, unsupported charsets,
unresolved contained resources, and every resource fetch. The writer emits
CRLF, stable source order, 76-column base64, deduplicated image resources, and
a collision-scanned content-derived boundary.

ODP uses the same non-OPC ownership rule in `rpptx`. Its reader requires the
first stored presentation mimetype, indexes all safe unique entries before XML
projection, and enforces caller-selected entry, part, and total expansion
limits. Its writer emits `mimetype`, fixed-prefix `content.xml`, sorted image
entries, and the exact manifest with deterministic ZIP metadata. Path saves
stage and sync a sibling file before portable atomic replacement.

EPUB output is also ZIP but not OPC. The private `rdocx` writer emits the
uncompressed `mimetype` entry first, followed by the container, package,
navigation, stylesheet, source-ordered spine items, and deduplicated media.
Entry names, order, compression, timestamps, identifiers, XML attributes, and
metadata fallbacks are deterministic. Output is bounded while ZIP seeks and
writes. Input, auxiliary projections, relationships, media, list depth, and
intermediate XHTML are bounded before their export allocations. Generated XML
rejects forbidden XML 1.0 characters, and external hyperlink targets require a
syntactically valid allowlisted absolute URI. A path save publishes fully
serialized bytes through a same-directory atomic replacement.
Heading labels are assembled only from bounded direct runs that survive the
projection, so dropped content-control text cannot enter navigation or spine
metadata. Referenced media is accepted only when byte sniffing and structural
validation agree on core PNG, JPEG, or GIF. PNG validation requires four-letter
ASCII chunk type codes with an uppercase reserved byte, one first IHDR, legal
critical-chunk order and counts, contiguous IDAT chunks, and one terminal IEND.
An indexed PNG palette cannot exceed the capacity declared by its IHDR bit
depth. JPEG validation permits exactly one leading SOI marker, requires a valid
frame before the first scan, and requires a terminal EOI. Baseline and
progressive frame types are accepted. GIF image descriptors require nonzero
width and height. GIF image data requires an LZW minimum code size from 2
through 8 and at least one non-empty data sub-block. Extension fallback is not
used. SVG and every malformed or unsupported image are diagnosed and omitted.
Heading-to-spine assignment and source-anchor lookup are linear in the accepted
source size. Ordered-list counters retain their numbering identity across
ordinary block interruptions, while deeper counters restart when a parent item
advances. Hyperlink spans are validated against their paragraph before the HTML
projection can expand them.
Page-break elements are lifted out of paragraph and inline formatting before
the XHTML documents are packaged, so spine items retain conforming flow
content.

## Generalising the constructors

The only docx-specific code in the existing crate is two constructors. They are
replaced by:

```rust
impl OpcPackage {
    /// Empty package: minimal content types, no parts, no relationships.
    pub fn new() -> Self;

    /// Package whose officeDocument relationship points at `part_name`.
    /// Package-relative, no leading slash: "word/document.xml", "ppt/presentation.xml".
    pub fn with_main_part(part_name: &str, content_type: &str) -> Self;
}

impl ContentTypes {
    /// Only the two universal defaults, "rels" and "xml".
    pub fn minimal() -> Self;
}
```

The Word presets remain private package helpers in
`crates/rdocx/src/document.rs`. The public `WordCreationProfile` chooses the
minimal or Word-compatible graph, while `WordPackageClass` supplies the exact
main-part content type. The PowerPoint presets remain in
`crates/rpptx/src/package.rs`.

Rejected alternatives, recorded so they are not revisited: a `PackageKind` enum
forces the leaf crate to carry every format's content-type table and grows a
variant per format. Feature-gated `new_docx` / `new_pptx` helpers make a leaf
crate feature-conditional for two functions' worth of string constants, and
features are additive so a workspace containing both compiles both anyway.

## What transfers unmodified

**`main_document_part()`** keys off the `officeDocument` relationship type,
which PowerPoint uses for `/ppt/presentation.xml`. It reads a `.pptx` today.

**`resolve_rel_target(source_part, target)`** joins a relative target against
the source part's directory and collapses `.` and `..`. It already handles
`../slideLayouts/slideLayout1.xml` from `/ppt/slides/slide1.xml` correctly, and
it clamps a traversal that escapes the root rather than allowing zip-slip.

**`rels_path_to_part_name`** and its inverse are generic path algebra.

`Relationships::from_xml` retains the complete source bytes beside the parsed
relationship list. `to_xml` returns those bytes while the semantic list is
unchanged and validates identifiers before either path. Adding, removing, or
editing a relationship switches to schema-ordered serialization. This keeps a
no-op package save byte-preserving without allowing stale relationship XML
after a graph mutation.

## Relationship types

`rel_types` stays one flat module, grouped by comment. The existing thirteen
constants are kept. Added:

```rust
// Package-level. Note the package namespace, not officeDocument.
CORE_PROPERTIES, THUMBNAIL, DIGITAL_SIGNATURE_ORIGIN, DIGITAL_SIGNATURE

// Shared officeDocument
EXTENDED_PROPERTIES   // docProps/app.xml
CUSTOM_PROPERTIES     // docProps/custom.xml
COMMENTS              // Word comments part
WEB_SETTINGS          // Word web settings part
GLOSSARY_DOCUMENT     // Word glossary document part
DIAGRAM_DATA, DIAGRAM_LAYOUT, DIAGRAM_QUICK_STYLE, DIAGRAM_COLORS
DIAGRAM_DRAWING       // Microsoft 2007 cached diagram drawing
OLE_OBJECT, CONTROL, STRICT_OLE_OBJECT, STRICT_CONTROL
ACTIVEX_CONTROL_BINARY, VBA_PROJECT
VBA_PROJECT_SIGNATURE, VBA_PROJECT_SIGNATURE_AGILE

// PresentationML
SLIDE, SLIDE_LAYOUT, SLIDE_MASTER, NOTES_SLIDE, NOTES_MASTER,
PRES_PROPS, VIEW_PROPS, TABLE_STYLES, HANDOUT_MASTER,
POWERPOINT_COMMENTS, POWERPOINT_AUTHORS, AUDIO, VIDEO, POWERPOINT_MEDIA
```

The four `dgm:relIds` values resolve only in the relationship scope that owns
their graphic frame. A schema-position-owned `dsp:dataModelExt/@relId`, when
present, resolves in that same scope through the Microsoft 2007
`diagramDrawing` relationship. Checked node editing and cross-presentation
transfer require every present role to be internal and to have its exact
relationship type before staging package changes.

A `content_types` constants module is added alongside, so neither format crate
hand-types the long MIME strings. It includes the modern PowerPoint comments
and authors MIME types as well as the handout-master type.

The Word facade resolves at most one glossary relationship from the main
document. The relationship must be internal, its normalized target must not
escape the package root, the part must exist, and its override must use the
Word glossary content type. Duplicate, external, traversal-shaped, missing,
wrong-type, and malformed-root graphs fail before document mutation.
Glossary creation reserves one internal relationship, part and override as
one bundle. Last-entry removal retains that valid empty part. Creation,
replacement, removal and placeholder binding enter canonical staged
preparation and provenance-reconciling reopen before publication. Fragment
content resolves its dependency closure from the physical glossary owner.
Typed creation rejects dependency references requiring a source package.
Fragment creation, fragment body update, reverse capture and insertion omit
qualified comment markers before review dependency capture. Omission includes
selected note and textbox content and retains mixed run properties, wrappers,
foreign lookalikes, comments and PIs. New direct replacement bodies follow the
same marker omission policy. Dependency-free typed creation retains its
existing refusal of comment-bearing bodies. Metadata-only updates keep the
existing body and review state.

A glossary comments relationship identifies an independent physical review
owner only when exactly one valid internal target has qualified comment
content. The same strict thread and companion graph handles main and local
cleanup. Shared physical review dependencies, ambiguous, external, missing or
malformed local mappings refuse. A physical note source with qualified comment
markers cannot belong to both review owners. Shared unannotated notes remain
valid. Local companion relationships without local
comments definitions also refuse. With no local review edges, legacy
main-owned glossary markers retain strict main ownership proof. Local producer
comment parts and relationships remain retained after complete owned cleanup.
Forward transfer bounds marker omission and initial note selection to the
namespace-complete prepared body content being transferred. Retained root
producer payload stays outside that interval and glossary dependency capture.
Generic transfer keeps its existing dependency admission. Final section properties join
omission only when the fragment includes them, retaining their physical
section owner and existing generic identity admission. Forward transfer derives its recursive note closure from the prepared selected
main body, omitting only referenced note owners and their selected textbox
content. Unselected owners in the same physical note part retain their bytes
and do not affect omission refusal. Reverse capture uses the same bounded
main-owned note traversal before constructing a main-owned fragment. A glossary-local note
relationship refuses projection through main numeric note ownership. It never
selects an unrelated equal-id main note.


Both facades resolve core properties through the package-level
`CORE_PROPERTIES` relationship and retain its normalized target. Immutable
property access leaves the source part bytes untouched. Mutable access marks
the typed `CoreProperties` model for serialization to that target with its
content-type override. A package that creates metadata without an existing
relationship uses `/docProps/core.xml` and adds the missing package
relationship. If that conventional part name is already occupied without the
core-properties relationship, serialization returns an error before changing
the package. `CoreProperties` models all fifteen elements of the core-properties
schema, from title, subject, creator, keywords, description, last modified by,
created and modified to category, content status, identifier, language, last
printed, revision and version, each as text. A rewrite after one
change therefore keeps every other property. Unset values are not written, so
a part holding only the original eight serializes byte for byte as before. The
Word `Document` and the layout engine context it holds twice keep the model
behind a `Box`, so the larger model does not grow `Document` against the debug
test-thread stack budget that `CT_PPr` boxing also protects.

The Word facade applies the same package-level ownership rule to application
and custom properties. New property families reserve collision-safe part and
relationship identities before publishing typed state. Removing a whole
family deletes only its resolved part, exact package relationship, and exact
content-type override. Removing the final custom property prunes that graph
only when the facade created it. Producer-owned empty parts remain present.

The Word facade owns external hyperlink relationships at the document part
boundary. `Document::add_hyperlink_relationship` allocates the relationship,
and `Paragraph::add_hyperlink` writes a schema-ordered `w:hyperlink` that
references it. The same paragraph writer emits explicit hard breaks as run
content. Both operations use the existing package-preserving save path, so
unmodelled parts and relationships remain intact.

Numbering state is also fail-closed at this boundary. Definition and instance
create, update, and remove operations run on a staged candidate. They validate
the complete numbering graph, including level ranges, placeholders, unique
identifiers, override ownership, style references, and live paragraph
references, before package mutation. Rejecting an invalid operation does not
create a numbering part, relationship, content-type entry, or consumed
identifier.

Paragraph-style numbering links are one cross-part transaction. The facade
writes the style's `w:numPr` and the effective definition or replacement
level's `w:pStyle` on a staged document, validates both style and numbering
graphs, serializes the candidate, and then publishes it. Unlinking requires the
same exact style, instance, and level tuple and removes both edges together.
A style `w:numPr` that holds a `w:ilvl` and no `w:numId`, as python-docx's
default `Subtitle` does, owns no link. It takes the instance of its `basedOn`
chain, as layout and Word do, and has no numbering when the chain has none.
Linking such a style replaces its level, and unlinking it later removes its
`w:numPr` rather than restoring the level. When a numbering level names such a
style back, unlinking removes only that name. Unlinking one that no level names
changes nothing. A paragraph `w:ilvl` without a
`w:numId` takes the instance of its style chain in the same way, which is how a
repeated template paragraph is checked.

Numbering parsers retain namespace declarations and compatibility attributes
from modelled containers. Unknown level and override children use their schema
slots, while abstract-definition, instance, and root children keep
insertion-aware boundaries. Typed mutation therefore preserves producer
extensions and unchanged imported overrides byte for byte. `CT_Lvl` and
`CT_NumLvl` serialize their standard children in schema sequence. Identifier
allocation uses the next value after the maximum when available and the first
unoccupied value when the maximum is `u32::MAX`.

Paragraph properties retain imported `w:numPr` leaves and unmodelled children
in their original namespace and schema positions. Typed `w:ilvl` and `w:numId`
updates remain before retained `w:numberingChange` and insertion properties.
Self-closing `w:numPr` carriers copy any inherited namespace binding required
by a retained root attribute onto the serialized carrier.
An unchanged plain numeric leaf may use the typed serializer's layout,
while malformed or extended source leaves remain byte-exact.

Every standard `w:numFmt` token has a typed representation. Producer-defined
values remain typed as their original token rather than being substituted with
decimal numbering. Numbering serialization writes either form back unchanged.
Layout and text exporters emit no marker for a format whose rendering
semantics are not implemented.

The main document reader retains root, body, and modeled-owner namespace facts
that preserved raw descendants depend on. Save replays those declarations on
their logical owners through insertion, removal, and reordering without
rewriting the raw subtree bytes. Prefix aliases, nested shadows, and ordinary
namespace URI escaping are resolved by the XML parser. Serialization fails
closed when owner identity or a serializer prefix binding cannot be preserved
safely, leaving the opened package bytes authoritative.
The main document, header, footer, comments, footnotes, endnotes, and styles
roots retain their other attributes, such as `mc:Ignorable`, in source order.
A rewrite keeps each compatibility attribute with declarations for every
prefix it lists. The typed Word part serializers write without indentation, so
rewritten document, story, comment, style and numbering parts retain compact
layout and gain no whitespace-only text between elements. An unchanged part
keeps its exact producer bytes, including an empty self-closed comments root.
Story insertion reads the retained main-part XML while it matches the typed
model. Picture insertion adds one paragraph at its body boundary. Canonical
relationship and drawing identifiers are patched into that retained XML, so
unrelated producer toggles, empty properties, and default root declarations
stay in the package after the picture is saved.

Modeled paragraph, run, table-row, and section-property owners retain every
ordered root attribute, including producer identity, revision-session, foreign,
and unqualified attributes. Retention uses the existing raw-preservation carriers
without exposing the attribute record as child XML. Expanded names govern
duplicate rejection and authored paragraph identity precedence, so an authored
`paraId` replaces only the retained attribute with the same namespace and
local name. Typed child mutation leaves all other retained root attributes in
source order. A table row keeps its record in its raw-child list at a position
no cell boundary reaches. Every reader that treats raw row children as content
skips it: comparison row signatures and boundaries, the retained table layout
cache, the row diagnostics of the MHTML, ODT, RTF and EPUB writers, and the
rich merge row-region markers. The rich merge region-marker and whole-paragraph
fragment checks skip the record of a paragraph the same way, so a paragraph
Word wrote still holds a region marker or a fragment field.

Retention covers attributes, not namespace bindings. A declaration is recorded
only when a retained attribute uses its prefix, because the alias machinery
already materializes a binding onto every element that needs one, and recording
a declaration a child carries for itself would emit it twice. A root carrying
nothing but declarations retains no record at all. On the way back out, the
canonical `w` and `w14` bindings are not copied onto the written element, since
the part root that owns the element declares them in what Word and python-docx
write, and the element's own `w:` name and the authored identity write make the
same assumption. A binding of either prefix to another URI is still copied. Together these keep a
reopened save byte identical to the save it was read from. A part root that
does not declare `w14`, such as one rdocx wrote or one under an element that
declared the prefix itself, gains the canonical declaration when the written
content uses the prefix. The serializers of the document, header, footer,
note and comment parts and the comparison output of every story add it, so the
written part stays namespace well formed. During complete part serialization,
a retained attribute also omits same-URI bindings already guaranteed by that
part root: `r` and `mc` for the main document, canonical `wp` when that root
actually writes it, `r` and conditionally canonical `wp` for headers and
footers, and `r` for footnotes and endnotes. Standalone paragraphs and
comments keep their own bindings. A scoped serializer context restores its
previous guarantees after nested calls and errors. The retained-attribute
regression checks one declaration of each canonical prefix at the document
root, alongside producer attributes, a new binding, and an unchanged
unmodelled child.

A paragraph cut out of its part and parsed on its own carries none of the
declarations of its part. The table-of-contents rebuild adds the bindings the
instruction paragraph inherits to its start tag before it parses it, so every
run attribute resolves as it does inside the part. The text-box anchor reader
does the same for each text-box paragraph. The text-box replacement and
template walkers still parse such a paragraph in the default scope, which names
`w` as the Word prefix without binding it. For that scope the capture resolves
a Word prefix that the scope names without a binding to the WordprocessingML
namespace, and an explicit binding always wins. A run attribute under any other
prefix, such as `w14`, a foreign namespace or a second WordprocessingML alias,
still fails those two walkers. The replacement walker parses a text-box
paragraph without its start tag. It writes an edited paragraph back under a
start tag that carries the attributes it was read with, identities and local
declarations included, since the paragraph returns to the scope it came from.

An unknown default namespace declared on the document root is classified by
its effective lexical scope before canonical serialization. An unused root
default may be omitted without blocking a typed mutation. An unprefixed element
that inherits it keeps the binding live and blocks modified serialization.
Nested default declarations shadow the root declaration, including when they
repeat the same URI, and unprefixed attributes never use a default namespace.
Malformed or ambiguous declarations fail closed. A story splice into the main
part also publishes canonical XML, so it applies the same root default
classification before it publishes. Like a modified save, it also refuses any
declaration on `w:body`, which the canonical body drops, and a root `w`, `r` or
`mc` declaration bound to another URI, because the canonical root rebinds those
prefixes. After a successful canonical publication, by a save or by a story
splice, the document refreshes its root and body namespace facts from the
published main-story bytes so a later save applies the same classification.

Paragraph line spacing retains the signed integer path required by
WordprocessingML and accepts one bounded producer deviation. A plain signed
decimal `w:spacing/@w:line` value is normalized with exact decimal arithmetic
to the nearest integer twip, with exact halves rounded away from zero.
Exponent notation, malformed forms, non-finite spellings, and numeric values
outside the signed 32-bit range remain errors. The modeled value serializes as
one canonical integer while other modeled integer measurements use the same exact rounding rule.

Direct paragraph `m:oMath` and `m:oMathPara` children use that same owner and
boundary discipline. The reader accepts any prefix bound to the Transitional
OfficeMath namespace. Canonical typed writes use `m:` and replay the inherited
bindings needed by retained raw content. A conflicting `m` binding, malformed
grammar sequence, foreign same-local-name node, or legacy Equation Editor
object remains unmodelled raw XML. Run-boundary collapse rebases both the raw
child position and the equation projection so later mutation still replaces
the correct source node. Repeated-child raw slots remain ordinal boundaries
when callers insert or reorder typed values. If callers shorten a collection,
every now-unreached higher slot is emitted at the retained owner's tail rather
than discarded.

Presentation MathML conversion is a separate facade boundary rather than an
OPC part. Its reader resolves the W3C MathML namespace by expanded name, rejects
DTD and unresolved entity input, and applies byte, event, depth, node, text,
matrix, and diagnostic limits. Its writer emits one default MathML namespace,
stable attributes, ordered children, and explicit `mo` fences. `mfenced` is
accepted only on input. Unsupported safe content is diagnosed, and descendants
are retained only at the declared transparent `semantics` boundary.

Logical owner identity includes exact normalized raw marker multiplicity and
the resolved namespace facts of owner-dependent element and attribute uses.
A same-URI declaration already local to a retained subtree remains independent
and does not inflate the owner marker set. Candidate replay promotes only the
captured marker multiset, including when preservation made an inherited use
self-contained. Same-URI and different-URI decoys, duplicate owners, fixed
prefix shadows, and undeclared nested prefixes remain fail-closed.

Word table readers carry ancestor bindings through tables, rows, cells,
content controls, borders, and raw properties. Their writers retain unknown
table, row, and cell facts at insertion-aware schema boundaries and keep
malformed row revision markers in their original slots. Drawing relationship
projection requires the direct WordprocessingML and DrawingML picture path and
the Office relationships namespace. Foreign attributes, descendant
lookalikes, ambiguous pictures, and ambiguous blips remain opaque.
The typed `tblHeader`, `cantSplit`, and `noWrap` values use the shared on-off
vocabulary. Absence remains distinct from explicit true and false, and present
values write canonically in their existing row or cell property slots.

Raw Word run children receive semantic classification only at the OXML parse
boundary. A WordprocessingML `pict` is classified as a legacy horizontal rule
only when its in-scope expanded names identify exactly one VML `rect` with an
enabled Office `hr` attribute and whitespace otherwise. The classification is
stored in the existing raw-child position sidecar, while the subtree bytes and
ancestor namespace ownership remain unchanged. Foreign, malformed, numeric,
false, visible, or structurally ambiguous content stays unmodelled raw XML.
Authored mixed runs keep typed children and raw boundaries in one logical
sequence. Because `w:fldSimple` and complex field sequences are paragraph
children, the writer emits physical `w:r` segments around each field. Direct
run properties are copied to authored segments and the cached field result.
Positioned raw children stay on the same side of each field boundary. Legacy
unordered raw children remain at the final run boundary. Readers remain prefix
tolerant and writers use the fixed Word prefixes.

Run splitting uses the same encoded raw-child position sidecar. A second
private classification bit associates each parsed `mc:AlternateContent`
drawing projection with its verbatim compatibility block, so both move to the
same split run while serialization still writes the raw block exactly once.
Tabs, breaks, field markup, drawings, references, symbols, and unknown children
remain zero-width and keep their relative schema positions around literal text.

The Word facade resolves an existing comments part through the main document's
`COMMENTS` relationship and retains the normalized target. Saving serializes
the typed comments model back to that target with its content-type override
once the model no longer matches the part. Every public output shares that
test, so a save without a comment edit keeps the part byte for byte and leaves
a package signature over it valid. The model preserves unmodelled attributes
and children at their insertion boundaries, and a rewritten root keeps its
producer declarations and compatibility attributes in source order. Comment
range and reference anchors remain ordered among neighbouring paragraph and
run XML. A document without a comments relationship does not gain a comments
part, relationship, or override during an ordinary save.

Story-qualified marker edits resolve both endpoints through the same body,
table-cell, text-box, header, footer, note or comment owner. Block content
controls expose their contained paragraphs under a second path segment in
each owner. Namespace-resolved scans pair bookmark, comment, permission and
proofing markers in accepted-view order. Rewrites splice only the affected
paragraph bytes into its owning part and retain unrelated siblings and
unsupported markers. Invalid owner paths, crossing or unmatched pairs, stale
snapshots, duplicate names and exhausted identifiers reject the staged edit
without changing the live package.

The Word facade resolves an existing settings part through the main document's
`SETTINGS` relationship and retains the normalized target instead of assuming
`/word/settings.xml`. `rdocx-oxml` projects valid document protection metadata
while retaining the complete settings bytes as the serialization source.
Saving writes those bytes only to the resolved existing target. Unsupported
protection modes, unsupported algorithm enum values, and malformed numeric
metadata remain opaque and byte-identical. A document without a settings
relationship does not gain a settings part, relationship, or content-type
override during an ordinary save.

The settings model also projects one valid `m:mathPr` child. Typed defaults
cover the math font, display justification, unsigned margins and spacing,
small-fraction and display toggles, and integral and n-ary limit placement.
Mutation replaces every existing math-properties occurrence with one
schema-positioned subtree while preserving unrelated settings bytes. Creating
defaults without a settings relationship allocates a collision-safe settings
part and adds the relationship and content type through the existing package
path.

The bounded settings authoring surface also projects document variables,
compatibility settings, default tab stop in integer twips, character spacing
control, and theme font languages. Reads accept in-scope Word namespace
aliases. Each mutation replaces only its modeled child or repeated child set,
uses fixed `w:` prefixes for new XML, and inserts at the schema position.
Unmodeled children inside `w:compat` and `w:docVars`, plus every unrelated
top-level settings child, retain their bytes and namespace context.

One valid `w:updateFields` child is a typed optional on-off value. Absence reads
as `None`, a bare element reads as true, and the complete shared false
vocabulary reads as false. Setting a value writes one fixed-prefix child after
`w:characterSpacingControl` and before `w:compat`. Setting `None` removes only
the modeled occurrence, while an already absent value is byte-identical and
does not allocate a settings graph. Duplicate or malformed producer forms remain
unmodelled and byte-identical, and their mutation returns an error rather than
collapsing ownership. Every facade change uses the staged settings candidate,
including collision-safe part and relationship allocation, before commit.

The bounded surface closes over the thirty-one top-level names in
`SUPPORTED_SETTINGS`, the complete `CT_Compat` on-off family, and the thirteen
authored `w:mailMerge` members. Mail-merge members are written one at a time at
their own schema positions, because `w:dataSource`, `w:headerSource` and
`w:odso` sit among them and stay preservation-only. Every typed member has a
paired remover that deletes only its own occurrence, and removal of a
duplicated or malformed member returns an error and changes no byte. Reads
accept in-scope Word aliases, and a same-local-name element in a foreign
namespace is never taken as the modeled child.

The Word facade resolves an existing web settings part through the main
document's `WEB_SETTINGS` relationship and retains the normalized target
instead of assuming `/word/webSettings.xml`. A document without that
relationship gains no web settings part, relationship, or content-type override
during an ordinary save, and the fresh Word-compatible profiles keep their exact
existing part inventory. The first authored value allocates a collision-safe
part through the same staged boundary, and removing the last value prunes the
part, its relationship, and its override when this facade created them.
`w:frameset` and `w:divs` are preservation-only. A read-only division-identifier
projection over the retained `w:divs` subtree reports whether a paragraph
`w:divId` resolves.

The font-table reader accepts any in-scope Word and relationship namespace
prefixes. It models font names, alternate names, family, pitch, and the four
embedded-face slots. New XML uses fixed `w:`, `r:`, `rdocx:`, and `mc:`
prefixes in schema order. The root merges producer MCE tokens and declares
`rdocx` ignorable. Producer attributes and unmodeled children remain in their
original relative slots. The `rdocx:` attributes retain the caller's explicit
authorization fact and exact license identity across save and reopen.

Watermark authoring follows the document-to-header graph rather than assuming
conventional header names. The facade materializes a missing default, first, or
enabled even header only at the first section that needs that same-type variant.
Later omitted references keep Word's same-type inheritance and do not receive a
blank override. Each image relationship belongs to its owning header part, and
its target is relative to that part even when a producer uses a custom header
path. Settings values controlling even headers are namespace checked and XML
decoded before selection.

Header and footer text replacement retains the section-reference kind through
enumeration, load, and save. A referenced relationship is eligible only when it
is internal and has the exact header or footer type implied by that reference.
Cross-type relationships and unrelated parts with header-shaped XML remain
untouched and cannot shadow a later valid relationship. Text, raw XML, image,
and background-image setters apply the same eligibility rule before reusing a
referenced part. An ineligible slot receives a fresh collision-safe part name.

Per-section story operations resolve only internal relationships whose type and
target root match the requested header or footer family. Linking reuses one
exact-type relationship. Inheritance removes the direct reference and exposes
the preceding same-type story. Unlink and replacement copy the effective story
XML into a collision-safe part, copy its part-local relationship set, rebase
internal relative targets, and allocate fresh drawing identities. Removal
installs an explicit empty story so it cannot expose inherited content. Pruning
removes only unreachable facade-owned relationships, parts, content types, and
media. Shared, producer-owned, opaque, and still-reachable graphs remain intact.
Every operation publishes only after the staged package serializes and reopens.
The typed `w:evenAndOddHeaders` setting reads namespace aliases and writes a
fixed `w:` prefix in its schema slot.

Section removal preserves the following section's effective header and footer
behavior before deleting a non-final boundary. For each default, first, and even
variant, resolution takes the first reference whose relationship has the exact
header or footer type, is internal, reaches an existing part with the expected
story root, and parses successfully. Missing, external, cross-type, absent-part,
malformed-target, and later duplicate references cannot displace an earlier
usable story. A missing usable override on the following section receives the
effective reference before package pruning. Pruning deletes only a facade-owned
header or footer relationship and part that no modeled or opaque reference can
still reach. Shared and producer-owned targets remain intact. The boundary,
relationship, part, content-type, and authored-identity changes publish only
after the staged package serializes and reopens successfully.

An authored watermark owns only a VML shape whose expanded name is `v:shape`
and whose unqualified id is `rdocx-watermark`. Replacement patches that exact
byte range in the original header, leaves tables, controls, root attributes,
namespace declarations, unrelated VML, and other producer bytes in place, and
keeps every emitted shape-type reference resolvable. Text and image operations
stage all header, relationship, media, and content-type changes on a cloned
package. A missing part, invalid dimension, parse error, or serialization error
leaves the live document and package unchanged.

Story-scoped picture options are validated before media publication. The
owning story receives the image relationship, while crop rectangles remain in
the picture blip fill and floating geometry remains in the WordprocessingDrawing
anchor. Authored text boxes carry local namespace declarations and a complete
VML shape-type definition so either compatibility branch is independently
resolvable. Section-aware text watermark replacement touches only the selected
API-owned shape. It neither synthesizes unrelated header variants nor enables
the document-wide even-and-odd setting.

Comment deletion uses namespace-qualified raw source markers across every
relationship-resolved Word story, including preserved wrappers and revisions.
The same inventory validates orphan roots without interpreting accepted-view
ranges. Point references are valid and linked replies need no separate anchor.
Deletion of a demonstrably undefined marker follows the existing editing
lifecycle because there is no definition to orphan. Absence is proved from
qualified raw definition entries at the actual relationship target. Missing
or ambiguous owned sources refuse, and CLI validation still reports dangling
markers.
Complete thread removal splices only owned definition and companion rows,
using last-paragraph ids and commentsIds durable linkage for
commentsExtensible. Actual internal relationship targets remain authoritative.
Ambiguous linkage, a surviving endpoint or reference, or unrelated anchors
inside a removed comment definition refuse atomically. Unknown root and entry
payload remains unchanged outside the selected owned rows.

Threaded comments add a document relationship using the Microsoft
`commentsExtended` relationship type. The facade retains its resolved target
and writes the standard
`application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`
content type at that exact part. New comment state creates both relationships
and both overrides together. An ordinary save retains an accepted standard
override and unrelated comments-extended sidecar XML. Existing custom targets
remain authoritative, and removal of the final API-owned thread removes only
the parts, relationships, and overrides created by the typed model only when
no opaque root attributes or children remain. Imported empty parts and
unrelated companion payload are retained.

Modern PowerPoint collaboration follows two independent relationship scopes.
The presentation part owns at most one Microsoft authors relationship, and
each commented slide owns one Microsoft comments relationship referenced by
the slide's typed `p188:commentRel`. Existing internal targets are normalized
and retained instead of being replaced with conventional names. External,
missing, wrong-type, duplicate, or shared comment-part ownership fails before
the live presentation changes.

Creating the first author uses `/ppt/authors.xml` only when that path is free.
Creating the first comment on a slide allocates a free positive
`/ppt/comments/commentN.xml` suffix through `MediaNamer`. Both operations stage
the part, relationship, content-type override, typed XML, and reopen before
commit. A matching MIME type does not make an unlinked conventional part safe
to overwrite. The notes-master and handout-master roots are likewise resolved
from the presentation relationship graph without assuming their filenames.

Notes and handout export requires exactly one internal relationship of the
expected type and content type at each required edge. This includes
presentation to master, master to theme, notes slide to notes master, and notes
slide back to its source slide. Noncanonical internal part names remain valid.
Before layout, notes-master and notes-slide relationship scopes are copied into
a collision-free transient owner scope with absolute normalized targets and
rewritten relationship ids. This keeps equal source ids independent and never
changes the opened package.

Content-control data binding follows the existing package graph rather than a
conventional filename. The main document's custom XML relationship resolves an
item part. That item's custom XML properties relationship resolves the
properties part whose root `ds:itemID` identifies the datastore. Matching is
case-insensitive and accepts the optional braces used by producers, while a
same-local-name attribute or element in another namespace is not metadata.

The binding evaluator accepts namespace-aware absolute child paths with
optional one-based numeric child indices. Functions, wildcards, descendant
axes, and general predicates are outside the contract. Prefix mappings and
in-scope namespace shadowing resolve every path step by expanded name. The
selected final element must be unique and contain only simple text or be empty.
The facade replaces only that element's text span, expands an empty selected
element when needed, and retains every unrelated custom XML byte exactly.
Display and custom-part changes are staged together and committed only after
all selected bindings serialize and reparse successfully.

## Part naming

**Numeric suffixes are allocated after the greatest positive parsed suffix,
never as `count + 1`.** Deleting slide 2 and then adding a slide must not
collide with slide 3. `MediaNamer` scans positive decimal suffixes in the
requested directory and stem, including `usize::MAX`, and ignores missing,
signed, zero, nonnumeric and unrelated suffixes. Ordinary packages allocate
`1 + max(existing suffix)`. At the finite boundary, checked increment wraps
from `usize::MAX` to 1 and skips every occupied parsed suffix until a free
positive number is found. The presentation facade uses this format-neutral
allocator. The Word facade applies the same suffix rule through its document
identifier owner so media, charts, workbooks, comment parts, and content-type
entries are reserved with their related identifiers as one candidate. Neither
facade creates `image0` or overwrites an existing numbered image part.

Word relationship identifiers are independent per source part. New owners
retain the established `rId0` first allocation, while an occupied owner
continues after its greatest numeric identifier and avoids nonnumeric producer
identities. Bookmark and comment identifiers start at zero. Drawing and
numbering-instance identifiers start at one, and abstract-numbering identifiers
start at zero. Imported definitions are scanned by expanded XML name, including
the unqualified `id` attribute on `wp:docPr`. Producer `wp:docPr` definitions
must be unique within one physical XML part, but the same normalized value may
occur in another part. Every accepted value joins the package-wide occupied set
so later authored drawings remain globally fresh. Other duplicate definitions,
exhausted ranges, and pending collisions fail before a staged candidate is
published.
Comment ids are facade identities rather than serialization-order values.
Authored values take the lowest unused nonnegative slot, and rdocx save and
reopen do not renumber them or their reply links.

Canonical part layouts:

```
docx                      pptx
/word/document.xml        /ppt/presentation.xml
/word/styles.xml          /ppt/slides/slideN.xml
/word/media/imageN.ext    /ppt/slideLayouts/slideLayoutN.xml
/word/charts/chartN.xml   /ppt/slideMasters/slideMasterN.xml
/word/embeddings/         /ppt/notesSlides/notesSlideN.xml
  WorkbookN.xlsx          /ppt/theme/themeN.xml
/word/comments.xml        /ppt/media/imageN.ext
/word/commentsExtended.xml /ppt/media/mediaN.ext
                          /ppt/charts/chartN.xml
                          /ppt/embeddings/WorkbookN.xlsx
                          /ppt/authors.xml
                          /ppt/comments/commentN.xml
```

Diagram parts retain producer-selected names while they remain owned by their
original scope. Slide duplication and bounded SmartArt transfer allocate a
fresh positive suffix beside each source diagram part when its resolved target
would collide in the destination. Equal image bytes may reuse a compatible
destination media part, but diagram XML parts never alias an unrelated
destination part merely because their names or bytes match.

Word comment part creation uses the conventional names when free and scans
numbered alternatives when either path is occupied. It never overwrites an
unrelated part merely because the conventional comment path exists.

Word chart assembly follows the same independent suffix rule as PowerPoint.
The document relationship targets `/word/charts/chartN.xml`, and that chart's
package relationship targets `/word/embeddings/WorkbookN.xlsx`. Both parts and
their content-type overrides are staged with the drawing on complete typed
document state. The mutation becomes visible only after the typed ChartML,
SpreadsheetML workbook, relationships, content types, structured drawing,
theme projection, and shared identifier owner all validate.

An authored Word chart requires one effective internal document-theme edge.
Its target must exist, carry the exact theme content type, and parse as a
DrawingML theme. A valid related theme is reused. Otherwise the facade stages
the Office default under a collision-safe `/word/theme/themeN.xml` name and
retargets the ineffective theme edge or allocates a new one. Source theme bytes
remain unchanged, and preserved charts do not synthesize themes.

## Media

`oxml-media` owns image-byte interpretation and bounded, format-neutral media
classification and naming.

```rust
pub enum ImageFormat { Png, Jpeg, Gif, Bmp, Tiff, Webp, Svg, Emf, Wmf }

impl ImageFormat {
    pub fn sniff(data: &[u8]) -> Option<Self>;
    pub fn from_extension(ext: &str) -> Option<Self>;
    pub fn extension(self) -> &'static str;
    pub fn content_type(self) -> &'static str;
}

/// Sniff first, fall back to the extension, default to PNG.
pub fn resolve(data: &[u8], filename: &str) -> ImageFormat;

pub struct ImageInfo {
    pub format: ImageFormat,
    pub width_px: u32, pub height_px: u32,
    pub dpi_x: Option<f64>, pub dpi_y: Option<f64>,  // None means the file declares none
    pub bit_depth: u8, pub channels: u8, pub has_alpha: bool,
}
pub fn probe(data: &[u8]) -> Option<ImageInfo>;

pub struct NativeSize {
    pub width_emu: i64, pub height_emu: i64,
}

impl ImageInfo {
    pub fn native_size(&self, default_dpi: f64) -> Option<NativeSize>;
}

pub struct MediaNamer { /* dir, stem, next */ }
impl MediaNamer {
    pub fn scan<'a>(dir: &str, stem: &str, existing: impl Iterator<Item = &'a str>) -> Self;
    pub fn next_part_name(&mut self, ext: &str) -> String;
}
```

**Sniffing beats the extension.** `resolve` uses detected image bytes before a
filename extension, so a `.png` that is really a JPEG receives the JPEG
extension and content type. Unknown bytes fall back through a recognised
filename extension and finally to PNG for compatibility.

Audio and video helpers keep safe MIME grammar strict while comparing known
type and subtype names case-insensitively. MP3, WAV, and ISO base media inputs
must carry their expected container signature. Unknown safe content types and
extensions remain opaque payloads rather than acquiring a decoder claim.

`rdocx::Document` scans existing `/word/media/imageN.ext` parts into a
`MediaNamer` when it opens. Every body, header, footer, and raw-XML image path
uses the allocator and registers the sniffed canonical extension and content
type before adding its relationship. HTML and layout extraction resolve MIME
from the stored bytes first, so a misleading package part name cannot override
the actual image format.

**`native_size` takes the DPI rather than baking one in**, because the right
default differs by consumer: python-docx assumes 72 when a file declares none,
while Word assumes 96. Each declared finite positive axis DPI takes precedence
over the caller default. Missing or invalid declared DPI falls back per axis.
The conversion multiplies pixels by 914400 EMU per inch and truncates toward
zero. The method returns `None` if either effective DPI is not finite and
positive, or if a converted dimension is outside the `i64` range.

`NativeSize` keeps the result dependency-free and exposes explicit EMU fields.
The PresentationML picture insertion path supplies 72 for python-pptx parity
without adding an `oxml-core` edge to `oxml-media`.

`rdocx::Document::add_picture_auto` is an additive convenience API that probes
the image and calculates `native_size(72.0)` before changing document state.
It converts the shared EMU result with `Length::emu` and delegates successful
insertion to the existing explicit-size `add_picture` path. Unavailable
dimensions return `rdocx::Error::UnavailableImageDimensions` with the supplied
filename before a media part, relationship, drawing, or paragraph is added.

`rpptx::Presentation` scans `/ppt/media/` into a content-hash `MediaStore` when
it opens. Insertion compares the complete byte string inside each hash bucket,
reuses an equal package-wide media part, and otherwise allocates the next
numbered part after the greatest occupied suffix. The sniffed canonical
extension and content type are registered with the package. Each source slide
creates or reuses its own internal image relationship to that shared part, with
a relative target resolved from the slide part name.
Cross-presentation slide import uses the same byte equality check for images,
audio, and video, including media stored outside `/ppt/media/`. The import
stages the destination package and remaps each copied part's relationship ids
once. An unsupported internal relationship graph fails before the destination
changes, while external targets retain their target mode and address.

The presentation HTML importer accepts image bytes only through
`HtmlImageResource`. The HTML source string is an exact lookup key. Missing
resources are diagnosed, duplicate keys and aggregate byte overflow fail
closed, and no URL or filesystem path from markup is fetched. Successful
images enter the normal presentation media insertion path with caller-supplied
filenames and explicit CSS geometry.

The Word HTML fragment importer uses the corresponding native
`rdocx::HtmlImageResource` slice plus self-contained image data URIs. It never
resolves a URL or filesystem path from markup. Explicit width and height use
the existing CSS pixel conversion. A missing dimension uses the probed native
size at 96 DPI. Malformed bytes, MIME mismatches, count overflow, and aggregate
byte overflow fail before the live document changes.

The Word MHTML importer accepts contained PNG and JPEG resources only when the
declared MIME type agrees with byte sniffing. Missing CSS pixel dimensions use
the Word 96 DPI native-size rule and explicit width or height attributes retain
their exact CSS pixel projection. Repeated equal image bytes reuse the same
MIME resource on export while each drawing keeps its own display geometry.

The PDF importer routes decoded JPEG and PNG image content through the same
package-wide presentation media store. JPEG bytes with `DCTDecode` are retained
directly. Bounded 8-bit `DeviceGray` and `DeviceRGB` image streams become PNG.
Each imported slide owns its image and external URI relationships, while equal
image bytes still deduplicate package-wide.

Slide removal considers only `/ppt/media/` targets reached from the removed
slide and its removed notes relationship scopes. A candidate part is deleted
only when no remaining internal package relationship reaches it. Pre-existing
orphan media outside that candidate set is left untouched. The facade rebuilds
its content-hash media index after the graph change, so a later insertion sees
the surviving package state.

Non-image PowerPoint media uses `/ppt/media/mediaN.ext`. Complete-byte hashes
provide candidate buckets, and reuse requires an exact byte match plus a
compatible extension and content type. Each picture retains independent poster
ownership. Embedded audio and video use the standard relationship plus the
Microsoft Office media relationship when the model requires both. Linked
sources preserve the exact external target and never fetch it.

Media facade mutations stage a cloned package and slide, serialize and reopen
the result, and publish it only after every picture, timing, relationship,
content-type, and payload change succeeds. Replacement preserves the shape id,
geometry, poster, and bounded playback settings. Removal prunes only payload
candidates made newly unreachable by relationships removed in that operation.
Shared targets, retained raw references, and producer orphans survive.

Header parsing is lifted from the PDF crate, where `jpeg_dimensions` and the
PNG IHDR reader are currently private. The JPEG walk classifies SOI, TEM and
RST0 through RST7 as standalone markers with no length field, and EOI terminates
the codestream. Length-bearing segments are validated for a present length of
at least two bytes and bounds within the input before the walk indexes or
advances. Fill bytes and truncated input return safely. Preserve these
invariants when the reader moves to `oxml-media`.

## Package integrity

Both facades expose a `validate()` that is cheap, non-panicking, and run
automatically under `debug_assertions` before `save`. It checks dangling
relationship ids, missing content-type overrides, relationship targets that
resolve to no part, and orphan media. `rpptx` adds its own presentation-specific
checks, listed in `06-presentationml-model.md`.

Presentation HTML conversion publishes only a serialized, reopened, validated
candidate. Projection failure returns `Error::Html` or an existing package
error before any partial presentation escapes. Default master, layout, and
theme parts remain byte-identical while new slide children follow the existing
fixed-prefix PresentationML serializers and shape-tree sequence.

Word MHTML conversion also publishes atomically. Import projects the selected
HTML root only after all resource references pass bounded preflight, then saves
and reopens the generated DOCX. Export reparses the complete MIME result and
serializes and reopens the source document before returning bytes. Path writes
stage the complete result through the shared portable atomic replacement path,
so an error cannot truncate an existing destination.

Word package-class and Flat OPC path saves use that same atomic replacement
helper. Both byte writers stage, serialize, and reopen their output before it
is returned or published. Flat OPC import constructs the complete package and
validates its Word class before a `Document` becomes observable.

Fresh Word-compatible construction validates the complete staged graph before
the `Document` becomes observable. The validator requires the exact owned part
and content-type inventory, one relationship of every required type, no extra
relationship scope, normalized internal targets, and an existing target for
every internal edge. Public conformance then serializes and reopens all four
classes and compares normalized package semantics.

PDF conversion has the same publication boundary. It builds a fresh candidate,
adds every source page in order, serializes, reopens, validates, and only then
returns `PdfImportResult`. Mixed effective page sizes, malformed graphs, active
JavaScript, and any declared resource-limit failure return `Error::PdfImport`
before a partial presentation can escape.

Executable presentation payloads are selected only through normalized internal
relationships with the exact expected type. OLE, control, ActiveX binary, VBA
project, and VBA signature targets reject external or root-escaping paths,
duplicate relationship identities, missing parts, and wrong relationship
types before inventory or mutation. Package-signature inventory requires one
internal origin whose relationship set contains only one or more internal
digital-signature relationships with distinct existing targets. Unrelated,
misplaced, duplicate, external, missing, or traversal-shaped signature graph
edges fail closed before signature evidence can be retained or removed.

The Word facade applies the same package rules from schema-positioned owners in
supported story parts. An OLE payload requires exactly one relationship-owned
`o:OLEObject` inside a valid run-owned `w:object`. An ActiveX control requires
one valid `w:control`, an exact properties content type, and exactly one
internal ActiveX binary relationship. A VBA project is owned by at most one
main-document relationship and may own at most one legacy or Agile signature.
Targets must be normalized internal Pack URI references with the exact
relationship and content types. Missing relationship scopes, ambiguous or
overlapping owners, shared removal targets, malformed signature graphs, and
malformed owner XML fail before a package mutation becomes observable.

Replacement preserves the normalized source-part and relationship identity,
target name, content type, and owner XML. Removal patches only the validated
complete owner range or the main-document VBA relationship, then deletes only
owned candidates made newly unreachable by that operation. Signature policy
either retains exact package and VBA signature bytes behind deterministic
invalidation markers or removes only the validated signature infrastructure.
Every mutation serializes, reopens, and re-inventories a staged package before
commit, so failure leaves the original package byte-identical.

The bounded SmartArt copy graph accepts only the five diagram relationship
types and internal images whose parts own no relationships. Traversal has cycle
protection and a shared 128-part ceiling in preflight and copy. Unsupported
internal charts, packages, media, OLE, custom parts, missing targets, and
external slide layouts reject before any destination part or relationship is
published.

Word table widths, cell widths, table indents, and default cell margins share
one exact signed-integer projection. Modeled integer measurements accept decimal
lexical values and round to the nearest integer, with exact halves away from
zero. The unrounded value must fit the target type. Exponent forms, empty
fractions, overflow, percentages, universal measures, and malformed input fail
with an element and attribute diagnostic instead of becoming zero. Missing
widths retain their existing default. Attributes are selected by the bound
WordprocessingML namespace, and serialization writes the canonical integer with
fixed `w` attributes in schema order while unmodelled table content retains its
stored bytes.

Word table grids recognize `tblGrid`, active `gridCol` children, their width
attributes, and `tblGridChange` by the bound WordprocessingML namespace.
Foreign same-local children remain unmodelled and retain their exact bytes.
One nonempty historical grid-change subtree is preserved, while a second
modeled change fails parsing rather than discarding history. A structurally
empty `tblGridChange` remains unmodelled in its original slot. Serialization
writes active columns first and the modeled historical change after them in
schema order.

A table style's conditional regions project all five `CT_TblStylePr` layers.
`w:pPr`, `w:rPr`, `w:tblPr`, `w:trPr` and `w:tcPr` are modeled and serialized
in that schema sequence, and a region whose five layers are unchanged since
parsing serializes back as its original bytes, unmodelled children and foreign
attributes included. `w:type` is projected to a closed region enum. A value the
workspace does not recognize projects as absent, round-trips from its preserved
bytes, and is never a reason to refuse the file. A style's own base `w:trPr`
and `w:tcPr` are modeled at their schema ranks beside its `w:tblPr`.
`w:tblStyleRowBandSize` and `w:tblStyleColBandSize` are modeled at their
`w:tblPr` schema slots and are counts rather than measurements, so no unit
conversion applies. `w:cnfStyle` on `w:pPr` is modeled at its schema slot
between `w:divId` and `w:rPr`, and the source element is retained as an
attribute carrier so the per-region attribute form of `CT_Cnf` survives being
modeled.

The native table facade authors auto, fixed-twip, and percentage widths through
one typed width mode. Checked physical measurements must be nonnegative and fit
the signed twip representation after the repository's pinned truncating unit
conversion. Percentages must be finite and between zero and 100, inclusive,
and serialize in fiftieths of a percent. Checked shading and border colors are
`auto` or six hexadecimal digits. A visible border has a checked nonzero width,
while an invisible edge remains an explicit `none` value rather than becoming
an absent child.

Complete grid replacement validates the active column count, positive widths,
signed total, row omissions, cell spans, and row coverage before mutation. A
successful update writes the active grid, fixed table width, and every covering
cell width together. Table property setters retain unrelated raw property
slots, border extensions, namespace aliases on read, and canonical fixed `w`
prefixes on changed modeled children.

Checked row authoring covers minimum and exact heights, repeating-header and
split toggles including explicit false and removal, alignment, conditional
regions, and leading or trailing grid omissions. Checked cell authoring covers
width, borders, margins, shading, vertical alignment, text direction,
conditional regions, wrapping, horizontal spans, vertical merges, and nested
tables. Topology-changing operations clone the complete table, reconcile only
untouched empty cells, validate positive grid widths, exact row coverage, and
immediately adjacent vertical merge ranges, then replace the live table.

Existing direct table rows clone and remove through a staged `Document`
mutation. A clone retains the complete row, cell, nested-content, relationship,
and raw XML model, then freshens bookmark, content-control, and drawing
identities, drops the `w14:paraId` and `w14:textId` of the row and of its
paragraphs, and omits copied comment anchors. Revision-save identities stay.
Body namespace declaration names are converted to fragment prefixes before
freshening, including the empty prefix for a root default namespace.
Table-level raw XML and content controls move with their logical row boundary.
Removing a vertical-merge restart promotes a matching continuation below, and a
table always retains one direct row. Invalid indexes, topology, XML, or reopen
results discard the candidate.

Row and cell property readers select modeled elements and attributes by their
bound WordprocessingML namespace. Foreign same-local children remain raw in
their exact schema slots. Changed modeled children use canonical `w` prefixes
and row or cell `xsd:sequence`, while unrelated row, cell, and border extension
bytes remain exact. Nested tables retain the required trailing cell paragraph
because Word refuses a cell ending with a table. Checked nested tables are also
nonempty.

Paragraph property readers select every modeled `w:pPr` child and attribute by
its bound WordprocessingML namespace, and a foreign same-local child stays
unmodelled in its own schema slot. Serialization writes canonical values with
fixed `w` attributes at each child's `xsd:sequence` position and replays every
retained raw child at its recorded slot and occurrence. `w:framePr` and each
border edge write their retained attributes before the modeled ones, so typing
those elements drops no producer attribute and keeps the retained ones in
source order. A modeled toggle whose source element carried an attribute the
model does not own replays that element in place of the canonical form, which
keeps a parse and save of an untouched paragraph byte identical.

Run property readers follow the same contract over the complete `EG_RPrBase`
sequence. `w:rFonts`, `w:color`, `w:bdr`, `w:fitText`, `w:eastAsianLayout`, and
`w:shd` write their retained attributes before the modeled ones, so typing
those elements drops no producer attribute. The one normalisation is
`ST_UcharHexNumber`. `w:themeTint`, `w:themeShade`, `w:themeFillTint`, and
`w:themeFillShade` hold a byte, so they are re-serialized as the two upper-case
hex digits Word writes, and a value that is not two hex digits is retained
verbatim through the element's ordered attribute vector instead.

Run content readers type `w:sym` and the four special characters in their
source positions among text, tabs, breaks, drawings, and fields. A `w:sym`
whose `w:char` is not four hex digits, a `w:ptab` missing one of its three
required attributes, and a `w:cr`, `w:noBreakHyphen`, or `w:softHyphen`
carrying any attribute stay in positioned raw capture, because a partial
projection would drop bytes a no-op save must return.

`w:ruby` is a paragraph-content sibling of `w:r` rather than a run child. The
paragraph reader accepts any prefix bound to the WordprocessingML namespace on
`w:ruby`, `w:rubyPr`, `w:rt` and `w:rubyBase`, including a binding declared on
the `w:ruby` element itself, and the writer emits the fixed `w` prefix at each
child's `xsd:sequence` position. The base runs are spliced into the paragraph's
own run list and the annotation records the half-open span they occupy, which
is what `w:hyperlink` already does, so every run index the paragraph maintains
across a split, an insertion, or a complex-field collapse moves the span with
it. The phonetic runs stay inside the annotation.

The typed ruby model admits only what it can write back exactly. A `w:ruby`
carrying an attribute, a `w:rt` or `w:rubyBase` child that is not a run,
non-whitespace character data between children, or an empty base line stays in
positioned raw capture instead, for the same reason a partial `w:sym`
projection does. Whitespace between children is this crate's own indentation
and is read and dropped. Inside `w:rubyPr`, a child outside the six modeled
ones is retained and written after them, and a modeled child missing its
required `w:val` is retained rather than normalized into an empty slot, which
is the rule `w:kern` already follows.

Section property readers select every modeled `w:sectPr` child by its bound
WordprocessingML namespace and replay each retained raw child at its recorded
schema slot and sub-slot. The sub-slot is what keeps `xsd:sequence` intact now
that `w:vAlign` sits between `w:formProt` and `w:noEndnote`,
`w:textDirection` between `w:titlePg` and `w:bidi`, and `w:docGrid` between
`w:rtlGutter` and `w:printerSettings`. `w:footnotePr`, `w:endnotePr`,
`w:paperSrc`, `w:pgBorders`, `w:lnNumType`, `w:vAlign`, `w:textDirection` and
`w:docGrid` are modeled. `w:formProt`, `w:noEndnote`, `w:bidi`, `w:rtlGutter`
and `w:printerSettings` stay byte-preserved at their slots, and `w:bidi` and
`w:rtlGutter` belong to the bidirectional family of F-266. `w:docGrid` reads
its `w:type`, `w:linePitch` and `w:charSpace` prefix-tolerantly, keeps any
other attribute in source order, and writes the retained attributes before the
modeled ones under the fixed `w:` prefix. A value outside `ST_DocGrid`, or a
pitch or space that is not an integer, is retained as an unmodelled attribute
rather than typed in part. `w:pgBorders`, `w:paperSrc` and `w:lnNumType` write their retained
attributes before the modeled ones, exactly as a border edge does. A modeled
child whose source element carried an attribute the model does not own is
replayed in place of the canonical form rather than typed in part, which is the
choice `w:pgSz` already makes and which keeps a parse and save of an untouched
section byte identical.

Word table styles parse modeled children and attributes by expanded name.
Base table properties and conditional regions retain self-contained source XML
with every inherited namespace binding they use. Typed table, cell, border,
shading, and paragraph projections drive layout, while unrelated producer
children remain at their schema positions. Unchanged projections reuse the
preserved subtree. A typed mutation writes one canonical modeled child in
`CT_Style` sequence order and reinserts unmodelled direct children once.

Rendering can use a default table style found only through the main document's
internal `stylesWithEffects` relationship. The target resolves relative to the
actual main part. This fallback reads a sole self-contained default table style
only when the main styles have no table default and its ID does not collide.
It requires a Word `styles` root, complete direct style-owner projection and
unique IDs. Missing, malformed, ambiguous or unprojected catalogues contribute
no fallback. Other effects styles are not merged. Rendering uses an owned
style copy and leaves both source parts, relationships and opaque XML intact.

The style projection also owns `link`, automatic redefinition, visibility,
gallery priority, quick-format, and locking values. Readers accept any prefix
bound to the WordprocessingML namespace. Writers emit the fixed `w` prefix and
place these children between `next` and property groups in schema order. A
facade update retains unmodelled style, paragraph, run, table, and conditional
region XML while applying modeled changes.

The default-off `oxml-opc/agile-encryption` feature reads and writes
password-protected OOXML packages. Readers parse the CFB `EncryptionInfo` and
`EncryptedPackage` streams, accept namespace aliases, and reject elements that
violate the Agile descriptor sequence. Supported read combinations are
AES-CBC with 128, 192, or 256-bit data and password-encryptor keys, varied
independently, and SHA-1, SHA-256, SHA-384, or SHA-512. The parent `keyData`
size governs the encrypted package key, while the password encryptor size
governs its wrapping key. Descriptor sizes, salt lengths, spin counts,
ciphertext lengths, and algorithm names are validated before expensive work
begins.

Password verification releases no package key on failure. A matching password
decrypts the data-integrity material and authenticates the complete encrypted
package stream, including its declared plaintext length, before any package
bytes reach the ZIP parser. Authentication is streamed and package decryption
uses 4096-byte Agile segments, so bounded constructors keep their ordinary ZIP
limits without adding a whole-package ciphertext buffer. Word emits a
hash-sized encrypted HMAC key, while the ECMA shape permits a salt-sized key,
so both validated lengths are accepted. Integrity, truncation, or length
failure is reported before ZIP parsing and leaves no partially opened package.

The writer stages the deterministic OPC ZIP and the complete CFB envelope
before publishing bytes. Its fixed profile is AES-256-CBC, SHA-512, 100,000
password iterations, 4096-byte encrypted package segments, and an HMAC over
the size-prefixed encrypted package stream. Every save draws independent
salts, package key, verifier, and HMAC key from the operating system random
source. `EncryptionInfo` uses schema child order. The version 3 CFB also
contains the complete DataSpaces map, definition, and strong-encryption
transform streams expected by Microsoft Word. Validation, random-source,
serialization, encryption, authentication, or staging failure leaves the live
package unchanged. The package API reserves capacity for the complete envelope
before appending it to the caller's byte buffer. A failed reserve or staging
step leaves existing output bytes unchanged.

The default-off `oxml-opc/digital-signatures` feature discovers signature
origins and signature parts only through normalized internal package
relationships. It parses XML Signature elements by expanded name and accepts
the strict RSA-SHA256, SHA-256 digest, and exclusive-canonicalization profile.
The OPC relationship transform selects exact declared relationship IDs,
rejects missing, duplicate, external, and absent targets, and emits canonical
ID order. Unsupported or weak algorithms fail closed.

Each report keeps cryptographic validity separate from complete declared
coverage. Cryptographic validity authenticates `SignedInfo`, every direct
reference, and only those manifest references reachable through an
authenticated same-document reference graph against the embedded X.509 public
key. Exclusive canonicalization retains processing instructions in their XML
child position. Coverage is complete only when every non-signature part,
content types part, and non-signature relationship is declared.
Certificate-chain trust is not inferred and remains caller policy.
Verification is read-only. A loaded package retains the original content-types
bytes while its typed content types remain unchanged, so saving does not
invalidate a signature by reserializing equivalent XML.

Signature creation accepts only a PKCS#8 DER RSA private key and an X.509 DER
certificate with the matching public key. It builds the signature origin,
signature part, content-type overrides, and internal relationships on a cloned
package. Collision-free names never replace occupied parts. The manifest uses
content-type-qualified references in deterministic part and relationship
order, authenticates every non-signature part and internal relationship, and
signs canonical `SignedInfo` with RSA-SHA256. The package object carries the
schema-ordered OPC `SignatureTime` property. Before allocating signature
infrastructure, creation rejects external, duplicate, dangling, misplaced
signature-typed, or untyped package graph entries and relationship sets whose
source is not an existing normalized part. A package that already declares a
signature origin is rejected instead of creating a second origin. The
candidate replaces the live package only after every shared verifier report
has both cryptographic validity and complete declared coverage.

Comment mutations validate coordinates and allocate every required id before
changing package or document state. Saving keeps the comments and
comments-extended relationship graph reachable from the main document, with
matching overrides and namespace declarations. A failed validation or
allocation leaves anchors, typed parts, relationships, and overrides unchanged.

Comment moves stage marker removal, narrowly qualified Google wrapper pruning,
destination placement and package reopen together. Unchanged definition and
companion parts retain their owned identities. Typed run references retain
unmodeled source through boundary-owned raw carriers that never emit twice or
survive typed removal or replacement. Numeric identity alone does not make an
independent raw element a typed carrier. Unordered raw references remain
verbatim, while plain id-only Word aliases retain canonical serialization. Literal comment matching qualifies Word markers and text in the actual staged main-part source, including ancestor-local table, cell, paragraph and control scopes. Complete-source namespace validation refuses unresolved element or attribute prefixes even outside the selected range. Comment moves replay existing qualified nested-owner declarations before replacing their main source and before the literal safeguard. Only the uniquely proved selected owner may refresh its private logical snapshot after authorized marker changes, with its exact raw namespace markers and namespace facts preserved. Ambiguous owners refuse, while all unselected owners retain strict replay. Source removal derives retained owners from the original package after the same qualified selected-marker edits, so transported whole runs disappear while mixed runs retain their declarations and non-reference source sequence. The captured reference remains namespace-closed for restoration. This preserves raw alias fields without changing the story fingerprint axis. Rich anchor projection consumes the namespace-closed original paragraph source and preserves those bindings through its private endpoint-fidelity probe. Tabs and breaks remain zero width only for the literal matching safeguard, while the rich reader retains their display characters.
Mixed reference runs retain neighboring
raw children, comments and processing instructions in source order.

Tracked-revision resolution is also staged above the package boundary. The
facade resolves selected revision placements in the main document, headers,
footers, comments, normal footnotes, endnotes, and nested text boxes. It patches
each affected source part once and reparses the complete candidate package
before replacing live typed state. Main-story resolution starts from the
prepared package's authoritative document bytes rather than a typed
reserialization. Namespace declarations carried only by a removed revision or
property owner remain available to retained raw descendants.
Any selector, revision-shape, namespace, parse, or serialization failure leaves
all package part bytes and live document state unchanged. The ordinary
deterministic save path writes the validated result later and preserves every
unrelated part and relationship.

Document comparison uses the same package boundary. It clones the complete
typed document and package state, resolves identical story shells and
relationships, and aligns modeled owners in each nonignored story. The prepared
package's main-document bytes are authoritative. Exact source spans flow through
body items, paragraphs, tables, rows, cells, controls, and runs, including when
another child of the same owner changes. An unchanged drawing-bearing run keeps
its complete wrapper, local namespace declarations, extended drawing children,
and relationship identifier. Compared pictures align by the bytes of their
relationship targets. A changed picture keeps the original media for rejection
and carries the edited media only when tracked content references it. A
repeated image may use a new relationship to an existing media part, so the
redline keeps one part per distinct payload. Imported media retains the edited
content type, and the complete package is reopened before either resolution is
checked. Policy projection removes only the selected comparison facts. Ignored
formatting,
textual whitespace, fields, comments, and story categories retain the original
bytes. Character and word alignment carries source ownership and raw-child
boundaries, keeps non-text content atomic, and emits each preserved child once.
Tables with different active grids use one deleted-table record followed by
one inserted-table record at the aligned boundary. Row markers carry the
revision metadata, so acceptance retains only the edited grid and rejection
retains only the original grid. Equal-grid tables keep row and cell comparison.
An attribute-free empty `w:pPr` or paragraph-mark `w:rPr` carries no formatting
in comparison. An attributed empty element remains opaque so its producer
attributes survive. A changed unmodelled paragraph or table property reports
a formatting diagnostic and retains the original bytes. Ignored main stories
are excluded from acceptance and rejection postconditions after each staged
revision resolution.
Generated revisions use canonical `w`, `xml`, and `mc` prefixes in schema
order, while reparse remains prefix tolerant. Source-span patching interleaves
changed owner bytes with the exact original gaps, preserving unowned
whitespace, comments, processing instructions, foreign elements, prefix
bindings, raw property children, and relationship targets. Nested text-box
projection uses one collision-safe marker selection across both staged inputs
and restores only matched owned subtrees. The staged package is accepted and
rejected independently to prove both package-wide policy postconditions. Any
metadata, policy, alignment, unsupported-shell, parse, serialization, or
postcondition failure leaves the original package, typed state, and caches
unchanged.
Comparison staging closes only the inherited bindings used by detached inline
and anchor wrappers. Bindings already present on the story root remain there,
while bindings declared on an outer drawing owner travel with the detached
wrapper. Dirty typed inputs recover matching package drawing payloads before
their staged flush. Physical complex-field runs project onto one modeled owner,
including several sibling fields that share one physical run.
Comparison-only story projections instead close every required drawing binding
on the inline or anchor root before equality. Story-root, paragraph, run, and
drawing-owner declaration placement is therefore equivalent in the comparison
model without changing package serialization. Retained drawing payload remains
significant after declaration placement normalization, so a real drawing
change is still tracked.

Literal redaction also uses the complete package boundary. The Word facade
flushes a staged clone, removes one non-empty exact literal from relationship-
resolved Word stories, comments, revisions, core and custom properties,
ChartML caches, and internal embedded workbooks, then serializes and reopens
the candidate. Sensitive XML is matched by expanded name. Unchanged byte
ranges and unrelated parts remain intact. External workbook relationships,
malformed sensitive XML, missing content types or internal targets, and ZIP
limit failures reject the candidate. Before publication, every inflated outer
and nested-workbook entry is scanned for both UTF-8 and UTF-16LE forms of the
literal. Any residual trace leaves the live package, typed state, and layout
caches unchanged.

Template rendering follows the same staged package boundary. A stack parser
pairs nested controls within one body or table-row container before evaluation.
The evaluator clones typed body entries and rows into candidate sequences, so
section properties, row properties, and ordered raw-child sidecars travel with
their owner. A row loop may clone several adjacent template rows per iteration.
Every paragraph and table row a loop renders is a copy, so it drops the
`w14:paraId` and `w14:textId` of its retained root-attribute record. Text-box
paragraphs stay inside the raw XML of their drawing and keep theirs.
The original table and its properties, grid, raw boundaries, content controls,
and relationships remain in place. Cloned row and cell property sequences keep
grid spans, vertical merge state, and unmodelled children byte for byte.
Repeated numbered paragraphs keep their existing numbering part reference and
level. No numbering relationship, instance, or abstract definition is added.
Markers are removed only from the candidate. Scalar syntax and JSON values are
resolved against lexical loop scopes before replacement reaches typed body
content, relationship-resolved headers and footers, the notes of the footnotes
and endnotes parts, raw text boxes, or chart parts. Replacement values pass
through collision-free sentinels, so a value that contains template syntax is
not evaluated recursively. The live typed
document and package are replaced only after every discovered tag is accounted
for, every repeated numbering reference resolves, and the candidate document
serializes successfully. Any control, lookup, numbering, scalar-type, parse, or
serialization failure leaves package parts, typed content, and layout caches
unchanged.

Mail merge uses the same fail-closed package boundary. Separate mode clones the
typed document and complete package for each record, applies the merge-local
field policy, serializes, and reopens every candidate before returning the
record-ordered outputs. Section mode serializes the main document and scans it
by expanded name for every header and footer reference, including references
inside content controls and preserved wrappers. Relationship-namespace ids are
resolved through the document relationship graph, and only the resulting
internal header and footer parts join relationship-resolved footnotes and
endnotes in the merge-dependency scan. A referenced non-body `MERGEFIELD` that
varies across records rejects the operation before candidate assembly.

Combined output reuses the first validated package and replaces only its main
body with the record bodies and their schema-ordered section boundaries.
Bookmark, content-control, and drawing identities are allocated without
collision across those bodies. Simple and complex bookmark field targets plus
hyperlink anchors follow renamed bookmarks in typed and preserved raw XML.
Clean parsed footnotes remain source-backed. An actual footnote field update
patches only the field-source spans in the relationship-resolved part, so
unmodelled siblings remain byte-preserved. Any rejected record, XML parse, or
identity-allocation failure leaves the source and all prospective outputs
uncommitted.

Rich mail merge extends this staging boundary for typed values and repeated
body regions. An image is embedded only after its exact positive EMU dimensions
validate. A whole-paragraph DOCX fragment contributes body content without its
final section properties. Before candidate mutation, the importer validates
every main-body relationship reference, rejects dangling or external targets,
and discovers the complete internal descendant closure. It allocates every
destination part name first, copies bytes and content types, preserves
part-local relationship ids, and rewrites main-body relationship ids to the
new document scope. Reachable styles and numbering receive deterministic
collision maps, and each repeated region or fragment occurrence receives fresh
document identities before insertion. Any value-kind, marker, relationship,
identity, callback, allocation, serialization, or reopen failure discards the
entire prospective result.

Cross-document block fragments carry their source package so retained XML stays
authoritative across every supported owner. Local namespace declarations and
unrelated destination XML survive insertion. The closure starts at selected
story references and required note and comment companions, follows internal
OPC targets through cycles, and copies reachable payloads with exact content
types and part-local relationship IDs. External edges retain their target,
mode and type and are never fetched.

Style aliases and links, numbering overrides and custom XML binding stores
are resolved with their companions. Store collisions rewrite the selected
bindings and copied item-property IDs together. Note references are rewritten
from complete maps, including forward, backward and cyclic references. Note,
comment-thread, bookmark, paired-marker and revision identities are fresh.
Equivalent reuse is limited to proven style and numbering comparisons and
relationship-free leaf payloads. Identity-bearing item properties that need a
store-ID rewrite are never reused.

Selected XML is changed only inside one staged transaction. The package is
serialized and reopened before the live destination changes. Missing targets,
malformed graphs, incomplete ownership, exhausted allocators, unsafe embedded
part-name rewrites and integrity-bound signatures that cannot remain valid
reject atomically. Existing destination signature coverage is invalidated by
the normal mutation policy. Opaque graphs are preserved without new decoding
or rendering support. Rich mail merge retains its separate prohibition on
external fragment edges.

Dynamic table-of-contents rebuild uses the same staged package rule. It scans
the relationship-resolved main document by expanded WordprocessingML names,
correlates the existing complex TOC begin, separator, and end markers, and
records exact byte offsets for the owned cached-result range. Bookmark markers
are inserted at schema-valid unowned boundaries by byte-position edits. Source
selection retains paragraph, run, and raw-child positions for bookmark scope.
Each required built-in entry level resolves a paragraph style by the
case-insensitive built-in name `toc N` and retains the producer's style id. A
style with an empty id is never chosen, because entries reference their style
by id. An existing canonical `TOCN` id is the collision-safe fallback, and a
canonical style is created only when neither form exists. Effective paragraph
properties decide whether the style already owns a right tab. Style-graph
validation and styles-part serialization complete inside the staged candidate,
so unrelated styles and unmodelled style children retain their source bytes.
One final empty component in a custom-style list is a tolerated producer
separator. Interior empty names, missing levels, and invalid levels remain
malformed. TOC discovery resolves duplicate style identifiers from the first
source definition and reports each duplicated identifier once. Several
defaults of one style type resolve to the one layout applies, the first in
source order, and each such type is reported once. That staged TOC view ignores
later definitions and later defaults without deleting or rewriting them.
Validation compares the view before and after entry styles are staged. Every
defect the source view already has is one that open, save, layout, and text
replacement accept, so the rebuild retains it and reports it once in check
order. A defect only the staged view has was introduced by the rebuild, such
as a new canonical style completing a dangling producer link into a one-way
link, and it rejects the rebuild. Public style mutation also compares counted graph defects before and after
staging. Existing repeated style IDs and multiple defaults remain in their
source order. An edit by repeated ID changes its first definition, while
removal deletes every definition with that ID. A newly introduced defect, or
an additional occurrence of an existing defect, rejects the candidate. Strict
whole-graph validation remains available to report producer defects.
Old-result exclusion adds a total nested-run order within each accepted
revision or content-control owner, so fields on opposite sides of a marker in
one wrapper remain distinguishable. The outer coordinate is the typed
paragraph owner's actual run boundary and raw-child slot, including terminal
hyperlink revisions and owners after preserved raw children. Hyperlink child
shapes retained as raw XML do not advance that run boundary. Sources before
the begin marker or after the end marker in a boundary paragraph remain
eligible. Retained comments and processing instructions consume raw-child
slots exactly as they do in the typed paragraph parser. A direct simple field
advances a modeled run boundary only when its parsed instruction is nonempty.
Hyperlink revisions and direct runs sharing one outer coordinate receive
distinct nested ordinals.
Generated entry paragraphs replace only the recorded range. The instruction
runs, matching field markers, neighbouring raw XML, relationships, and every
other package part remain outside the edit set. The provisional package and
the final page-substituted package both parse and reopen before one atomic
commit. Placeholder substitution requires one match in its recorded result
span and cannot search or replace elsewhere in the part. Overlapping TOC field
ranges fail before edits are built. Unsupported valid TOCs remain byte-identical with a reported
diagnostic. Malformed ownership, ambiguous bookmark identity, or any package,
layout, or reopen failure leaves the original package untouched.
Malformed or unprojected Word wrapper chains remain outside the ownership
scan even when their element names use the WordprocessingML namespace.
Each modeled content control owns only its first `w:sdtContent` child. A later
same-namespace content container remains opaque. The scan applies typed block
grammar and the 32-level revision nesting bound, counting property-change
revision elements as well as content revisions.
The first supported `w:sdtPr` type child also owns its producer attributes and
ordered child payload. Prefix-tolerant parsing records the typed discriminator,
while serialization uses the fixed type prefix and retains that payload only
when the discriminator is unchanged. Duplicate type children remain ordered
raw properties.
When a supported instruction is wrapped by inline ownership elements, staged
parsing and replacement close that exact balanced owner chain before emitting
the following paragraph content. Isolated instruction-run projection injects
the inherited namespace bindings required by every copied qualified name. It
locates the start-tag boundary with the XML parser and does not repeat a
declaration already local to the run.

Pagination-aware cache publication uses the same staged package boundary. One
deterministic layout records PAGE, NUMPAGES, and resolved PAGEREF values against
the owning paragraph node and top-level field position. Main-story fields are
updated through typed document content. Referenced headers, footers, footnotes,
and a uniquely owned endnotes part are patched through their relationship-
resolved source spans. Unsupported switches, unresolved targets, ambiguous
story ownership, and non-decimal section page formats retain their original
cache. Every written field becomes clean. A parse, layout, source-correlation,
serialization, or reopen failure publishes neither package bytes nor typed
state.

Generated-table source projection retains every producer namespace binding,
including a foreign binding for the conventional `w` prefix. A safe alias for
Word content rebuilds through the checked inventory. If checked target insertion
requires canonical serialization under a retained conflicting prefix, the
existing serializer boundary refuses atomically and preserves the complete
package and cache. Projection never rebinds opaque content to make insertion
possible. Simple cache expansion closes each replayed subtree over the owner's
namespace scope and validates the staged part before publication. Opened and
self-closing simple owners use the same ownership and refusal checks.
