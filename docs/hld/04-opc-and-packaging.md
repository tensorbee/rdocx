# 04, OPC and packaging

Owner: `oxml-opc`, with media naming in `oxml-media`.

## The package

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
constructions remain byte-identical.

Word story discovery begins with the main body and its nested cell and text-box
owners, then follows header and footer references in document order and note
and comment relationships in package relationship order. Ordinary footnotes
and endnotes keep their source order, while conventional nonpositive separator
records are not public story owners. Each `StoryId` includes the normalized
source part, owner kind, source-order ordinal, and a structural fingerprint.
Any changed owner makes a retained identity stale before indexed content can be
resolved.

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

The main Word document reader accepts one namespace-correct `document` root
and one namespace-correct `body` child, rejects truncation, duplicate roots,
foreign lookalikes, and non-whitespace content outside that root, and retains
the first body `sectPr`. Later section owners remain opaque rather than
disappearing. Self-closing paragraphs, tables, and cells are modeled as empty
typed owners. An attribute-free self-closing paragraph-property element or
paragraph mark is modeled like its start-and-end form, while an attributed one
stays raw with its exact attributes. Start-and-end forms are modeled only when
their complete attributes and content satisfy the same grammar. Otherwise their
exact namespace-complete subtree remains opaque. Header and footer references
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
occurrence receives its own global `wp:docPr` identity. Producer-owned raw
references remain fixed occupants, and image part naming follows final
serialized relationship order.

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
Building-block replacement enters the canonical staged preparation and
provenance-reconciling reopen path before publication.

Both facades resolve core properties through the package-level
`CORE_PROPERTIES` relationship and retain its normalized target. Immutable
property access leaves the source part bytes untouched. Mutable access marks
the typed `CoreProperties` model for serialization to that target with its
content-type override. A package that creates metadata without an existing
relationship uses `/docProps/core.xml` and adds the missing package
relationship. If that conventional part name is already occupied without the
core-properties relationship, serialization returns an error before changing
the package.

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
An unchanged plain numeric leaf may use the typed serializer's indentation,
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

Modeled paragraph, run, and section-property owners retain every ordered root
attribute, including producer identity, revision-session, foreign, and
unqualified attributes. Retention uses the existing raw-preservation carriers
without exposing the attribute record as child XML. Expanded names govern
duplicate rejection and authored paragraph identity precedence, so an authored
`paraId` replaces only the retained attribute with the same namespace and
local name. Typed child mutation leaves all other retained root attributes in
source order.

Retention covers attributes, not namespace bindings. A declaration is recorded
only when a retained attribute uses its prefix, because the alias machinery
already materializes a binding onto every element that needs one, and recording
a declaration a child carries for itself would emit it twice. A root carrying
nothing but declarations retains no record at all. On the way back out, the
canonical `w14` binding is not copied onto the written element, since the part
root that owns the element already declares it and the authored identity write
makes the same assumption. Together these keep a reopened save byte identical
to the save it was read from.

An unknown default namespace declared on the document root is classified by
its effective lexical scope before canonical serialization. An unused root
default may be omitted without blocking a typed mutation. An unprefixed element
that inherits it keeps the binding live and blocks modified serialization.
Nested default declarations shadow the root declaration, including when they
repeat the same URI, and unprefixed attributes never use a default namespace.
Malformed or ambiguous declarations fail closed. After a successful canonical
publication, the document refreshes its root and body namespace facts from the
published main-story bytes so a later save applies the same classification.

Paragraph line spacing retains the signed integer path required by
WordprocessingML and accepts one bounded producer deviation. A plain signed
decimal `w:spacing/@w:line` value is normalized with exact decimal arithmetic
to the nearest integer twip, with exact halves rounded away from zero.
Exponent notation, malformed forms, non-finite spellings, and numeric values
outside the signed 32-bit range remain errors. The modeled value serializes as
one canonical integer without widening decimal acceptance to sibling measures.

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
the typed comments model back to that target with its content-type override.
The model preserves unmodelled attributes and children at their insertion
boundaries, while comment range and reference anchors remain ordered among
neighbouring paragraph and run XML. A document without a comments relationship
does not gain a comments part, relationship, or override during an ordinary
save.

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

Threaded comments add a document relationship using the Microsoft
`commentsExtended` relationship type. The facade retains its resolved target
and writes the standard
`application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`
content type at that exact part. New comment state creates both relationships
and both overrides together. An ordinary save retains an accepted standard
override and unrelated comments-extended sidecar XML. Existing custom targets
remain authoritative, and removal of the final API-owned thread removes only
the parts, relationships, and overrides created by the typed model.

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
one exact signed-integer projection. The parser accepts integer lexical forms
and decimals only when the nonempty fractional portion contains zeroes. It
checked-parses the integer portion into `i32` without floating point. Fractional
values, exponent forms, empty fractions, overflow, percentages, universal
measures, and malformed input fail explicitly instead of becoming zero. Missing
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
identities and omits copied comment anchors. Body namespace declaration names
are converted to fragment prefixes before freshening, including the empty
prefix for a root default namespace. Table-level raw XML and content controls
move with their logical row boundary. Removing a vertical-merge restart
promotes a matching continuation below, and a table always retains one direct
row. Invalid indexes, topology, XML, or reopen results discard the candidate.

Row and cell property readers select modeled elements and attributes by their
bound WordprocessingML namespace. Foreign same-local children remain raw in
their exact schema slots. Changed modeled children use canonical `w` prefixes
and row or cell `xsd:sequence`, while unrelated row, cell, and border extension
bytes remain exact. Checked nested tables are nonempty and retain the required
trailing cell paragraph.

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
and relationship identifier. Policy
projection removes only the selected comparison facts. Ignored formatting,
textual whitespace, fields, comments, and story categories retain the original
bytes. Character and word alignment carries source ownership and raw-child
boundaries, keeps non-text content atomic, and emits each preserved child once.
Tables with different active grids use one deleted-table record followed by
one inserted-table record at the aligned boundary. Row markers carry the
revision metadata, so acceptance retains only the edited grid and rejection
retains only the original grid. Equal-grid tables keep row and cell comparison.
An empty paragraph-property element or paragraph mark compares like an absent
one. A differing unmodelled paragraph or table property child keeps its
original bytes and yields a formatting diagnostic, even when the modeled
properties match. A tracked paragraph-property change records `CT_PPrBase`
only, so the paragraph mark stays out of `w:pPrChange`. A changed mark keeps
its original run properties and yields a formatting diagnostic.
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
change is still tracked. A drawing that a redline story writes twice, as a
move or a deletion beside an insertion, keeps its `wp:docPr` id on the first
copy and gets a fresh id above every compared drawing id on later copies, as
Word does. Both postconditions read each fresh id as the id it copies.

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
The original table and its properties, grid, raw boundaries, content controls,
and relationships remain in place. Cloned row and cell property sequences keep
grid spans, vertical merge state, and unmodelled children byte for byte.
Repeated numbered paragraphs keep their existing numbering part reference and
level. No numbering relationship, instance, or abstract definition is added.
Markers are removed only from the candidate. Scalar syntax and JSON values are
resolved against lexical loop scopes before replacement reaches typed body
content, relationship-resolved headers and footers, raw text boxes, or chart
parts. Replacement values pass through collision-free sentinels, so a value
that contains template syntax is not evaluated recursively. The live typed
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

Cross-document body fragments carry their source package so retained XML stays
authoritative while the selected closure is rebuilt in the destination. The
closure starts from selected main-story and comment-story relationship
references, follows internal targets recursively, and copies only reachable
parts with their exact content types. Style and numbering graphs are pruned to
the selected references and their transitive links before deterministic reuse
or renaming. Every destination relationship id, part name, comment id,
bookmark id, drawing id, style id, and numbering id is reserved before any
selected XML is rewritten. An external edge, missing target, malformed
relationship part, invalid ownership range, or exhausted allocator rejects the
candidate without publishing package or typed state.

Dynamic table-of-contents rebuild uses the same staged package rule. It scans
the relationship-resolved main document by expanded WordprocessingML names,
correlates the existing complex TOC begin, separator, and end markers, and
records exact byte offsets for the owned cached-result range. Bookmark markers
are inserted at schema-valid unowned boundaries by byte-position edits. Source
selection retains paragraph, run, and raw-child positions for bookmark scope.
Each required built-in entry level resolves a paragraph style by the
case-insensitive built-in name `toc N` and retains the producer's style id. An
existing canonical `TOCN` id is the collision-safe fallback, and a canonical
style is created only when neither form exists. Effective paragraph properties
decide whether the style already owns a right tab. Style-graph validation and
styles-part serialization complete inside the staged candidate, so unrelated
styles and unmodelled style children retain their source bytes.
One final empty component in a custom-style list is a tolerated producer
separator. Interior empty names, missing levels, and invalid levels remain
malformed. TOC discovery resolves duplicate style identifiers from the first
source definition and reports each duplicated identifier once. Validation of
that staged TOC view ignores later definitions without deleting or rewriting
them. Public style mutation retains strict duplicate rejection.
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
