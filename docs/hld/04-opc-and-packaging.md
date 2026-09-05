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

**Saves are deterministic.** Both `part_rels` and `parts` are emitted in sorted
key order, so writing the same package twice produces byte-identical output.
That property is load-bearing for the round-trip corpus and must not regress.

Modern presentation package identity is the main presentation part's exact
content type. `oxml-opc` names the ordinary presentation, macro-enabled
presentation, ordinary template, macro-enabled template, ordinary slideshow,
and macro-enabled slideshow values. `rpptx` accepts only those six values.
Changing an output class replaces only that override in a staged package.
Executable and unrelated parts remain opaque and byte-preserved.

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

The docx presets become a short private helper in `crates/rdocx/src/document.rs`,
and the pptx presets one in `crates/rpptx/src/package.rs`.

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

Both facades resolve core properties through the package-level
`CORE_PROPERTIES` relationship and retain its normalized target. Immutable
property access leaves the source part bytes untouched. Mutable access marks
the typed `CoreProperties` model for serialization to that target with its
content-type override. A package that creates metadata without an existing
relationship uses `/docProps/core.xml` and adds the missing package
relationship. If that conventional part name is already occupied without the
core-properties relationship, serialization returns an error before changing
the package.

The Word facade owns external hyperlink relationships at the document part
boundary. `Document::add_hyperlink_relationship` allocates the relationship,
and `Paragraph::add_hyperlink` writes a schema-ordered `w:hyperlink` that
references it. The same paragraph writer emits explicit hard breaks as run
content. Both operations use the existing package-preserving save path, so
unmodelled parts and relationships remain intact.

Numbering state is also fail-closed at this boundary. Updating a known list
level marks the existing numbering model for serialization. Rejecting an
unknown list identifier or an invalid level does not create an empty numbering
part, relationship, or content-type entry. Numbering parsers retain namespace
declarations and compatibility attributes from modelled containers. Unknown
level children use their `CT_Lvl` schema slots, while abstract-definition,
instance, and root children keep insertion-aware boundaries. Mutating or adding
a definition therefore preserves producer extensions, identifiers, templates,
and level overrides verbatim. Identifier allocation uses the next value after
the maximum when available and the first unoccupied value when the maximum is
`u32::MAX`.

Producer-defined `w:numFmt` values remain typed as their original token rather
than being substituted with decimal numbering. Numbering serialization writes
that token back unchanged. Layout and text exporters emit no marker for a
format whose rendering semantics are unknown.

The main document reader retains root, body, and modeled-owner namespace facts
that preserved raw descendants depend on. Save replays those declarations on
their logical owners through insertion, removal, and reordering without
rewriting the raw subtree bytes. Prefix aliases, nested shadows, and ordinary
namespace URI escaping are resolved by the XML parser. Serialization fails
closed when owner identity or a serializer prefix binding cannot be preserved
safely, leaving the opened package bytes authoritative.

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

Raw Word run children receive semantic classification only at the OXML parse
boundary. A WordprocessingML `pict` is classified as a legacy horizontal rule
only when its in-scope expanded names identify exactly one VML `rect` with an
enabled Office `hr` attribute and whitespace otherwise. The classification is
stored in the existing raw-child position sidecar, while the subtree bytes and
ancestor namespace ownership remain unchanged. Foreign, malformed, numeric,
false, visible, or structurally ambiguous content stays unmodelled raw XML.

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

Watermark authoring follows the document-to-header graph rather than assuming
conventional header names. The facade materializes a missing default, first, or
enabled even header only at the first section that needs that same-type variant.
Later omitted references keep Word's same-type inheritance and do not receive a
blank override. Each image relationship belongs to its owning header part, and
its target is relative to that part even when a producer uses a custom header
path. Settings values controlling even headers are namespace checked and XML
decoded before selection.

An authored watermark owns only a VML shape whose expanded name is `v:shape`
and whose unqualified id is `rdocx-watermark`. Replacement patches that exact
byte range in the original header, leaves tables, controls, root attributes,
namespace declarations, unrelated VML, and other producer bytes in place, and
keeps every emitted shape-type reference resolvable. Text and image operations
stage all header, relationship, media, and content-type changes on a cloned
package. A missing part, invalid dimension, parse error, or serialization error
leaves the live document and package unchanged.

Threaded comments add a document relationship using the Microsoft
`commentsExtended` relationship type. The facade retains its resolved target
and writes the comments-extended content type at that exact part. New comment
state creates both relationships and both overrides together. Existing custom
targets remain authoritative, and removal of the final API-owned thread removes
only the parts, relationships, and overrides created by the typed model.

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
positive number is found. Both facades use this allocator, so allocation never
creates `image0` or overwrites an existing numbered image part.

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
their content-type overrides are staged with the drawing on cloned package and
document state. The mutation becomes visible only after the typed ChartML,
SpreadsheetML workbook, relationships, content types, and structured drawing
all serialize successfully.

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
One historical grid-change subtree is preserved, while a second modeled change
fails parsing rather than discarding history. Serialization writes active
columns first and the historical change after them in schema order.

Word table styles parse modeled children and attributes by expanded name.
Base table properties and conditional regions retain self-contained source XML
with every inherited namespace binding they use. Typed table, cell, border,
shading, and paragraph projections drive layout, while unrelated producer
children remain at their schema positions. Unchanged projections reuse the
preserved subtree. A typed mutation writes one canonical modeled child in
`CT_Style` sequence order and reinserts unmodelled direct children once.

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
before replacing live typed state. Namespace declarations carried only by a
removed revision or property owner are promoted to retained raw descendants.
Any selector, revision-shape, namespace, parse, or serialization failure leaves
all package part bytes and live document state unchanged. The ordinary
deterministic save path writes the validated result later and preserves every
unrelated part and relationship.

Document comparison uses the same package boundary. It clones the complete
typed document and package state, resolves identical story shells and
relationships, and aligns modeled owners in each nonignored story. Policy
projection removes only the selected comparison facts. Ignored formatting,
textual whitespace, fields, comments, and story categories retain the original
bytes. Character and word alignment carries source ownership and raw-child
boundaries, keeps non-text content atomic, and emits each preserved child once.
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

Dynamic table-of-contents rebuild uses the same staged package rule. It scans
the relationship-resolved main document by expanded WordprocessingML names,
correlates the existing complex TOC begin, separator, and end markers, and
records exact byte offsets for the owned cached-result range. Bookmark markers
are inserted at schema-valid unowned boundaries by byte-position edits. Source
selection retains paragraph, run, and raw-child positions for bookmark scope.
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
When a supported instruction is wrapped by inline ownership elements, staged
parsing and replacement close that exact balanced owner chain before emitting
the following paragraph content. Isolated instruction-run projection injects
the inherited namespace bindings required by every copied qualified name. It
locates the start-tag boundary with the XML parser and does not repeat a
declaration already local to the run.
