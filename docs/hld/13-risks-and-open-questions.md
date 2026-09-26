# 13, Risks and open questions

## Settled decisions

### Preset shape definition provenance

The 187 preset geometries use the official
[ECMA-376 fifth-edition Part 1 archive](https://ecma-international.org/publications-and-standards/standards/ecma-376/),
`ECMA-376-1_5th_edition_december_2016.zip`. The precise source is
`OfficeOpenXML-DrawingMLGeometries.zip/presetShapeDefinitions.xml`, whose
SHA-256 is
`2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939`.
The exact count and digest are generator inputs, not advisory metadata.

The [Ecma software policy](https://ecma-international.org/policies/by-ipr/ecma-international-policy-on-submission-inclusion-and-licensing-of-software/)
classifies XML data sets as source code. Software incorporated in an Ecma
standard is available under the policy's three-clause BSD licence. Any checked-in
table derived from this XML must retain the Ecma copyright notice, all three
conditions, and the disclaimer. This third-party notice does not change the
repository's MIT OR Apache-2.0 licence.

LibreOffice's preset table must not be used. Its MPL-2.0 file-level copyleft is
outside the repository's approved licensing model. Derivation from the
specification text remains the fallback if the official archive, count, digest,
or notice cannot be reproduced exactly.

### PyPI distribution names

The `rdocx` and `rpptx` distribution names are controlled by the authenticated
publisher and are live on PyPI at 0.13.2 and 0.11.0 respectively. Both use the
`tensorbee/rdocx` repository, `wheels.yml` workflow, and `pypi` environment as
their trusted-publisher identity. The distribution and import names are the
same for each project.

## Open questions, to settle before the milestone that needs them

None.

## Risks, ranked

### R1, silent output drift during the extraction

**The top risk.** M4 and M5 change unit conversion, text-shaping input types and
content-type derivation. All three alter output without failing to compile, and
the existing 320 tests cannot see it.

*Mitigation*: the hash harness, built first and gating every PR. Every
intentional change lands as its own labelled commit with a reviewed delta. Do
not fold behaviour changes into moves. PDF alpha is additionally gated on
ExtGState structure and a deterministic midpoint compositing pixel.

### R2, the PDF coordinate-system flip

Replacing the per-element Y flip with one global CTM touches every element type
in a shipped, working renderer.

*Mitigation*: its own reviewable commit, landed **before any pptx code exists**,
gated on golden-PNG diffs comparing pixels rather than PDF bytes. The reviewed
Poppler 26.01.0 baseline declares exactly four stroke-antialias pixel changes
across `invoice` and `quote`. Normal check mode remains exact with no tolerance.

### R3, `Group`-blind collection passes

Three passes in the PDF backend must stay aligned with recursive content
emission. Missing one produces PDFs with absent fonts, images or link
annotations for grouped content.

*Mitigation*: all three passes use the mandatory `walk()` helper. Image and
annotation resources share depth-first leaf ordinals with recursive emission,
link rectangles apply the accumulated transform, and each pass has an explicit
nested-target regression test.

### R4, inheritance correctness

The layout-to-master match is by type, not `idx`. The text chain is seven levels
deep and level-indexed. The colour map is per-master and a dark master inverts
it. Wrong here means subtly wrong fonts, positions and colours on real
templates, which users notice and cannot describe.

*Mitigation*: M9 is a standalone milestone with visual differential tests and a
table of sampled theme-colour resolutions asserted to exact RGB. Word style
graphs are validated for type-compatible inheritance, cycle freedom,
reciprocal links, legal next styles, and one default per style type before a
mutation publishes. A TOC rebuild instead rejects only the defects its staged
entry styles introduce and reports the producer defects it retains.
Paragraph, character, and table default resolution has
focused deterministic render coverage. Effective numbering resolution applies
style inheritance, concrete replacement formatting, base-level start and
restart controls, and counter projection once, then shares that result with
visible markers, TOC rebuild, and REF fields.

Presentation latent placeholders use physical ownership as a policy boundary.
Slide-owned date, footer, and number placeholders are direct content. An
inherited layout or master placeholder requires its own source part's `p:hf`
container and flag. The source-built pinned LibreOffice differential covers
the absent, all-disabled, and slide-number-only master cases.

### R5, schema child ordering

Diffuse, because it touches every writer, and violations are silent until
PowerPoint refuses the file.

*Mitigation*: `OrderedRawChildren`, typed numbering children with explicit XSD
ranks, ordered `w:numPr` overlays that retain producer attributes and children,
focused `CT_Lvl` and `CT_NumLvl` order assertions, plus the corpus-wide
"opens without repair" gate at M8 and M11.

### R6, raw-XML preservation against relationship remapping

The preservation strategy that makes lossless round-tripping possible is exactly
what makes deep copy dangerous, because `r:id` attributes hide inside preserved
blobs.

*Mitigation*: `rewrite_rel_ids`, and `add_slide` synthesising rather than
deep-copying so the common path never needs it. The Word facade scans
relationship definitions by owner and identifier definitions by expanded XML
name before it allocates. Story-scoped authored occurrences are registered only
after insertion and reconciled against live expanded-name XML on each staged
serialization. Zero-use occurrence provenance retires without deleting its
relationship definition. Same-owner clones may share a relationship, every
live reference is remapped simultaneously, and each live authored picture gets
a distinct global drawing identity. Imported raw owners and their identifiers
remain fixed occupants. Producer drawing-definition uniqueness is checked per
physical XML part, then every accepted value joins package-global authored
occupancy. This accepts legitimate cross-part reuse without permitting authored
collisions or same-part aliases. Ambiguous preserved identity ranges and
relationship ownership reject before publication.

Comparison and revision resolution treat prepared package XML as authoritative
for the main story. Exact source spans pass through every changed owner so a
stable drawing run keeps namespace scope and opaque payload even when a sibling
changes. Staged accept and reject checks cover the package graph, and malformed
drawing identity fails without publishing candidate state.

An unknown root default namespace is safe to omit only when namespace-aware
scope analysis proves that no unprefixed element inherits it. Nested default
declarations shadow the root even when their URI is identical, while
unprefixed attributes do not consume it. Used, malformed, or ambiguous cases
reject before publication. Successful canonical serialization refreshes the
cached namespace facts from the emitted main-story bytes, which keeps repeated
saves consistent after an unused declaration is omitted.

### R7, scope

Charts plus full parity in one release is **17 to 18 months solo** with nothing
shipping until the end. Both halves of that were chosen deliberately. The
duration was not: the plan was estimated at phase granularity and said nine to
twelve months, and the story-level sizing in `14-development-backlog.md`
disagrees. The story-level number is the trustworthy one.

*Mitigations, neither requiring rework*: a second developer takes it to roughly
9 to 11 months, since M7 and M8 parallelise once M6 lands. Or cut a
read-plus-render release at the end of M10, about 12 months solo, which is the
point where the library becomes useful. Charts are self-contained and nothing
depends on them, so M12 can move independently at any time.

### R8, `oxml-layout` packaging

The bundled fonts are 6.8 MB outside `src/`, published today only because
`cargo publish --no-verify` skips the build-from-archive check, and there is no
`include` or `exclude`.

*Mitigation*: an explicit `include`, drop `--no-verify`, and assert the `.crate`
size against the 10 MiB crates.io limit in CI. Roughly 3.5 to 4 MB compressed is
expected, but verify rather than assume.

### R9, index-path aliasing in bindings and story locations

An index path addresses a position, not an object, so a handle held across a
structural mutation would silently read the wrong element.

*Mitigation*: the revision counter making it a loud `StaleElementError`, lazy
collections so the idiomatic loop never holds a stale handle, and stable ids in
v0.2. Native story locations also carry the owner fingerprint and item kind.
Structural operations accept an existing boundary only from the canonical
flattened item that resolves to an actual direct owner child. The separate
`ContentLocation::end` marker represents the boundary after final content
without aliasing an item path.

### R10, the `rpptx-oxml` scope surface

PresentationML is enormous and there is no natural stopping point.

*Mitigation*: the raw-XML passthrough discipline. **If you do not render it, do
not model it.** Stated as a rule in `06-presentationml-model.md` and enforced in
review.

### R11, modern DOCX completeness becomes unbounded

WordprocessingML, DrawingML, package extensions, and producer-specific
compatibility markup do not provide a natural meaning of feature complete.
Without a closed capability vocabulary, an audit can continually discover new
work and a public facade can appear complete while omitting one half of a
cross-part invariant such as style-linked numbering.

*Mitigation*: the closed property-level capability matrix uses stable
`DOCX-001` through `DOCX-085` identifiers and classifies each row as complete,
partial, preserve-only, unsupported, or a permanent non-goal. Every partial or
unsupported row names exactly one live owner. M23 closes the five-document
from-scratch gate. M24 closes the declared modern DOCX authoring matrix. A new
capability must change the matrix and its one owning story explicitly rather
than hiding the gap behind raw XML.

### R12, private client documents escape into repository history

The five M23 reference documents are customer data. A useful differential gate
needs their exact parts and renders, but committing a document, derived page,
text extract, filename, or identifying manifest would create an irreversible
confidentiality failure.

*Mitigation*: the corpus stays in a configured ignored private directory.
Tracked tests contain only synthetic fixtures, anonymous P1 through P5 family
requirements, and non-identifying capability assertions. Local required-corpus
mode records hashes, pure Rust generator sources, embedded approved raster
evidence, and comparison output outside the repository. It scans staged and
tracked paths for forbidden artifacts and fails closed when the configured
private corpus is incomplete. Each generator depends only on public `rdocx`,
starts from `Document::new()`, cannot import HTML or raw package content, and
must produce identical bytes twice. Public CI proves the same API boundary
through synthetic documents.

### R13, a public facade writes only half of an OOXML invariant

Several Word features span parts or stories. Numbering linked to a paragraph
style, fields linked to bookmarks, drawings linked to media, and section
headers linked through relationships can all produce a valid ZIP that behaves
incorrectly in Word when only one side is written.

*Mitigation*: each partial or unsupported matrix row has one operation-level
owner across F-243 through F-310 rather than several stories exposing
uncoordinated XML fragments. Conformance tests inspect every owned part, reopen
through `Document`, compare fresh layout, and require authored public-API
content to report no unexplained unmodeled properties.

The Word facade's private document identifier owner reserves related part,
relationship, content-type, and XML identities together. Final-order
canonicalization runs on a staged clone, so a collision or exhausted range
cannot publish half of a package invariant.

Generic content insertion, removal, cloning, and movement use the same staged
boundary. A clone freshens known document identities before insertion.
Relationship-bearing content remains restricted to its unchanged story owner,
and preserved identity ownership that cannot be proved rejects. Serialization
and reopen complete before the candidate replaces live state.

Cross-document main-body import removes that unchanged-owner restriction only
through an owned `DocumentFragment`. The importer discovers selected body and
comment dependencies before mutation, closes internal relationships
recursively, reserves every destination identity, and rewrites exact retained
XML only from complete maps. Caller policy may reuse equivalent styles,
numbering, and related leaf parts. Unsupported external, dangling, malformed,
split-range, and exhausted graphs fail without changing the live document.
All-story import and the broader dependency classes remain owned by F-276.

Picture, hyperlink, and relationship lookup operations resolve the exact OPC
owner from `StoryId`. Cells and text boxes inherit their containing part, and
related stories use their resolved part. Media, relationship, content-type,
XML, and drawing changes publish together only after the complete candidate
reopens. Shared same-owner clone relationships and per-occurrence drawing
identities are canonicalized from final serialized order.

Ordered section removal uses that same staged boundary. Before removing a
non-final owner, it resolves the first usable same-variant internal header and
footer relationship and materializes inherited behavior on the following
section. It prunes only facade-owned parts that no modeled or opaque reference
can reach. Missing, malformed, cross-type, external, shared, and producer-owned
edges cannot publish a half-updated section and story graph.

Per-section header and footer mutations use the same invariant boundary.
Lookup follows only the requested variant through preceding sections. Link
accepts only an existing internal exact-type story. Unlink and replace copy the
story XML and complete part-local relationship set, rebase internal relative
targets, and freshen drawing identities before publication. Inherit removes the
direct reference, while remove installs an explicit empty story. Pruning is
limited to facade-owned graph nodes that are unreachable from modeled and
opaque references. The document-wide even-page setting is an explicit typed
operation, and first-page creation enables the section title-page state.

Style graph mutations use the same rule. Adding or updating one side of a
legal paragraph and character link updates the reciprocal edge in the staged
candidate. Missing targets, incompatible types, duplicate defaults, cycles,
and live references abort before the document changes.

Numbering graph mutations use staged definition and instance values. They
preserve and expose imported paragraph-style links.
`link_style_to_numbering` and `unlink_style_from_numbering` clone the complete
document, validate both graphs and the exact reciprocal tuple, update the
numbering-level link and style numbering properties together, serialize the
candidate, and publish once. Failure leaves both parts and allocation state
unchanged.

Visible counters are keyed by concrete `numId` and level. This makes a new
instance independent even when it shares an abstract definition. Microsoft
Word 16.112.3 continues one captured shared-definition sequence across distinct
instances, so that behavior is an intentional documented divergence. The
pinned differential uses separate abstract definitions where both systems
agree, and a focused regression guards the approved concrete-instance rule.

### R14, Python publication exposes a partial or unaudited package family

The `rdocx` and `rpptx` distributions have independent native versions and
release boundaries, while their twelve platform wheels and two source
distributions share one build matrix. A count-only upload, manual publication
path, stale artifact, mismatched project version, or unverified
trusted-publisher identity could publish the wrong project or mix both families.
A project that omits its README from package metadata can also publish working
files whose PyPI page has no usable installation or API guidance. Published
release files are immutable, so that omission requires a new version.

*Mitigation*: `py-rdocx-vX.Y.Z` and `py-rpptx-vX.Y.Z` each select one
distribution at its native crate version. Manual workflow dispatch has
build-only authority for both distributions. The release preflight binds the
selected seven downloaded artifacts to the reviewed SHA, validates exact names,
embedded versions, `cp39-abi3` tags, platforms, and source distribution, and
requires the reviewed summary, author, keywords, classifiers, project URLs,
Markdown content type, and README guidance in wheel and source metadata. It
installs the selected project under Python 3.9 and 3.12. The tag-only `pypi`
environment receives OIDC authority after a separate final approval.
Completion requires the selected PyPI version, all seven files, authenticated
owner or maintainer roles, an exact reviewed GitHub release body, and every
planned contributor comment.

### R15, raster media exhausts memory or disappears silently

Compressed picture dimensions can imply much larger decoded storage. An
unbounded decoder can exhaust memory, while a low silent ceiling can remove an
ordinary screenshot or print-resolution figure from both PDF and raster
output. Straight-alpha bytes passed to a premultiplied surface can also paint
colour from pixels that should be transparent.

*Mitigation*: presentation admission keeps a 16 MiB encoded ceiling and a
64 MiB decoded ceiling with checked arithmetic before allocation. Rejected
relationships carry a stable scoped failure into one diagnostic and a visible
bounds fallback. The shared raster backend premultiplies decoded PNG channels
exactly once at the tiny-skia boundary. Deterministic format, large-image,
malformed-header, overflow, diagnostic, and fallback tests gate these rules.

## Assumptions that would invalidate the plan if wrong

- **That a slide is a page.** The entire rendering reuse argument rests on it.
  Verified: `output.rs` is docx-free and `rdocx-pdf` depends on nothing else.
- **That `OpcPackage` reads a `.pptx` unmodified.** Verified through
  `main_document_part()` keying off `officeDocument`. Worth an actual test in M2
  rather than continued confidence.
- **That preset geometry is a data problem once the evaluator exists.** Verified
  by the official 187-definition ECMA-376 data set. The specification-text
  derivation remains available if the pinned source cannot be reproduced.
