# 12, Testing strategy

## Test taxonomy

Six categories. Every story's design plan picks the applicable ones and names
exactly one as its test gate.

| Category | Purpose | Where |
|---|---|---|
| `unit` | Pure logic, no I/O | `crates/<crate>/src/*.rs` under `#[cfg(test)]` |
| `integration` | Multi-crate behaviour through the public API | `crates/<crate>/tests/` |
| `regression` | Locks down one named past failure | `crates/rdocx/tests/regression_test.rs` and the rpptx equivalent |
| `round-trip` | Parse, serialise, reparse, compare | corpus-driven |
| `golden` | Byte or pixel comparison against a recorded baseline | the hash harness |
| `differential` | Compared against an external oracle | LibreOffice for renders, python-docx and python-pptx for the bindings |

The existing repository convention is preserved: **no binary fixture files.**
Fixtures are constructed in code, including hand-assembled PNG and JPEG headers
with precomputed CRCs. It keeps the `.crate` payload small and the diffs
readable. The corpus in the next section is the one deliberate exception, and it
lives outside the published crates.

The ODT reader differential constructs its ZIP, XML, and PNG input in source,
then converts the same input with the exact pinned LibreOffice build in an
isolated profile. The normalized comparison covers body order, effective text
and paragraph formatting, list kind and level, table grid and spans, and image
bytes and dimensions. It deliberately ignores package bytes, relationship ids,
part names, and namespace prefixes.

The ODT writer round-trip gate builds its document in source, writes ODF 1.3
through the native facade, and reopens the result through the ODT reader. The
normalized comparison covers body order, effective text and paragraph
formatting, list kind and level, table grid and horizontal and vertical spans,
and image bytes and truncating EMU dimensions. Focused tests lock the stored
first `mimetype`, fixed-prefix XML and manifest order, byte-identical repeated
writes, exact whitespace elements, output and diagnostic bounds, stable lossy
paths, and atomic path replacement. The writer does not use LibreOffice as a
package-byte oracle.

The ODP differential builds its presentation and ODF package in source. The
exact pinned LibreOffice 26.2.5.2 build converts both directions in isolated
profiles, and the gate checks slide count, supported text, successful PDF
production, page count, and rendered text. Focused tests cover expanded-name
aliases, schema-position lookalikes, unsafe entries, duplicate expanded
attributes, archive and output limits, deterministic bytes, exact manifest
ownership, diagnostic exhaustion, and atomic failure publication. Existing
presentation hashes must remain 49 of 49 unchanged.

The presentation HTML differential builds one HTML document, image, and font
reference in source and sends that same input to Google Chrome 152.0.7977.65
and the native importer. Chrome runs headless at 1,280 by 720 with an isolated
profile and host resolution disabled. After save and reopen, the gate compares
shape kind and order, exact text and link relationships, selected run
formatting, geometry within one CSS pixel, and deterministic 96 DPI output at
a full-image luminance SSIM floor of 0.95. A paired regression passes real PNGs
through the same SSIM helper and proves structural, text, two-pixel geometry,
and calibrated pixel mutations fail the shared acceptance predicate. Focused
tests cover every importer limit, CSS cascade and EMU conversion, stable lossy
diagnostics, editable shape, table, image, and link projection, schema order,
and byte-identical default template parts. The 49-entry hash harness remains
unchanged.

Regression tests are named as sentences describing the failure they prevent, so
a reintroduction is obvious from the test name alone rather than from a diff.
The existing file is the model: `zero_column_tables_do_not_panic`,
`saving_is_reproducible`.

The legacy form and glossary round-trip gate constructs every package in
source. It covers typed text, checkbox, and drop-down values, deterministic
part-scoped ordinal identity across supported internal Word stories, AutoText
classification, and selected building-block replacement. Prefix aliases,
fixed-prefix changed output, schema-order insertion, structural reopen,
byte-exact unsupported subtree retention, unsafe relationship graphs, stale
identities, wrong value kinds, bounds, and atomic failure are focused checks.
The Python, WASM, and CLI surfaces and the 49-entry hash set remain unchanged.

The glossary, embedded-content, and package-story malformed XML matrices run
through the single strict XML 1.0 lexical validator in `oxml-core`. A shared
unit matrix covers declarations, literal characters, names, namespaces,
duplicate expanded attributes, references, comments, and processing
instructions. Consumer regressions pin `OxmlError::InvalidValue`,
`Error::InvalidEmbeddedMutation`, and `Error::Other` mapping, while the
embedded cases also require byte-identical rollback after rejected mutation.

The Presentation collaboration round-trip gate is
`modern_comments_replies_sections_and_handout_settings_survive_ordered_mutation_save_and_reopen`.
It builds a noncanonical package in the existing `rpptx` integration binary,
authors and reorders comments and replies, replaces section membership, moves
a slide, changes notes-master and handout-master header-footer values, saves,
and reopens. The reopened model must retain comment and reply order, producer
slide ids, section membership, both header-footer changes, relationship
targets, and the modern authors and comments content-type overrides.

Adjacent `rpptx-oxml` round trips cover namespace aliases, inherited default
namespaces, fixed-prefix shadows, schema child order, structural reparse, and
byte-exact unsupported attributes, direct events, anchors, text bodies,
extension lists, and section sidecars. Facade regressions reject duplicate or
unknown ids, invalid membership, external or wrong-type relationships, shared
comment parts, occupied conventional paths, and unserializable staged changes
without mutating the opened package. All fixtures remain source-built in the
two existing integration binaries.

The Presentation executable-content regression gate is
`embedded_inventory_reports_exact_hashes_relationships_and_signature_state`.
It constructs OLE, ActiveX, VBA, package-signature, and VBA-signature graphs in
the existing `rpptx` integration binary and requires stable kind, source,
relationship, target, content type, byte length, SHA-256, and signature state.
Adjacent source-built tests prove exact extraction, transactional replacement
and removal, ownership-aware reachability, retained raw XML and signature
evidence, producing-scope selection, namespace and compatibility handling,
malformed graph rejection, and byte-for-byte failure atomicity. No payload is
decoded or executed.

The Word executable-content regression gate is
`word_embedded_inventory_reports_exact_hashes_relationship_paths_and_signature_state`.
It source-builds OLE, ActiveX, VBA, package-signature, legacy VBA-signature,
and Agile VBA-signature graphs in the existing `rdocx` regression binary. The
gate requires deterministic ordering and stable kind, source, relationship,
target, content type, byte length, exact SHA-256, and signature state. Adjacent
tests cover extraction, staged replacement, ownership-aware removal,
newly-unreachable cleanup, both signature policies, schema-position and
markup-compatibility grammar, strict XML lexical and namespace validation,
prefix-alias and unsupported-subtree preservation, save and reopen, and
byte-for-byte failure atomicity. Payload bytes are never decoded or executed.

The Presentation timing round-trip gate is
`the_corpus_timeline_preserves_every_unsupported_sibling`. It walks every
slide, layout, and master in the 50-deck corpus. The gate requires nonzero
coverage for timing roots, transitions, typed nodes, conditions, builds, set
values, transition parameters, compatibility-wrapped transitions, and
effect parameters. It serializes and reparses each model, compares the typed
projection, and inventories any unsupported node bytes before and after. A
category that becomes fully typed does not require a synthetic raw node.
Source-built regressions cover namespace aliases, compatibility choice and
fallback selection, singleton and schema order rejection, lexical owner-tag
preservation, atomic mutation, morph metadata, and the narrow PowerPoint empty
layout-transition compatibility shape. A parse-state size guard and the full
`rpptx` corpus package test keep the default test-thread stack sufficient.

The Presentation media package round-trip gate is
`embedded_audio_and_video_corpus_media_round_trip_without_duplication`. It
opens the configured `EmbeddedAudio.pptx` and `EmbeddedVideo.pptx` corpus
decks, inspects their audio and video sources, extracts exact embedded bytes,
saves and reopens, and requires stable relationship types, targets, content
types, poster ownership, playback settings, unsupported metadata, and package
part counts. The gate requires the external corpus in completion verification.
Source-built cases in the existing `rpptx-oxml` and `rpptx` integration
binaries cover embedded and linked add, replacement, extraction, removal,
failure atomicity, relationship ownership, shared and orphan payloads, raw XML,
namespace shadows, schema position, timing IDs, command fallbacks, duplication,
and same-presentation duplicate-slide shape-reference remapping. Tests assert
independent expected package facts rather than comparing only pre-mutation and
post-mutation views.

The deterministic media timeline golden gate is
`static_poster_output_and_timestamped_playback_state_match_source_built_oracle_fixtures`.
It builds poster, audio, video, link, codec, trigger, trim, volume, loop,
pause, seek, stop, and unknown-duration cases in the existing `rpptx`
integration binary. A valid poster and independent labelled Audio and Video
fallbacks use deterministic fonts at literal 150 dpi and pin exact decoded
RGBA SHA-256 values. Normalized timestamp rows pin the synchronized playback
state. Adjacent regressions preserve ordinary sibling diagnostics, distinguish
equal local shape ids across slide and inherited scopes, and require exact
legacy diagnostic strings and bytes from both existing render entry points.

The deterministic animation golden gate is
`animated_gif_and_motion_jpeg_avi_match_the_reviewed_two_machine_manifest`.
Its source-built two-slide deck has distinct visible content, an incoming fade,
click-triggered video playback, a separate click-controlled shape, explicit
outgoing-slide selection, and segments on both sides of the click boundary.
The exact GIF manifest pins timestamps, six decoded frame hashes, dimensions,
loop metadata, and the complete container hash. The exact AVI manifest pins
timestamps, RIFF duration and dimensions, every encoded JPEG payload size and
hash, every independently decoded frame hash, the complete container hash, and
all media diagnostics in order. Independent structural parsing verifies RIFF,
LIST, stream headers, every padded `movi` chunk, and every `idx1` record.
One-field mutations prove sensitivity for every required identity.

Adjacent regressions prove integer segment sampling, cumulative GIF delay,
single preparation with at most one retained resolved frame, output-cap failure
during codec writes, validation before rendering, JPEG quality sensitivity,
and unchanged static PDF and raster output. The same exact locked golden runs
in the Ubuntu workspace job and the macOS presentation-fidelity job. Both use
deterministic bundled fonts and require identical reviewed constants.

The ordered-body integration gate opens an in-code package through the public
Word facade and compares the exact direct sequence of paragraphs, a table, a
body content control, preserved producer XML, and a final paragraph. It also
proves that recursive paragraph and table accessors retain their existing
results. The adjacent low-level regression uses default and aliased Word
namespaces for self-closing paragraphs, tables, and final section properties,
while foreign same-local-name empty children remain byte-preserved raw XML.

The ordered compatibility reader regression builds its package in source and
compares every public direct item variant across body, cell, paragraph,
hyperlink, and run boundaries. It covers namespace aliases, typed field and
drawing facts, preserved raw subtrees, and legacy flattened accessors. Its
save and reopen matrix compares the ordered public facts and every exposed raw
subtree, including namespace shadowing and owner insertion, removal, and
reordering. Unsafe namespace replay must fail closed without changing the
opened package.

The Word reader-fact regression gate combines source-built drawing,
document, table, numbering, revision, and field fixtures. It requires Office
relationship expanded names, inherited namespace replay, table schema slots,
default-style numbering association, direct numbering overrides and
cancellation, bounded nested revisions, and ordered complex-field display
segments. Mixed owner-dependent and independently self-bound same-URI element
or attribute uses retain exact marker cardinality. Same-URI and different-URI
decoys still fail closed. The complete `rdocx-oxml` and `rdocx` suites, public
package dry runs, archive ceiling, and unchanged 49-entry hash harness complete
the gate.

The legacy horizontal-rule reader regression classifies canonical, aliased,
default, locally shadowed, and ancestor-bound Word, VML, and Office names by
expanded namespace URI. Its negative matrix keeps false, numeric, missing,
foreign, multiple-shape, visible-child, comment, and malformed forms as
unsupported XML. Adjacent regressions preserve the exact raw subtree and item
boundary through save and reopen, retain the earlier public `CT_R` literal
shape, include classification in equality, and prove ordinary modeled runs
retain no namespace-classification allocation. The pinned Word corpus, public
package dry run, archive ceiling, and unchanged 49-entry hash set complete the
gate.

Adjacent parser regressions require producer-defined numbering formats to
round-trip without an invented marker and require malformed encoded `w:t` or
`w:delText` values to fail before a partial document is published. Binding
coverage confirms that malformed document XML keeps the existing `XmlError`,
while HTML and ODT import failures keep the generic `RdocxError`.

The table-measurement parser gate covers integer and whole-valued decimal forms
at table width, cell width, table indent, and default cell-margin sites. Aliased
Word attributes are accepted, foreign same-local attributes are ignored, and
fractional, exponent, empty-fraction, overflow, percentage, universal-unit,
malformed, and empty forms return errors rather than zero. Round-trip evidence
requires canonical integer output in schema order and byte-identical unmodelled
table content. The pinned Word corpus and the 49-entry hash harness remain the
integration and output-stability gates.

The tracked table-grid gate parses canonical and aliased Word grid elements,
rejects foreign same-local names from the modeled projection while preserving
their bytes, and fails closed on duplicate modeled grid changes. Save and reopen
evidence requires active columns before one byte-identical historical subtree.
Facade and deterministic layout regressions prove presence inspection without
allowing historical widths to alter the active grid. The public package dry run,
pinned Word corpus, and unchanged 49-entry hash set complete the gate.

The release-notes regression gate validates both release tag families through
the same deterministic parser used by publication. It requires one exact
version section, the complete ordered heading set, meaningful rendered
Markdown in every section, and no placeholder tokens. Raw HTML cannot satisfy
semantic emptiness, while escaped element-like Markdown, visible link labels,
and real code content remain valid. Check and render modes leave the changelog
unchanged, and rendering returns the reviewed body byte for byte. Workflow
mutation matrices also require this validator before either crates.io publish
path, preserve the rendered artifact until a fresh byte comparison immediately
before GitHub release creation, and bind the release command to the same
preflight and post-publication body check.

Release review also reconciles one selected-family contribution inventory
against the rendered notes. Every included GitHub issue and pull request must
survive as a direct link, every authenticated external contributor must receive
specific credit, and cross-family records must remain excluded. Workflow
mutation tests pin the inventory, approval-report, and post-publication
notification requirements. After the release body verifies byte for byte, the
release records the comment URL posted to each included issue and pull request.

The sprint-workflow regression also covers ordinary and release dependencies
inside one sprint. An A to B to C regression requires each reviewed dependency
prefix to complete before its consumer starts, then returns the same state to
implementation. Each checkpoint commits its clean review file, records review
at that resulting HEAD, and repeats full verification because the evidence
commit changed HEAD. It does not create a confirmation review solely for that
review-file commit. Release cases extend the ordinary checkpoint with prepared
release and post-publication evidence HEADs plus separate immediate approval.
The ordinary final gate remains bound to the latest integrated HEAD. Review
pass numbers remain global, but each scheduled evidence boundary has its own
bounded remediation loop. Passing the global counter limit solely because
earlier boundaries finished clean uses the recorded extension path and does not
weaken the current boundary's limit.

The resume regression changes an existing F-ID's canonical title and size and
adds another story after run state exists. Reinitialisation with `--resume`
must refresh those two metadata fields and discover the new F-ID while
preserving phase, feature state, owner, wave, worker evidence, review records,
and verification records.

The Word field regression matrix records Microsoft Word 16.104 build
16.104.25121423 with an en-US locale, Gregorian calendar, period decimal
separator, comma grouping separator, and UTC clock context. Its readable
in-code `F-161-readable-field-matrix-v1` input covers every supported field
family and compares normalized document-order outcomes with exact literals.
Focused tests cover recursive IF operands, story-isolated SEQ state, typed
paragraph traversal, package properties and variables, explicit external
inputs, formatting pictures, and stable cached-display fallbacks. The oracle
is test metadata only. It is not a runtime dependency and adds no binary
fixture.

The extended Word field matrix uses the same pinned Word build and environment.
Its source-built `F-231-readable-field-matrix-v1` input covers formula, TOC,
TC, mail-merge control, display barcode, and merge barcode outcomes in exact
document order. Focused tests cover formula precedence, nested operands,
postfix percentages, format pictures, resource bounds, malformed input, stable
decimal normalization, story-isolated merge state, bare and explicitly
selected TOC forms, normalized TOC separators, sequence-prefixed page numbers,
and style lists, validated TC selections, recursive positional
and switch operands, shared quoted escapes, typed barcode options, the `CASE`
alias, and every barcode value, height, scale, rotation, and colour boundary.
Equivalent compact and spaced formulas share the 512-token parser limit. The
differential matrix includes supported outcomes and exact ordered fallback
diagnostics. The cache matrix covers resolved text, every structured outcome,
pagination deferral, unavailable explicit context, and unsupported
instructions. Every non-text or fallback result retains the original
instruction and stored display. Fallbacks also retain a stable diagnostic.

The dynamic TOC differential uses the same pinned Word build and locale. Its
source-built `F-232-dynamic-toc-rebuild-v1` input combines built-in heading,
custom-style, direct-outline, and selected TC sources. It compares exact entry
order, levels, internal links, and displayed page values. Focused regressions
cover final deterministic page targets for table sources, collision-free
bookmark reuse and allocation, sequence and separator behavior, multiple TOCs,
unsupported valid instructions, malformed ownership, duplicate bookmarks,
foreign ancestor and indirect-marker rejection, nested TOC ranges, lazy
maximum-id allocation, collision-safe owned placeholder substitution, and the
no-TOC no-op. Opaque-wrapper cases cover both complex and simple Word-shaped
fields. Nested full-paragraph bookmarks prove repeatable first-marker reuse,
mixed-case TOC and SEQ identifiers prove shared normalization, and one rebuild
binds two exact entry targets to distinct final pages. Same-namespace malformed
control and revision wrappers remain opaque. Accepted insertion coverage binds
inserted SEQ and TC fields plus inserted-only and mixed heading text to exact
entries. The maximum parsed outline value stays inside checked conversion, and
a control before a partial bookmark prevents whole-paragraph reuse. The
preservation case keeps alternate-prefix field scaffolding and neighbouring raw
XML while proving an untouched custom part is byte-identical after save and
reopen. Failure cases compare the live document before and after the rejected
rebuild.
Block-owner coverage keeps a direct body `sdtContent` opaque. It also places
invalid control paragraphs before a valid TOC to prove the byte scanner and
typed projection retain identical paragraph coordinates, and keeps invalidly
owned complex and simple TOCs opaque without diagnostics or mutation. Public
parser coverage round-trips standalone controls with each paragraph, table,
row, cell, and run child from the context-free union. Malformed shell coverage
uses an invalid modeled control id and a missing required property value at
block and inline placements. Nested complex and simple fields plus bookmarks
remain byte-identical and cannot shift a later valid rebuild. Revision coverage
rejects field ownership below the 32-wrapper parser ceiling and counts simple
TOCs inside accepted insertions and move-to revisions. Wrapped instruction
cases cover insertion, hyperlink, and content-control owners. Same-boundary
SEQ and TC cases exercise both control and accepted-revision orders. A bookmark
whose end precedes its start at one run boundary proves raw-position range
validation and atomic rejection. Content-control ownership tests keep a second
`sdtContent` child opaque and project accepted revisions inside the first
content child. Revision-depth coverage includes a property-change wrapper at
the overflow boundary. Alias-prefixed wrapped instruction runs retain the
namespace binding inherited from their wrapper. A revision-only hyperlink
before a positioned control proves raw-owner ordering. Inline-control text and
accepted revisions participate in layout, with exact later-page TOC targets
that move when that projection is omitted. Namespace injection covers a quoted
greater-than sign and a declaration repeated locally on the run. A terminal
hyperlink revision precedes a following same-boundary control. Same-paragraph
bookmark scope retains only the field between its exact markers. Old-result
exclusion keeps sources before the begin marker and after the end marker while
rejecting sources on both boundary interiors. Accepted insertion and move-to
coverage composes each revision with a nested content control, and the layout
case makes that composition move a later exact page target. Boundary tests put
selected SEQ and TC fields, the separator, and the end marker inside accepted
revision and content-control owners. They require the total nested position to
retain sources before the result while excluding sources inside it. Tracked
layout tests compose revision and inline control wrappers in both orders and
require exact visible text plus one paragraph change bar. Outer-coordinate
cases place revision and control owners after preserved raw paragraph children
and put a TOC in a terminal hyperlink revision. A hyperlink-retained simple
field before the TOC proves raw child shapes do not advance typed run
boundaries. The tracked control case requires a direct run-property revision
to render exact text and one paragraph change bar. Comment and processing
instruction prefixes exercise retained raw-slot accounting. Missing and empty
simple-field instructions prove raw fields do not advance modeled boundaries.
A selected TC inside a hyperlink revision immediately before the direct end
run remains owned. A styled end-boundary paragraph proves stale cached text is
removed from saved and reopened XML while its structural prefix remains valid,
and excluded from both the generated heading title and its bookmark target.
Begin and end boundary headings with pre-existing whole-paragraph bookmarks
prove generated hyperlinks and optional PAGEREF fields instead target exact
surviving-fragment bookmarks with no dangling generated reference. The same
cases prove every original crossing bookmark is repaired around that fragment
for hyperlink plus PAGEREF, hyperlink-only, and target-free entries. Partial
same-paragraph and cross-paragraph ranges exercise the same repair rule when
exactly one marker is consumed. End markers wrapped by a hyperlink, accepted
insertion, or inline content control keep heading text after the field inside
the same wrapper. Each wrapper case asserts the exact generated entry,
hyperlink and PAGEREF target, bookmark text range, and structurally valid XML.
After save and reopen, each generated wrapper-local target remains a matched
public bookmark, resolves through deterministic PAGEREF layout, and permits a
second rebuild. Repaired original wrapper-local markers exercise the same
reopen and repeat-rebuild contract under hyperlink plus PAGEREF, hyperlink-only,
and target-free entry policies. Parser coverage proves accepted owner order,
single projection of direct markers, and exclusion of foreign and malformed
same-namespace wrappers.
Adjacent bookmark-coordinate coverage asserts accepted public range direction
and text for revision-local and control-local markers. Tracked layout keeps
nested marker positions around deleted and accepted text, while TOC bookmark
scope selects only the nested TC, SEQ, and heading sources inside its accepted
half-open boundaries. A nested control with a local alternate Word prefix
retains its markers after reopen. Complex-field collapse, comment reference
insertion, and new bookmark mutation retain live REF and PAGEREF targets before
reopen. Generated targets in table and block-control paragraphs appear through
the public bookmark facade with the same recursive paragraph order used by TOC
discovery and layout.
An alias-prefixed self-closing paragraph-property element retains exact producer
attributes. An end-marker content control retains modeled identity, binding,
type, end properties, and ordered raw property slots after save and reopen.

The RTF reader differential records Microsoft Word 16.104 build
16.104.25121423 as the oracle. Its checked input is source-encoded RTF that
covers body order, run and paragraph formatting, tables, list overrides, PNG
images, diagnostics, and generated-DOCX reopen behavior. The normalized record
binds run formatting, including all-caps, small-caps, hidden text, breaks, and
tabs, to the generated and reopened body runs rather than relying on global
markers. The ignored regeneration gate opens the DOCX saved by that exact Word
build and compares the same structural record.

The RTF writer round-trip gate builds its document in code, writes RTF through
the native facade, reads the bytes back through the RTF reader, and compares
the normalized public structure for text, run formatting, paragraph
formatting, tables, lists, and PNG and JPEG images. Focused writer regressions
lock deterministic header table order, signed UTF-16 escaping, formatting
resets, table cell boundaries, multilevel list emission, truncating EMU to twip
image dimensions, atomic path saves, output bounds, diagnostic caps, and exact
location-aware diagnostics for unsupported body, paragraph, run, table, row,
cell, image, field, note, comment, bookmark, hyperlink, and raw XML cases.
All fixtures stay in source, and a DOCX preservation regression proves the
writer does not mutate unmodelled XML in the source package.

The HTML import regression gate uses source-built browser fragments and CMS
documents in the existing `rdocx` unit, integration, and regression binaries.
It compares exact paragraph, run, list, table, grid-span, vertical-merge, CSS,
and diagnostic facts in source order. Focused cases cover HTML5 parser repair,
collapsed and preformatted whitespace, saved `w:br` line breaks, semantic run
formatting, embedded and inline cascade precedence, unsupported constructs,
external resources, separate list identity, nine list levels, row groups,
multiple cell paragraphs, and nested-table loss diagnostics. Limit tests fail
closed for input, bounded path reads, retained text, DOM nodes and depth,
blocks, all Word runs, rows, columns, cells, and diagnostics. The integration
gate serializes and reopens the generated DOCX before comparing its public
structure. No binary fixture or sample is added, so all 49 hash entries remain
unchanged.

The MHTML gate stays in the existing `rdocx` HTML unit tests and Word
integration binary. Source-built MIME cases cover folded headers, root
selection, Content-ID and Content-Location resolution, every supported transfer
encoding, unsafe or unresolved resources, image MIME sniffing, 96 DPI sizing,
all parser and writer ceilings, deterministic CRLF output, source order,
deduplication, boundary collision avoidance, 76-column base64, and atomic path
saves. The integration record compares body order, formatting, tables, lists,
images, links, and ordered loss diagnostics after MHTML reparse and DOCX reopen.
Default HTML bytes and all 49 hash entries remain unchanged.

The ignored MHTML differential authenticates Microsoft Word 16.104 build
16.104.25121423 before opening one source-built MHTML document and saving DOCX.
It compares normalized public structure rather than package bytes. Independent
mutations to body text, formatting, table content, list identity, hyperlink,
image, and diagnostic records each fail the same acceptance predicate. Word's
`Strong` run style is accepted as the normalized representation of source
`strong` markup. Word drops the source-built contained PNG while rdocx retains
it under the declared MHTML image contract. The shared predicate compares all
other normalized fields and pins both sides of that intentional image
difference, so removing the rdocx image still fails acceptance.

The PDF import regression gate builds PDF objects, content streams, embedded
Carlito bytes, paths, text, and URI annotations in source. Unit coverage locks
strict parsing, page and object bounds, aggregate decompression, operation,
pixel, shape, and diagnostic limits, CropBox and rotation normalization, exact
12,700 EMU point conversion, comment-safe content decoding, explicit font
substitution, ordered unsupported-operator diagnostics, strictly positive dash
arrays at representable boundaries, zero-member and interior-phase isolation,
positive stops that convert to zero, and valid sibling recovery. Both modes
save, reopen, and validate in the existing `rpptx` integration binary. The ignored
differential pins Poppler 26.01.0 at 150 DPI, exact dimensions, exact editable
text and link facts, and a raw full-image luminance SSIM floor of 0.995.
Pixel-aligned representative geometry includes a 38.4-point styled square so
renderer-only antialiasing does not weaken the acceptance metric. An unchanged
source passes, while a 1.01-point geometry change, one-pixel imported geometry
shift at 150 DPI, and calibrated
pixel, text, and link mutations fail the same final predicate. No binary
fixture or hash baseline is added.

The EPUB regression gate builds the complete publication in source. It checks
the stored first `mimetype` entry, fixed timestamps and metadata, byte-identical
repeated output, front matter, outline-root spine splitting, nested navigation,
semantic XHTML, distinct and unmarked lists, bounded list depth, exact image
attribute correlation, referenced-only image bytes, absolute URI validation,
XML 1.0 character rejection, stable typed and raw loss diagnostics, and atomic
destination replacement. Projection limit cases cover source text, styles,
numbering, relationships, media bytes, and image occurrences before export
cloning or expansion. Focused cases cover uncloned table grids, named style and
deep-heading diagnostics, retained Roman and letter list formats, custom marker
losses, marker alignment, table-cell list diagnostics, interrupted-list
continuation, nested-list restarts, numbered heading elements and anchors,
bounded hyperlink spans, explicit no-underline formatting, image alternative
text and drawing-property diagnostics, alternate drawings, preserved text
spacing, column breaks, page breaks in formatted runs and field displays, IPv6
and IPvFuture hosts, malformed authority and fragment delimiters,
namespace-aware revision diagnostic deduplication, and every dropped custom
document property. Recovery cases also cover direct-only bounded heading text,
paragraph-local Word and foreign namespace aliases at one raw boundary with a
conflicting document-root binding, rejected extension-only and active SVG media,
duplicate PNG headers, invalid chunk type codes, illegal critical-chunk order,
indexed palettes beyond the declared bit-depth capacity, structurally validated
raster media, repeated JPEG start markers, scans before JPEG frames, baseline
and progressive JPEG controls, invalid GIF LZW minimum code sizes, empty GIF
image data, zero-sized GIF image descriptors, style-derived deep headings,
final section properties,
document backgrounds, visible and inert document defaults and default paragraph
styles, revision-only inert defaults, invalid HTTP user information, non-basic
underline styles, patterned, foreground, and invalid paragraph, run, and cell
shading, and the two distinct losses on preserved deleted text. The external test remains
ignored in ordinary local runs. The tracked CI test job downloads the reviewed release,
verifies the archive and JAR digests, sets `EPUBCHECK_JAR`, and invokes the exact
ignored test as a required gate. That gate requires exact
EPUBCheck 5.3.0 from the W3C release and accepts no validation error. The
reviewed distribution ZIP has SHA-256
`6c07e68584b2e2ce2f89fe06e1246dfead3eb36b46b340e7d93524f29dcff6c5`.
The gate also verifies the extracted validator JAR has SHA-256
`f7f96617c929371821609b88c8484d6dc9f24fe916499863c46094c5fb778a65`
before execution. One source-built oracle fixture combines front matter,
multiple outline roots, nested headings and lists, and media while asserting
source-ordered spine and navigation entries. It also includes a page break so
the external validator covers the lifted XHTML structure, plus an interrupted
ordered list, a numbered heading, a table-cell list projection, explicit
no-underline text, image alternative text, rejected active SVG, a style-derived
deep heading, non-basic underline and patterned shading diagnostics, preserved
deleted text, final section and background diagnostics, a visible default
paragraph style and visible document defaults, patterned cell shading, rejected
duplicate-IHDR and oversized indexed-palette PNGs, an invalid HTTP
user-information target, invalid PNG, JPEG, and GIF structures, valid baseline
and progressive JPEG and GIF controls, and paragraph-local Word and foreign
revision aliases under a conflicting root binding.
The writer does not mutate the source document, and a save and reopen check
proves retained unmodelled XML remains byte-preserved. No binary EPUB fixture
or runtime oracle dependency is added.

The digital-signature regression gate constructs its DOCX and signature XML
in source. A fixed RSA certificate produced by OpenSSL 3.6.3 on 9 June 2026
and precomputed RSA-SHA256 signatures cover both `ds` and `sig` namespace
prefixes. Focused tests verify prefix tolerance, strict algorithms, exclusive
canonicalization, relationship selection and order, complete declared
coverage, named part mutation, malformed or partial coverage, and read-only
save and reopen behavior. Creation tests use a fixed PKCS#8 RSA key and X.509
certificate produced by OpenSSL 3.6.3 on 22 August 2026. They assert schema
order, content-type-qualified canonical references, collision-free allocation,
key and certificate rejection, complete round-trip verification, and atomic
failure for invalid relationship graphs. The optional interoperability oracle
is Microsoft Word for Mac 16.104 build 16.104.25121423. The ignored gate writes
the generated DOCX, reopens those exact bytes for local RSA-SHA256 and complete
coverage verification, then requires an explicit human-evidence token after
Word recognizes the embedded digital signature and protects the document from
editing. Word for Mac does not expose a Windows certificate-trust verdict. The
oracle does not establish certificate trust or replace the Rust cryptographic
and coverage assertions.

The scalar template unit gate splits one tag across five differently formatted
runs and proves that the first matched run supplies replacement formatting
while unmatched prefix and suffix formatting remain intact. The structural
regression gate combines a nested body loop, conditional inclusion, and table
row loop in one readable JSON fixture. It compares exact paragraph and row
order. Focused tests cover dotted and lexical lookup, loop shadowing and scope
restoration, every JSON truthiness case, malformed and crossed controls, and
atomic rejection. The round-trip test clones section-ending paragraphs and
table rows, then proves their section properties and unmodelled XML remain in
schema order after saving and reopening. Headers, footers, text boxes, and
chart labels retain the scalar-only coverage shared with literal placeholder
replacement. The repeating-table regression gate expands three adjacent
template rows over ten records and compares all thirty rows in order. It also
checks banding, grid spans, vertical merge restarts and continuations, and
atomic rejection of an invalid repeated numbering reference. The continuous
numbering regression proves that mixed list and ordinary paragraphs retain one
`numId` and level without changing the numbering part. The paired round-trip
test compares row, cell, table, numbering, and raw XML state after reopen and
checks the raw children at their schema boundaries.

The flat mail-merge regression gate builds one readable record set in code and
compares exact separate and sectioned outputs in record order. It proves that
an absent record value becomes empty only under merge policy. Focused
regressions retain ordinary field fallback, switches, atomic failure, empty and
single-record boundaries, section-property order, tables, lists, and producer
XML after reopen. Scanner regressions cover simple and complex non-body fields,
nested header references, relationship-resolved footnotes, preserved raw body
references, entity-escaped bookmark names, collision-safe identity allocation,
and foreign same-local-name attributes. The methods are opt-in and no sample
invokes them, so the 49-entry hash harness must remain unchanged.

The rich mail-merge gate builds nested and named source records in code and
compares exact paragraph, list, table, image, fragment, and formatted-run order
with no remaining merge fields or region markers. Its unit matrix covers
lexical shadowing, named-source fallback, sibling isolation, empty regions,
crossed and missing markers, and local versus global callback counters. Focused
regressions repeat images and DOCX fragments across conflicting relationships,
styles, numbering, hyperlinks, bookmarks, content controls, and drawing ids.
They also cover formatter isolation and errors, exact EMU dimensions,
fixed-prefix schema order, byte-identical unrelated raw XML, external and
dangling fragment relationships, wrong value kinds, invalid XML characters,
and atomic failure. A compatibility regression compares both flat APIs against
their F-166 bytes. The rich API is opt-in and no sample invokes it, so the
49-entry hash harness remains unchanged.

The document-comparison regression gate compares the main document,
relationship-resolved headers and footers, comments, normal footnotes,
endnotes, fields, and nested text boxes. It covers paragraphs, tables, cells,
nested tables, lists, and modeled content inside existing content-control
shells. Accepting the generated revisions must reproduce every edited modeled
story, while rejection must reproduce the originals and leave no residual
tracked containers. The policy matrix fixes exact normalized revision kind,
content, order, story, and owner records for run, Unicode-scalar character, and
three-class Unicode word granularity. It independently covers formatting,
textual-whitespace, field, comment, and every story-category ignore, with a
rejecting mutation for each record dimension. Focused coverage locks down
deterministic repeated content alignment, same-story move pairing,
cached-result and complete-field changes, escaped metadata and collision-free
ids, minimal adjacent wrappers, direct inline-content-control alignment,
hyperlink shells, text-box source ownership, significant non-text boundaries,
supported run, paragraph, table, and section property revisions, stable story
diagnostics, atomic failure, and absent property-owner cleanup. The source-span
round-trip gate proves all unowned whitespace, comments, processing
instructions, foreign elements, prefix bindings, raw property children, and
relationships remain byte-exact and appear once in tracked, accepted,
rejected, saved, and reopened views. The legacy/default compatibility test
keeps `Document::compare` byte-identical to default options. No sample invokes
comparison, so the 49-entry hash harness remains unchanged.

The source-built external differential pins Microsoft Word 16.104 build
16.104.25121423 and locale `en-US`. Its 24 normalized records cover every
supported story, move, field owner, and property revision, and calibrated
mutations reject kind, order, story, pair, owner, and formatting changes. Two
producer representations are intentionally normalized. Word represents a
comment replacement with duplicated comments and anchors rather than nested
revision wrappers, while both forms carry the same deletion and insertion
semantics. Word pairs moves with one shared range name and distinct wrapper
ids, while rdocx uses one shared wrapper id. No sample invokes comparison, so
the 49-entry hash harness remains unchanged.

The redaction regression gate constructs one package in code with body,
table, content-control, header, footer, footnote, endnote, comment, inserted,
deleted, core-property, and custom-property occurrences. A second authored
chart fixture requires the exact literal to disappear from both ChartML caches
and its relationship-resolved workbook. The gate scans every inflated outer
and nested entry for UTF-8 and UTF-16LE forms after reopen. Focused cases prove
prefix-tolerant expanded-name matching, foreign same-local-name preservation,
unrelated part and relationship stability, malformed XML rejection, external
workbook rejection, nested ZIP limits, and atomic residual-scan failure. The
native-only API is absent from Python, WASM, and CLI wrappers. No sample invokes
redaction, so all 49 hash entries remain unchanged.

The watermark golden gate builds a five-page document in code, renders with
bundled fonts, and compares the exact PNG-byte digest for every page. It also
requires the selected watermark group to precede ordinary header and body
elements. Focused tests cover aliased VML projection, raw header preservation,
canonical generated child order, package-visible first and even variants,
same-type section inheritance, displayed page-number parity, header-scoped
image relationships, forced media-id collisions, blank selected variants,
entity-decoded settings, unsupported colour and media diagnostics, atomic
rejection, and margin-relative centering. No sample authors a watermark, so the
49-entry hash harness remains unchanged.

The Word glyph-provenance regression resolves every attributed run through its
result-local `WordSourcePath` and requires the selected paragraph's exact
Unicode-scalar slice to equal the displayed run. Its in-code fixture covers
ASCII, emoji, CJK, wrapping, body paragraphs, nested tables,
headers, footers, footnotes, endnotes, and accepted and tracked revision views.
Focused tests split sourced text in both Word and shared line stages, distinguish
duplicate paragraphs from repeated story layout, and keep generated or
non-bijective text unattributed. A repeated-text field regression places a
parsed complex cache beside literal text and a public simple field, proving
that actual projection ownership determines later scalar offsets. Compatibility
tests compare the existing low-level layout results after stripping provenance.
Both WASM targets and the package dry-run cover the intentional exhaustive
public literal change. All 49 hash entries remain unchanged.

The native full-layout regression resolves every positioned glyph-run font id
through the returned `WordLayoutResult`, resolves every attributed source id
through the same bundle, and proves PDF reuses the accepted layout `Arc`. A
caller-font fixture rewrites the family records of an in-memory TTF so neither
its name nor bytes can be supplied by bundled fonts, then requires every
sourced run to resolve to that exact owned font. The cache boundary populates
accepted layout before tracked calls and proves they neither replace its `Arc`
nor add an accepted invocation. Public integration coverage renders different
accepted and tracked revision text through the caller-font option accessor.
The workspace package dry-run, 10 MiB archive ceiling, and WASM target check
cover the additive published API. All 49 hash entries remain unchanged.

The bundled-fallback caller-font regression supplies an incomplete caller set
and requires requested caller faces to retain their exact bytes while missing
families resolve from the deterministic bundled inventory. The strict
caller-only path must still fail on the same incomplete set. Repeated edits and
checked transfer prove retained work remains reusable only for the exact font
bytes and complete document context. Rejection preserves both private engines.
Warm and fresh results compare pages, fonts, diagnostics, provenance, outlines,
revision options, and rendered PDF bytes. Staged mutation and poisoned-lock
cases prove publication and recovery remain safe. Both WASM targets, the
package dry run and archive ceiling, and the unchanged 49-entry hash harness
are required riders.

The caller-font comparison regression generates five valid font files totalling
exactly 22 MiB and supplies 40 aliases through the deterministic
bundled-fallback facade. Test-only structural accounting surrounds only the
retained-context exact font equality. An unchanged warm layout must report zero
bytes through that second comparison, retain bounded page identities, and equal
a fresh layout across pages, fonts, diagnostics, outlines, provenance, and PDF
bytes. Equal-length changed font bytes must still invalidate normal reuse, and
checked engine transfer must still compare the ordered family names and bytes
exactly.

The `document_facing_aliases_share_one_caller_font` gate uses caller bytes that
differ from bundled same-family bytes. Multiple document-facing names must
select that one caller face with exact bytes, diagnostics, provenance, and
shared ownership. Focused alias regressions cover exact-family precedence,
case-only labels, constructor metadata retained across changed additional-font
loads, CSS-like candidate choice, equal-context reuse, changed-context misses,
checked-transfer rejection, and warm and cold output equality. Entry and byte
boundary cases require oversized explicit alias slices to produce the same
deterministic identity in the font manager and reusable engine context.

The relayout-cache gate compares a warm normal-font result with a fresh cold
engine after editing one safe body paragraph. Pages, font table order and ids,
font bytes, diagnostics, revision view, and every resolved provenance span must
match exactly, while only the changed safe paragraph rebuilds. Focused cases
cover actual mutation invalidation, style and theme context changes, unsafe
numbering, fields, hyperlinks, media, relationships, ordinary and
`AlternateContent` drawings, nonempty diagnostic replay, late transactional
failure, paragraph reorder and insertion, caller-font isolation, TTC indices,
and a legitimate active set larger than 256 faces. Exact shaping tests require
newest-first lookup without FIFO refresh, force a fingerprint collision to
remain a miss until complete key equality, and prove that deriving spacing
once per parent segment leaves subsegment glyph ids, advances, and
Unicode-scalar source ranges unchanged.

The editor-scale paragraph-cache regression retains 700 distinct safe
paragraphs, edits one paragraph, and requires 699 warm hits with only the edit
rebuilt. The complete warm result and source map equal a fresh cold result, and
restart pagination reports a bounded rebuilt range. A forced fingerprint
collision still requires exact typed paragraph equality. Focused cases prove
that an early direct footnote or endnote reference still permits 699 hits and
one rebuild, while fields, numbering, drawings, and raw-child prefixes disable
later reads. Changing the reference ID misses its paragraph key, changing a
note part invalidates the exact retained context, note-bearing table and header
or footer content remains conservative, a late failure publishes nothing, hits
preserve insertion order, and FIFO eviction holds at the independently pinned
4,096-entry and 50 MiB paragraph limits. Cacheable
active paragraph and table blocks share immutable cache payloads through a
private representation. Warm and fresh results must retain exact pages,
structure, provenance, and nested table paths while public block APIs remain
unchanged. Compile-time checks also pin the 5,216-entry and 64 MiB combined
envelope.

The restart-pagination regression gate compares warm edits at the start,
middle, tail, and a retained page boundary with a fresh deterministic engine.
It requires complete equality of pages, fonts, diagnostics, provenance,
numbering, notes, fields, outlines, and rendered inputs. A middle edit must
share the unchanged page prefix and tail while reporting only its bounded
rebuilt range. Insertions and deletions have the same complete-equality check.
Source-built 700-paragraph cases prove unchanged footnotes, endnotes, default
headers, and page-number footers keep bounded restart work through both the
engine and bundled-fallback facade. They also prove endnotes append once,
changed related stories and note-reference sequences invalidate reuse, and a
footnote continuation cannot publish a dirty checkpoint. Multi-section
content, note-bearing tables, floating drawings, backgrounds, and mismatched
boundary state must use the full paginator. Ordinary multi-line prose,
headings, `keepNext`, and `keepLines` must publish complete-boundary restart
records. A deterministic Issue 67 fixture requires 175 naturally wrapped
four-line paragraphs to span 16 pages, keep the completed recorded pass, and
publish no checkpoint on a page ending inside a paragraph. Ten middle edits
must each produce 174 paragraph-cache hits and one build, recompute at most two
pages, and equal every field of a fresh deterministic result, including
metadata, logical structure, and the result-local Word source map. Late edit,
insert, delete, undo, note-bearing split, and displayed PAGE footer cases
remain exact. A 700-paragraph source-built case requires late edit, insert,
delete, and undo
results to equal fresh deterministic layout while recomputing only a bounded
page region.

The Issue 67 release-performance rider is an ignored timing-only regression.
It runs identical 175-paragraph and 700-paragraph sources through the reusable
native and deterministic bundled-fallback paths in four alternating rounds.
Before timing, each run authenticates a deterministic manifest of every
tracked crate and workspace-manifest byte, the surrounding regression source,
and the exact benchmark harness. Reference runs also require the pinned commit
identity. The current manifest is content-bound instead of parent-SHA-bound,
so committing or integrating the reviewed bytes does not invalidate it.
The median of each build's per-edit medians must be no worse than 1.25 times
the immutable v0.11.1 release and at most 0.75 times pinned commit `0582da0`.
Wall-clock thresholds do not run in the normal unit-test pass.

The incremental-layout scale gate builds 1,000 one-page paragraphs through the
public deterministic bundled-fallback facade, edits paragraph 500, and compares
the warm result with a fresh layout. It requires exactly 1,000 pages, at most
two warm page-layout invocations, at least 998 retained page-frame `Arc`
identities, 999 paragraph-cache hits, one paragraph build, and complete result
equality. The paired engine gate requires a 1,024-page restart record to remain
within the aggregate cache budget and a 1,025-page record to fall back safely.
An additional candidate larger than the former 8 MiB limit must publish when
the actual aggregate remains within 64 MiB. A candidate above the aggregate or
an arithmetic overflow must fail closed without changing output.

The substituted-page regression gate proves that unchanged PAGE, NUMPAGES, and
PAGEREF pages reuse their prior substituted frame only through pristine `Arc`
identity and a complete exact substitution key. Focused mismatches cover page
index, displayed page number, total-page count, bookmark targets, pristine
content, font identity, and revision view. Field-bearing blocks retain zero
pagination checkpoints. Field-free pairs share one frame, oversized pair sets
drop the record, and warm output, PDF bytes, and raster pages equal a fresh
deterministic engine. The unchanged hash harness covers the sample backend set.

The empty-paragraph attribution regression covers body, nested table, header,
footer, footnote, and endnote stories. It requires exactly one empty,
zero-width segment with the paragraph source and scalar range `0..0`, while
ordinary layout keeps the same structure without a source. Direct paragraph
mark formatting and paragraph-style defaults select the carrier font and
size, and the segment ascent and descent equal that font's resolved metrics. A
compatibility case keeps non-empty text unchanged, proves ordinary and
attributed layout structure agrees after removing source ids, and proves PDF
and raster output is unchanged when the no-glyph carrier is removed. The
sample page-one raster and resource hashes remain unchanged.

The safe-table cache gate proves an unchanged recursive table hits, diagnostics
and font traces replay, and outer and nested provenance rebind after an earlier
body insertion. Numbering and other traversal-sensitive table content bypass
the cache. A late font failure after staged paragraph and recursive table work
publishes neither queue. Focused bounds checks cover both published and pending
entry and byte ceilings, including the complete nested retained payload.

The safe header and footer cache gate covers default, first, even, inherited,
header, footer, and watermark variants. It requires exact hits to replay
diagnostics and font traces and to rebind current Word source ids. Part text,
resolved image bytes, watermark geometry, same-width page-height changes,
styles, numbering, notes, theme, revision view, additional fonts, section
properties, and provenance mode must miss. Traversal-sensitive parts bypass
reuse. A late failure after staged header work publishes nothing. Published and
pending queues remain within 64 entries and 4 MiB, oversized entries bypass
retention, and warm deterministic layout and PDF bytes equal a fresh engine.

Boundary tests exercise the exact shaping identity, process font discovery,
canonical file-byte identity, lock poison recovery, bounded resolution and
coverage state, the 256-entry and 64 KiB caller-alias identity, bounded and
shrunk per-paragraph font traces, and both pending and published block queues.
Structural byte tests use retained capacities for
owned keys, rows, cells, blocks, glyph data, diagnostics, font traces, restart
pages, and reflow parameters including tab stops. The combined retained state
must stay within 5,216 entries and 64 MiB, with paragraph state capped at 4,096
entries and 50 MiB, table state capped at 32 entries and 2 MiB, header and
footer state capped at 64 entries and 4 MiB, and restart state capped at 1,024
entries. Restart candidates use the checked 64 MiB aggregate budget instead of
an independent byte cap. Oversized entries must bypass retention.
Repeated and concurrent focused tests preserve `Document: Send + Sync`. The
no-default feature test, both WASM checks, committed-graph package dry-runs,
archive-size ceiling, and reviewed 49-entry hash harness are required riders.

The OfficeMath round-trip gate is
`officemath_corpus_parses_mutates_saves_and_reopens_without_losing_supported_or_raw_siblings`.
Its source-built corpus covers all thirteen supported expression variants and
opaque root, property, and argument siblings through typed mutation and
reopen. Focused grammar tests cover inline and display equations, aliases,
fixed-prefix writes, schema child order, property defaults and domains,
malformed sequences, XML depth, text decoding, conflicting namespace bindings,
and legacy Equation Editor isolation. Existing integration targets cover
paragraph item order, collapsed raw-boundary rebasing, full-corpus authoring,
display justification, settings relationships, and mutable facade reopen. The
published-crate riders deny rustdoc warnings, dry-run both packages, and keep
each archive below 10 MiB. The 49-entry hash harness remains unchanged.

The OfficeMath layout gate is
`officemath_baselines_and_glyph_geometry_match_the_pinned_word_pdf_oracle`.
Its source-only harness builds one DOCX that covers all thirteen supported
expression families and pins Microsoft Word 16.104 build 16.104.25121423,
Poppler 26.01.0, and 150 DPI. It requires exact Word text tokens and derives
each expression's ink width and vertical bounds directly from the Word and
deterministic Rust PDFs. Fixed raster windows separate only the delimiter and
accent that Poppler coalesces into one Word token. Aggregate and
per-expression geometry use a 1.0 point tolerance. The complete-page raster
uses 64 by 64 pixel block luminance with a 0.99 SSIM floor. A 1.01 point
rendered-group perturbation proves the geometry and raster path is
mutation-sensitive. The source DOCX digest and tool identities live in the
text manifest, while both PDF outputs remain untracked.

## The hash harness

The single highest-value mechanism in the plan is
`scripts/hash_harness.py --check`. It deletes the expected generated outputs,
runs `generate_all_samples`, and records the flushed `word/document.xml`,
`word/styles.xml`, and `word/numbering.xml` state, the page-one PNG, and a
three-part fingerprint of the deterministic PDF for each of the seven samples.
PNGs are rendered at 150 dpi through the deterministic font path.

PDF is fingerprinted because it is a first-class output written by a different
code path from the PNG. Rasterising page one exercises none of the writer's
glyph positions, CID font subsets or ToUnicode CMaps as bytes, and nothing at
all beyond page one. Three entries per sample:

| Entry | Covers |
|---|---|
| `<sample>:pdf/pages` | The page count, each page's `/MediaBox`, and each page's inflated content stream, in `/Kids` order |
| `<sample>:pdf/resources` | Inflated font subsets, ToUnicode CMaps, image XObjects, and other non-content streams except `/Type /Metadata` |
| `<sample>:pdf/bytes` | SHA-256 of the file as written |

The first two hash inflated bytes, so they say **what** moved and survive a
change of Deflate implementation or level. The third says **that** something
moved and cannot be evaded, including by a change that is purely in
compression. A fingerprint of extracted text and page geometry alone was
rejected, because the dependency refresh in F-X020 moved all seven sample PDFs
while `pdftotext` output stayed identical in 7 of 7.

Document metadata streams are excluded only from `pdf/resources`. They are not
page resources, and their complete bytes remain covered by `pdf/bytes`. A
focused scanner test adds a `/Type /Metadata` stream and requires only the byte
entry to move. The existing changed-resource test continues to require a real
font-like stream change to move `pdf/resources`.

The harness reads a PDF with a scanner over the object syntax, using the
standard library alone, and raises rather than skipping anything it does not
understand. A missing PDF is an error and not an absent entry, because `null`
means "this optional XML part is absent by design" and a sample whose PDF failed
to generate is not that.

The sorted `scripts/hash_baseline.json` manifest has 49 entries. Each entry is
either a SHA-256 digest or JSON `null` when an optional XML part is absent.
Check mode reads the manifest without modifying it and reports added, removed,
and changed entries. Baseline writes require `--update --reason <text>`, and an
empty reason is rejected. Generated PNGs remain ignored under `samples/`.

The current reviewed table-fidelity delta changes exactly
`feature_showcase:pdf/pages` and `feature_showcase:pdf/bytes`. That sample's
later PDF page contains a valid vertical merge and a bordered nested table, so
correct merge-edge suppression and recursive grid painting change its page
stream. Its PDF resources, page-one PNG, selected OOXML parts, every other
sample entry, and the manifest cardinality remain unchanged.

It exists because the extraction changes unit conversion and text-shaping input
types, and both alter output **without failing to compile**. Structural
round-trip tests cannot see that class of defect.

Rules:

- Every PR in M1 through M6 gates on it.
- Baseline updates require a non-empty review reason.
- An intentional behavioural change lands as its own labelled commit with its
  expected delta stated in the message and reviewed.
- An unexplained delta blocks the merge.

## The golden-PNG gate

`python3 scripts/golden_png_harness.py --check` generates deterministic PDFs
for the seven `samples/` documents, rasterises page one at 150 DPI with
`pdftoppm`, and compares decoded RGBA pixels. The rasteriser is test
infrastructure only. Its exact version is printed on every run and recorded in
`scripts/golden_pixel_manifest.json`. The current manifest records
`pdftoppm version 26.01.0`.

Each readable manifest entry contains the page width, height, and SHA-256 digest
of the decoded RGBA buffer. There are no committed PNG fixtures. Check mode
requires identical dimensions and a zero-pixel-difference digest, then reports
the first differing sample precisely. Reviewed updates use `--update --reason
<text>`, and an empty reason is rejected.

The gate deliberately compares pixels rather than PDF bytes. The operator
stream legitimately changes when the per-element Y flip becomes one global
CTM. The reviewed Poppler 26.01.0 baseline includes exactly four
stroke-antialias changes. In `invoice`, pixels `(112, 397)` and `(112, 398)`
swap `fcf5f5ff` and `ffffffff`. In `quote`, pixels `(112, 303)` and
`(112, 304)` swap `f4fafaff` and `ffffffff`. The other five samples remain
exact. This is a baseline, not a tolerance, so check mode still requires exact
equality for all seven buffers. The regression proof runs `--check
--inject-one-pixel <sample>`, copies that generated PNG to a temporary
directory, changes exactly one decoded pixel, and requires check mode to fail
with the sample name.

The pull-request `test` job runs the golden-PNG check after the full workspace
suite. That job installs the pinned Poppler 26.01.0 oracle first, so the decoded
pixel comparison is unconditional, failure-propagating, and bound to the
reviewed rasteriser identity.

The `feature_showcase` page-one golden is also the automatic-hyphenation
acceptance case. It enables the Word setting, assigns `en-US` to
`representation`, and constrains the line so deterministic output is `repre-`
followed by `sentation`. The reviewed LibreOffice Writer 26.2.5.2 oracle makes
the same break. The hash and pixel manifests may move only for that declared
sample after the source-built layout, XML round trip, deterministic raster, and
pinned Writer evidence agree.

## The SVG page golden gate

`svg_page_rasterises_like_the_png_backend` constructs one representative page
entirely in code. It includes exact-font searchable text, an embedded image, a
path and normalized gradient, three recursive groups with noncommuting scale,
rotation, and skew, a clip, opacity, nonzero shadow blur, a safe link, marked
content, and a diagnosed paint fallback. The PNG side uses deterministic
bundled layout at 150 dpi. The SVG side is rasterised at the same exact 300 by
300 dimensions by development-only resvg 0.48.1, whose font database receives
only the layout result's explicit font bytes and exact face identities.

The comparison composites both RGBA buffers over white and requires global
luminance SSIM of at least 0.99. A one-point view-box perturbation must score
below 0.99, which proves the calibrated threshold rejects a visible placement
regression. No PNG, SVG, or font fixture is committed. Focused regressions also
cover deterministic definition order, searchable complex text, XML and link
safety, recursive sibling preservation, transform composition order, singular
effect omission for unprovable text ink, and non-clipping singular geometry
bounds.

Revision-view rendering has a separate deterministic two-view golden gate. An
in-code Word fixture renders accepted and tracked views with bundled fonts at a
fixed DPI. Accepted pixels must equal the same document after `accept_all`
removes the wrappers, while the tracked pixels must differ. Regression coverage
also parses a revision whose paragraph splits across pages, requires one
outside-margin change bar on every fragment, and compares text positions with
an unchanged control. The general hash baseline remains unchanged at 49
entries.

The cross-family native-chart golden constructs one Word document and one
PowerPoint presentation from the same `ChartData`, page size, chart rectangle,
and effective theme. The `rdocx` test target takes `rpptx` as a development-only
dependency, so production dependency trees retain no Word to PowerPoint edge.
Both deterministic PDFs use bundled fonts. Poppler
`pdftoppm version 26.01.0` rasterises the chart rectangle at 150 DPI into
750 by 450 pixel crops, and the decoded RGBA comparison requires exactly zero
differing pixels.

The SHA-bound Word artifact is
`e50845637449e2af4b8e2dbf16f5f6f53e5f598a00401fcc34c13f5d5716a1c4`.
The SHA-bound PowerPoint artifact is
`7525e9a088c5fbf58fa1ed98cdfa0ec2fabf998662112ced7a6b6521f2c4edfc`.
The recorded crop result is `750x450 differing=0`.

## The deck corpus

Fifty real `.pptx` files are stored outside the published crates and fetched by
`scripts/fetch_pptx_corpus.py` into the ignored `corpus/pptx` directory. The
tracked manifest pins each URL, producer, relative path, and SHA-256. It
contains 49 Apache POI slideshow test decks at commit
`11ede1db13c554b4341266faeb84e327fc316379` and one public Google Slides export.
`--check` verifies the complete directory without changing it. The set spans
producers because non-Microsoft writers are where parser assumptions break:

- PowerPoint 2016 and Microsoft 365
- Google Slides export
- Keynote export
- LibreOffice Impress
- A multi-master corporate template
- Decks containing SmartArt, charts, embedded audio and video, and ink

The read-facade differential runs `dump_deck` over all fifty decks and compares
its normalized records with python-pptx 1.0.2. The executable test command pins
that exact oracle version with `uv run --with python-pptx==1.0.2` and rejects a
different resolved version. Records cover slide id and name, recursive shape
path and structural kind, ordinary shape text, row-major table text, aggregate
slide text, and optional speaker-note text. Empty python-pptx names and shape
text capability without a stored `p:txBody` are normalized to the facade's
explicit `Option` contract.

Eight gates run against it:

1. **DrawingML structural round-trip**: every `a:txBody` and `a:spPr` parses,
   serialises and reparses to a structurally equal value. The pinned corpus has
   6,898 text bodies and 8,643 shape-property elements. This is the carried M7
   exit gate at the first point where the external corpus exists. Every
   `ppt/tableStyles.xml` part also parses, serialises, and reparses through the
   typed table style model while retaining unsupported XML at its boundary.
2. **Raw round-trip**: open and canonically save with every document part
   treated as opaque. Every decompressed part stays byte-identical, while
   content types and relationships stay structurally equal. ZIP metadata and
   compression are not model state. This proves the OPC layer and the corpus
   harness before any PresentationML modelling exists.
3. **Modelled round-trip** (M8 exit): parse and serialise the presentation,
   slide, layout, master, notes slide, notes master, and theme roots. Reparse
   each canonical result and compare it structurally. Build the expected
   package from those exact modelled bytes, retain the original bytes for all
   unmodelled parts, save through deterministic OPC output, reopen, and compare
   content types, relationships, part names, part counts, and every part byte
   against that expectation. The gate requires nonzero corpus coverage for all
   seven root types.
4. **Timing model round-trip**: every slide, layout, and master timing or
   transition subtree projects the supported model while unsupported siblings
   retain exact bytes. Coverage counters must remain nonzero for every bounded
   timing category, including compatibility transitions. Any raw nodes present
   are inventoried before and after without requiring the corpus to contain one.
5. **Opens without repair** (M8 and M11): every saved deck opened manually in
   PowerPoint once per milestone. Not automatable, and not skippable.
6. **Media package round-trip**: the tracked embedded audio and video decks
   preserve exact media bytes, relationships, content types, poster ownership,
   playback settings, unsupported metadata, and part counts through save and
   reopen without duplication.
7. **SmartArt typed projection and preservation**: every diagram data, layout,
   quick-style, colour, and cached drawing part projects through its bounded
   namespace-aware model without rewriting source bytes. Source-built fixtures
   prove producing-scope ownership, relationship-role validation, checked node
   editing, complete graph remapping, schema-position sensitivity, and failure
   atomicity. The gate requires nonzero coverage for all five part families
   when the pinned corpus is required.
8. **Executable-content producing scopes**: the tracked
   `alterman_security.pptx` deck inventories relationship-owned OLE objects in
   slides, layouts, and masters without interpreting their payloads. The gate
   requires exactly two layout owners and one master owner and proves each
   reported payload extracts to the reported byte length.
9. **Notes and handout deterministic export**: a source-built package uses
   noncanonical notes-master, handout-master, notes-slide, theme, and media
   targets. PDF and PNG checks cover page order, `notesSz`, exact slide-image
   edges, header and footer metadata, a slide without notes, all six handout
   grids, master-behind-thumbnail z-order, three-up note rules, and package-byte
   preservation. Relationship negatives cover missing, external, duplicate,
   wrong-type, malformed, and equal-id cross-scope cases. Placeholder cases
   prove index-first and type-fallback matching, ambiguity rejection, and
   unmatched rejection. The geometry unit gate covers exact targets, clipping,
   five rules, and rejection of a 1.01-point displacement. The 49-entry render
   hash manifest remains unchanged.

The portable M21 core test source-builds one macro-enabled signed deck that
combines modern comments, sections, the self-contained minimal SmartArt
preservation fixture, exact embedded audio and poster bytes, a typed fade
timeline, three notes pages, and a three-up handout. Save and reopen preserve
those semantic surfaces and real source mutations change the static, animated,
notes, and handout outputs. This portable test classifies the minimal SmartArt
render as an unsupported fallback and does not claim authentic SmartArt raster
fidelity. A separate ignored macOS reference-only writer reads SHA-256-pinned
authentic layout, quick-style, and colour resources and emits corrected signed
and signature-free sources for manual oracle capture. The mandatory ignored
release oracle does not read those installed resources. It reads the captured
signed source and four outputs from one configured oracle directory. Its
embedded manifest records PowerPoint 16.104 build 16.104.25121423 and
AppleScript build 1214, `signed=true`, exact source and output SHA-256 values,
and `open_no_repair=true` for an observation whose active name was exactly
`m21-corrected-signed.pptm` and whose hash was the canonical signed source hash.
The signed macro-enabled source
with SHA-256 `74fe838af835fbf9852d232d1eb39683bfbb1381b86095073e9e96974b50aac9`
is canonical. Every Rust semantic and output check starts from its bytes. The
shared semantic assertion repeats after save and reopen. It pins slide IDs and
order, notes ownership, comment author and reply cardinality, section identity
and slide membership, WAV and poster bytes and content types, playback
settings, typed fade target and duration, complete signature coverage, and the
exact SmartArt relationship IDs, data topology and text, layout identity, and
style and colour identities. Authentic mode also rejects any unsupported
SmartArt fallback. The
portable source-built signature-free package
differs only by the empty signature origin, the XML signature, their owning
relationship, the package origin relationship, and their two content-type
overrides. Every other package part and relationship is byte-identical,
including presentation, media, theme, SmartArt, timing, notes, handout, macro,
and render relationships. The directly bound signed outputs have SHA-256
values `aebe97df20d029a611afa935fad0e96653e0b515396ce7ec1f5e2c665d92f8de`
for the three-page static PDF,
`4643c6cb25222b343067364a8983673c79962e32378809206b7f9e6f5306e5e9`
for the movie,
`d940316865a28e626c2cc7756d9bef4f132c516d03cba63387e1f6f0ca0dba2a`
for the three-page A4 portrait notes PDF, and
`77345fd00914bb2b233bf548530bd2f6de05c25b53a08cd7392bf38be696d05f`
for the one-page A4 portrait handout PDF. The movie is sampled at 0, 297,
and 594 of 600 ticks. The
static visual gate compares all three pages and masks only the declared
audio-poster rectangle on page one. Each page has exact normalized token
cardinality and order, full-page ink geometry within 6 pixels at 150 DPI, and
at least 0.45 SSIM in each union ink region. Page three must retain the complete
SmartArt graph, relationships, three-node text, and visible ink. The movie
manifest records the exact visible token vector and ink-band count observed at
each of its three samples. The movie gate compares that observation with the
actual Rust frame text and applies the same page-one visual boundaries.
The low regional SSIM floor accounts only for deterministic Carlito versus
PowerPoint font rasterization and is paired with exact text and geometry gates.
Real shifted-raster geometry, source text, extra-token, duplicate-token,
reordered-token, token-containing, and solid-raster mutations fail their
applicable predicates across the three static pages, animated output, notes,
and handout output. Notes require three Rust pages and three A4 portrait
PowerPoint pages with exact per-page token vectors and exact bounded
monochrome-band cardinality. The semantic note component on each side compares
by normalized width and height within 0.06 of one page dimension and by
monochrome ink occupancy within 0.35. Absolute component placement is not
equated because the two notes masters and page sizes differ. Extending or
solid-filling the semantic component on either side fails the combined notes
predicate. All three handout thumbnail bounds compare in normalized page
coordinates within 0.05 of one page dimension. Geometry mutations beyond that
boundary fail.

The earlier minimal SmartArt source remains a local regression for an explicit
unsupported fallback. An independent ignored classification requires only its
one recorded static PDF to prove that PowerPoint 16.104 renders a blank third
page where Rust renders `Unsupported SmartArt`. Neither result is
representative acceptance evidence. The corrected captured-source oracle and
the native SmartArt render differential below supply the visible-render
evidence for authentic SmartArt resources.

The native SmartArt render differential uses six one-slide source decks built
from exact SHA-256-pinned PowerPoint 16.104 layout, quick-style, and colour
resources. The same source deck feeds PowerPoint and the native facade. The
manifest binds the source, PowerPoint PDF, normalized PNG, exact shape and text
ownership, bounds, diagnostics, dimensions, and ordered text line counts.
Every shape edge must remain within 1 point. Symmetric text-masked non-text
SSIM must be at least 0.90. Owner-centered horizontal ink edges, raw vertical
ink edges, and line widths must remain within 3 points. Full-image SSIM is a
diagnostic because the oracle uses Calibri and deterministic mode uses the
bundled metric-compatible Carlito font. A 1.01-point displacement and a
calibrated decorative paint and size mutation prove the thresholds reject
material divergence. Required-corpus mode fails when any manifest artifact or
provenance hash is absent.

Five supported layout programs use the bounded typed instruction evaluator.
The exact three-node `cycle1` resource uses one private PowerPoint 16.104
compatibility profile because its producer diameter and curved-connector solve
is not specified by OOXML or the resource program. Production rendering
requires the exact layout identity and resource SHA-256. A changed identity,
resource byte, instruction, or node count fails closed. This exception adds no
facade, binding, layout, or renderer API.

The M9 resolver gate selects `WithMaster.pptx`, `backgrounds.pptx`,
`placeholder-layout-color.pptx`, and
`bug58144-headers-footers-2007.pptx` for native visual acceptance. Its
repeatable normalized differential also includes `60810.pptx`, whose master
picture appears exactly once on every enabled slide and zero times on its two
`showMasterSp="0"` layouts. The executable oracle pins python-pptx 1.0.2 and
compares ordered source shape kind, bounds and text. Concrete resolver evidence
retains RGBA fills and unsupported diagnostics, including exact cyan `#00FFFF`
for the inherited placeholder run. The one-time PowerPoint record in the
integration test names build 16.104.25121423, the exact original paths and the
clean no-repair verdict.

These automated visual tests use the same external-corpus policy as the other
corpus gates. A missing configured corpus skips them when
`RDOCX_PPTX_CORPUS_REQUIRED` is unset and fails them when it is set. The
one-time native acceptance record does not require the external files to remain
present after review.

The native timeline differential uses a source-built deck and Microsoft
PowerPoint 16.104, Info.plist build 16.104.25121423, and AppleScript build 1214.
The ignored source, movie, raw PNGs, and manifest are bound by SHA-256. The
source hash is
`a1f610feab5ee9ba0629c1b4731cf5e5b1453f6980a7574770036010b1f833fa`,
the movie hash is
`28514432f4aafae9d6c5ddd522d23e87458e08ec59ebf5555398a74e712fa83e`,
and the nine-case manifest hash is
`cd0e3f582e55546432c83528e1967926c45f3ba6258a5c8c195c0421974d611c`.
Each case binds a source slide, Rust slide-local timestamp and click count, and
an exact rational movie sample. Gate-side AVFoundation re-extraction must
return the same sample and the same 1920 by 1080 raw bytes. One deterministic
in-memory resize normalizes only the verified oracle image to the unchanged
2001 by 1125 Rust raster produced at literal 150 dpi.

The required gate covers five automatic timeline states, a terminal outgoing
state, fade, morph, and push. Exact click and fill one-millisecond boundaries
stay in Rust regressions because the 600-timescale movie cannot distinguish
them. Zoom remains covered by Rust direction and composition regressions because
the source movie's zoom interval is byte-identical and therefore cannot supply
independent external evidence. Foreground geometry permits at most 1 point
error and global luminance SSIM must be at least 0.99. The recorded nine cases
have geometry errors from 0.48 to 0.96 points and SSIM from 0.997866 to
0.999953. Mean geometry error is 0.533333 points and mean SSIM is 0.999543.
Required-corpus mode fails closed on missing artifacts, hashes, provenance,
sample identity, dimensions, normalization provenance, or case coverage.

## The Word corpus

The modern Word package-class gate source-builds DOCX, DOCM, DOTX, and DOTM
from one valid WordprocessingML graph. ZIP, `Document`, Flat OPC, and converted
ZIP round trips must retain the exact class, VBA bytes, relationship scopes,
and unrelated XML payloads. A four-way conversion comparison proves the main
override is the only package difference and that the live document is
unchanged. Malformed expanded names, duplicate parts, unsafe paths, wrong data
kinds, invalid base64, malformed relationship owners, permissive-parser
lookalikes, and each resource limit fail before publication. Microsoft Word
16.104 build 16.104.25121423 supplies only the ignored no-repair acceptance
fact. Source-built assertions remain the structural authority.

The M22 completion gate composes its feature families in one source-built
macro-enabled template. It authors and deterministically renders OfficeMath,
rebuilds a dynamic table of contents and field caches, performs sectioned mail
merge and full document comparison, inventories the retained VBA project, and
round-trips through Flat OPC. The final package must retain its DOTM identity,
exact executable bytes, equations, and unsupported XML. Separate focused tests
cover inherited Flat OPC payload namespaces in qualified names and
markup-compatibility values, plus the required binary treatment of Transitional
and Strict alternative-format import targets. The composed predicate inspects
the rebuilt TOC cache, section boundary, and body and header comparison output
so each milestone operation is mutation-sensitive.

Five real `.docx` files are stored outside the published crates and fetched by
`scripts/fetch_docx_corpus.py` into the ignored `corpus/docx` directory. The
tracked manifest pins one document for each of `business-letter`, `report`,
`form`, `legal-revision`, and `multi-script`, with its producer, SPDX licence,
immutable licence URL, immutable source URL, relative path, and SHA-256.

Four documents come from Apache POI at commit
`11ede1db13c554b4341266faeb84e327fc316379` under `Apache-2.0`. The tracked
revision contract comes from `sontanon/docx-mcp` at commit
`891aabaa6b33eb93d867b5d69adb5991bdfbde69` under `MIT`. It is a Microsoft Word
Act of Engagement contract containing tracked insertions and deletions.

The fetcher accepts only the exact five categories and the reviewed
`Apache-2.0` and `MIT` licence and licence-URL pairs. It rejects an unsafe or
duplicate leaf path, duplicate source URL, missing producer, non-HTTPS URL,
invalid lowercase SHA-256, incomplete category coverage, missing file, extra
file, and digest mismatch. Downloads use a temporary sibling and replace the
destination only after its digest matches. `--check` verifies the complete
directory without changing it. The primary workspace-test and MSRV jobs fetch
both pinned corpora before running Cargo tests.

## The Word render fidelity gate

The five-document Word corpus is rendered at 150 dpi through the production
`rdocx-cli render` command and its bundled-font deterministic layout. Exact
LibreOffice Writer 26.2.5.2 build
`cd7284b4cbbfeb507e630c1aac019f4157393acb` opens an accepted copy through an
isolated profile and exports `writer_pdf_Export`. The copy is prepared through
the existing `rdocx::Document::accept_all` API by a locked, offline, untracked
helper. The harness reopens it and rejects remaining modeled revisions before
Writer runs. Exact pdftoppm 26.01.0 rasterises every oracle page at 150 dpi.

The harness scores the union of page indices through the larger Rust or oracle
page count for each document. It composites unequal dimensions at the top left
of a shared white canvas sized to the maximum width and height. A page missing
from one side receives a blank white counterpart. The TSV records document,
category, page index, both original dimensions, normalization action, SSIM, and
both paths. JSON evidence records per-document Rust, oracle, and union page
counts plus dimension mismatch and unmatched-page totals.

The reviewed calibration contains 18 Rust pages and 16 oracle pages across 18
union indices. Fourteen paired pages need one-pixel white-canvas normalization,
and two Rust-only pages use blank oracle counterparts. One of 18 pages reaches
0.95 SSIM. Coverage is 0.055555556, minimum is 0.020380485, median is
0.067698319, and maximum is 0.974913551. These results expose current fidelity
and do not change the advisory target.

**Trend reference: at least 0.95 SSIM on at least 80 percent of pages. Hard
automatic gate: the exact corpus and tool identities match, both renderers
succeed with nonzero output, and the expected TSV and JSON artifacts exist and
are nonempty.** Page-count differences, dimension differences, and a missed
trend are scored evidence rather than orchestration failures.

The same harness source-builds five one-page Word fixtures for Arabic,
Devanagari, Thai, Simplified Chinese, and mixed bidirectional text. Each uses an
approved deterministic Noto family, complete `w:lang` attributes, 24 point
text, and exact 24 point line spacing. The bidirectional page also carries an
RTL paragraph base, explicit RTL and LTR runs, logical start and end indents,
and start justification. Unlike the five-document trend, these five pages have
a hard raw luminance SSIM gate of at least 0.95 on at least 80 percent of pages.
The reviewed evidence passes five of five pages with scores 0.956809869,
0.972241230, 0.997558968, 0.997294132, and 0.992810907 respectively. The TSV
records each page and the JSON records fixture identities, coverage, threshold,
and per-document page counts.

The Writer oracle registers the three fixed Noto files and one checked-in
oracle-only static Thin instance of the exact bundled Noto Sans SC subset. The
fixture inventory test binds the source and output SHA-256 values, requires the
adjacent OFL licence and provenance record, and rejects extra files. The static
instance changes neither product font bytes nor the deterministic Rust input.

## The M11 cross-viewer acceptance gate

The M11 gate uses one deterministic ten-slide deck built from the checked-in
default template by `build_f116_ten_slide_deck` in the existing `rpptx`
integration binary. No generated deck is checked in. The reviewed temporary
candidate is `/private/tmp/rdocx-f116-m11-write-api.pptx`, with SHA-256
`d36da6e8849eabd4487d2572baea19c3716ee7d0fe03aaa4714a28ce3c41de4f`.
Its ordinary `.pptx` and slideshow `.ppsx` forms reopen through the facade,
use the correct main content types, and return no validation issues.

The one deck covers the complete M11 write surface:

| Story | Candidate coverage |
|---|---|
| F-107 | ten slides synthesised from the bundled template |
| F-108 | clean structural validation before and after reopen |
| F-109 | position, size, rotation, name, fill, line, and adjustment mutation |
| F-110 | textbox, preset shape, three connector forms, and group construction |
| F-111 | package-deduplicated pictures with slide-scoped relationships |
| F-112 | paragraphs, bullet, run properties, and direct Latin font |
| F-113 | cells, fill, margins, banding, width, merge, and split |
| F-114 | image-bearing duplication, removal, and final-index move |
| F-115 | slide size, core properties, hidden state, background set and clear, and slideshow save |

Every viewer receives that exact SHA. Microsoft PowerPoint checks its pinned
version, Info.plist build, and AppleScript build before opening, counting ten
slides, and closing without saving. Keynote records a user-confirmed open and
ten-slide inspection against its installed version and bundle build.
LibreOffice runs a headless import and `impress_pdf_Export` with hidden slides
enabled, then `pdfinfo` must report ten pages. Google Slides imports through a
signed-in browser, reports ten slides without a conversion error, and exports
once. Its row records the acceptance date and browser build rather than an
application version. The ignored gate reruns the automatable PowerPoint and
LibreOffice checks and validates all four SHA-bound evidence rows. It does not
replace the Keynote or Google human-action evidence with unsupported UI
automation.

The modern package-class gate builds PPTM, POTX, POTM, PPSX, and PPSM packages
from one valid PresentationML graph. Every class reopens with its exact main
content type. Macro classes retain the relationship-owned VBA project and
signature bytes, ordinary classes retain opaque producer parts, and all
relationship scopes remain equal. A six-way conversion comparison proves the
main override is the only package difference. A signed conversion separately
proves evidence is retained and reported invalidated.

The evidence bound to the reviewed SHA is:

| Viewer | Version or date | Build | Result |
|---|---|---|---|
| Microsoft PowerPoint | 16.104 | Info.plist 16.104.25121423, AppleScript 1214 | clean, opened ten slides and closed without saving |
| Apple Keynote | 14.4 | 7043.0.93 | clean, user-confirmed human-action open of ten slides without a conversion error, then closed |
| Google Slides | accepted 2026-08-09 | Google Chrome 151.0.7922.76, build 7922.76 | clean, saved to Drive, showed slides 1 through 10 without a conversion error, and started one Microsoft PowerPoint download |
| LibreOffice Impress | 26.2.5.2 | cd7284b4cbbfeb507e630c1aac019f4157393acb | clean, headless import and ten-page PDF export |

All four rows record clean observations against the same artifact SHA. The
Keynote row is user-confirmed human-action evidence. The Google Slides row is
bound to the acceptance date and browser build without recording the private
import URL.

## The render fidelity gate

The 50-deck pinned corpus is rendered through bundled fonts at 150 dpi.
LibreOffice 26.2.5.2 with build
`cd7284b4cbbfeb507e630c1aac019f4157393acb` exports PDF through the
`impress_pdf_Export` filter with `ExportHiddenSlides=true`, then pdftoppm
26.01.0 rasterises every page at the same 150 dpi. The hidden-slide option is
part of the asserted command because a default PDF export omits five corpus
slides.

Clean Ubuntu 24.04 workspace jobs install that exact LibreOffice build from the
official Linux x86-64 Debian archive with reviewed SHA-256
`2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`.
The installer bounds the download, archive member count, and expanded bytes,
rejects unsafe members and populated prefixes, and checks the exact runtime
identity before the `oxml-chart` viewer gates execute. The installer supplies
the explicit NSS, NSPR, D-Bus, Cairo, GLib, X11, CUPS, font, and Kerberos
runtime libraries required by the official build.

The harness decodes both PNGs through the existing strict decoder and computes
global luminance SSIM after compositing RGBA over white. It uses population
variance and covariance with the standard 8-bit constants `K1=0.01`,
`K2=0.03`, and `L=255`. Dimensions must match exactly. Per-slide scores and
paths are written to TSV, while the summary reports coverage, minimum, median,
and maximum.

**Trend reference: at least 0.95 SSIM on at least 80 percent of slides. Hard
automatic gate: every slide renders without panic, missing output, dimension
mismatch, or a dropped bounded shape. Hard manual gate: the pinned native
PowerPoint representative review is recorded and accepted.**

The current complete 50-deck run covers 421 slides and passes both corpus
orchestration tests. Twenty-five slides reach at least 0.95 SSIM, or 5.938
percent. The minimum is -0.091350, the median is 0.512539, and the maximum is
1.0. `target_met` is false. This is recorded trend evidence, not a claim that
the advisory quality target passed.

The trend line is not a PowerPoint-conformance threshold. A calibration over
all 34 slides of the ecodesign representative uses Microsoft PowerPoint 16.104
and the same pdftoppm 26.01.0 raster path. Native PowerPoint against the pinned
LibreOffice oracle produces zero slides at or above 0.95 SSIM, with median
0.650406194 and maximum 0.940934972. Slide 25 reproduces the recorded native PNG
hash, which confirms that the calibration uses the accepted native pipeline.
An implementation can therefore agree more closely with PowerPoint and still
move away from the LibreOffice trend line.

LibreOffice is the oracle only because PowerPoint is not scriptable on CI
runners, and LibreOffice has its own rendering bugs. **SSIM regressions are
therefore review-required, not automatic failures.** Spot-check against real
PowerPoint output once per milestone. The CI comparison records whether the
trend reference was met but does not fail solely because it was missed. Exact
oracle versions, full corpus coverage, valid dimensions, successful rendering,
and zero dropped bounded shapes remain enforced.

CI retains `gate-evidence.json`, `render-manifest.tsv`, and
`ssim-results.tsv` as the Presentation fidelity evidence artifact. The image
trees stay job-local because the TSV identifies every deck, slide, score, and
paired path without uploading hundreds of redundant raster files.

Stand this harness up in M10 alongside the first text rendering, not afterwards.

The M10 native spot-check uses Microsoft PowerPoint 16.104, Info.plist build
16.104.25121423 and AppleScript build 1214. PowerPoint PDF exports are
rasterised by pdftoppm 26.01.0 at 150 dpi. The low representative is
`sample_pptx_grouping_issues.pptx` slide 1 at LibreOffice SSIM -0.177170506.
PowerPoint confirms the white background and complete grouped geometry, while
the Rust render has a wrong red background and missing or misplaced groups.
The median representative is
`at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx` slide 25 at
SSIM 0.172346895. PowerPoint confirms a full chart and product image which are
absent from the Rust render. The high representative is `crop-to-0.pptx` slide
2 at SSIM 1.0, an intentionally blank white slide matching at a glance.
LibreOffice follows PowerPoint for the substantive low and median content.

The temporary native PDF SHA-256 values are
`bd1511f546c970cddb9602f6b5421a3490e3ff22e5da74ca183e2e57b73a8f24`,
`99503f6dce0773c64da5b52e917d0d3f1f21aaddb0214532dda4c5131fdaa320`,
and `d5ce8e607f805914768d314ed6bb0f7f8fb762f9f62d680666e119f5c1afdf65`
for low, median, and high. Their 150 dpi PNG SHA-256 values are
`6ee02b21b8ee7ec1dd741ffd3a4b0bc2fe7a0d917c5b3c1d6c1b2aa69d7a088b`,
`85610d4b6778432355ab498f2a5da3bce6831cf502703d08caae70988307a49c`,
and `100875bd72e1c1ebe08263aac08bfb28dfd974a7f0f270ea98e0bbf9b9c7cbd2`.

Table rendering has an additional deterministic gate. A banded table with a
two-dimensional merge must produce the expected sampled fills, visible text,
merged bounds, and exactly one physical stroke per border segment. Separate
regressions prove that continuation cells emit no duplicate fill, border, or
text and that cell margins feed the shared fixed-box text path. Raster evidence
uses deterministic font mode.

The dense Word form golden is a readable OOXML document constructed in the
existing regression entrypoint. It combines recursive tables, exact and
minimum rows, a vertical merge, based-on and conditional table styles,
cell-relative foreground and page-behind anchors, outer and interior `nil`
borders, and a 7 point empty paragraph mark. It requires one Letter PDF page,
exact outer, merge, and nested-grid line bounds, an absent crossing merge edge,
all readable text, a zero-glyph mark carrier, and identical repeated PDF and
PNG renders. The 96 dpi raster is 816 by 1056 pixels and pins the complete RGBA
checksum plus the non-white, foreground-fill, and behind-fill pixel counts. No
binary fixture is committed.

## New tests the extracted crates need

These crates have never seen a non-docx package, so the existing tests do not
cover the cases that matter now.

**`oxml-opc`**
- `with_main_part("ppt/presentation.xml", ...)` then `main_document_part()`
  resolves, and the package round-trips.
- A pptx-shaped package: package rels to `presentation.xml`, slide rels to
  `slide1.xml`. Assert
  `resolve_rel_target("/ppt/slides/slide1.xml", "../slideLayouts/slideLayout1.xml")`
  resolves correctly. The `..` traversal is currently exercised only by docx
  headers.
- Every `rel_types` constant is unique and well-formed. Cheap, and the only
  thing that catches a copy-paste typo among the new constants.
- Zip-slip: a part named `../../etc/passwd` and an absolute-path entry are
  normalised or rejected. The code handles it, nothing tests it, and the crate
  is about to become a public shared component.

**`rdocx-oxml` OfficeMath**
- Prefix aliases and default namespaces resolve by expanded name, while
  unqualified attributes and conflicting fixed-prefix bindings stay untyped.
- Every supported expression and property sequence writes in schema order and
  reparses with opaque siblings in the same logical slots.
- Inline and display equations retain paragraph item order through mutation,
  raw-boundary collapse, save, and reopen.

**`rdocx` MathML and LaTeX conversion**
- Source-built unit cases cover every supported Presentation MathML element,
  LaTeX command family, normalization rule, limit, and ordered loss path.
- Round trips compare the normalized `MathArgument` tree rather than format
  bytes. Perturbations swap fraction operands, detach scripts, change delimiter
  scope, reorder matrix cells, and remove a diagnostic.
- The ignored live differential verifies exact Pandoc 3.10 identity and uses
  its bundled texmath engine in both directions. Pandoc `display` and
  `semantics` wrappers, its conversion of pre-scripts to an empty-base
  post-script followed by a run, its insertion of explicit n-ary `\limits`,
  and its removal of explicit delimiter scope are asserted as intentional
  divergences.

Agile encryption has a source-encoded Microsoft Word 16.104 oracle package so
the regression gate does not depend on an opaque binary fixture. The gate opens
that package only with its password, checks every supported AES data-key and
password-encryptor key-size pairing independently across every supported SHA
algorithm with deterministic synthetic packages, rejects malformed descriptors
and wrong passwords, and proves tampering fails before ZIP parsing. A
package-preservation round trip also checks that unrelated parts survive after
authenticated decryption.

Agile encryption writes have separate fixed-profile, round-trip, randomness,
and failure-atomicity coverage. Tests inspect descriptor child order and every
required CFB and DataSpaces stream, decrypt through the production reader, and
compare every part, relationship, content type, and unmodelled XML byte.
Injected deterministic random sources make secret separation and random-source
failure testable without replacing the operating system source in production.
The external gate opens one produced document in pinned Microsoft Word 16.104,
records correct-password success and wrong-password rejection, and treats both
outcomes as mandatory manual evidence.

The presentation security gate stays in the existing `rpptx` integration
binary and builds every package and certificate fixture from source. Encrypted
round trips check the correct and wrong passwords, bounded reads, unrelated
parts, relationships, content types, raw PresentationML subtrees, empty
password rejection, and atomic destination failure. Signature tests cover
complete trusted-certificate fixture coverage, untouched producer-shaped bytes,
failed signing without live-state mutation, and invalidation after nested
slide, shape, text, core-property, and package-graph mutations. Feature
isolation checks the ordinary facade and the Python, WASM, and CLI manifests.
The ignored external gate binds its generated artifact to SHA-256
`a0d33171c63ec084231daeef3b35718f5a2d709a5c92c9c0e2017ccaf9fa52d6`.
Pinned Microsoft PowerPoint 16.104 build 16.104.25121423 opened that artifact
with the correct password and rejected a wrong password. Both observations are
required, and no binary fixture enters the repository.

**`oxml-core`**
- New unit round-trips: `Centipoints::from_pt(18.0).0 == 1800`,
  `Angle::from_degrees(90.0).0 == 5_400_000`,
  `Percent1000::from_percent(75.0).to_fraction() == 0.75`.
- Existing `Length`, `Twips` and `Emu` constructors have positive and negative
  truncation-pinning tests that move with the units into `oxml-core`.
- `xml_text` becomes public API, so add CDATA, mixed content, nested elements,
  unknown entities, and the `GeneralRef` split case.
- `AppProperties` parses a Word `app.xml` **and** a PowerPoint one, leaving the
  other format's fields `None`, and omits them on write.

**`oxml-media`**
- Sniffing every format from magic bytes, and **sniff beats extension**: a
  `.png` that is really a JPEG resolves to JPEG.
- MP3, WAV, and ISO base media signature checks accept valid containers and
  reject invalid bytes for known MIME names, including uppercase type and
  subtype spelling. Safe MIME grammar remains strict, and unknown safe types
  remain opaque.
- DPI from PNG `pHYs` with unit 1 and unit 0, and from JPEG JFIF density units
  1 and 2, including a file with EXIF before the SOF.
- **A truncation loop per format**: `for n in 0..data.len()`, assert no panic.
  Cheap, and it catches every slice-index bug in one shot.
- The counter fix, named as a sentence:
  `next_image_name_uses_the_highest_existing_index_not_the_part_count`.

**`oxml-layout`**
- `Transform` composition order matches the PDF `cm` operator.
- `walk()` flattens nested groups and accumulates the transform correctly.
- `FontManager` with no fonts returns an error rather than panicking, and
  `--no-default-features` is in the CI matrix so the system-font-discovery-off
  path is exercised while bundled deterministic fonts remain available.
- Multilingual fixtures cover Arabic joining, Indic cluster integrity, Thai
  word boundaries, CJK prohibited punctuation, conditional hyphens, and mixed
  bidi lines. Every fixture compares logical source intervals separately from
  line-local visual order and requires equal glyph-array lengths.
- Word direction regressions type and round-trip `w:bidi` and `w:rtl` around
  retained unknown siblings. Structural layout checks cover paragraph base and
  exact run overrides, direction-relative start and end alignment and indents,
  leading-edge numbering markers, line-local L1 then L2 ordering, unchanged
  logical source spans, and PDF visual paint with logical `ActualText`.
- The DrawingML direction round trip rejects a foreign same-local-name
  attribute and preserves unknown attributes, children, and schema order.
- The rich PowerPoint fixture exercises the shared PDF, raster, and SVG paths
  with deterministic bundled fonts. PDF extraction and SVG text stay logical
  while painted bidi positions are visual. The stable source fixture and full
  49-entry hash harness protect the legacy Latin path.
- The native Word rich-layout fixture covers the same four scripts, exact
  language projection, valid clusters and offsets, resolvable Word source
  intervals, deterministic PDF and raster output, and searchable logical SVG
  text. An exact-line regression fixes every complex script to the Word 0.8em
  baseline while the Latin path and all 49 hashes remain byte-identical.

**`oxml-pdf`**
- Three-deep groups balance `q` and `Q`, emit each `cm` before child content,
  and apply the declared clip rule and shared opacity state before recursion.
- `Path` with solid fill only, solid stroke only, and both, produces `f`, `S`
  and `B`. The combined case also proves `q`/`Q` counts balance, which catches
  the classic unbalanced graphics-state bug.
- Repeated equal alpha values produce one ExtGState with matching `CA` and
  `ca`, while distinct values remain distinct and opaque content emits none.
- A 50 percent black fill over white produces the exact midpoint pixel in the
  deterministic raster path.
- Shared raster option tests construct deterministic in-code pages and decode
  PNG, JPEG and multi-page TIFF output. They prove selected-page order,
  dimensions, distinct page pixels, transparent PNG behavior, JPEG quality
  validation, TIFF cardinality and byte-identical opaque PNG compatibility
  wrappers.
- Linear and radial path gradients produce type 2 patterns, type 2 or type 3
  shadings, and type 3 stitching functions over interval type 2 functions.
  Structural tests also pin stop normalization, fill and stroke pattern
  operators, mixed solid paint, and page-local pattern resources.
- A 90 degree group rotation turns a linear gradient's sampled colour change
  vertical when rasterised at 72 dpi with the recorded Poppler 26.01.0.
- **`Group` containing `Text` finds the font.** The regression test for the
  recursion hazard.
- Tagged-PDF structure tests cover headings, nested lists, table headers and
  cells, figures with alternate text, artifacts, deterministic MCIDs, and the
  parent tree. A raster equality test compares the exact PNG bytes before and
  after adding `MarkedContent`.
- The Word-to-PDF regression renders all six heading levels, three real list
  depths, and a table whose two header cells repeat across pages. It follows
  each `TH` to its paragraph child, checks parent-tree ownership on every page,
  and requires one MCR for every emitted semantic MCID. Source-compatibility
  tests construct the unchanged image and group variants directly.
- The external accessibility oracle is veraPDF 1.30.2 with profile `ua1`. Its
  source installer and signature are pinned outside the repository. The
  ignored differential test requires that exact version and a conforming
  report for an in-code deterministic fixture before feature completion.
- The archival regression gate renders one tagged in-code fixture with an
  actually embedded bundled-font subset through PDF/A-2b and PDF/A-3b. The
  pinned veraPDF 1.30.2 oracle must pass profiles `2b` and `ua1` on the first
  file, then `3b` and `ua1` on the second. Focused tests assert matching XMP,
  output intent, ICC linkage, deterministic identifiers, named preflight
  errors, retained headings, lists, tables, and alternate text. The ordinary
  path has a byte pin and the complete 49-entry hash harness remains unchanged.
- `Group` containing `Image` registers the XObject.
- `Group` containing `LinkAnnotation` emits it with a transformed rectangle.
- A preceding leaf proves nested XObject registration and recursive emission
  use the same depth-first ordinal.
- Raster: a rotated rectangle at 72 dpi has a filled interior pixel and an empty
  corner, and phase-zero line and path dashes have exact painted runs and gaps.
  Nested group samples pin transform order, clip intersection, and subtree
  opacity. Fill-rule, linear and radial gradient, gradient-domain, and page
  background samples pin the remaining paint translations. These are
  deterministic unit tests with no golden files.

## Binding tests

MHTML remains native Rust only. The existing Python, WASM, and CLI surface
inventories therefore assert no new method, error, dependency, or feature. The
exhaustive Python error adapter maps native MHTML and invalid embedded-mutation
failures to the established generic `RdocxError` class. The
published-crate riders compile both WASM graphs, deny rustdoc warnings, verify
the patched workspace package graph, and enforce the 10 MiB archive ceiling.

The parity suites are worth more than any number of Rust-side assertions,
because the whole value proposition is compatibility:

- The rdocx gate asserts exact `python-docx==1.2.0`, then executes the explicit
  seventeen-example S33 documentation manifest from stable v1.2.0 tagged
  sources. Sixteen bodies change only the import namespace. The exact
  Quickstart held-row body uses one declared public row re-fetch before its
  second cell assignment to respect strict global revision invalidation. Each
  manifest entry pins its source URL, heading, exact source statements,
  transformation and normalized structural assertion. The two-way
  differential authors the same paragraphs, runs, direct formatting, tables
  and cells with each writer, reads both files through both libraries, and
  directly compares normalized public records including distinct relative and
  absolute line spacing, units, enums, and saved table style.
- The same for `rpptx` and `python-pptx`.

The rpptx binding gate executes the seven python-pptx 1.0.2 Getting Started
workflows with the import namespace changed from `pptx` to `rpptx` and the
minimal public re-fetches required after structural writes. Its differential
rider asserts the exact oracle version, compares each writer through both
readers, and directly compares the normalized rpptx-authored and
python-pptx-authored records. It never compares package bytes and the oracle is
not a runtime dependency.

Both libraries are test-only CI dependencies. Neither oracle is a runtime or
published-crate dependency, and neither differential compares package bytes or
commits binary fixtures.

Each package has a strict typing smoke program that consumes its installed
public surface. Fresh cp39-abi3 wheels must contain the native-extension stub
and `py.typed` marker, pass exact `mypy==2.3.0 --strict`, and pass `stubtest`
against both installed packages. Strict mypy also checks every inline-typed
pure-Python source in each installed wheel. Representative enum-input,
return-type, inline-source, constructor, and member mutations must make those
gates fail, so hand-written stubs cannot drift.

The document WASM wrapper has a package-preservation Node gate and a PDF gate
in its single defaults-off profile. The PDF gate calls generated `toPdf`
through reflection and requires `%PDF-` through `%%EOF`, a Type 0 font, a
`FontFile2` stream, and the bundled Carlito base font. This proves the public
JavaScript name, complete output, and embedded fallback font at the generated
boundary.

The presentation WASM wrapper has one Node round-trip gate in its default
profile and a second Node gate with `render` enabled. The first crosses the
generated JavaScript `Uint8Array` boundary and proves that facade-owned slide
mutation preserves the complete package. The second produces a complete PDF.
The final normal-default artifact is built with exact wasm-pack 0.15.0,
optimized with reviewed wasm-opt 125, compressed with `gzip -n -9`, and
rejected at 1,000,000 decimal bytes. The wrapper manifest keeps render out of
defaults while its facade dependency selects the bundled template explicitly.
A padded artifact or render-enabled default must make the exact named size gate
fail.

The `rpptx` CLI integration gate corrupts a relationship and requires
`validate` to exit nonzero. It then requires all 50 manifest decks to validate
with a zero exit and never skips a missing corpus. The primary workspace-test
job and the MSRV job fetch and verify both pinned corpora before running Cargo
tests. Both jobs install exact uv 0.10.2 through the reviewed official setup
action, isolate its cache under the runner temporary directory, and give Rust
test threads an explicit 8 MiB stack for the largest corpus round trip. Command
regressions also prove bounded DPI, bounded diff work, zero-slide PNG failure
without output, and one-slide-at-a-time PNG conversion.
The thumbnail and outline gate requires an exactly 320-pixel-wide proportional
slide-one PNG and recursive paragraph output with stable level indentation.
Regressions cover nonstandard aspect ratios, shared output defaulting, grouped
text order, embedded paragraph-break normalization, and field-only title
identity so the title appears exactly once.

The `rdocx` CLI has one integration binary that invokes the compiled executable
through `CARGO_BIN_EXE_rdocx`. Its tests cover `inspect`, `text`, `convert`,
`diff`, `replace`, `validate`, and `render` with in-code DOCX and
corrupt-package fixtures. The assertions bind schema 1, default paths, exact
stdout, exit-status verdicts, output validity, replacement persistence,
document-order text, bundled-font deterministic render bytes, legacy
zero-based `render --page`, one-based `render --pages`, shared image format
extensions, invalid range rejection and no partial output. Process ID and an
atomic counter isolate temporary workspaces across concurrent runs.

All 27 workspace packages explicitly declare one distinct README. The root
README is the high-level `rdocx` guide. Each crate-local document states the
package purpose, direct-use guidance, adjacent package relationship,
publication status, and a concrete Rust, CLI, Python, or JavaScript example.
The compatibility shims direct users to their shared replacements.
Internal binding and WASM crates state that they are not crates.io packages.

`scripts/readme_doctests.py` validates the exact package-to-README inventory,
the documented CLI argument names, Python and JavaScript surface names,
deterministic feature guidance, and matching dependency and import names. It
builds the applicable libraries with locked dependencies and Cargo JSON
messages, locates each emitted rlib from one package build graph, and invokes
rustdoc with the 2024 edition, warnings denied, the dependency search path, and
every matching `--extern` binding. It compiles 27 Rust examples across the 21
Rust-library READMEs. It also creates all 22 publishable archives and
byte-compares their single packaged README with the declared source. Archive
creation uses the same exact 22-package local source patch set as the release
dry run, so a reviewed version can be checked before its internal dependencies
exist on crates.io. The patches never enter an archive and upload nothing. The
docs job and canonical non-fast verification call this same runner.
The stable 0.13.0 carrier regression pins all eleven inherited version
carriers, both Python project versions, both rdocx WASM dependency assertions,
the stable CI package literal, the seven publishable crates, and every stable
README requirement. It also proves the current incubating workspace carriers
are 0.10.0 while `rpptx-wasm` remains ineligible for publication.
The paired incubating regression pins all sixteen explicit manifests, fifteen
workspace dependency requirements, sixteen lockfile entries, publication
flags, README examples, Rust assertions, the CI WASM literal, and the exact
15-package publication preflight at 0.10.0. It separately proves the stable
workspace remains at its prepared 0.13.0 boundary and `rpptx-wasm` remains
ineligible for publication.
The current stable shared-family gate packages and verifies
`rdocx-layout@0.13.0`, requires its normalized archive dependency on
`oxml-layout@0.10.0` to contain no local path, and compiles the packaged crate
against the exact shared registry version without an `oxml-layout` patch.
That registry consumer is excluded from the incubating tag preflight because
0.10.0 does not exist before its own publication. The stable tag preflight runs
it only with explicit published-shared authority. The earlier F-X068
post-publication proof against 0.8.0 remains immutable release evidence.
A separate recovery gate constructs an isolated registry consumer of exact
`rdocx-layout@0.10.1` and inspects its unpatched normal dependency tree. It
requires registry `oxml-layout@0.6.0` and rejects 0.7.0, so the immutable
published proof remains independent of current workspace pins.
The 0.6.0 release gate verified every selected registry entry and owner, the
annotated tag target, byte-identical GitHub release notes, and selected record
notifications at reviewed SHA
`55fb2f54caf91d7dedc8936b4c7b116354590628`. The failed stable 0.10.0
attempt is not a passing release gate because only two packages published and
no GitHub release was created.
The 0.10.1 release gate verified all seven selected registry entries under sole
owner `mantissaman (Atul Sharma)`, the annotated tag at reviewed SHA
`ae0dcb162a7805e59e5890464b226765645ad547`, byte-identical GitHub release
notes, nine contribution notifications, and six authorized unmerged
pull-request closures.
The 0.7.0 release gate verified all 15 incubating registry entries under sole
owner `mantissaman (Atul Sharma)`, the annotated `rpptx-v0.7.0` tag at
reviewed SHA `1b076c16fb494fe47b054d761e061181a1ea0b15`, the stable-family
exclusion, byte-identical GitHub release notes, and the absence of
`rpptx-wasm@0.7.0` from crates.io. Its selected contribution inventory is
empty, so it requires no external notification.
The 0.8.0 release gate verified all 15 incubating registry entries under sole
owner `mantissaman (Atul Sharma)`, the annotated `rpptx-v0.8.0` tag at
reviewed SHA `7f4414b0aeef1ec2cbae75fcb5aa96ab6dee6d70`, stable-family exclusion,
byte-identical GitHub release notes, the published stable shared-family graph,
and the absence of `rpptx-wasm@0.8.0` from crates.io. Its selected contribution
inventory is empty, so it requires no external notification.
The 0.9.0 release gate verified all 15 incubating registry entries under sole
owner `mantissaman (Atul Sharma)`, immutable annotated tag `rpptx-v0.9.0` at
reviewed SHA `45b4f277ff5fd6d1b032e929c5dcee7fb9d2c550`, byte-identical GitHub
release notes, selected-family exclusion, and absent `rpptx-wasm@0.9.0`. Its
selected-family inventory is empty, so it requires no notification.
The 0.10.0 release gate verified all 15 incubating registry entries under sole
owner `mantissaman (Atul Sharma)`, immutable annotated tag
`rpptx-v0.10.0` at reviewed SHA
`1e409c553b950eb8029e3e78e39ff775f18ba3ab`, byte-identical GitHub release
notes, stable-family exclusion, and absent `rpptx-wasm@0.10.0`. Its selected
diff contains no external issue or pull request, so its reviewed contribution
inventory is empty and no notification is required.
The stable 0.13.0 preparation gate pins the exact seven-package family, the
published shared 0.10.0 source dependency boundary, all binding exclusions,
and the reviewed empty selected-family contribution inventory. Repository
history and GitHub records since `v0.12.0` contain no external issue or pull
request that implements the selected M22 changes. It requires full
verification and clean review at one SHA before a separate release approval.
The failed stable 0.11.0 release gate is not a passing family gate. Its
annotated tag targets reviewed SHA
`25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4`, and its preparation, full
verification, notes, and archive preflights passed. Publication then stopped
after `rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0` because packaged
`rdocx-layout@0.11.0` could not compile against registry
`oxml-layout@0.7.0`. The other five stable packages, GitHub release, and six
notifications are absent. The shared recovery gate proves all 15 shared 0.8.0
entries. The stable 0.11.1 recovery gate verified all seven selected registry
entries under sole owner `mantissaman (Atul Sharma)`, the annotated tag at
reviewed SHA `5a850ce9ae6c31f8365594ed2970193266f8b2a6`, byte-identical GitHub release
notes, the published `oxml-layout@0.8.0` dependency, and all six leave-open
notifications.
A separate cleanup gate proved all seven 0.11.1 entries live and unyanked under
the authenticated owner, exactly `rdocx-opc@0.11.0` and
`rdocx-oxml@0.11.0` present, the other five 0.11.0 entries absent, the
immutable v0.11.0 tag target unchanged, and no v0.11.0 GitHub release. After
another final approval, it yanked only those two incomplete entries and read
back their yanked flags. A regression pins that exact allowlist and forbids
tag, release, notification, closure, and other-version mutations. Complete
coherent releases remain live.

The large-document regression source-builds 1,000 one-page paragraphs and
measures deterministic pagination separately from direct PDF rendering. Its
test-binary allocator is inactive outside an explicit measurement generation.
The release gate requires exactly 1,000 pages, nonempty PDF output, layout at
or below 64 MiB and at or above 250 pages per second, and PDF rendering at or
below 16 MiB additional peak and at or above 1,000 pages per second. Workflow
mutation tests reject a missing, unlocked, debug, non-ignored, non-exact,
parallel, or failure-swallowing invocation.

## What CI runs

| Job | Command |
|---|---|
| changes | On pushes and pull requests, classify changed paths for the nine filtered jobs with `dorny/paths-filter` v4.0.3 pinned to reviewed commit `ceb8a2b8f2d89434be7ff52d3de7ec3738c5cc9d` |
| test | Install exact uv 0.10.2, Poppler 26.01.0, LibreOffice 26.2.5.2, and Pandoc 3.10, fetch both pinned corpora, run the pinned Pandoc texmath differential and exact locked release-mode 1,000-page performance regression with one test thread, run `cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py` with an isolated uv cache and 8 MiB Rust test-thread stack, run the exact locked deterministic animation golden, then run `python3 scripts/golden_png_harness.py --check` |
| no-default-features | `cargo test -p oxml-layout --no-default-features` |
| wasm | Locked `wasm32-unknown-unknown` checks, `wasm-pack test --node`, and local bundler pack and fresh-install gates for `rdocx-wasm` and `rpptx-wasm` |
| prose | `python3 scripts/prose_check.py` and `python3 scripts/sync_agent_skills.py --check` |
| release-regressions | Install cargo-release 1.1.3 with its locked dependency graph, then run `python3 -m unittest scripts.test_sprint_workflow` |
| hash-harness | `python3 scripts/hash_harness.py --check` |
| presentation-fidelity | Fetch the pinned corpus, run the exact locked deterministic animation golden, then run `python3 scripts/pptx_ssim_harness.py --check` on the pinned macOS render stack |
| word-fidelity | Restore the pinned Rust cache, run `cargo fetch --locked`, fetch the pinned Word corpus, then run `python3 scripts/docx_ssim_harness.py --check` on pinned Ubuntu 24.04 LibreOffice and Poppler with its locked offline helper |
| clippy | `cargo clippy --workspace --all-targets --all-features --exclude rdocx-py --exclude rpptx-py -- -D warnings` |
| fmt | `cargo fmt --all -- --check` |
| doc | `cargo doc --workspace --no-deps --all-features --exclude rdocx-py --exclude rpptx-py` with `RUSTDOCFLAGS=-D warnings`, then `python3 scripts/readme_doctests.py` |
| package-oxml-layout | Verify the exact 24-font and six licence-and-notice-file inventory, then build and size-check the verified archive |
| msrv | Install exact uv 0.10.2, fetch both pinned corpora, then run `cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py` under Rust 1.93 with an isolated uv cache and 8 MiB Rust test-thread stack |
| python-bindings | On pull requests, build each Python package with `maturin develop --locked` in its own Python 3.12.9 environment, then run its complete pytest directory |
| supply-chain | `cargo-deny check` |
| ci-gate | Always validate that every selected filtered job succeeded and every unselected filtered job was skipped |
| python-wheels | On manual dispatch or a `py-v*` tag, build six cp39-abi3 wheels for each Python package and one source distribution per package, then install and test every compatible artifact in a fresh environment |

MHTML uses the existing test, clippy, fmt, doc, wasm, hash-harness, and package
routes. Its Microsoft Word differential remains an explicit ignored local
oracle because that exact Word build is not available on the Ubuntu CI runners.

The checksum-pinned Pandoc 3.10 installer admits the authenticated
162,406,703-byte archive under an exact 160 MiB extracted-size ceiling. It skips
without materializing only the archive's two exact in-root executable aliases,
`pandoc-lua -> pandoc` and `pandoc-server -> pandoc`, while every other
symlink, hardlink, device, FIFO, and unsupported member type remains rejected.

The `changes` job routes `test`, `msrv`, `wasm`, `python-bindings`,
`presentation-fidelity`, `word-fidelity`, `hash-harness`, `supply-chain`, and
`prose` through inline fail-safe path filters. Every filter selects `ci.yml`,
so a routing edit cannot suppress its own gate. Product and toolchain paths
include each job's transitive workspace inputs. A documentation-only HLD
change selects `prose` and skips the filtered product jobs. The supply-chain
job also runs on the weekly schedule without change detection.

`ci-gate` has `if: always()` and depends on the detector plus every filtered
job. It accepts only `success` for a selected job and only `skipped` for an
unselected job. Failure, cancellation, an unexpected skip, or a failed change
detector makes the aggregate gate fail. On the scheduled route, it requires
the detector to be skipped and the supply-chain job to succeed. The stable
aggregate check exists in the tracked workflow. Active repository ruleset
`21823007` protects the default branch with exact required status `CI gate`.
The check does not require a current-base SHA and applies when the ref is
created. The effective `main` rules contain only that required check. The
ruleset has exactly one bypass actor, repository role `admin` with numeric
actor ID 5 in `always` mode. This permits the reviewed direct sprint-close
push while ordinary pull requests remain subject to the aggregate gate.

The protection proof is bound to reviewed and verified S58 SHA
`31c51f04f1a9e7c6a198ef16eebba0d782a5827a`. Docs-only PR
[59](https://github.com/tensorbee/rdocx/pull/59) at
`aee0808a37a3afcc46c6ca236df096198c9601e4` reached clean mergeable state.
Hosted run `33275852961` reported successful Detect changes job `99162308288`,
Prose job `99162325899`, and CI gate job `99162339881`. Test, MSRV, WASM,
Python bindings, Presentation fidelity, Word fidelity, Output stability, and
Supply chain were skipped as unselected. Deliberately failing PR
[60](https://github.com/tensorbee/rdocx/pull/60) at
`ee1c0ae09d676498a594a77601e36240d0199a2b` produced failed hosted run
`33276064981`. Detect changes job `99162895790` succeeded, selected Prose job
`99162911436` failed, and CI gate job `99162924862` failed. The pull request
reported `mergeStateStatus=BLOCKED` and `viewerCanMergeAsAdmin=true`. Both
proof pull requests are closed and unmerged. Their remote refs were verified
at the named heads before deletion and are now absent. Their disposable
worktrees and local branches were removed cleanly.

The Word fidelity job has one explicit Cargo network boundary. Its exact
`cargo fetch --locked` step follows the pinned Rust cache and precedes corpus
and harness work. The later acceptor build remains `--locked --offline`.
Workflow regressions require that order and cardinality and reject a missing,
unlocked, duplicated, post-harness, or wrong-job fetch. The accepted
contribution evidence is PR 58 at source SHA
`c8fed1d1268fd765d602bac2da6524900c1c1cfd`, hosted run `33025657609`, Word
job `98366252284`. That job uploaded both required evidence files in one
nonempty 1,420-byte artifact. The integrated hosted Word job remains a separate
sprint-completion rider.

The `--exclude` pair on every all-feature command is required, not cosmetic:
`pyo3/extension-module` tells the linker that Python symbols come from the host
interpreter, which is false for a test binary, and on Linux this is an
unresolved-symbol link failure that is easy to misdiagnose.

The dedicated release regression job runs the complete standard-library test
module after checkout. It is unconditional and failure-propagating, so stale
stable or incubating version carriers fail on pull requests before a release
tag can reach the publication workflow. The same module holds the reviewed
release-notes parser, command, publication-order, exact-body, and generated
skill contracts.

Every Poppler-dependent CI job builds the reviewed 26.01.0 command-line oracle
from the official source archive. `scripts/install_pinned_poppler.py` enforces
the exact source SHA-256, an 8 MiB download ceiling, streaming extraction with
2,048-member and 64 MiB expanded-size ceilings, safe member paths and types,
and exact runtime identities for `pdftoppm`, `pdfinfo`, and `pdftotext`. A
successful run always starts with an empty prefix and rebuilds the reviewed
source. Test, MSRV, both Python binding rows, and Presentation fidelity invoke
the same unconditional failure-propagating installer before use. Platform
package managers provide build dependencies only, never a moving Poppler
binary package.

The wheel workflow runs the installed `rdocx` suites except the
Poppler-versioned rendering gate, which belongs to its pinned render job. It
runs the installed `rpptx` documented-example and differential suite. Native
cells also check the inline Python sources with exact `mypy==2.3.0 --strict`
and run `stubtest` across every public and native-extension module. The
musllinux cells install into clean Python 3.9 Alpine environments and run the
same package parity suites. Repository unit tests
parse the exact two-package, six-target product and use negative mutations to
prove that package, target, clean-install, parity, artifact dependency, and
tag-only OIDC requirements are sensitive before the hosted matrix runs.

The pull-request binding job has one matrix row for `rdocx` and one for
`rpptx`. It uses Python 3.12.9 with exact `maturin==1.13.3` and
`pytest==9.1.1`, installs `python-docx==1.2.0` or `python-pptx==1.0.2` for the
applicable row, and installs the Poppler toolchain required by the full rdocx
rendering suite. Each row creates a fresh environment, builds the extension,
then runs every test in that package's binding test directory. The build and
pytest commands are separate ordinary steps with no successful fallback or
`continue-on-error`, so either failure makes the pull-request check fail.
The operative top-level `pull_request` trigger schedules change detection, and
the binding job runs only when its complete input closure is selected. Neither
its build nor pytest step has an environment or condition that can suppress
execution after selection. Root permissions are exactly `contents: read`.
Only the change detector adds `pull-requests: read`, which is required to list
changed pull-request files. No job grants `id-token: write`. Checkout v6.0.2,
setup-python v6.2.0, rust-cache v2.9.1, and the selected stable rust-toolchain
revision are bound to full reviewed commit SHAs. Their operative input maps are
exact and cannot be satisfied by comments.

The pull-request WASM job uses exact Node 24.11.1 and wasm-pack 0.15.0. It
installs the official Binaryen version 125 Linux archive only after verifying
its pinned SHA-256, places that optimizer on `PATH`, and requires the exact
official identity `wasm-opt version 125 (version_125)`. It target-checks both
WASM packages with `--locked`, then runs both inline suites through
`wasm-pack test --node`.

Both manifests bind release optimization to `-Oz`,
`--enable-bulk-memory`, and `--enable-nontrapping-float-to-int`. The last flag
is required by nontrapping conversion operations emitted by the Rust 1.93
standard library. CI builds the exact `@tensorbee/rdocx-wasm` and
`@tensorbee/rpptx-wasm` release bundler packages with locked dependencies. Each
package is packed locally, installed into a separate fresh consumer through an
isolated npm cache with scripts disabled, and checked for its exact name,
version, WASM, JavaScript glue, public declaration, and import. The steps are
unconditional and propagate ordinary non-zero command status. Structured
regressions reject optimizer, checksum, package, target, scope, locking,
installation, authentication, publication, and tag mutations.

The job retains root `contents: read` permission and has no npm publication,
registry authentication, token, OIDC, release, or tag authority. Checkout
v6.0.2, setup-node v6.5.0, rust-cache v2.9.1, and the selected stable
rust-toolchain revision are bound to full reviewed commit SHAs.

## Gaps being closed

Stated plainly, because they are why two shipped defects went unnoticed:

- **Command-level output contracts need explicit coverage.** The published
  `rdocx-cli` surface has one compiled-binary integration test for each of its
  seven commands.
- **PDF and PNG output is only checked for non-emptiness**, so layout
  regressions are invisible. The hash harness closes this.
