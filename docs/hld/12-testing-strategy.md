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

The complete run-property and inline story compares against the recorded
WordprocessingML for `EG_RPrBase` rather than a fresh Word save, because Word
GUI capture is not available on the development machine. The no-repair
confirmation is a tracked human action under the Milestone 24 end-of-milestone
gate, and no automated test skips without it. One disagreement is recorded as
deliberate under rule 5 of `.claude/skills/differential-testing.md`. Word
resolves the theme attribute when a producer `w:rFonts` presents both an
explicit and a theme font for one slot, and rdocx resolves the explicit name.
The facade clears one form when the other is set, so the divergence is
reachable only on a producer document the caller never edited, where leaving
the bytes as written is what the no-op save contract requires.
`explicit_font_priority_divergence_from_word_is_deliberate` asserts our side so
a later change cannot drop the decision silently.

The table style and conditional formatting story compares against a pinned
`w:tblStylePr` reference tree built as source XML, for the same reason. Word
GUI capture is not available on the development machine, so confirming that
Word reopens the authored package without offering to repair it is a tracked
human action under the Milestone 24 end-of-milestone gate rather than an
automated gate. The raster half of
`every_conditional_table_region_matches_word` converts the same package with
the pinned LibreOffice 26.2.5.2 build in an isolated profile and rasterizes
both sides with the pinned pdftoppm 26.01.0 at 150 dpi in deterministic font
mode. One disagreement is recorded as deliberate under rule 5 of
`.claude/skills/differential-testing.md`. ECMA-376 orders the conditional
regions with the vertical bands before the horizontal bands and a later region
overriding an earlier one, so the horizontal band outranks the vertical band.
Word resolves it that way and this workspace follows Word. LibreOffice
26.2.5.2 inverts it and paints the vertical band. The test asserts both sides
of the divergence, so neither our resolution nor a future oracle change can
move without failing.

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
in the Ubuntu workspace job and the Ubuntu presentation-fidelity job. Both use
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

The numbering package-model gate enumerates every standard number format and
round-trips every public level property through `Document`. It separately
checks typed start and replacement-level overrides, `CT_Lvl` and `CT_NumLvl`
schema order, clean public-authored diagnostics, whole-graph atomic rejection,
style-link inspection, and byte-exact producer-extension retention across a
modeled update. The complete `rdocx-oxml` and `rdocx` suites, public package dry
runs, archive ceiling, and unchanged 49-entry hash harness complete the gate.

The style-linked numbering differential uses a source-built document and pins
Microsoft Word 16.112.3 build 16.112.26083020 on macOS 26.6.2 build 25G83 with
the `en-GB` locale. Twenty-seven exact Word records cover three levels across
body paragraphs and table cells, concrete-instance continuation, a new
instance, section restart, and REF `\n`, `\r`, `\w`, `\t`, and `\p` output.
The REF records include same-parent and different-parent relative context.
Three normalized TOC records cover numbered headings in the body and a table.
Focused tests separately cover replacement-level restart, ordinary level
restart, legal numbering, and `numId` zero suppression. A mixed Unicode
level-text probe uses `章节 %1 - Part %2)` and requires Word and rdocx to produce
`1-1)` for the text-suppressed full and level-aware forms.

The ignored render oracle exports the source-built document to PDF from Word
and rasterizes it with `pdftoppm` 26.01.0 at 150 DPI. rdocx renders its pages
directly at the same resolution in deterministic font mode. The two-page gate
requires at most one pixel of raster dimension difference, at least 0.95 paired
ink coverage, at most 0.08 normalized ink bounding-edge delta, at most 0.27
total variation across a 32-region ink distribution, and at most 0.04 row or
column projection distance on every page. A synthetic five percent row or
column shift must exceed the projection threshold. The recorded minimum
coverage is 0.983131. The recorded maximum edge delta is 0.062144, the maximum
distribution delta is 0.257661, and the maximum projection distance is
0.038314. The recorded minimum synthetic-shift distance is 0.049960. Exact
object-model records remain the text authority because global SSIM is unstable
for sparse text pages. The public differential uses separate abstract
definitions for independent instances. A focused regression records the
intentional contract that distinct concrete `numId` values start independent
counters even when they share one abstract definition, while this pinned Word
build continues the shared definition across those instances.

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
The entry-style and geometry regression supplies localized style ids, an A4
section with 1417 twip side margins, a style-owned 7777 twip right tab, one
missing TOC style, and a numbered heading with a tab suffix. It proves
localized id reuse, canonical creation only for the missing level, a 9072 twip
section fallback, structural tabs, safe default geometry under extreme parsed
values, and exact preservation of an unrelated unmodelled style child. The
normalized differential records are pinned to Microsoft Word 16.113 build
16.113.26091433 on macOS with the `fx114-toc-entry-records-v1` contract.
The producer-variant regression
`toc_rebuild_accepts_trailing_style_separator_and_duplicate_style_ids`
combines a final empty custom-style component with a duplicate `Normal` style.
It requires two rebuilt entries, one ordered diagnostic, both style definitions
after save and reopen, and unchanged strict failure from the public style-graph
validator. Unit controls retain stored display for interior empty components.
The Word-default instruction case retains argument-free `TOC \\z`, rebuilds
through an existing content-control payload, and saves and reopens the result.
A mixed simple and unsupported complex TOC case asserts exact diagnostic text
in physical source order, the derived native count, the immutable Python tuple,
strict typing, and runtime stub agreement.
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
The content-control type-payload gate covers Word and extension namespace type
children with producer attributes and nested children. It proves unchanged
payload retention under the fixed output prefix, canonical replacement after a
typed discriminator change, and stable ordering of the other property slots.

Word identifier allocation has a source-built regression gate. Two documents
created through different request orders must produce identical complete DOCX
bytes and the declared relationship, bookmark, comment, drawing, numbering,
part, and content-type identities in final recursive document order. Repeated
save and reopen must not advance allocation. Collision cases cover per-owner
relationships, bookmarks, comments, `wp:docPr`, abstract numbering, numbering
instances, content-type keys, and normalized ZIP part names. Namespace aliases
are accepted only through their expanded names, unrelated categories remain
independent, overflow is atomic, and imported raw owners retain exact bytes.
Content-type cases include `png` against `PNG` and case variants of a full part
name. They require linear case-insensitive duplicate detection, deterministic
lookup for directly mutated invalid public maps, serialization rejection of
those conflicts, and exact unchanged producer bytes and order under default and
all-feature builds. ZIP entries, package parts, and resolved relationship
targets reserve case-insensitive part identities while retaining producer
spelling. Mutation-history cases include building-block replacement. Header and
footer cases reject cross-type or unrelated header-shaped relationship targets
for text, raw XML, image, and background-image setter families.
Cross-part drawing coverage opens a package whose body and header reuse one
normalized producer `wp:docPr` identity, preserves both drawing payloads,
allocates a later authored drawing outside the package-wide occupied union, and
reopens repeatedly. A same-part character-reference alias remains a duplicate,
and a foreign same-local-name element remains outside the drawing scope.
Current-graph relationship cases add an unreferenced theme edge after chart
authoring and require chart `rId1` followed by theme `rId2`. A producer theme
captured on package open keeps its original id, while unknown internal and
external authored edges retain their targets and modes under deterministic
ordering. The SHA-bound F-159 Word candidate remains unchanged.
Opaque relationship coverage places a late direct relationship behind a raw
body attribute and allocates a modeled relationship in both opposite orders.
The raw bytes and raw relationship id remain fixed while the modeled edge
canonicalizes identically. OPC coverage also exercises mixed-case special ZIP
parts and relationship owners, case-equivalent get, set, contains, and remove
operations, Word and PowerPoint main-part reopen, signature discovery and
coverage, duplicate relationship ids, direct part-map conflicts, and a sentinel
destination that remains unchanged after pre-serialization validation fails.

The story-scoped relationship gate authors equal picture and hyperlink content
in body, header, footer, footnote, and text-box owners, saves, reopens, and
resolves every reference only through its exact `StoryId` relationship scope.
Atomic negatives cover stale and wrong owners, missing identifiers, wrong
types, external images, internal hyperlinks, and missing targets. Main-part
body, text-box, and table-cell lifecycle matrices remove and clone the authored
paragraphs themselves. Removal leaves zero live XML references while retaining
the relationship definition for the owned fragment. Clones serialize two
references to one relationship, and picture clones receive distinct global
`wp:docPr` identities. Namespace sensitivities cover aliases, character
references, producer-shadowed `r` and `wp` bindings, and inherited raw
relationship attributes in standalone footnotes and endnotes. Semantic reorder
coverage extracts references from the exact body and text-box paragraphs and
resolves them to their original media bytes and URLs.

The picture-replacement binding gate keeps body, header, and footer drawing XML
byte-identical while replacing PNG and JPEG targets through native Rust and an
installed Python wheel. It proves relationship-local copy-on-write for a shared
media target, deterministic target extensions, exact content types, and orphan
cleanup after an unshared format change. Missing, external, wrong-type, and
unsupported-byte requests compare the complete saved document before and after
the rejected operation. Save and reopen must retain every replacement and the
same relationship identifier in its original owner scope.

The run-splitting binding gate proves the native and Python direct-body entry
points with ASCII and multibyte text, endpoint no-ops, copied formatting, an
exact comment range, invalid-coordinate rollback, and save-reopen behavior.
OXML regressions split hyperlink-owned and mixed runs containing tabs, breaks,
field markup, alternate-content drawings, symbols, and unknown raw children,
then compare their exact serialized order and layout-projection ownership.

`comments_keep_standard_content_type_ids_dates_and_threads` creates comments
out of document order, replies to the first allocated identity, resolves its
thread, and saves and reopens. It requires the standard comments-extended
content type, stable comment and parent ids, exact optional RFC 3339 dates, the
no-date deterministic default, invalid-date rollback, and retained unrelated
sidecar XML. Installed Python tests, strict mypy, and stubtest cover the same
optional date keywords and frozen snapshots.

The cross-document fragment gate selects a main-body range containing custom
styles, direct and style-carried numbering, a bookmark and REF field, a
picture, an editable chart and workbook, an exact foreign subtree, and a
resolved comment thread with a relationship-bearing payload. Two imports must
reopen with all references resolved, exact selected raw XML retained twice,
unused source dependencies absent, and collision-free package identities.
Equivalent-reuse and rename policies must produce deterministic style,
numbering, media, chart, and workbook results. Atomic negatives cover external
relationships, split ownership, dangling targets, malformed relationship XML,
and relationship-id exhaustion, with destination bytes unchanged.

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

The rich-fragment gate
`rich_html_fragments_match_word_in_every_supported_container` inserts the same
source-built CSS, nested-list, table, link, explicit-image, and data-URI subset
into body, cell, header, and footer owners. Every result serializes and reopens
before the test compares content, refreshed story identity, direct range,
owner-local image and hyperlink relationships, and exact ordered diagnostics.
A companion boundary regression inserts before a cell paragraph and proves an
adjacent opaque child remains byte-preserved and in order. Malformed image
bytes are rejected against a byte-identical live-document snapshot. The gate
adds no sample, so all 49 hash entries remain unchanged.

The ignored render regeneration authenticates Microsoft Word 16.112.4 build
16.112.26090911, LibreOffice 26.2.5.2
cd7284b4cbbfeb507e630c1aac019f4157393acb, and Poppler `pdftotext` 26.09.0.
Both viewers render the clean source-built four-container candidate as one
612 by 792 point page. Poppler layout extraction removes whitespace and the
viewer-specific bullet glyph, then requires the same complete ordered token
record for header, body, cell, and footer fragment content. Both pages are
rasterised at 150 DPI and compared with the shared global luminance SSIM
metric. The floor is 0.75 because Word and LibreOffice use different list
indentation, paragraph spacing, and table width while the exact token, page,
and page-size checks independently forbid missing or reordered content. The
reviewed pair scores 0.7853882556046246.

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

`comparison_tracks_changed_table_grids_as_table_replacement` source-builds
tables that gain a column, lose a column, and resize both columns. Each tracked
document saves and reopens with deletion before insertion. Acceptance compares
cleanly with the edited table and rejection compares cleanly with the original.
The surrounding comparison matrix retains focused row, cell, and table
property revisions when the active grid is unchanged.

Drawing preservation coverage keeps inline and anchored drawings in changed
main-story and header owners. It checks exact wrapper bytes, local prefix
bindings, extended `docPr` payloads, relationship targets, media bytes, and
schema child order through compare, save, reopen, accept, and reject. Run,
word, and character granularity each retain a stable multi-unit sibling once.
A malformed drawing without `docPr/@id` fails before the original document or
package changes.
Inherited-binding coverage adds story-root and outer-drawing declarations plus
an unrelated typed edit before comparison. It checks dirty staging, complex
fields, sibling fields sharing one physical run, and both revision outcomes.
`text_only_comparison_with_body_header_and_footer_drawings_accepts_exactly`
moves the `wp` binding from each story root to its drawing run while making six
body edits and one footer edit beside a complex PAGE field. Run, word, and
character comparison must accept and reject to exact visible text, retain the
body, header, and footer relationship targets and media bytes, and preserve raw
drawing payload after declaration placement normalization. Self-comparison is
a byte-exact no-op, while a changed `docPr` identity remains visible to both
revision outcomes.

Terminal paragraph comparison coverage inserts at the start, middle, and end
of the main story, including one, two, and three appended paragraphs. It proves
exact accept and reject reconstruction when the original ends in a
self-closing empty paragraph and when retained content includes a field,
drawing, media relationship, and unrelated opaque package part.

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

The portable authored-chart gate constructs line, bar, pie, and doughnut
documents through the public Word API. Focused regressions check axis titles,
explicit visibility, value-axis formats, series and indexed point colours,
one-series legends, percentage labels, the doughnut hole, exact workbook data,
complete staged identifiers, typed theme reuse, collision handling, malformed
theme replacement, failure atomicity, and save/reopen preservation. A facade
test proves that `oxml-chart`, `rdocx`, and `rpptx` expose the same `RgbColor`.
The line portability case requires each authored axis title to order `c:tx`,
`c:layout`, and false `c:overlay`, then requires false `c:marker` and
`c:smooth` before the line plot's axis ids. Parse and rewrite retains those
defaults and the palette. Bar, pie, doughnut, area, scatter, and radar omit the
line-only defaults, while every non-scatter workbook remains byte-identical to
the line workbook for the same source data.

The ignored external oracle generates the exact
`ab67b50393fc5258f7a3e9719344639d665feccc2615b13cab1915ea9a84566b`
candidate in code. Microsoft Word 16.112.3 build 16.112.26083020 must open it
without repair. Pages Creator Studio 15.1.1 build 7044.0.273 must render the
declared axes, colours, percentages, legend, and doughnut shape, then export
DOCX bytes whose ChartML and relationship-owned workbooks retain the source
semantics across all four editable workbooks. The test opens the candidate
through macOS LaunchServices before AppleScript export so Pages receives the
sandbox-scoped file URL. Each run records the exported
DOCX digest. That digest is not a fixed expectation because Pages recalculates
manual chart layout coordinates and drawing extents between exports. The gate
compares the parsed chart and workbook semantics instead of producer bytes.
This gate hardens Kevin Brown's PR 71 contribution without importing its stale
delivery records. No standard sample authors a chart, so all 49 hash entries
remain unchanged.

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

The advanced table geometry golden gate is
`fixed_autofit_and_nested_table_geometry_matches_reviewed_word_pages`. It
builds one document in code holding a fixed-grid table, an auto-width autofit
table, and a nested table, lays it out in deterministic font mode, and pins the
page count together with the origin, width, and height of every painted cell,
which is where the per-row origins and the resolved column widths are visible.
Focused tests cover the `w:tblpPr` attribute matrix under an alias prefix, the
autofit engagement predicate, minimum-plus-slack distribution, bidirectional
column reversal against logical cell order, row grid offsets, a conditional
region's row height reaching layout, and checked-setter rejection. No standard
sample authors a floating, bidirectional, spaced, or auto-width table, and the
engagement predicate keeps every authored `dxa` and `pct` table on the declared
grid, so all 49 hash entries remain unchanged.

The floating table golden gate is
`floating_tables_match_reviewed_word_page_geometry_and_pagination`. It builds
one document in code holding a margin-anchored, a page-anchored, and a
text-anchored float, lays it out in deterministic font mode, and pins the page
count, the three float origins, and the line boxes of the text inside each
float's keep-out band. Focused tests cover the anchor frame mapping and the
inline spelling, a float taking its origin from the anchor rather than the
indent, a float that does not fit moving whole to the next page without
repeating a header row, the two-pass convergence of a text-anchored float, the
look-ahead that pushes the text above a float aside, and float against float
resolution for `w:tblOverlap`. The named guard that the wrap extensions stay
inert is `a_document_with_no_floating_table_still_paginates_in_one_pass`, which
holds the two-pass predicate false for an ordinary table and for a float framed
by the page or a margin. No standard sample authors a floating table, so all 49
hash entries remain unchanged.

The M23 drawing gate is `m23_drawings_text_boxes_and_watermarks_match_word`.
It authors cropped inline and floating pictures, every wrap family, rotated and
vertical text boxes, compatibility fallbacks, and section-selected watermarks
through the public facade. Save, reopen, and repeat-save retain exact owner
relationships, schema order, source rectangles, wrap polygons, WPS geometry,
and self-contained VML references. Companion tests cover the complete drawing
option matrix, selected default, first, and even header variants, preserved
leading and trailing text-box spaces, and atomic rejection of invalid or
overflowing geometry. Deterministic bundled-font PDF and PNG output is checked
with Poppler 26.01.0, while LibreOffice 26.2.5.2 provides the portable reopen
and visual inspection. The Word 16.112.4 structure record is reviewed
statically when GUI automation is unavailable. No sample opts into the new
options, so all 49 hash entries remain unchanged.

The M24 section page-semantics gate is
`section_page_semantics_match_pinned_libreoffice_render`. It source-builds one
document carrying two columns with a rule, a page border, line numbering and
mirrored margins, renders it through the deterministic font manager, converts
the same package with LibreOffice 26.2.5.2, and rasterises both at 150 DPI with
`pdftoppm` 26.01.0. The comparison is the horizontal ink blocks of the page
interior, which are the line-number band, each column track and the rule
between them, and each block edge must agree within six pixels, or 2.9 points.
Global-window luminance structural similarity is reported and floored at 0.15
as a collapse guard rather than a similarity claim, because over a page of 11
point prose it measures glyph rasterization far more than layout: the measured
agreement is 0.21 while the same page against a blank sheet scores 0.02.

**Word GUI automation is not available on the machine that produced this
gate**, so no Word-authored pinned record set was recorded for it. The Word
confirmation lands instead as the mandatory `#[ignore]` capture test
`capture_f269_word_section_evidence`, which asserts the installed Word build
before it records anything and is never part of the automated gate, plus a
named follow-up in `docs/hld/14-development-backlog.md`. That is the same
pattern every other GUI-only evidence path in this document uses. F-251's
`pdftotext` and `pdfinfo` 26.09.0 pins are left exactly as they are. The only
two tests that invoke those binaries are `#[ignore]`d regeneration helpers, so
nothing fails today, and re-pinning them to the installed 26.01.0 would
invalidate recorded evidence no human here can re-capture.

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
footnote or endnote part retains at least 698 of 700 ordinary hits and rebuilds
at most two paragraphs across text, insertion, and deletion changes. Warm and
fresh deterministic layouts and source paths remain exact. A third transaction
after the note edit and one ordinary paragraph edit still records 699 hits and
one build, proving unaffected entries survived publication. Note-bearing table
and header or footer content remains conservative, full restart reuse stays
disabled after a note change, a late failure publishes nothing, hits preserve
insertion order, and FIFO eviction holds at the independently pinned 4,096-entry
and 50 MiB paragraph limits. Cacheable
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
The sourced body-length regression inserts and deletes near block 640 of 700,
then exercises Enter, adjacent merge, and multi-block selection deletion. Each
operation restarts from a safe retained prefix and recomputes at most three
pages while every layout field and Word source path equals a fresh deterministic
result. Page-frame identity checks require retained identities to form only a
contiguous prefix, so no shifted sourced tail can survive. The existing
source-free insert, delete, and undo matrix retains exact suffix attachment.

The restart-identity memo regressions build a 715-block mixed body and require
each candidate identity to be serialized at most once across all restart scans
and restart-record publication in one layout. Warm and fresh deterministic
results remain exact. A forced fingerprint collision still compares complete
serialized bytes, while a fingerprint miss leaves the candidate memo slot
uncomputed. Test-only counters bound peak populated slots by the body length and
peak retained identity capacity by the published restart identities. The memo
itself is local to one layout, while the existing 5,216-entry and 64 MiB checks
continue to govern persistent retained work.

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

## The mixed-script geometry golden gate

`mixed_script_page_matches_the_pinned_geometry_and_reading_order` builds one
page through the public facade with Arabic, Hebrew, Korean, Japanese, Latin,
and one Kanji paragraph, each setting its font through the `w:rFonts` slot
Word uses for it, and lays it out with `FontManager::new_deterministic`. The
digest is taken over a canonical serialisation of every painted run in paint
order, carrying its font family, point size, origin, logical text, glyph ids,
and advances, and for a rich run also its direction, script, bidi embedding
level, both offset axes, and its cluster ranges.

The recorded digest is
`516ebb6e45438731d3cb0983707ad00c9de55068401e073ef2a069a56f397402`.

The Kanji paragraph is what makes slot resolution load bearing in the gate. It
names `Noto Sans SC` on `w:ascii` and `Noto Sans JP` on `w:eastAsia`, and both
faces cover its text, so coverage fallback cannot choose between them. Every
other paragraph has exactly one bundled face that covers it, so those prove
script identity, reading order, and geometry rather than slot resolution.

The gate is a geometry digest and not a pinned-oracle raster, because every
property under test is stated exactly by the layout result. A rasteriser would
add an external dependency and prove less. The serialisation is host-stable
because the pipeline is f64 throughout with no FMA contraction, the shaper is
pure Rust, and every face is bundled, so each coordinate is an integer font
unit scaled by one multiply and summed in a fixed order. Printing four decimal
places is not what makes it stable. It is a guard band that keeps an ordinary
representation difference away from the printed digits, and it is applied to a
value whose sign of zero has been normalised first, because a negative zero
would otherwise print a leading minus and move the digest for no geometric
reason. Reading order, script identity, resolved family, and bidi embedding
level are asserted separately and readably before the digest, so a failure
names the property that broke rather than only reporting that a hash moved.
Logical order is asserted as exact reassembly per source paragraph, not as
containment, so a dropped span fails.

Adjacent regressions prove that a right-to-left paragraph keeps logical order
in its rich runs, its cluster ranges, its SVG text, and its round-trip XML,
that Arabic shaping applies contextual joining forms inside one run against a
zero-width non-joiner control, that a numbering marker stays on the leading
right-hand edge of a right-to-left paragraph, and that East Asian text reads
`w:eastAsia` rather than `w:ascii` where both named families cover the text.

Korean text now reaches the rich shaping path, and a paragraph on that path
cannot enter the paragraph block cache, because the retained size of a rich
inline item cannot be bounded. That has always been true of Arabic, Hebrew,
and CJK. `a_complex_script_paragraph_is_never_admitted_to_the_paragraph_block_cache`
asserts it for all five rather than leaving it to be rediscovered, and the
incremental relayout tests use Latin fixtures so they still measure reuse.

## The ruby and emphasis geometry golden gate

`ruby_and_emphasis_page_matches_the_pinned_geometry_and_reading_order` builds
one page through the public facade carrying a ruby-annotated Japanese word, an
emphasis-marked Japanese run, an emphasis-marked Korean run, an
emphasis-marked Latin run, and a Latin control, and lays it out with
`FontManager::new_deterministic`. It reuses the canonical serialisation the
mixed-script gate records, which walks the element tree rather than the top
level, so the runs inside an annotation group reach the digest with everything
else.

The recorded digest is
`b119714501d061f912bf9c05224f66dc8d4a30f3bdd195040038b89157e6fbf6`.

The same test lays out the mixed-script page again and asserts
`516ebb6e45438731d3cb0983707ad00c9de55068401e073ef2a069a56f397402` is unmoved,
so a change that quietly moved the sibling story's baseline fails here rather
than at the next re-record.

Before the digest the test asserts the properties one at a time, so a failure
names what broke. The base line is painted before its phonetic line, both
emphasis glyph inventories appear once per non-space base character, and the
saved document's paragraph text carries the ruby base and never the phonetic
line.

The gate is a geometry digest and not a pinned-oracle raster, for the reasons
the mixed-script section gives. Nothing here needs a rasteriser.

Adjacent regressions prove that paragraph text extraction returns the ruby base
alone, that redaction still steps over `w:ruby` now that it is modelled, that a
run split before an annotation carries the span with it rather than leaving it
on the neighbouring run, and that a marked run advances exactly as far as an
unmarked one.

## The grid and vertical geometry golden gate

`grid_and_vertical_page_matches_the_pinned_geometry_and_reading_order` builds
one page through the public facade carrying a `linesAndChars` gridded Japanese
paragraph, a combined run inside round brackets, a table with a `tbRl` cell and
a `btLr` cell on either side of an `lrTb` control, and a Latin control, and
lays it out with `FontManager::new_deterministic`.

Its serialisation is the mixed-script one with the accumulated group transform
added. A rotation never reaches a glyph run's own origin, because the run keeps
group-local coordinates and the rotation lives on the group above it, so a
digest over origins alone would pass with every rotation removed. Recording the
six transform coefficients is what makes this gate prove the subject it exists
for.

The recorded digest is
`cb3043d53719f5dd9e16b61a001aff8c8827c19f96536972d4a17b9a626d2164`.

The same test lays out both sibling pages again and asserts
`516ebb6e45438731d3cb0983707ad00c9de55068401e073ef2a069a56f397402` and
`b119714501d061f912bf9c05224f66dc8d4a30f3bdd195040038b89157e6fbf6` are unmoved,
so this story cannot move either recorded baseline while recording its own.

Before the digest the test asserts the properties one at a time. Both rotated
cells paint their text and reach the page through a transform, the horizontal
control reaches it untransformed, and the reopened document returns its
paragraphs and its cells in grid order, because rotation and combining are
painting concerns and the saved bytes stay logical.

The gate is a geometry digest and not a pinned-oracle raster, for the reasons
the mixed-script section gives. Nothing here needs a rasteriser.

Adjacent regressions prove that a `default` grid produces geometry identical to
no grid at all, that adding a rotated neighbour does not move a horizontal
cell, and that a rotated cell's text is still extracted, rendered to SVG and
saved in logical order.

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
   source-ordered diagnostics for skipped unmatched notes-slide overlays. The
   source-built Google Slides index variant must produce the same notes PDF and
   PNG as its matched control while retaining the hard-failure cases. The
   geometry unit gate covers exact targets, clipping, five rules, and rejection
   of a 1.01-point displacement. The 49-entry render hash manifest remains
   unchanged.

The notes-owner regression removes both reverse slide relationships from the
same source-built control and requires byte-identical deterministic notes PDF.
Companion mutations require conflicting, multiple, and external reverse owners
to fail closed. Notes mutation saves and reopens while retaining the body
placeholder identity, first-run formatting, and unmodelled run XML. Missing
notes and missing body placeholders fail without changing package bytes.

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

The source-built latent-placeholder differential is
`slide_owned_latent_placeholders_ignore_master_header_flags`. It copies the
bundled blank layout's slide-number field onto a slide and tests an absent
master `p:hf`, every master flag disabled, and only `sldNum` enabled. The exact
LibreOffice 26.2.5.2 and Poppler 26.01.0 path must agree with deterministic Rust
PDF text, retain bottom-right slide-number ink, and leave the inherited date
region empty. A one-pixel PDF page-width difference is accepted because the
two PDF writers round the 10-inch page boundary differently at 72 DPI.

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

The container-neutral story gate source-builds one package containing body,
cell, text-box, header, footer, footnote, endnote, and comment owners. Its
integration predicate applies the same location resolver and text mutation to
every kind, then checks identical wrong-owner, kind, bounds, and stale errors.
The round-trip predicate pins deterministic owner and item order for paragraphs,
tables, controls, fields, drawings, and preserved nodes, including exact raw
XML after save and reopen. The atomicity regression records package bytes
before every invalid location and requires them to remain unchanged. Focused
regressions also cover prefix aliases, typed admission, nested and same-run
fields, namespace-complete owned projections, borrowed subtree semantics,
opaque preservation boundaries, separator-note filtering, lifecycle state,
and mutation of empty text nodes.

The ordered mixed-run gate authors text, a tab, all three typed breaks, an
inline picture, a field, a Unicode symbol, and trailing text through one native
run handle. It applies the public formatting setters, saves and reopens the
package, and compares the flattened logical child order across the physical
field boundary. The cached field result and both surrounding run segments keep
the authored properties. A companion regression inserts a field beside a
positioned foreign run child, proves the raw bytes stay on their original side
of the field boundary, and proves an invalid field instruction is atomic. The
reopened document must render through deterministic bundled fonts.

The Word paragraph-property round-trip gate is
`every_public_paragraph_property_reopens_and_preserves_unrelated_xml`. It opens
a source-built document whose `w:pPr` carries unmodelled children at four
schema slots, authors every paragraph property through `Paragraph`, saves,
reopens, and reads each value back through `ParagraphRef`. The retained
children must survive byte identical. Focused unit coverage proves the frame
attribute round trip, the schema sequence of the newly typed children among
retained raw siblings, the attribute carrier of each newly typed toggle,
prefix-tolerant reads with a foreign same-local element left unmodelled, the
border edge attribute retention that section page borders consume, and style
inheritance for the new members. Integration coverage authors, reads, and
clears each border edge, each tab stop by index, and the paragraph mark, and
pins the accepted outline-level range.

The Word owner-attribute round-trip gate is
`paragraph_run_and_section_identity_attributes_survive_noop_save`. It
source-builds paragraph, run, and section-property roots with aliased modern
identities, revision-session values, foreign attributes, and an unqualified
attribute. Save, reopen, typed text mutation, and another save retain the exact
values, source attribute order, child schema order, and deterministic package
bytes. Focused unit coverage rejects duplicate expanded names and proves that
authored `paraId` replaces only its expanded-name match. The public run-shape
regression prevents the internal retention record from changing the existing
`CT_R` struct literal surface.

The run-level page-break differential gate authors both the break-only
paragraph written by python-docx and a break between two pieces of text. Its
source-built record pins Microsoft Word 16.104 build 16.104.25121423, the
en-GB locale, and deterministic bundled fonts, then compares page count and
page text. A line-layer unit matrix retains distinct line, page, and column
break kinds. Paginator regressions cover two breaks, keep-lines, an overflow
continuation, a trailing break, and the non-page break kinds. The consumer
regression proves one split paragraph produces ordered body fragments on two
physical pages, resolves PAGE, NUMPAGES, and PAGEREF from that sequence, emits
two deterministic PNG pages, and reports two pages from the deterministic PDF
through pinned Poppler. No binary fixture or runtime oracle dependency enters
the published crates. The adjacent-break differential pins the reporter's
LibreOffice Writer 26.2.5.2 page and text result for a trailing run page break
followed by `pageBreakBefore`. Single-break controls remain two pages. Focused
paginator controls retain separate transitions across intervening content,
visible continuation formatting, line breaks, and column breaks.

The contributor reader-fact regression combines strict document and body
boundaries, first section properties, missing revision authors, empty simple
fields, bounded nested tables, marker child-content facts, effective complex
field properties, cell margins, empty cells, historical grid preservation, and
self-closing TOC coordinates. Namespace aliases and decoys, malformed owners,
schema order, repeated save and reopen, and the exact seven-entry Word XML hash
delta are mandatory. Every PDF and PNG fingerprint remains unchanged.

The generic content-mutation regression interleaves insert, remove, clone, and
same-owner move operations across body and table-cell content. It asserts final
direct-child order, fresh clone identities, valid relationship scope, exact
unmodelled XML, save and reopen, and byte-identical rollback for invalid or
stale inputs. Focused source-built cases distinguish canonical actual-item
anchors from nested flattened projections and use `ContentLocation::end` for
empty, self-closing, and body section-property boundaries. The content-control
matrix covers expanded-name aliases, schema order, retained private slots,
direct nested block controls, opaque foreign subtrees, and exact XML 1.0
whitespace and character-reference handling without panics.

`clone_content_scales_linearly_and_names_invalid_arguments` builds 50- and
100-paragraph packages without an external producer. It bounds the doubled
clone workload, verifies the cloned text, and pins the source-handle and
destination-index type errors. The surrounding clone regressions retain fresh
identities, relationship scope, stale handles, and byte-identical rollback.

The ordered section round-trip gate is
`ordered_section_mutations_preserve_independent_story_references`. It creates
four ordered owners, removes one boundary, and reopens three portrait,
landscape, and portrait sections without moving body content or unmodelled
section XML. Every surviving explicit or materialized story reference is
checked by variant, relationship id, exact relationship type, resolved target,
and parsed story text. Focused tests cover total immutable and mutable lookup,
bound ordinal and final-owner identity, dimension-normalizing orientation
mutation, final-owner promotion, inherited default and first-page headers and
footers, shared authored targets, producer-owned targets, last-owner pruning,
and sole-section or out-of-range atomicity. The relationship hazard matrix
rejects missing, external, cross-type, absent-part, and malformed-target
overrides, keeps the first usable duplicate, and proves a pruning-scan failure
leaves the live typed document and package graph unchanged.

The complete section-geometry differential gate is
`mixed_orientation_sections_match_word_geometry_and_page_numbers`. One
source-built document produces portrait letter, landscape letter, and A4 pages
with physical identities 1, 2, and 3 and displayed PAGE values 1, 12, and 27.
The ignored regeneration route pins Microsoft Word 16.112.3 build
16.112.26083020 and Poppler 26.09.0, checks the PDF's total page count, and
requires one complete geometry and displayed-number record per page. Its exact
unique artifact directory is removed on success and unwind.

The per-section story differential gate is
`section_header_footer_variants_match_word_width_and_inheritance`. One
source-built document produces three pages at each of 612, 720, and 540 point
page widths. Every group selects the expected default, first, and even header
and footer through direct, inherited, and replaced references. The ignored live
route pins Microsoft Word 16.112.4 build 16.112.26090911 on macOS 26.6.2 build
25G83 and Poppler 26.09.0. It requires all nine exact page records and removes
its unique artifact directory on success and unwind.

The related-story picture gate is
`header_and_footer_pictures_render_from_story_relationships`. It reserves the
same local image relationship identifier in the main story, a header, and a
footer, then proves layout returns each owner's exact bytes after save and
reopen. Companion layout tests cover inline and anchored scoped identifiers,
collision-safe media identities, and one deterministic diagnostic for a
missing scoped relationship without body fallback. The sample harness records
the existing 400 by 40 feature-showcase header logo as a PDF image, while the
pinned LibreOffice 26.2.5.2 and Poppler checks confirm the same logo dimensions
and 200 DPI placement in the external render.

The installed Python structure gate is
`word_structure_snapshots_preserve_order_ownership_and_types`. A source-built
three-section package includes inherited and independent header and footer
owners, rich story content, a custom style, and equal hyperlink relationship
identifiers in the body and two headers. It proves exact frozen records,
source order, physical ownership, item paths, owner-scoped URL resolution,
single ownership for a hyperlink in nested content controls, physical ordering
when a nested link precedes an ancestor-owned link, snapshot stability after
mutation, strict mypy, and stubtest after reopen. The body-owner matrix adds a
direct paragraph, table, and body-level control plus nested field, drawing, and
run-level control items. Native and installed Python assertions keep the flat
path while returning the containing direct body coordinate. Header items and
final section properties remain unanchored.
The native companion
`story_item_links_resolve_only_through_the_checked_owner` resolves an equal
identifier to distinct body and header targets and checks the same interleaved
source order. All 49 hash entries remain unchanged.

The Python story-scale gate is `python_story_inventory_scales_linearly`. It
doubles a corpus containing paragraphs, table cells, and hyperlinks, requires
exact doubled inventory counts, and bounds the elapsed ratio without relying
on an absolute machine speed. A native counted companion requires one complete
story-source build for each item or hyperlink snapshot. A second counted
companion, `story_link_snapshots_read_each_link_from_its_own_span`, doubles a
linked corpus and bounds the XML events the story text walker reads, so reading
each link from the head of its part fails it.
`story_link_snapshots_match_story_links_when_a_part_binds_word_twice` requires
the snapshot, story and item hyperlink projections to agree on the text of a
header link when the part binds Word to two prefixes. The binding companion
interleaves direct, inserted, inline-control, and deleted runs, then proves
StoryItem text, Paragraph text, and live run handles use the same accepted
order. Nested formatting and splitting survive save and reopen, while an old
run handle fails after the structural split.

The indexed Python content gate is
`python_indexed_content_mutation_is_counted_and_atomic`. It maps live direct
Paragraph and Table handles across interleaved content, inserts a paragraph,
pops and reuses an opaque fragment, clones and moves content, then reopens the
package in final source order. Nested, stale, foreign, and out-of-range inputs
must fail without changing bytes or invalidating a live handle. Literal and
regular-expression replacements return exact cross-run counts, a zero count
keeps handles live, and invalid syntax is atomic. A native companion covers
paragraph ordinals that include nested block-control paragraphs. The complete
45-test installed cp39-abi3 wheel suite, strict mypy, stubtest, both WASM
checks, the workspace gate, and the unchanged 49-entry hash set complete the
binding proof.

The Python story mutation gate is
`python_story_revision_field_and_xml_operations_are_typed_and_atomic`. It
accepts a compared revision set, authors default header and footer stories,
adds a relationship-scoped hyperlink, reads exact immutable item XML, and
reopens the edited package. The same gate proves stale StoryItems fail,
replacement removes only newly empty hyperlinks, cloned comment anchors do not
multiply, and direct paragraph lookup precedes an enclosing content control.
Companion tests cover every revision filter, field context evaluation, revision
invalidation, and GIL release. The complete 57-test source-tree suite runs
against pinned Poppler 26.01.0. A fresh cp39-abi3 wheel installs and runs on
Python 3.9, while strict mypy and stubtest validate that same wheel on Python
3.12.

The round-three installed binding gate is
`python_round_three_authoring_and_inspection_is_typed_and_lossless`. Its Word
half inserts an in-memory picture after a checked story item and anchors a
comment in a table-cell paragraph. Its presentation half checks shape bounds
and identity, direct run font facts, autofit, notes mutation, and run text
replacement after reopen. The package assertions retain run properties,
foreign children, media relationships, comment anchors, and unrelated XML.
Strict mypy and stubtest cover the same optional values, frozen ranges, and
setters.

Table-property round-trip coverage opens `0`, `false`, and `off` for
`w:tblHeader`, `w:cantSplit`, and `w:noWrap`, then saves and reopens the table.
Existing bare-element and absent-property cases retain true and inherited
semantics, and ordered raw siblings remain in their schema slots.

`rich_section_stories_survive_reopen_replace_and_unlink` covers paragraphs,
tables, fields, block controls with nested tables, hyperlinks, images, drawings,
fresh drawing identities, relationship rebasing, title-page enablement, save,
and reopen. `removing_one_variant_retains_shared_and_inherited_stories` covers
explicit empty removal, same-type inheritance, shared targets, facade-owned
pruning, and byte-exact retention of unrelated stories. Stale, wrong-kind,
out-of-range, and absent-story operations prove byte-identical rollback. The
typed settings test covers namespace aliases, fixed write prefixes, schema
order, and duplicate projection behavior. All 49 hash entries remain unchanged.

The companion round-trip gate is
`section_geometry_round_trips_with_unsupported_children_in_order`. It covers
every M23 geometry property, typed page-number start, schema-slot replay,
prefix aliases and shadows, duplicate page-number elements, retained M24
attributes, repeated reference boundaries, and save and reopen. Distinguishable
references retain predecessor or successor placement. Indistinguishable equal
duplicates use deterministic source ordinals and produce byte-stable output
after value-preserving String replacement, reorder, removal, clone, and reopen.
`rejected_section_geometry_is_atomic` covers zero, negative, and signed-range
failures without changing the document. No standard sample authors this state,
so all 49 hash entries remain unchanged.

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

## The private from-scratch DOCX conformance corpus

M23 uses five client reference documents that are never committed, published,
or fetched by repository automation. They live in a configured ignored private
directory outside the tracked corpus tree. The repository must not contain the
documents, their rendered pages, extracted text, identifying filenames,
customer metadata, or a manifest that would disclose their provenance.

F-240 inventories each private document locally at package-part, relationship,
section, story, paragraph, run, table, numbering, field, drawing, and layout
levels. Its capability matrix stores only non-identifying classifications in
tracked documentation. Exact hashes, filenames, XML extracts, and differential
artifacts remain beside the private corpus. A staged-path and tracked-path scan
fails when a private artifact is about to enter repository history.

The anonymous audit records only the capability families needed by each local
reference. Counts, values, source names, text, relationship targets, media,
rendered pages, and extracted markup are not part of this summary.

| Anonymous reference | Non-identifying required families |
|---|---|
| P1 | Section-scoped header and footer variants, fields, drawings, text boxes, content controls, custom data bindings, custom properties, and web add-in declarations |
| P2 | Section-scoped stories, dense tables, fields, drawings, text boxes, links, custom data bindings, custom properties, and web add-in declarations |
| P3 | Numbering, tables, section-scoped stories, drawings, text boxes, theme, font table, settings, and web settings |
| P4 | Multiple sections, tables, drawings, links, theme, font table, settings, and web settings |
| P5 | Modern comment metadata, numbering, tables, section-scoped stories, drawings, text boxes, theme, font table, settings, and web settings |

These are requirements for the source-built conformance documents, not
fingerprints. The matrix maps them to F-243 through F-310 without recording the
combination of package counts or property values from any source.

`scripts/docx_authoring_conformance.py` is the single owner of the M23
conformance gate. It provides self-test, public-only, optional-private, and
required-private entry modes. The two evidence paths are:

1. Public CI source-builds a synthetic fixture through a temporary consumer
   whose sole dependency is `rdocx`. Construction begins exactly once with
   `Document::new()` and uses public `rdocx` APIs only. The harness validates
   package inventory, content types, relationships, schema child order,
   modeled reopen state, unsupported-content diagnostics, unmodeled-part
   preservation, and repeated deterministic 150 DPI output. A base package,
   raw XML injection, private OXML facade, or prebuilt document fails the gate.
   The same consumer source-builds Word-compatible DOCX, DOCM, DOTX, and DOTM
   profiles. Each profile must retain its exact main-part identity, complete
   normalized support graph, deterministic relationship identities, omitted
   fresh timestamps, and absence of a synthesized VBA project.
2. Local required-corpus mode builds the five target documents through their
   pure Rust generators and compares them with the configured private
   references. Each temporary consumer depends only on the public `rdocx`
   facade, starts exactly once from `Document::new()`, and reopens only its own
   output. Static boundary checks reject HTML import, raw package access,
   private crates, source templates, and runtime access to the private corpus.
   Every generator runs twice and must produce identical DOCX bytes. Missing
   input, unexpected input count, a digest change, or missing evidence fails
   closed. An optional-private invocation reports a skip when the ignored
   directory is absent, while required-private mode rejects that absence. This
   mode never prints document text or embeds source XML in a tracked report.

The local comparison first normalizes ZIP metadata that is not document state.
Generated packages must resolve every internal relationship and expose exactly
one package-root office-document relationship to `word/document.xml`. The gate
then checks the modeled projection, schema child order, required compatibility
branches, and successful reopen. Pinned LibreOffice opens and exports every
generated DOCX through an isolated profile. A conversion failure, repair or
corruption diagnostic, missing output, invalid package graph, or invalid root
relationship fails the gate. The harness renders both reference and candidate
with deterministic bundled fonts at the pinned resolution and records page
count, exact dimensions, and image similarity against a reviewed per-case
threshold outside the repository. The ignored manifest binds the anonymous P1
through P5 aliases to exact local digests, page expectations, and tool
identities. Any approved embedded raster evidence stays inside the ignored pure
Rust generator sources. Tracked-path and staged-path scans reject private
document formats without echoing a path or digest. Feature-level tests remain
authoritative when byte identity is not a valid expectation.

The feature-level property gate authors core, application, and custom
properties plus the bounded settings defaults through public APIs. It saves and
reopens every typed value, proves selective removal leaves unrelated graphs
intact, and compares two equivalent fresh outputs byte for byte.

Settings mutation fixtures bind the Word namespace through aliases and place
foreign subtrees before, inside, and after modeled settings. They assert exact
foreign subtree bytes after mutation and assert the schema order of default tab
stop, character spacing control, compatibility settings, document variables,
and theme font language.

`public_authored_settings_package_reports_no_unmodeled_supported_children` is
the closed-set gate. It authors every supported settings and web settings member
through the public `Document` surface, saves, reopens, and asserts that both
diagnostic lists are empty and that foreign subtrees placed before, inside
`w:compat`, inside `w:mailMerge`, after the last child, and inside
`w:webSettings` are byte identical.
`settings_children_serialize_in_schema_sequence_order` authors members in
reverse schema order and proves the `xsd:sequence` result at the settings,
`w:compat`, `w:mailMerge` and `w:webSettings` levels.
`supported_settings_is_a_strict_subsequence_of_the_order_table`,
`every_supported_name_is_projected_by_from_xml`,
`duplicate_and_malformed_supported_children_report_diagnostics`,
`unsupported_settings_children_are_never_diagnostics` and
`compatibility_option_covers_the_complete_compat_on_off_set` keep the constant,
the parser and the diagnostics from drifting apart.
`web_settings_part_is_created_on_demand_and_pruned_when_empty` and
`fresh_package_profiles_gain_no_web_settings_part` hold the packaging rule that
an ordinary save adds nothing.
`document_default_tab_stop_drives_implicit_tab_positions` runs in deterministic
font mode and pins both the document interval and the exact `36.0` point
fallback.

`update_fields_on_open_is_typed_optional_and_schema_ordered` adds absent, bare
true, explicit false, namespace alias, foreign lookalike, set, clear, remove,
and save-reopen coverage for `w:updateFields`. Companion low-level cases keep
duplicate and malformed producer forms byte-identical and reject mutation.
The native allocation test exhausts the relationship identifier space and
proves setting a value fails atomically while removing an absent value remains
a successful byte-identical no-op. Installed Python tests, strict mypy, and
stubtest cover the optional property and XML error surface.

The theme and font feature gate authors a shared DrawingML theme, language
defaults, descriptive font records, and an explicitly licensed caller font
through public `rdocx` APIs. Save and reopen must return the same typed values,
font key, license identity, relationships, and original font bytes. Rejected
authorization, missing license identity, and XML-normalized identity whitespace
must leave package bytes unchanged. XML syntax in an accepted identity returns
unchanged after reopen. An aliased-prefix producer font table keeps foreign
children and relationship attributes in schema order after a modeled edit,
including merged MCE ignorable tokens. Equivalent construction and repeated
saves are byte-identical. The differential check pins Microsoft Word
16.104 build 16.104.25121423 and LibreOffice 26.2.5.2 build
cd7284b4cbbfeb507e630c1aac019f4157393acb, then compares deterministic pixels
for the publicly authored embedded face against the bundled oracle face.

The style feature gate authors paragraph, character, and table defaults,
based-on inheritance, reciprocal links, next styles, UI flags, and conditional
table regions through the public facade. Atomic regressions reject missing and
wrong-type targets, duplicate defaults, and inheritance cycles with unchanged
document bytes. Update and removal cases retain producer XML exactly. The
differential pins Microsoft Word 16.104 build 16.104.25121423 and LibreOffice
26.2.5.2 build cd7284b4cbbfeb507e630c1aac019f4157393acb. The sanitized Word record
must match normalized effective properties and deterministic 150 DPI bytes
with a zero-byte difference threshold. The pinned LibreOffice build must save,
reopen, and retain the valid graph, reciprocal links, and effective formatting.

The M23 table-property gate is
`m23_layout_and_data_tables_match_word`. It starts with `Document::new()`, uses
only the public table facade, saves and reopens typed width, grid, indentation,
layout, shading, border, margin, and look values, and checks canonical table
property sequence. It compares the sanitized fixed-grid geometry and reviewed
pagination record, then renders twice with deterministic bundled fonts and
requires identical PNG bytes. Companion tests cover all width modes, valid
spanning rows, aliased Word prefixes, legacy `w:tblLook` masks, explicit
invisible borders, exact raw property and border-extension retention, row
coverage, and atomic rejection of malformed values and overflow. The final
five-document comparison and its pinned LibreOffice 26.2.5.2 and Poppler
26.01.0 evidence remain owned by the required-private F-263 gate.

The M23 row-and-cell gate is `m23_nested_rows_and_cells_match_word`. It uses
only the public native facade to author checked heights, explicit toggles,
alignment, omissions, horizontal and vertical merges, cell appearance, all six
text directions, conditional regions, and a nested table. Save and reopen must
retain every typed value and canonical row and cell property sequence. The
sanitized record is pinned to Microsoft Word 16.112.4 build 16.112.26090911.
The comparison normalizes Word's removal of explicit false, default direction,
and contextual conditional markers plus its inferred vertical-text row height.
Exact source-form assertions cover those intentional writer differences.
Deterministic bundled-font layout and two same-process PNG renders must agree.

Companion regressions cover exact and minimum height pagination, repeated
headers, split policy, every cell border edge and text direction, merge and
omission rollback, nonempty-cell protection, malformed existing topology,
namespace aliases, foreign same-local children, and exact row, cell, and border
extension retention. The final private corpus comparison remains F-263 and is
pinned to LibreOffice 26.2.5.2 and Poppler 26.01.0.

The caller-width measurement gate is
`independent_nested_tables_measure_to_one_final_height`. It measures two
independent tables with nested tables through checked body locations. The
fixtures cover wrapping, cell margins, borders, horizontal spans, and recursive
row height. The larger fractional point result is rounded upward to one twip
minimum and applied to both outer rows. Deterministic whole-document layout
must then report equal table fragments at that exact rounded height. A companion
paragraph test proves narrower wrapping is taller, rejects zero width and kind
mismatch, preserves package bytes, and retains the same cached deterministic
layout `Arc`. A second companion authors two preserved OfficeMath fallbacks and
requires their two production diagnostics in exact source order.

The row mutation gate is
`table_rows_clone_remove_and_clear_through_native_and_python`. It clones a row
with exact height, direct toggles, and a nested table, clears the copied
toggles, removes the source, reopens, and renders twice with deterministic
fonts. Identity coverage adds a picture, bookmark, comment anchors,
content-control ID, producer XML, and a Word-style unused root default
namespace. The copy has fresh document identities, one shared valid image
relationship, no copied comment anchors, and no `xmlns:xmlns` declaration.
Companion regressions cover vertical-merge promotion, table-level raw and
content-control boundary order, the one-row guard, invalid coordinates,
malformed existing topology, byte-atomic failure, stale Python handles,
negative indexes, strict mypy, and stubtest.

The fresh-profile round-trip gate adds an unrelated unmodelled XML part and
package relationship before reopen and repeat-save. The part bytes and
relationship identity must survive exactly. The optional repair gate is pinned
to Microsoft Word 16.104 build 16.104.25121423 and records whether that local
GUI check ran.

F-263 closes M23 only when all five generators pass local required-corpus mode,
the synthetic public suite passes in CI, opening and saving requires no repair,
and every unexplained structural or visual delta has an owning story. F-240 can
revise the M23 and M24 sprint plan if a later approved audit discovers a missing
public authoring capability. The completed initial audit found the existing
F-243 through F-310 boundaries sufficient for the anonymous requirements.

M24 extends the synthetic conformance matrix to all declared modern DOCX
authoring rows. Each row proves public construction, save and reopen equality,
part ownership, binding parity where exposed, deterministic allocation, and
honest losslessness diagnostics. Preservation-only and permanent non-goal rows
use explicit diagnostic tests rather than an authoring test that silently
falls back to XML.

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

The picture-transparency differential is
`picture_alpha_mod_fix_matches_presentation_renderers`. Its source-built deck
places a 30 percent picture on the slide and another on its layout beside an
opaque control. The exact LibreOffice 26.2.5.2 build exports the deck to PDF,
and Poppler rasterises it at 72 DPI for bounded channel comparisons. The
regular companion gate checks the same resolved opacities, exact repeated PNG
bytes, PDF `/ExtGState` alpha, and modelled slide and layout round trips.
Unit gates cover inherited DrawingML prefix aliases, invalid amounts, raw
sibling and duplicate preservation, picture backgrounds, cached previews, and
multiplication with animation opacity. The existing shared SVG group-opacity
gate covers the same backend-neutral group contract.

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
- `straight_alpha_images_composite_with_premultiplied_pixels` proves that
  transparent stored white and black pixels compose identically over navy in
  PNG, JPEG and TIFF output. `large_pictures_render_or_report_the_decode_limit`
  admits the reported 4000 by 1500 and 2100 by 2100 PNGs through presentation
  raster and PDF output. It also requires one 64 MiB rejection diagnostic and
  one visible bounds fallback. Focused unit gates reject malformed headers,
  overflowing dimensions, and over-limit decoded sizes before allocation.
- Linear and radial path gradients produce type 2 patterns, type 2 or type 3
  shadings, and type 3 stitching functions over interval type 2 functions.
  Structural tests also pin stop normalization, fill and stroke pattern
  operators, mixed solid paint, and page-local pattern resources.
- A 90 degree group rotation turns a linear gradient's sampled colour change
  vertical when rasterised at 72 dpi with the recorded Poppler 26.01.0.
- Rich PDF text tests require one initial matrix per multilingual run and exact
  relative glyph positioning. Same-line tests require one logical
  `ActualText`, unchanged paint traversal, and hard boundaries at owner,
  baseline, source, duplicate-index, and index-gap changes.
- The public Word and PowerPoint regression builds 120 numbered Word lines and
  48 numbered slide lines split across Latin and CJK runs. Pinned Poppler
  26.01.0 must extract every complete line once and in source order. Its 72 DPI
  first-page PNGs must retain the exact pre-change digests for both formats.
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
- Fractional Word paragraph line-spacing coverage opens observed producer
  decimals through the public facade, checks exact positive and negative half
  boundaries and signed overflow, saves canonical integer values, and reopens
  them. Aliased prefixes and sibling spacing attributes remain modeled, while
  equivalent fractional and integer documents render byte-identical pages in
  deterministic font mode.
- The same for `rpptx` and `python-pptx`.

The rpptx binding gate executes the seven python-pptx 1.0.2 Getting Started
workflows with the import namespace changed from `pptx` to `rpptx` and the
minimal public re-fetches required after structural writes. Its differential
rider asserts the exact oracle version, compares each writer through both
readers, and directly compares the normalized rpptx-authored and
python-pptx-authored records. It never compares package bytes and the oracle is
not a runtime dependency.

The rpptx text-property gate `test_text_properties_agree_with_python_pptx_in_both_directions`
writes text frame, paragraph, and run font values with each library and reads
them back with the other. `test_text_enums_match_python_pptx_member_values_and_xml_tokens`
pins every exported text enum member to its python-pptx 1.0.2 value and to the
XML token python-pptx writes for it. Both skip when the oracle is absent, like
the Getting Started rider.

The rpptx extension gate
`presentation_render_comments_and_notes_match_native_snapshots` compares the
single-slide convenience with the ordered all-slide result, checks complete PDF
and PNG signatures for slides and notes, and reads optional notes text. It adds
an author, comments, and ordered replies through native GUID and RFC 3339
validation, saves, reopens, and compares frozen snapshots exactly. A valid GUID
that names an unknown author must fail without invalidating the held slide.
Dedicated thread assertions prove both slide and notes raster calls release the
GIL. The native companion gate compares every convenience PNG byte for byte
with the resolved layout raster path and covers invalid DPI and missing slide
indices.
The binding notes-mutation gate assigns through `Slide.notes_text`, proves the
held slide becomes stale after success, and checks text, placeholder identity,
first-run formatting, and unmodelled run XML after save and reopen. A slide
without notes rejects assignment while retaining both bytes and handle
validity. Strict typing and stub checks require the writable property.

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

The rdocx binding gate
`priority_word_operations_return_typed_snapshots_and_remain_atomic` exercises
frozen range and result values, comment threads, redline comparison,
deterministic layout fragments, page lookup, and TOC counts through save and
reopen. Exact stale-handle revisions prove that each successful structural
mutation advances the binding revision once. Invalid ranges and comparison
metadata prove that failed staged operations preserve both package bytes and
live handles. Dedicated thread tests cover comparison, layout, TOC rebuild,
and serialization GIL release. The installed strict mypy and stubtest gates
cover every added class, nullable field, tuple return, and method signature.

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

The guarded replacement gate refuses existing and input destinations, rejects
unexpected zero and mismatched expected counts without output, accepts an
explicit expected zero, and counts slide plus speaker-note matches exactly.
Successful output retains notes formatting and unmodelled XML, reports exact
stdout, reopens with replaced slide and notes text, and leaves no staged file.

The `rdocx` CLI has one integration binary that invokes the compiled executable
through `CARGO_BIN_EXE_rdocx`. Its tests cover `inspect`, `text`, `convert`,
`diff`, `replace`, `validate`, `render`, nested comment and revision commands,
comparison, and TOC rebuild with in-code DOCX and corrupt-package fixtures. The
assertions bind schema 1, explicit story scope, default paths, exact stdout,
exit-status verdicts, output validity, replacement persistence, comment thread
round trips, id, author, and paired date revision filters, accept and reject
comparison equivalence, document-order text, bundled-font deterministic render
bytes, legacy zero-based `render --page`, one-based `render --pages`, shared
image format extensions, invalid range rejection and no partial output. Process
ID and an atomic counter isolate temporary workspaces across concurrent runs.

The same binary binds the schema-1 structured automation surface. Its exact
text record asserts direct style, numbering, nullable run formatting, and typed
row, cell, and paragraph path segments. The layout record requires positive
point-space extents and one-based physical and displayed page numbers for
paragraphs and tables. Guarded replacement proves both an exact-count publish
and a mismatch that creates no output. Focused deterministic layout tests
separately require a page-spanning body item to produce ordered fragments on
multiple pages, require empty-cell tables and image-bearing paragraphs to keep
real extents, and compare warm restart fragments with a fresh pagination.

All 27 workspace packages explicitly declare one distinct README. The root
README is the high-level `rdocx` guide. It leads with the complete native
document workflow and a seven-row implemented-outcome summary before examples,
installation, alternatives, or boundaries. Its three Rust examples cover blank
authoring, read and mutation, and render and export. The detailed property
boundary remains in the modern DOCX capability matrix rather than a status
table on the product front page. The dated comparison accepts only reviewed
official evidence for rdocx, python-docx, docx-rs, docx4j, and Aspose.Words.
`ND` means not documented in that evidence, and no row makes a volatile
performance, popularity, price, or footprint claim.

Every published measurement uses a fixed nine-column table that names the
value, version, platform, build mode, input, command, statistic, and date. The
root plus all 22 publishable crate pages carry their re-derived `.crate`
archive footprint. Exactly the root, `rdocx-layout`, `oxml-pdf`, and
`rdocx-py` pages also carry the four large-document layout and PDF rows. The
rows state the enforced floor or ceiling and one dated observation. A row
shared by several pages is byte-identical on each page.

Each crate-local document leads with an outcome and at least three implemented
capabilities, then states direct-use guidance, adjacent package relationships,
publication status, and a concrete Rust, CLI, Python, or JavaScript example.
The compatibility shims direct users to their shared replacements. Internal
binding and WASM crates state that they are not crates.io packages. The
presentation Python binding does not claim rendering, `rpptx-render` owns
layout lowering rather than fixed output, and `rpptx-oxml` promises part-level
serialization rather than complete-package preservation.

`scripts/readme_doctests.py` validates the exact package-to-README inventory,
the documented CLI argument names, Python and JavaScript surface names,
deterministic feature guidance, scoped WASM build and import names, and
default-off encryption and signature features. It
derives the root dependency and CLI requirements from Cargo metadata, checks
every local path and Markdown anchor from all 27 source locations, and rejects
root narrative, section-order, workflow-claim, crate-audience, or boundary
drift. The root outcome gate also checks the canonical matrix classifications
that support package I/O, encrypted package I/O, preservation, and permanent
non-goals. Comparison evidence has exact official-URL use counts, exact row
claims, and one exact date-bounded uniqueness conclusion. Its focused
`--check-official-links` mode resolves those sources during implementation
review, while default CI remains network-independent. Mutation tests remove or
reorder the narrative sections, alter a workflow row, restore rejected crate
claims, and break non-root links. It builds the applicable libraries with locked
dependencies and Cargo JSON
messages, locates each emitted rlib from one package build graph, and invokes
rustdoc with the 2024 edition, warnings denied, the dependency search path, and
every matching `--extern` binding. It compiles 23 Rust examples across the 21
Rust-library READMEs. It also creates all 22 publishable archives and
byte-compares their single packaged README with the declared source. Archive
creation uses the same exact 22-package local source patch set as the release
dry run, so a reviewed version can be checked before its internal dependencies
exist on crates.io. The patches never enter an archive and upload nothing. The
docs job and canonical non-fast verification call this same runner. The
archive gate re-derives compressed bytes, member bytes, and member count. It
normalizes Cargo's generated `.cargo_vcs_info.json` to a fixed-length clean
revision before requiring exact source-determined member values. Compressed
size must stay within 64 bytes of the recorded observation under the pinned
toolchain, may grow by no more than the same allowance on another toolchain,
and always remains subject to the 10 MiB ceiling. The allowance covers the
generated commit hash and dirty marker, not tracked source growth.
`--record-measurements` prints the derived archive rows and the approved speed
rows in their exact Markdown form. Mutation coverage rejects incomplete or
stale provenance, row-to-page drift, a speed guarantee beyond its code gate,
an untracked deferred measurement, and unbounded superlatives across all 27
pages.
The stable 0.14.0 carrier regression pins all ten inherited version carriers,
the `rdocx` Python project version, both rdocx WASM dependency assertions,
the stable CI package literal, the seven publishable crates, and every stable
README requirement. It also proves the current incubating workspace carriers
are 0.12.1 while `rpptx-wasm` remains ineligible for publication.
The paired incubating regression pins all seventeen explicit manifests,
sixteen workspace dependency requirements, seventeen lockfile entries,
publication flags, README examples, Rust assertions, the CI WASM literal, and
the exact 15-package publication preflight at 0.12.1. It separately proves the
stable workspace remains at its prepared 0.14.0 boundary and both `rpptx-py`
and `rpptx-wasm` remain ineligible for crates.io publication.
The S73 release contract regression requires the 7 stable crates at 0.14.0,
the 15 incubating crates at 0.12.1, and both Python projects at their native
versions. It renders the `v0.14.0`, `rpptx-v0.12.1`, `py-rdocx-v0.14.0`, and
`py-rpptx-v0.12.1` notes and requires each rendered issue and pull-request set
and each credited handle to equal the reviewed contribution inventory. Every
linked record must also appear in its section's contributor credit. A reviewed
map assigns each included S71 to S73 story to the families that ship it, and
every GitHub record its backlog entry links or its AS_BUILT entry names as a
pull request must appear in each assigned family's inventory. The stable
`v0.14.0` package proof compiles packaged `rdocx` against registry-only shared
0.12.1, so `rpptx-v0.12.1` must be published first.
The completed S73 release gate verified all 15 incubating crates at 0.12.1 and
all seven stable crates at 0.14.0 under sole owner `mantissaman`, plus both
seven-file PyPI distributions under the same owner. All four annotated tags
target reviewed SHA `58ca5a279277f7cd8de0b8f250fb4650de14371b`, and every GitHub
release body is byte-identical to its reviewed changelog section. Clean Python
3.9 and 3.12 installs passed their priority suites, and Python 3.12 passed exact
mypy 2.3.0 strict checks and stubtest. All 51 reviewed contribution comments
were posted and verified before every fully addressed included record was
closed as authorized.
The Python metadata regression requires both projects to name a crate-local
Markdown README and provide their reviewed summary, author, keywords,
classifiers, and project URLs. Artifact validation repeats that check against
wheel `METADATA` and source-distribution `PKG-INFO`, including required
installation, quick-start, typing, and project-link sections in the embedded
long description.
The immutable v0.13.0 shared-family gate packages and verifies
`rdocx-layout@0.13.0`, requires its normalized archive dependency on
`oxml-layout@0.10.0` to contain no local path, and compiles the packaged crate
against the exact shared registry version without an `oxml-layout` patch. That
narrow proof did not cover the facade's newer `oxml-opc` API. Stable 0.13.1
therefore replaces it with a package-level `rdocx` proof against registry-only
shared 0.11.0. The earlier F-X068 post-publication proof against 0.8.0 remains
immutable release evidence.
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
The shared 0.11.0 release gate verified all 15 registry entries under sole owner
`mantissaman (Atul Sharma)`, immutable annotated tag `rpptx-v0.11.0` at reviewed
SHA `0b6bd622f8a14189d7d1281d011f81319ef8ad2a`, byte-identical GitHub release
notes, stable-family exclusion, and absent `rpptx-wasm@0.11.0`. Its preparation
gate pins all 16 incubating carriers, the exact publication set, source and CI
literals, and the selected notes. Its selected diff contains only the additive
`oxml-opc` Word main content-type constants required by F-238. The contribution
inventory is empty, so no notification was required.
The immutable stable 0.13.0 attempt passed its preparation gate, full
verification, and clean review at SHA
`05332b17f481741e7d5ab4e39699c6d1536475af`. Publication then stopped after
five packages because packaged `rdocx` could not find those constants in
registry `oxml-opc@0.10.0`. The stable 0.13.1 recovery gate packages and
compiles `rdocx` against registry-only shared 0.11.0 before publication.
The completed stable 0.13.1 release gate verified all seven registry entries
under sole owner `mantissaman (Atul Sharma)`, immutable annotated tag `v0.13.1`
at reviewed SHA `c391d12422c288be5db314bad8338dd08bb47d9a`, byte-identical
GitHub release notes, incubating-family exclusion, and unpublished binding and
WASM carriers. Its selected contribution inventory is empty, so no external
notification was required. Issue 69 is addressed by the completed S70 F-X084
through F-X086 mechanisms. F-X088 verified those fixes together at reviewed
S71 SHA `667416b1b54968b1524d57232c44f73a175fd27a`. All six focused
deterministic-font regressions and the complete gate passed, including 49 of 49
unchanged hash entries. The authenticated evidence comment credits
`@emptinessform`. The reporter fork does not contain a committed timing harness,
so closure evidence uses a temporary direct-engine reconstruction. It matches
the reported 700 four-line paragraphs, one 3 by 3 table every 50 paragraphs,
63-page prime, three body positions, warm cache, release mode, and deterministic
bundled fonts. Seven alternating measured rounds after warmup produce 21
samples per operation and build on one Apple M5 Max macOS 26.6.2 environment
with rustc 1.97.1. S71 min and median times in milliseconds are 13.832 and
15.310 for typing, 14.213 and 14.688 for footnote insertion, 19.928 and 22.269
for footnote deletion, 13.549 and 13.978 for Enter, 13.389 and 14.159 for merge,
and 13.360 and 13.956 for selection deletion. The corresponding v0.13.1
footnote medians are 48.157 and 52.845 milliseconds. Note insertion and deletion
change from zero paragraph-cache hits and 700 builds on v0.13.1 to 699 hits and
one build on S71. Direct model mutation on macOS excludes the reporter's
Windows editor and UI overhead, and page-layout invocations remain the full
output page count in both builds for this table workload.

The authenticated correctness and timing evidence states that v0.13.1 remains
affected while the fixes will be included in the next stable release without a
promised date. Issue 69 is closed as completed at
<https://github.com/tensorbee/rdocx/issues/69#issuecomment-5592205748>. The next
stable release contribution inventory retains Issue 69 and offered commits
`4777a74167495a5116289e1f905dfd9ad4dbe807`,
`eff0ea0c28b5eaf08180b09b58e0c0f486b7433b`,
`9e48bc86876c294b8daa314e577e84b6fcd7ac97`, and
`c8315b92857c951146fc866cd044b214194a09a8`.
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

The dated macOS 26.6.2 observation on an Apple M5 Max with rustc 1.97.1 uses
that exact release-mode invocation with one test thread. It records 31,019.1
layout pages per second at a 29.03 MiB peak and 60,058.0 PDF pages per second at
a 1.73 MiB additional peak. These values are observations. The lower throughput
floors and higher allocation ceilings remain the portable guarantees enforced
in CI.

## What CI runs

| Job | Command |
|---|---|
| changes | On pushes and pull requests, classify changed paths for the nine filtered jobs with `dorny/paths-filter` v4.0.3 pinned to reviewed commit `ceb8a2b8f2d89434be7ff52d3de7ec3738c5cc9d` |
| test | Install exact uv 0.10.2, Poppler 26.01.0, LibreOffice 26.2.5.2, and Pandoc 3.10, fetch both pinned corpora, run the pinned Pandoc texmath differential and exact locked release-mode 1,000-page performance regression with one test thread, run `python3 scripts/docx_authoring_conformance.py --public`, run `cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py` with an isolated uv cache and 8 MiB Rust test-thread stack, run the exact locked deterministic animation golden, then run `python3 scripts/golden_png_harness.py --check` |
| no-default-features | `cargo test -p oxml-layout --no-default-features` |
| wasm | Locked `wasm32-unknown-unknown` checks, `wasm-pack test --node`, and local bundler pack and fresh-install gates for `rdocx-wasm` and `rpptx-wasm` |
| prose | `python3 scripts/prose_check.py` and `python3 scripts/sync_agent_skills.py --check` |
| release-regressions | Install cargo-release 1.1.3 with its locked dependency graph, then run `python3 -m unittest scripts.test_sprint_workflow` |
| hash-harness | `python3 scripts/hash_harness.py --check` |
| presentation-fidelity | Prime locked Cargo dependencies, install exact LibreOffice 26.2.5.2 and Poppler 26.01.0 on Ubuntu 24.04, fetch the pinned corpus, run the exact locked deterministic animation golden, then run `python3 scripts/pptx_ssim_harness.py --check` |
| word-fidelity | Restore the pinned Rust cache, run `cargo fetch --locked`, fetch the pinned Word corpus, then run `python3 scripts/docx_ssim_harness.py --check` on pinned Ubuntu 24.04 LibreOffice and Poppler with its locked offline helper |
| clippy | `cargo clippy --workspace --all-targets --all-features --exclude rdocx-py --exclude rpptx-py -- -D warnings` |
| fmt | `cargo fmt --all -- --check` |
| doc | `cargo doc --workspace --no-deps --all-features --exclude rdocx-py --exclude rpptx-py` with `RUSTDOCFLAGS=-D warnings`, then `python3 scripts/readme_doctests.py` |
| package-oxml-layout | Verify the exact 24-font and six licence-and-notice-file inventory, then build and size-check the verified archive |
| msrv | Install exact uv 0.10.2, fetch both pinned corpora, then run `cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py` under Rust 1.93 with an isolated uv cache and 8 MiB Rust test-thread stack |
| python-bindings | On pull requests, build each Python package with `maturin develop --locked` in its own Python 3.12.9 environment, then run its complete pytest directory |
| supply-chain | `cargo-deny check` |
| ci-gate | Always validate that every selected filtered job succeeded and every unselected filtered job was skipped |
| python-wheels | On manual dispatch, build six cp39-abi3 wheels and one source distribution for each Python package. On a `py-rdocx-v*` or `py-rpptx-v*` tag, build, validate, and publish only the selected package's seven artifacts. Install and test every compatible built artifact in a fresh environment. |

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

The Rust CLI release-preparation contract parses `publish.yml` and both CLI
manifests. It requires the six exact native runner and target pairs, selected
family package and binary names, version and help smoke commands, exact archive
members, executable mode for tar archives, README and licence text equality
after CRLF-to-LF normalization,
the complete sorted SHA-256 manifest, and full commit pins for every external
action. It also proves the registry token exists only in the crates.io publish
job. Negative mutations remove a target, swap a family, bypass checksum
verification, remove either reviewed text comparison, start publication before
asset validation, or start release creation before publication and assets.
Each mutation must fail the contract.
The hosted matrix remains the execution proof for platforms unavailable to one
local machine.

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

The S72 publication proof is bound to reviewed SHA
`2b009243ed39ab66470d7484d490985368e865a8`. Manual build-only run
`34907492958` produced and validated both seven-file sets without publication.
Tag runs `34934221487` and `34939929652` published `rdocx 0.13.2` and
`rpptx 0.11.0` respectively through the `pypi` environment. Every live file
passed exact artifact validation. Fresh canonical-PyPI installs passed the
priority runtime suites under Python 3.9 and 3.12, and exact `mypy==2.3.0`
strict checks plus `stubtest` passed under Python 3.12. The release-note body
digests are
`3bf361a6fcc5a858d1f315f07ea766b0e60e3c0b3c7930a777e643f1bf62b728`
for rdocx and
`60fad5ee4003448082f1c14d0d7b3a5e9d159b21fa1ca7c07b0c7ac64300197f`
for rpptx.

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

The rdocx binding formatting gate sets and reopens paragraph style and
numbering, run character style, named Word highlight, and independent shading.
Its mixed-content run includes text, a tab, a page break, a field, a drawing,
a raw symbol, and trailing text, all of which retain their order after every
formatting mutation. Invalid highlight names leave the run unchanged. The
gate also runs strict typing and stub parity against a freshly installed
`cp39-abi3` wheel so the runtime properties and their nullable declarations
cannot drift apart.

The Word namespace regression matrix covers an unused unknown default on the
document root, an inherited use by an unprefixed element, unprefixed
attributes, nested different-URI and same-URI shadows, explicit undeclaration,
malformed XML, and duplicate default declarations. Native replacement proves
that a candidate is published only after successful serialization. The
compiled CLI proves the unsafe case exits without a panic or partial output.
The exact Issue 73 attachment is also replaced, saved, and reopened under the
pinned external-tool environment.

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
