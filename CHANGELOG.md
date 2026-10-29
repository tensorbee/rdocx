# Changelog

## v0.16.5

### Highlights

Word0.16.5 improves section page numbering, header and footer clearance, TOC navigation through merged tables, and comparison of documents containing existing revisions. The seven Rust packages, CLI and Python distribution share this version. Comparison and imported-format limits remain explicit below.

### Added

Existing comparison operations can retain unchanged revision-bearing owners while recording another reviewer's changes elsewhere. Existing field-update and rendering operations support decimal, lower and upper roman, and lower and upper letter section page numbers without a new public page-format authoring API.

### Fixed

- apply section page-number formats and restart or continuation values across PAGE, PAGEREF, selected footer placements and explicit field updates. Preserve physical page identity separately from displayed numbers and formatted text. Addresses [#423](https://github.com/tensorbee/rdocx/issues/423).
- clear body flow against the selected header and footer content, including paragraph spacing, tables and supported ordinary frames. Preserve written page geometry for margin-relative placement. Keep long-note fragments with their reference when space permits, retain authored note markers and special records, and advance note-only pages even when a continuation notice consumes their usable area. Addresses [#424](https://github.com/tensorbee/rdocx/issues/424) through a hardened equivalent of [PR428](https://github.com/tensorbee/rdocx/pull/428).
- preserve existing revision XML on unchanged supported owners, allocate new compared-story revision IDs without collisions, and report only newly created revisions in CLI comparison counts. Unsupported edits refuse atomically with an owner location. Addresses the bounded preservation alternative in [#425](https://github.com/tensorbee/rdocx/issues/425) through a hardened equivalent of [PR427](https://github.com/tensorbee/rdocx/pull/427).
- retain physical source navigation through vertically merged cells and accepted-view deleted rows. TOC rebuild skips proven omitted content across the supported selectors while keeping placed headings and stale-source checks. Addresses [#426](https://github.com/tensorbee/rdocx/issues/426) through the adapted navigation work in [PR428](https://github.com/tensorbee/rdocx/pull/428).
- update pulldown-cmark to0.13.5 for RUSTSEC-2026-0414. The public Markdown importer continues to leave wikilinks disabled. This dependency update does not establish exposure through that disabled parser option.

### Compatibility

No public signature migration is required. Word packages now require shared OOXML0.14.5. Retaining old revisions is bounded owner preservation, rather than unrestricted native Word Compare behavior. Edits, moves or deletion touching a revision-bearing paragraph, whole table or block control, or changed revised section properties remain refused. For those cases, retain the original redline separately, resolve the relevant author or revision IDs on independent copies with the checked APIs, then compare those baselines. Resolving those copies discards the resolved tracked markup.

Roman formatting remains bounded to1through3999 and alphabetic formatting to4096characters. Unsupported imported number formats and chapter-dependent numbering remain preserved and diagnostic. Native Word observations outside these bounds do not extend support. Imported notes without a source reference marker remain unmarked. Newly authored add_footnote content includes an explicit styled marker, matching the existing note creation operations. Ordinary frames exclude drop caps, and exact tbRl body corrections do not promise general vertical-text parity. General native Word layout equivalence is not claimed.

### Contributors

- [@set2374](https://github.com/set2374) reported the section PAGE-format behavior in [#423](https://github.com/tensorbee/rdocx/issues/423).
- [@hadim](https://github.com/hadim) reported header clearance, existing-revision comparison and merged-cell TOC failures in [#424](https://github.com/tensorbee/rdocx/issues/424), [#425](https://github.com/tensorbee/rdocx/issues/425) and [#426](https://github.com/tensorbee/rdocx/issues/426), and supplied [PR427](https://github.com/tensorbee/rdocx/pull/427) and [PR428](https://github.com/tensorbee/rdocx/pull/428). Their reference implementations were adapted and hardened with additional owner, source, pagination and refusal checks.

## rpptx-v0.14.5

### Highlights

PowerPoint0.14.5 aligns the fifteen shared OOXML and PowerPoint Rust packages, CLI and Python distribution for the post-S96 release. This family adds no new PowerPoint product feature in S96.

### Added

No new PowerPoint authoring or rendering API is included in this release.

### Fixed

No new PowerPoint-specific user-facing fix is included. Word comparison, page-number and layout repairs belong to the separate Word0.16.5 family.

### Compatibility

No public signature migration is required. The selected shared OOXML and PowerPoint packages and Python distribution use0.14.5 together. Existing rendering limitations remain. The separate unpublished rpptx-wasm carrier stays0.12.1, and npm publication is outside this release.

### Contributors

No new external PowerPoint issue or pull-request contribution is included in the reviewed S96 range. Credits in earlier published sections remain with their original releases.

## v0.16.4

### Highlights

Word0.16.4 adds bounded legacy fields, atomic form operations and mail-merge package authoring, and addresses the qualified S95.4 correctness and access-cost reports. The seven Rust packages, CLI and Python distribution share this version. The approved shared-buffer Rust source break and remaining qualification limits are explicit below.

### Added

- author legacy text, checkbox and drop-down form fields with separate current values and reset defaults, help/status text and typed options. Add default-value queries without changing the existing public metadata structs. Imported ordinary nested TIME, DATE and formula results remain bounded by the approved typed projection and retained-source rules.
- validate, fill, reset and protect Word forms across the qualified content-control and legacy-field owners. Operations validate the complete candidate before publication. Rules are caller inputs, not a stored second form model.
- author and inspect mail-merge package settings, canonical Word field-role mappings and recipient metadata. An offline rdocx custom-XML table can supply records. External source declarations and connection settings are stored without fetching files, HTTP or SQL. Word performs its own connector execution.

### Fixed

- preserve native legacy checkbox ownership when filling an empty result cache. A dedicated result run keeps the own-name bookmark boundary valid, and setter reopen checks reject a shifted neighboring field.
- write numeric line multiples consistently through defaults, paragraph formatting and styles, including exact stored-value round trips. List cross-cell comments in physical paragraph order and diagnose an individual unrepresentable anchor without losing the remaining comments. `part_names` and `part_xml` expose relationship parts and `[Content_Types].xml` through the same serialization and validation as saving. Addresses [#402](https://github.com/tensorbee/rdocx/issues/402), [#405](https://github.com/tensorbee/rdocx/issues/405) and [#406](https://github.com/tensorbee/rdocx/issues/406).
- preserve imported and returned picture relationship targets through staged edits and save/reopen. Resolve duplicate imported style IDs by their first definition, without redundant renamed copies for equal definitions or importing shadowed dependencies. Addresses the available [#404](https://github.com/tensorbee/rdocx/issues/404) cases and [#408](https://github.com/tensorbee/rdocx/issues/408).
- reduce repeated body scanning, package copying and normalization in qualified Word reads and edits while retaining current-owner, namespace, private-field and resource validation. Add checked Rust `try_add_paragraph`, `try_insert_paragraph`, `try_add_table` and `try_insert_table` paths. Python insertion uses those checked paths. Addresses [#407](https://github.com/tensorbee/rdocx/issues/407) and the captured Word follow-up performance contributions.
- use the bounded CJK font-metric policy derived from the font's OS/2 code-page bits. This does not install fonts or claim general native Word layout equivalence.

### Compatibility

- Python 3.9 remains supported for constructor calls and defaults. With the stable-ABI wheel, native class constructor signatures are available to `inspect.signature` on CPython 3.10 or newer. On CPython 3.9, pinned PyO3 exposes no class text signature, so constructor introspection raises `ValueError`. No constructor, stub or class identity changes.

- **Approved Rust source break in both prepared version lines:** `oxml_opc::OpcPackage.parts` now stores `Arc<Vec<u8>>`, and `remove_part` returns `Option<Arc<Vec<u8>>>`. The Word `rdocx_opc` re-export exposes the same change. Migrate direct map writes and owned removal results as described below. `set_part` accepts both `Vec<u8>` and `Arc<Vec<u8>>` through `Into`, and `get_part` still returns borrowed bytes. Python signatures are unchanged. These releases retain the explicitly requested 0.16.x and 0.14.x lines and do not claim unrestricted Rust source compatibility.

The approved Rust source break affects both `oxml_opc::OpcPackage` and its `rdocx_opc` re-export. Existing owned `set_part` calls continue to work:

```rust
use std::sync::Arc;

package.set_part("/word/example.xml", xml.into_bytes());

// Direct public-map insertion now needs shared storage.
package.parts.insert("/word/example.xml".to_owned(), Arc::new(bytes));

// Preserve sharing when reusing an already shared payload.
package.set_part("/word/example-copy.xml", Arc::clone(&shared_bytes));

// Removal returns shared storage. Clone the Vec only if ownership is required.
let shared = package.remove_part("/word/example-copy.xml").unwrap();
let owned_bytes: Vec<u8> = shared.as_ref().clone();
```

Readers can keep using `get_part`, which still returns `Option<&[u8]>`. Direct mutation of a retained part's bytes requires copy-on-write, such as `Arc::make_mut`, rather than mutating through a shared reference. Python signatures are unchanged.

- Bounded legacy form creation, form operations and mail-merge package authoring are included. Fillable PDF forms remain outside this release. Mail-merge custom XML is an offline rdocx table, not a claim that Word treats it as a native embedded database. External connectors remain declarations for Word to execute. Offline validation reports unsupported source or mapping states.
- Non-bool Python integers and floats mean line multiples in the affected formatting APIs. `Length` means exact spacing. Omitted values preserve formatting and `None` clears it. Imported `atLeast` spacing stays preserved until explicitly assigned. This changes raw integer interpretation in style calls.
- Save-equivalent part inspection can fail when the package cannot be serialized. Legacy infallible Rust body-insertion methods retain their existing behavior. Use the additive checked methods for atomic refusal.
- The original unreduced #404 picture-swap input was unavailable. Available reproduction cases are covered, but that original input remains unverified. The reporter should retest the original file after release and reopen if necessary.
- Performance work uses qualified workloads and safety fallbacks. It is not a promise that every public call is constant-time. Retained-data caches trade memory for speed. The [S95.4 delivery record](docs/sprints/AS_BUILT.md) retains measured timing, memory and retained-history results, including residual expensive calls and fallback costs.
The private unreduced Issue 390 document remains unqualified. Its reduced cases and constructed image-reinsert controls are included, so the report is partially addressed. Custom XML store removal, general XPath, citation/equation/bibliography/document-part control authoring, Python 3.14 and npm publication are excluded. Default paragraph line-spacing reads stored auto values as multiples and exact/atLeast values as Length. Assigning Length writes an exact rule, an omitted keyword preserves its element, and None clears only the selected property. The shared 0.14.4 family must precede Word publication where the internal pins require it.

### Contributors

- S95.4 thanks **Hadrien Mary ([@hadim](https://github.com/hadim))** for Issues [#402](https://github.com/tensorbee/rdocx/issues/402), [#404](https://github.com/tensorbee/rdocx/issues/404), [#405](https://github.com/tensorbee/rdocx/issues/405), [#406](https://github.com/tensorbee/rdocx/issues/406), [#407](https://github.com/tensorbee/rdocx/issues/407) and [#408](https://github.com/tensorbee/rdocx/issues/408), and assessed PRs [#409](https://github.com/tensorbee/rdocx/pull/409), [#410](https://github.com/tensorbee/rdocx/pull/410), [#411](https://github.com/tensorbee/rdocx/pull/411), [#412](https://github.com/tensorbee/rdocx/pull/412), [#414](https://github.com/tensorbee/rdocx/pull/414), [#415](https://github.com/tensorbee/rdocx/pull/415), [#417](https://github.com/tensorbee/rdocx/pull/417), [#418](https://github.com/tensorbee/rdocx/pull/418), [#419](https://github.com/tensorbee/rdocx/pull/419), [#420](https://github.com/tensorbee/rdocx/pull/420), [#421](https://github.com/tensorbee/rdocx/pull/421), [#422](https://github.com/tensorbee/rdocx/pull/422). Shared OPC work from PRs [#413](https://github.com/tensorbee/rdocx/pull/413) and [#416](https://github.com/tensorbee/rdocx/pull/416) also benefits Word. Contributor inputs were adapted and hardened to the approved local contracts, rather than represented as unchanged merges.

## rpptx-v0.14.4

### Highlights

PowerPoint0.14.4 bounds the qualified mutation workloads and repairs the named native animation and SmartArt cases. The fifteen shared OOXML and PowerPoint Rust packages, CLI and Python distribution share this version. The approved shared-buffer Rust source break and remaining qualification limits are explicit below.

### Added

- add `Effect::LinearOpacityMask` to the existing non-exhaustive Rust effect enum, with one shared validation query and raster/SVG composition. Invalid masks produce transparent output in those renderers.

### Fixed

- share unchanged package bytes and commit qualified edits in memory while refreshing the affected owner inventories. Preserve full checks for signed/embedded cases and operations that require shared contexts or whole-deck reports. Addresses [#403](https://github.com/tensorbee/rdocx/issues/403).
- Correct earlier visibility ownership when a later entrance is still pending, the four qualified native wipe directions, and native `ScaleX`, `ScaleY` and `style.rotation` numeric properties while retaining the legacy aliases. Scale materialized custom geometry and its text rectangle before page-clip projection, without double-scaling dynamic rectangle geometry.
- Correct the six qualified SmartArt family profiles and text owners, including native cycle and pyramid text frames, list insets, hierarchy/relationship sizing, matrix sizing and diagram text defaults. Explicit source formatting remains authoritative.

### Compatibility

- **Approved Rust source break in both prepared version lines:** `oxml_opc::OpcPackage.parts` now stores `Arc<Vec<u8>>`, and `remove_part` returns `Option<Arc<Vec<u8>>>`. The Word `rdocx_opc` re-export exposes the same change. Migrate direct map writes and owned removal results as described below. `set_part` accepts both `Vec<u8>` and `Arc<Vec<u8>>` through `Into`, and `get_part` still returns borrowed bytes. Python signatures are unchanged. These releases retain the explicitly requested 0.16.x and 0.14.x lines and do not claim unrestricted Rust source compatibility.

The approved Rust source break affects both `oxml_opc::OpcPackage` and its `rdocx_opc` re-export. Existing owned `set_part` calls continue to work:

```rust
use std::sync::Arc;

package.set_part("/word/example.xml", xml.into_bytes());

// Direct public-map insertion now needs shared storage.
package.parts.insert("/word/example.xml".to_owned(), Arc::new(bytes));

// Preserve sharing when reusing an already shared payload.
package.set_part("/word/example-copy.xml", Arc::clone(&shared_bytes));

// Removal returns shared storage. Clone the Vec only if ownership is required.
let shared = package.remove_part("/word/example-copy.xml").unwrap();
let owned_bytes: Vec<u8> = shared.as_ref().clone();
```

Readers can keep using `get_part`, which still returns `Option<&[u8]>`. Direct mutation of a retained part's bytes requires copy-on-write, such as `Arc::make_mut`, rather than mutating through a shared reference. Python signatures are unchanged.

- Native soft-wipe qualification covers direct leaf owners and the isolated plain-paint fixture. Group-target timeline wipes retain the existing hard-clip fallback with a diagnostic. The neutral mask composes an existing group scene once. Caller clips remain authoritative.
- Raster and SVG support the new linear opacity mask. Ordinary PDF retains its existing unsupported-effect behavior and checked PDF conformance refuses it. This release does not claim soft-wipe PDF parity, general shadow fidelity or expanded paint-domain support. Held entrance-wipe composition with later opacity remains outside the qualified interaction set.
- Custom-path emphasis scaling retains existing font, inset, stroke, shadow and precomputed-paint semantics. Qualified geometry repairs do not establish general native emphasis fidelity. Native profile claims must name the authenticated application/source profile and completed gates, rather than generalize across PowerPoint versions.
- Performance acceptance remains tied to the pinned 248-call catalogue and original 10/40/80-slide workloads. Costly correctness fallbacks and by-design whole-deck operations remain disclosed. The [S95.4 delivery record](docs/sprints/AS_BUILT.md) retains the normal-artifact measurements and remaining slow calls.

### Contributors

- S95.4 thanks **Hadrien Mary ([@hadim](https://github.com/hadim))** for [Issue #403](https://github.com/tensorbee/rdocx/issues/403) and assessed PRs [#413](https://github.com/tensorbee/rdocx/pull/413) and [#416](https://github.com/tensorbee/rdocx/pull/416). The shared package layer also supports the Word follow-ups credited in v0.16.4. These are adapted contribution inputs, not claims of unchanged PR merges.

## v0.16.3

### Highlights

Word 0.16.3 completes the bounded navigation, template and content-control work since v0.16.2, adds custom XML store authoring, and repairs reported owner-lifetime, package-preservation and access-cost defects. The exact seven Rust packages, CLI and `rdocx` Python distribution share this version. Approved Rust source compatibility changes and native Word differences are described below.

### Added

- Rust and Python author rich/plain content-control shells at qualified locations, edit metadata, move, clone, remove and unwrap them through checked transactions. Typed payloads cover dates, checkboxes, pictures, combo/drop-down lists, groups and repeating sections/items. Unsupported placements and invalid payloads refuse atomically. WASM and CLI gain no control-creation surface.
- Rust and Python create custom XML stores with properties, relationships and content types, bind existing controls through the supported absolute-child-path subset, and update the addressed store. Schema references are stored without fetching. Unknown store bytes remain verbatim.
- Numbering-aware REF, STYLEREF, PAGEREF, captions, outlines and TOC rebuilding share source-qualified numbering context. Explicit TOC heading ranges also bound outline-derived entries.
- Opt-in template grammar version 1 supports the approved paths, filters, escaping and container expansion across existing stories. Legacy calls retain their grammar. Empty and populated implicit scopes follow the same lookup policy.
- Read-only chart snapshots expose the qualified direct-body occurrences and finite data families. Optional display equations, persisted list restart/start and marker-font selection, plus note reference styles and current-id removal, complete the selected everyday authoring gaps.
- Python adds ordered custom-property iteration/lookup, read-only package part names/XML and default paragraph spacing access. Missing, binary and unsupported part encodings produce explicit diagnostics.

### Fixed

- Held Word paragraph/run handles survive resolvable table and cell property edits. Cached reads stay current after tab/break insertion and mirrored edits, and read/write lookup selects the same physical owner inside mixed controlled content.
- Named read-only handle operations and structural edits avoid unrelated package inventory scans. Half-open direct-body range removal publishes one transaction and returns the removed count.
- Style edits retain opaque root children in schema order. Image pruning keeps shared and reinserted resources, plain embedded font bytes stay intact, header validation names wrong targets, and qualified alternate text-box edits publish together or not at all.
- Comment resolution preserves the three states of true, false and unknown. Checked XML retains admitted namespaces and unchanged opaque sections while refusing invalid newly authored placement and malformed custom XML. DTD declarations are explicitly refused.
- Nested control relocation, unwrapping and repeating-section admission inspect decoded text and the actual destination content model before publication. Repeating types also retain their permitted physical level and section/item ownership across type changes and unwrap. Imported type-marker aliases retain namespace identity through modeled edits, including locally shadowed canonical prefixes.

### Compatibility

`rdocx_oxml::document::BodyContent` now boxes the `Paragraph`, `Table`, `ContentControl` and `RawXml` payloads. Wrap downstream constructors in `Box::new`, and account for boxes when consuming owned matches. `CommentRef::resolved` is `Option<bool>`, with Python `bool | None`. Missing extension metadata is unknown, not false. `ResolvedNumbering` also has the documented hidden definition/start-override metadata. Patch numbering does not imply unrestricted Rust source compatibility.

Repeating controls intentionally append a clone of the first item and refuse removal of the last item. Native Word defaults instead clone the addressed item and delete the whole section when its last item is removed. The approved API retains its explicit behavior. Authentic Word for Mac 16.114 (26101123) qualification compares equivalent native operations and retains both default differences. Native opening/saving and individual XML normalizations are distinct from whole-page rendering parity.

Native embedded-font qualification establishes genuine proper Bold-only use through initial, saved and reopened phases. Paired Italic can use Regular outlines with synthetic slant, and Word can empty the Italic part when saving. This is a documented native limitation, without a speculative writer workaround or a claim of general multi-face interoperability.

The private unreduced Issue 390 document remains unqualified. Its reduced cases and constructed image-reinsert controls are included, so the report is partially addressed. Custom XML store removal, general XPath, citation/equation/bibliography/document-part control authoring, later form work, Python 3.14 and npm publication are excluded. Default paragraph line-spacing reads stored auto values as multiples and exact/atLeast values as Length. Assigning Length writes an exact rule, an omitted keyword preserves its element, and None clears only the selected property. The shared 0.14.3 family must precede Word publication where the internal pins require it.

### Contributors

- [@hadim](https://github.com/hadim), Hadrien Mary, reported the remaining everyday compatibility groups in [Issue 385](https://github.com/tensorbee/rdocx/issues/385) and the font work underlying the bounded native qualification in [Issue 360](https://github.com/tensorbee/rdocx/issues/360). Earlier shipped APIs are not counted as new outcomes here.
- Hadrien reported custom-property iteration in [Issue 387](https://github.com/tensorbee/rdocx/issues/387), handles in [Issue 388](https://github.com/tensorbee/rdocx/issues/388), structural-edit costs in [Issue 389](https://github.com/tensorbee/rdocx/issues/389), package preservation in [Issue 390](https://github.com/tensorbee/rdocx/issues/390), and source/API gaps in [Issue 391](https://github.com/tensorbee/rdocx/issues/391). He supplied candidate [PR 393](https://github.com/tensorbee/rdocx/pull/393), [PR 395](https://github.com/tensorbee/rdocx/pull/395), [PR 396](https://github.com/tensorbee/rdocx/pull/396) and [PR 397](https://github.com/tensorbee/rdocx/pull/397). The integrated outcomes use reviewed adaptations and repairs, not an assertion that a candidate was merged unchanged.
- Hadrien reported read-only handle costs in [Issue 398](https://github.com/tensorbee/rdocx/issues/398) with candidate [PR 399](https://github.com/tensorbee/rdocx/pull/399), and TOC outline range behavior in [Issue 400](https://github.com/tensorbee/rdocx/issues/400) with candidate [PR 401](https://github.com/tensorbee/rdocx/pull/401). The accepted implementation preserves the approved ownership, timing and TOC selector contracts.
- Atul Sharma maintained the integration, qualification and release preparation. Publication retains separate approval and verification.

## rpptx-v0.14.3

### Highlights

PowerPoint 0.14.3 repairs reported authoring, handle and edit-cost defects, improves table text placement, and completes the selected checked-XML/value/CLI gaps since rpptx-v0.14.2. The exact fifteen shared OOXML and PowerPoint Rust packages, CLI and `rpptx` Python distribution share this version.

### Added

- Python TextFrame.clear retains one paragraph and its properties, drops its runs and retires those held runs while unrelated owners remain live. Bullet size distinguishes percentage floats, Length point values and clearing with None.
- Table-cell run hyperlinks use the slide relationship scope. Shapes and connectors support checked removal of their style element, while unsupported kinds refuse atomically.
- The regenerated preset table includes upArrow and the matching UP_ARROW enum value.

### Fixed

- Appending paragraphs, runs and fields preserves earlier live handles without silently refreshing previously materialized collections. Named scalar and structural edits avoid unrelated slide/package inventory walks.
- Slide placeholders iterate by ascending idx, including gaps. Layout and master placeholders retain tree order, and indexing remains lookup by idx.
- Transparency, gradient position, percentage bullet size and font baseline use the pinned ties-to-even conversion. Unrelated EMU, twips, time and adjustment truncation remains unchanged.
- Painted table rows grow for the approved text, span, margin and vertical-text cases so following rows do not overlap. Stored row heights and frame dimensions are preserved.
- Checked mutation refuses unsupported table/chart-frame flips, invalid slide dimensions, unknown styles, wrong output formats and malformed finite XML content models before publication. The CLI refuses no-match replacement-map operations under the documented expectation contract, and missing-family font removal produces an actionable refusal.

### Compatibility

Rust `ValidationIssue` gains `InvalidXmlChild`. Update downstream exhaustive matches. Slide placeholders change their iteration order, and the percentage tie rule changes intentionally. Length-valued bullet reads reflect the stored element, and invalid values or booleans refuse. A previously materialized collection is not automatically refreshed after an append.

Table growth is an intentional rendering change, not a rewrite of document row geometry. The pinned LibreOffice corpus comparison renders all 421 slides without dropped bounded shapes, while its SSIM trend remains below the advisory target. This is not a claim of full PowerPoint conformance or pixel parity. Python 3.14 and npm publication remain outside qualification. The unpublished `rpptx-wasm` carrier stays 0.12.1 with its separate npm authority. Shared/PowerPoint publication precedes Word where the prepared pins require it, with separate final approval for each family.

### Contributors

- [@hadim](https://github.com/hadim), Hadrien Mary, reported the selected remaining PowerPoint value, XML and CLI gaps in [Issue 385](https://github.com/tensorbee/rdocx/issues/385).
- Hadrien reported handle lifetime in [Issue 388](https://github.com/tensorbee/rdocx/issues/388), structural-edit costs in [Issue 389](https://github.com/tensorbee/rdocx/issues/389) and the authoring/rendering cases in [Issue 392](https://github.com/tensorbee/rdocx/issues/392). Candidate [PR 393](https://github.com/tensorbee/rdocx/pull/393), [PR 394](https://github.com/tensorbee/rdocx/pull/394) and [PR 397](https://github.com/tensorbee/rdocx/pull/397) supplied inputs to the reviewed integrated work. Candidates were not merged unchanged, and archive-measurement-only commits are not counted as product features.
- Atul Sharma maintained the integration, qualification and release preparation. Publication retains separate approval and verification.

## v0.16.2

### Highlights

Word 0.16.2 brings readable tracked comparisons, legal revision rendering, stronger formula layout and the bounded Python, signing and form repairs completed since v0.16.1. The exact seven Rust packages, CLI and `rdocx` Python distribution share this version. This patch release includes approved Rust and Python source compatibility changes described below.

### Added

- Legal revision styling selects blue insertions, red deletions and green moves through Rust, Python and CLI options. Plain remains the default, and stored formatting is unchanged.
- Python adds supported heading, break, paragraph, run and stream aliases, typed section enums and actionable alternatives where native semantics differ from python-docx. Default-template List Bullet and List Number still require the explicit numbering APIs. One-dimension picture sizing still names the supported two-dimension call.
- Python and `rdocx sign` / `verify-signatures` expose existing native package signing with DER or strict PEM inputs. Reports inspect the current staged package, including fill-then-sign and sign-then-fill operations. This is OOXML package signing, not PDF signing.
- Callable form snapshots, atomic value binding and `rdocx fill --values` cover supported plain-text, checkbox, drop-down and date controls plus legacy text fields. Typed SDT fills remove the qualified showingPlcHdr flag, write supported date metadata, resolve declared item values/display strings and refuse missing individual selectors.
- Python CustomProperties supports item assignment, deletion and membership with strict string keys. Section and TabStop snapshots retain immutable fields with live section or delete/re-add alternatives. Private paragraph and run XML names point to checked XML providers, and legacy drop-down selection names its integer index contract.
- Native nullable run width-scale and position setters clear their direct overrides. Native Word and Python chart legend controls expose visibility, position and overlay through the shared renderer.

### Fixed

- Comparisons preserve complete hyperlinks and chart/workbook resources, handle table-adjacent, empty and trailing paragraphs, and group consecutive word replacements without losing equal anchors or source ownership. Unsupported nested hyperlink cases remain explicit refusals.
- Paragraph border thickness and spacing participate in flow, joined/split borders and table reconstruction. The missing original border file was qualified through the approved inferred 15-point and `w:sz="15"` controls, not claimed as an exact replay of an unavailable input.
- Finite LaTeX conversion and selected-face ink/MATH metrics improve the reported symbols, fractions, scripts, radicals, integrals and baselines. The 73-input follow-up retains its independently observed degree fallback and unsupported diagnostics. This is not unrestricted TeX or whole-page native font parity.
- Repeated owner-inventory work during Python table access and scalar edits is reduced. The measured repair is not a universal timing or total cell-fill complexity promise. Six measured Rust inline annotations retain existing setter bodies and signatures.
- Identified paragraphs inside block controls validate without stripping `w14:paraId` or `w14:textId`. Newly invalid direct paragraph XML refuses atomically while supported inherited namespace/MC authority, original opaque exceptions and unchanged no-ops retain their bounded source-owned checks.
- Python tiny bare paragraph space-after and character-spacing values refuse before zero quantization. Integer and float `line_spacing` values denote line multiples, while explicit Length values denote absolute spacing. Public `update_section` requires Length for every nonzero geometry argument, with bare zero retained and the live section EMU route unchanged.
- DOCX save to a PDF path, duplicate stored tab stops, header page breaks, malformed language tags and authored fonts without a valid name table refuse before publication. Valid streams, ordinary fonts, language tags and body breaks remain supported. Word font CLI errors no longer suggest an unavailable `--force`.

### Compatibility

Update exhaustive Rust literals for `RenderOptions.revision_style` and `LayoutInput.revision_style`. `ParagraphBlock.has_visible_revision: bool` becomes `change_bar: Option<Color>`. `NoteLayout.revision_ranges` contains `(Range<usize>, Color)` pairs. Select `RevisionStyle::Plain` for the default paint style. These approved source breaks remain real despite the 0.16.x version line.

ZIP 9.0.3 changes the upstream error type carried by `OpcError::Zip`. Native signing features and WASM exclusions retain their existing boundaries. Python `font.size` now refuses nonzero integer EMU inputs that would store zero and names explicit units. New refusals do not redefine the truncating unit constructors. Python 3.14 is not qualified by the 3.9/3.12 release matrix.

Python source compatibility changes include `ColorFormat.type` returning `rdocx.enum.dml.MSO_COLOR_TYPE | None` instead of strings. Compare against enum members. Word `theme_color` remains `str | None`, and supported string setters, native color math and stored formatting remain unchanged.

The completed Issue 386 numeric, mapping, name and diagnostic repairs retain the documented default-list and two-dimension-picture alternatives. Whole-issue acceptance and closure require the actual final paired untouched 19-attempt reporter replay and independent 16-group semantic companion. A self-catching script exit zero alone is insufficient.

Prior intentional comparison grouping, legal/double-underline paint, border flow and legend layout changes remain attributed to their reviewed repairs. The S93 union changed eighteen hash entries and four pixel buffers. Version preparation introduces no further paint change. Native formula/border observations have the limits above. Issue 385 is partially addressed: chart read-back, display equations, list restart/font and note APIs plus remaining PowerPoint admission/CLI repairs remain in S95.11. Bibliography, navigation, fillable PDF, full-range watermark and Word multi-face native interoperability are excluded.

### Contributors

- [@Andrewkoro105](https://github.com/Andrewkoro105) reported the nullable setter inconsistency in [Issue 364](https://github.com/tensorbee/rdocx/issues/364), formula rendering in [Issue 365](https://github.com/tensorbee/rdocx/issues/365), borders in [Issue 366](https://github.com/tensorbee/rdocx/issues/366) and measured setters in [Issue 367](https://github.com/tensorbee/rdocx/issues/367).
- [@hadim](https://github.com/hadim), Hadrien Mary, supplied the comparison reports [368](https://github.com/tensorbee/rdocx/issues/368), [369](https://github.com/tensorbee/rdocx/issues/369), [370](https://github.com/tensorbee/rdocx/issues/370), [371](https://github.com/tensorbee/rdocx/issues/371) and [372](https://github.com/tensorbee/rdocx/issues/372), plus [PR 374](https://github.com/tensorbee/rdocx/pull/374). Selected donor slices and reviewed hardened equivalents landed. The closed PR was not merged unchanged.
- Hadrien reported legal revision styling in [Issue 373](https://github.com/tensorbee/rdocx/issues/373) and supplied [PR 375](https://github.com/tensorbee/rdocx/pull/375). Its adapted styling and source-preservation work retain the approved compatibility boundaries. The closed PR was not merged unchanged.
- Hadrien supplied [PR 376](https://github.com/tensorbee/rdocx/pull/376), adapted into job-local test optimization and bounded corpus/dependency routing without shortening acceptance suites. He reported table owner costs in [Issue 378](https://github.com/tensorbee/rdocx/issues/378), Python compatibility in [Issue 379](https://github.com/tensorbee/rdocx/issues/379), signing/forms in [Issue 380](https://github.com/tensorbee/rdocx/issues/380), chart-resource comparison in [Issue 382](https://github.com/tensorbee/rdocx/issues/382) and legend rendering in [Issue 383](https://github.com/tensorbee/rdocx/issues/383).
- Hadrien supplied the control-validation reproduction in [Issue 384](https://github.com/tensorbee/rdocx/issues/384) and the typed-fill, invalid-value and raw-XML groups of [Issue 385](https://github.com/tensorbee/rdocx/issues/385). Only the named groups are included here. The issue remains partially addressed.
- Hadrien reported the numeric, mapping, compatibility-name, color and legacy drop-down cases in [Issue 386](https://github.com/tensorbee/rdocx/issues/386). The implemented Word repairs and supported alternatives are described above.
- Atul Sharma maintained the release preparation and hardened integration. Publication retains its separate approval and verification.

## rpptx-v0.14.2

### Highlights

The exact fifteen shared OOXML and PowerPoint Rust packages, CLI and `rpptx` Python distribution prepare 0.14.2. Imported chart subtypes and master visibility now have independent native/LibreOffice controls, alongside bounded notes, signing and legend APIs.

### Added

- Python notes use `has_notes_slide` and owner-bound `notes_slide.notes_text_frame`. The existing color enum module adds `MSO_COLOR_TYPE` and the identity alias `MSO_THEME_COLOR_INDEX`.
- Python and CLI package signing/verification accept DER and strict PEM inputs, preserve staged package state and return owned reports. The approved ZIP 9.0.3 dependency is shared with Word. This adds callable surfaces over existing native signing, not PDF signing or AcroForm authoring.
- Python exposes registered slide IDs, checked-Element alternatives for private shape/text-body names and the 216-member `MSO_LANGUAGE_ID` module. `Font.language_id` maps declared language tags, with `None` or `NONE` clearing the direct override and existing `Font.language` string admission retained. Unsupported imported tags remain preserved and require explicit source inspection.
- Native and Python chart legend controls expose visibility, position and overlay. Shared drawing honors the corresponding plot-space reservation with absent overlay false.

### Fixed

- Imported stacked horizontal bars, stacked columns, marker-free lines and exploded pies report exact chart types 58, 52, 4 and 69. Valid unsupported forms return `None`, while invalid metadata refuses. Data, plot order and opaque chart XML survive save/reopen.
- Hidden master shapes are suppressed consistently in PDF, PNG, `convert` and `render`. Actual PowerPoint and LibreOffice controls distinguish master, layout and slide shapes, placeholders, styles and backgrounds. Master paint requires both slide and layout flags, while layout paint follows the slide flag.
- Four PowerPoint zero-value spacing assignments now refuse before mutation and name explicit units or positive line multiples. Existing valid zero before/after spacing, typed lengths and supported negative character spacing retain their contracts.
- Native presentation HTML import preserves whitespace, links, rich span edges and nonbreaking spaces. This is the new PowerPoint increment from the existing reference contribution, not a new Python or CLI HTML importer.
- Job-local test optimization and reviewed corpus/dependency routing retain the full acceptance scope without a universal timing claim.

### Compatibility

Python chart types are `int | None`. Python source compatibility changes include `ColorFormat.type` returning PowerPoint's `MSO_COLOR_TYPE | None` and `theme_color` returning `MSO_THEME_COLOR` instead of strings. A direct nonscheme color returns `NOT_THEME_COLOR`, absent direct color raises `AttributeError`, and unknown or `phClr` scheme tokens refuse projection with preserved-source inspection guidance. Existing supported string setters and native color math remain intact. The unpublished `rpptx-wasm` carrier stays 0.12.1 with its separate npm authority.

The completed Issue 386 color, language and name repairs retain checked owner identities. Whole-issue acceptance and closure require the actual final paired untouched 19-attempt reporter replay and independent 16-group semantic companion. A self-catching script exit zero alone is insufficient.

ZIP 9.0.3 changes the upstream `OpcError::Zip` payload type. Word's approved rendering-structure changes are documented in its separate 0.16.2 notes. Shared double-underline and legend layout corrections are intentional prior rendering changes, with no new paint delta from version preparation. The native empty-legend disagreement and font/whole-page parity limits remain explicit.

Issue 385 is partially addressed by the chart-type and master-render groups. Wrong-format saves, dimensions/flips/styles, checked PowerPoint XML, replacement-map and remaining font CLI repairs belong to S95.11. No full Issue 385 closure, full-range watermark interoperability, Word multi-face native interoperability, fillable PDF, Python 3.14 qualification or npm publication is claimed. Shared/PowerPoint publication precedes Word where its internal pins require it, with separate final approval per family.

### Contributors

- [@hadim](https://github.com/hadim), Hadrien Mary, supplied [PR 359](https://github.com/tensorbee/rdocx/pull/359). Only its new presentation HTML whitespace increment is credited here. The earlier Word import outcome already shipped.
- Hadrien supplied [PR 376](https://github.com/tensorbee/rdocx/pull/376), adapted into the bounded test/corpus routing work. The closed PR was not merged unchanged.
- Hadrien reported legend rendering in [Issue 383](https://github.com/tensorbee/rdocx/issues/383), notes/enum names and silent unit loss in [Issue 379](https://github.com/tensorbee/rdocx/issues/379), and Python/CLI signing in [Issue 380](https://github.com/tensorbee/rdocx/issues/380).
- Hadrien supplied the imported-chart and master-logo controls in [Issue 385](https://github.com/tensorbee/rdocx/issues/385). The remaining groups are excluded from this release's acceptance claim.
- Hadrien reported the PowerPoint name, color and language cases in [Issue 386](https://github.com/tensorbee/rdocx/issues/386). The implemented repairs are described above.
- Atul Sharma maintained the release preparation and hardened integration. Publication retains its separate approval and verification.

## v0.16.1

### Highlights

Word gains fuller Python authoring, safe checked XML owners, richer one-shot CLI operations and public HTML/Markdown import. The Rust crates, CLI and Python distribution share version 0.16.1.

### Added

- Python exposes complete run, paragraph, list, style, table, picture, internal-link, section, furniture, note and document operations through checked native providers. Physical owners and related package parts retain their source identities.
- Document.from_html and Document.from_markdown expose the supported import subsets, preserve semantic mark/code content and report unsupported input through diagnostics.
- The CLI adds ordered replacement maps, checked fields, physical image extraction, metadata, tag filling, and accepted-body inspection. Python can create Word charts.
- Checked paragraph, run, table, cell and section XML operations preserve opaque bytes, validate resource and relationship boundaries, and publish atomically. Licensed embedded font listing, removal and settings are available through native, Python and CLI surfaces.
- The unchanged 100-script compatibility corpus qualifies 95 correct files, with all 24 documentation cases and required XML, lifetime, native and rendering checks qualified. Five explicit divergences retain their documented boundaries.

### Fixed

- Nested HTML mark highlights survive rich inline parsing. Picture-only paragraphs no longer produce false empty-content diagnostics, and native accessibility follows visible accepted content.
- Scoped text-box replacements synchronize supported alternate copies. Table empty margins, signed indentation, physical picture ownership and source formatting survive save/reopen.
- Markdown export retains emphasis whitespace and table/list boundaries. New compatible documents explicitly name compatibility mode15. TOC rebuild preserves entry boundary paragraphs, and story ranges avoid repeated whole-source scans while valid empty comment owners remain accepted.
- Caller fonts reach supported render/export paths, variable defaults are respected, and deterministic metric-compatible bundled faces and qualified bullet fallbacks have explicit source/glyph evidence.

### Compatibility

Font.color exposes a live ColorFormat rather than the preceding scalar return. Use font.color.rgb or theme_color and preserve the distinct auto value where supported. Follow checked owner errors instead of reusing retired descendants or refetching them implicitly. Public Rust layout/model additions require updating exhaustive literals and matches according to the final accepted API inventory.

DrawingML text watermark opacity is restricted to0.5 with atomic refusal for other values. Existing VML support retains its full range. Native Word currently uses the approved one-font-per-name interoperability boundary, even though the API and own renderer retain four-face selection. Word multi-face interoperability, full-range watermark interoperability, bibliography catalogue parity, navigation and standalone authoring remain S92 work and are excluded from these release claims.

### Contributors

- [@Serkaine](https://github.com/Serkaine) supplied authenticated reports and reference contributions for these outcomes. Contribution slices are selectively adapted or delivered through reviewed hardened equivalents, rather than claiming every stacked PR merged unchanged.
  - [Issue 332](https://github.com/tensorbee/rdocx/issues/332), Markdown export puts spaces inside emphasis markers and joins a table to the list before it.
  - [Issue 334](https://github.com/tensorbee/rdocx/issues/334), New documents name no compatibility mode, so Word opens them in Compatibility Mode.
- [@hadim](https://github.com/hadim) supplied authenticated reports and reference contributions for these outcomes. Contribution slices are selectively adapted or delivered through reviewed hardened equivalents, rather than claiming every stacked PR merged unchanged.
  - [Issue 296](https://github.com/tensorbee/rdocx/issues/296), Caller fonts reach only the rdocx PDF: no --font-dir for rdocx images, nor anywhere in rpptx (convert --to png ignores it silently).
  - [Issue 297](https://github.com/tensorbee/rdocx/issues/297), A variable font from --font-dir renders bold runs with regular glyphs under a -Bold name.
  - [Issue 298](https://github.com/tensorbee/rdocx/issues/298), rdocx validate counts a paragraph holding only a picture as an empty paragraph.
  - [Issue 299](https://github.com/tensorbee/rdocx/issues/299), Python colours: RGBColor here, hex string there, and font.color.rgb (python-docx / python-pptx) fails in both bindings.
  - [Issue 300](https://github.com/tensorbee/rdocx/issues/300), rdocx Python: expose the table indent (w:tblInd), negative values included.
  - [Issue 301](https://github.com/tensorbee/rdocx/issues/301), replace_text_at on a text box edits the DrawingML copy only, the VML fallback keeps the old text.
  - [Issue 303](https://github.com/tensorbee/rdocx/issues/303), rdocx Python: run, paragraph, style and list formatting parity with the Rust API.
  - [Issue 304](https://github.com/tensorbee/rdocx/issues/304), rdocx: headers, footers, footnotes and page setup from Python, and two section bugs.
  - [Issue 305](https://github.com/tensorbee/rdocx/issues/305), rdocx: tables, pictures and links: Python parity, table column editing and cell split.
  - [Issue 306](https://github.com/tensorbee/rdocx/issues/306), rdocx Python: document-level features the CLI or Rust already have (export, templates, properties, validate).
  - [Issue 307](https://github.com/tensorbee/rdocx/issues/307), rdocx and rpptx CLI: the missing one-shot commands (batch replace, fields, images, metadata, slides, notes, fit).
  - [Issue 313](https://github.com/tensorbee/rdocx/issues/313), Fonts in rdocx and rpptx: caller fonts everywhere, variable fonts, more metric-compatible faces.
  - [Issue 314](https://github.com/tensorbee/rdocx/issues/314), Everyday coverage: Python and CLI parity for docx and pptx (Google Docs / Google Slides scope).
  - [Issue 316](https://github.com/tensorbee/rdocx/issues/316), rdocx and rpptx Python: first-draft script success, raw XML access, stable handles, no silent failures.
  - [Issue 321](https://github.com/tensorbee/rdocx/issues/321), rdocx layout: numbering instances of one definition should share their counters, as in Word.
  - [Issue 331](https://github.com/tensorbee/rdocx/issues/331), rdocx: deeply nested inline content controls crash the parser (stack overflow).
  - [Issue 337](https://github.com/tensorbee/rdocx/issues/337), rdocx: nested text boxes take exponential time and overflow the stack.
  - [Issue 339](https://github.com/tensorbee/rdocx/issues/339), rdocx: a tab is dropped from StoryItem.text and from the header, footer and JSON text views.
  - [Issue 340](https://github.com/tensorbee/rdocx/issues/340), rdocx: rebuild_toc adds an empty paragraph above the first entry and below the last.
  - [Issue 345](https://github.com/tensorbee/rdocx/issues/345), rdocx and rpptx Python: add_chart.
  - [Issue 350](https://github.com/tensorbee/rdocx/issues/350), rdocx: create a document from HTML in Python and the CLI.
  - [Issue 351](https://github.com/tensorbee/rdocx/issues/351), rdocx: create a document from Markdown.
  - [Issue 360](https://github.com/tensorbee/rdocx/issues/360), Font embedding: expose rdocx embed_font in Python and the CLI, add it to rpptx.
  - [Issue 362](https://github.com/tensorbee/rdocx/issues/362), rebuild_toc, update_page_fields and comments became quadratic in v0.16.0.
  - [PR 315](https://github.com/tensorbee/rdocx/pull/315), Pass caller fonts to every render and draw variable fonts at the requested weight.
  - [PR 317](https://github.com/tensorbee/rdocx/pull/317), Python binding gaps (cell text frames, one colour rule, table indent) and validate / text box replacement fixes.
  - [PR 319](https://github.com/tensorbee/rdocx/pull/319), Bundle Gelasio and Selawik, draw Symbol and Wingdings bullets and report the Aptos fallback.
  - [PR 320](https://github.com/tensorbee/rdocx/pull/320), Expose rdocx run, paragraph, style and list formatting in Python.
  - [PR 322](https://github.com/tensorbee/rdocx/pull/322), Edit rdocx table columns, split cells, and give Python pictures and internal links their everyday options.
  - [PR 325](https://github.com/tensorbee/rdocx/pull/325), Share rdocx text, export and validation with the CLI and expose document-level features in Python.
  - [PR 328](https://github.com/tensorbee/rdocx/pull/328), Write rdocx headers, footers, notes and page setup from Python, and fix new-section geometry and even footers.
  - [PR 330](https://github.com/tensorbee/rdocx/pull/330), Add batch replace, fields, images, metadata, fill, slide, notes and fit commands to the CLIs (stacked on #325).
  - [PR 333](https://github.com/tensorbee/rdocx/pull/333), Measure first-draft scripts, add raw XML access, keep rdocx handles stable and refuse silent failures.
  - [PR 336](https://github.com/tensorbee/rdocx/pull/336), Bound nested content controls when reading, and share list counters across numbering instances.
  - [PR 338](https://github.com/tensorbee/rdocx/pull/338), Read nested text boxes once and bound their nesting.
  - [PR 341](https://github.com/tensorbee/rdocx/pull/341), Keep TOC field markers in the first and last rebuilt entries, and rebuild one-paragraph TOCs.
  - [PR 342](https://github.com/tensorbee/rdocx/pull/342), Read run tabs, breaks and field instructions in story text as Paragraph.text does.
  - [PR 343](https://github.com/tensorbee/rdocx/pull/343), Embed an image from a header, footer or note part, so a header run can hold a picture.
  - [PR 354](https://github.com/tensorbee/rdocx/pull/354), Name compatibility mode 15 in new documents and write Markdown that CommonMark reads as intended.
  - [PR 356](https://github.com/tensorbee/rdocx/pull/356), Add charts from Python, and read and replace the data of an existing chart.
  - [PR 359](https://github.com/tensorbee/rdocx/pull/359), Create a document from HTML in Python and the CLI, and from Markdown.
  - [PR 361](https://github.com/tensorbee/rdocx/pull/361), Embed fonts in rdocx from Python and the CLI, and in rpptx everywhere.
  - [PR 363](https://github.com/tensorbee/rdocx/pull/363), Read each story source once when listing range paragraphs.
- [@pedroassumpcao](https://github.com/pedroassumpcao) supplied authenticated reports and reference contributions for these outcomes. Contribution slices are selectively adapted or delivered through reviewed hardened equivalents, rather than claiming every stacked PR merged unchanged.
  - [PR 294](https://github.com/tensorbee/rdocx/pull/294), Preserve mark highlights on nested inline runs.
  - [PR 302](https://github.com/tensorbee/rdocx/pull/302), Model self-closing empty table cell margins.

## rpptx-v0.14.1

### Highlights

PowerPoint gains richer Python authoring, current-owner XML edits, persistent autofit, charts, branded masters/layouts and semantic exports. The Rust crates, CLI and Python distribution share version 0.14.1.

### Added

- Python provides run/script properties, bounded gradient and pattern fills, table style/banding and cell text frames, properties, sections, media lifecycle, raster/PDF/notes/handout/ODP exports and read-only fit diagnostics.
- Shape authoring includes alt text, flips, groups, glued connectors and live owner checks. Footers, slide numbers, date fields, themes, transitions, editable masters and layouts use the actual owning parts.
- Chart creation and data replacement preserve actual chart/cache/workbook graphs and shared resources. Presentations export supported semantic HTML and Markdown with explicit resource and loss diagnostics.
- Nine checked XML owner routes validate qualified content before publication. Licensed embedded fonts retain actual producer payloads, checked relationships, bounded decoding and caller-font precedence.
- The unchanged 100-script compatibility corpus qualifies 95 correct files, with all 24 documentation cases and required XML, lifetime, native and rendering checks qualified. Five explicit divergences retain their documented boundaries.

### Fixed

- Cell vertical anchors and editable text frames follow the actual cell owner. Group transforms preserve resolvable siblings and report unsupported members instead of dropping the whole group.
- Explicit autofit stores the reviewed scale/spacing pair and fit_text chooses a supported checked size. Late unfit or unsafe owner updates refuse atomically. Qualified transitions retain their selected effect and direction.
- Recursive presentation parsing and traversal have checked resource limits and ordinary-stack refusal boundaries. Deterministic bundled faces and qualified Symbol/Wingdings fallbacks replace incomplete font substitution behavior.
- Caller fonts reach supported slide, notes, handout, raster/vector and authoring paths. Variable-font defaults and repeated diagnostic behavior retain the source-selected face.

### Compatibility

Font.color exposes ColorFormat instead of a hex string. Use font.color.rgb or theme_color. Length, colour and enum setters validate actual accepted forms and do not silently coerce unsupported values. Retired descendants fail before they can edit a replacement owner, while proven surviving ancestors retain their supported authority.

Autofit preserves the reviewed native allocation-versus-painted-bound distinction. A native floor pair is not a successful fit when content still overflows. Unsupported imports, fills, resources and exports keep their documented diagnostics or atomic refusal boundaries. Google Slides remains best effort, and no universal Office pixel parity is promised.

This family excludes the separate WASM npm release boundary and deferred S92 bibliography, navigation, standalone authoring and full-range watermark work.

### Contributors

- [@hadim](https://github.com/hadim) supplied authenticated reports and reference contributions for these outcomes. Contribution slices are selectively adapted or delivered through reviewed hardened equivalents, rather than claiming every stacked PR merged unchanged.
  - [Issue 295](https://github.com/tensorbee/rdocx/issues/295), rpptx Python: table cells have no text_frame, so table text cannot be formatted.
  - [Issue 296](https://github.com/tensorbee/rdocx/issues/296), Caller fonts reach only the rdocx PDF: no --font-dir for rdocx images, nor anywhere in rpptx (convert --to png ignores it silently).
  - [Issue 297](https://github.com/tensorbee/rdocx/issues/297), A variable font from --font-dir renders bold runs with regular glyphs under a -Bold name.
  - [Issue 299](https://github.com/tensorbee/rdocx/issues/299), Python colours: RGBColor here, hex string there, and font.color.rgb (python-docx / python-pptx) fails in both bindings.
  - [Issue 307](https://github.com/tensorbee/rdocx/issues/307), rdocx and rpptx CLI: the missing one-shot commands (batch replace, fields, images, metadata, slides, notes, fit).
  - [Issue 308](https://github.com/tensorbee/rdocx/issues/308), rpptx Python: expose the Rust text, fill, table, properties and export APIs.
  - [Issue 309](https://github.com/tensorbee/rdocx/issues/309), rpptx: shape authoring gaps (alt text, flip, group and ungroup, connector glue, picture placeholders, cell anchor).
  - [Issue 310](https://github.com/tensorbee/rdocx/issues/310), rpptx: autofit that sticks (store the shrink or the new height) and fit_text().
  - [Issue 311](https://github.com/tensorbee/rdocx/issues/311), rpptx: slide numbers, date and footer on slides, theme colours and fonts (read), transitions.
  - [Issue 312](https://github.com/tensorbee/rdocx/issues/312), rpptx: editable slide masters, layouts and themes (logo on every slide, backgrounds, text styles, apply a theme).
  - [Issue 313](https://github.com/tensorbee/rdocx/issues/313), Fonts in rdocx and rpptx: caller fonts everywhere, variable fonts, more metric-compatible faces.
  - [Issue 314](https://github.com/tensorbee/rdocx/issues/314), Everyday coverage: Python and CLI parity for docx and pptx (Google Docs / Google Slides scope).
  - [Issue 316](https://github.com/tensorbee/rdocx/issues/316), rdocx and rpptx Python: first-draft script success, raw XML access, stable handles, no silent failures.
  - [Issue 323](https://github.com/tensorbee/rdocx/issues/323), rpptx render: a group with a connector added inside a rotated, flipped, scaled group renders empty.
  - [Issue 344](https://github.com/tensorbee/rdocx/issues/344), rpptx: deeply nested groups, AlternateContent and timing nodes overflow the stack at open.
  - [Issue 345](https://github.com/tensorbee/rdocx/issues/345), rdocx and rpptx Python: add_chart.
  - [Issue 346](https://github.com/tensorbee/rdocx/issues/346), rpptx: read and replace the data of an existing chart.
  - [Issue 347](https://github.com/tensorbee/rdocx/issues/347), rpptx: import_slide rejects slides with charts, comments, embedded objects or slide jumps.
  - [Issue 348](https://github.com/tensorbee/rdocx/issues/348), rpptx: import_slide option to keep the source formatting (layout, master, theme).
  - [Issue 349](https://github.com/tensorbee/rdocx/issues/349), rpptx: merge presentations, and a slide import command in the CLI.
  - [Issue 352](https://github.com/tensorbee/rdocx/issues/352), rpptx: export a presentation to Markdown.
  - [Issue 353](https://github.com/tensorbee/rdocx/issues/353), rpptx: export a presentation to HTML.
  - [Issue 360](https://github.com/tensorbee/rdocx/issues/360), Font embedding: expose rdocx embed_font in Python and the CLI, add it to rpptx.
  - [PR 315](https://github.com/tensorbee/rdocx/pull/315), Pass caller fonts to every render and draw variable fonts at the requested weight.
  - [PR 317](https://github.com/tensorbee/rdocx/pull/317), Python binding gaps (cell text frames, one colour rule, table indent) and validate / text box replacement fixes.
  - [PR 318](https://github.com/tensorbee/rdocx/pull/318), Expose the rpptx Rust text, fill, table, presentation and export APIs in Python, and fix the default template.
  - [PR 319](https://github.com/tensorbee/rdocx/pull/319), Bundle Gelasio and Selawik, draw Symbol and Wingdings bullets and report the Aptos fallback.
  - [PR 324](https://github.com/tensorbee/rdocx/pull/324), Write slide numbers, footers and dates as PowerPoint does, read themes and author transitions in rpptx.
  - [PR 326](https://github.com/tensorbee/rdocx/pull/326), Author alt text, flips, groups, glued connectors and picture placeholders in rpptx, and flip before rotating.
  - [PR 327](https://github.com/tensorbee/rdocx/pull/327), Edit rpptx slide masters, layouts and themes: logos, backgrounds, text styles and apply_theme (stacked on #324).
  - [PR 329](https://github.com/tensorbee/rdocx/pull/329), Store rpptx autofit as PowerPoint does, resize shapes to their text and add fit_text.
  - [PR 330](https://github.com/tensorbee/rdocx/pull/330), Add batch replace, fields, images, metadata, fill, slide, notes and fit commands to the CLIs (stacked on #325).
  - [PR 333](https://github.com/tensorbee/rdocx/pull/333), Measure first-draft scripts, add raw XML access, keep rdocx handles stable and refuse silent failures.
  - [PR 335](https://github.com/tensorbee/rdocx/pull/335), Draw rotated and flipped groups as PowerPoint does, and keep an unresolvable shape's siblings drawn.
  - [PR 355](https://github.com/tensorbee/rdocx/pull/355), Bound slide nesting and read nested shapes and time nodes once.
  - [PR 356](https://github.com/tensorbee/rdocx/pull/356), Add charts from Python, and read and replace the data of an existing chart.
  - [PR 357](https://github.com/tensorbee/rdocx/pull/357), Import slides with charts, objects, comments and jumps, keep source formatting, and merge decks.
  - [PR 358](https://github.com/tensorbee/rdocx/pull/358), Export a presentation to Markdown and to semantic HTML.
  - [PR 361](https://github.com/tensorbee/rdocx/pull/361), Embed fonts in rdocx from Python and the CLI, and in rpptx everywhere.


## Unreleased

No changes have been recorded after the prepared family sections.

## v0.16.0

### Highlights

Word editing gains checked field updates, richer note policies, cross-story ranges and reusable document fragments. Comment and story edits preserve their physical owners and fail atomically when a safe update cannot be proved. The stable Word Rust crates, CLI and Python distribution share version 0.16.0.

### Added

- Native Rust note policies control separators, custom markers, numbering, restart and placement, including authored continuation content. Cross-story bookmark, comment, permission and proofing ranges can be checked, moved and removed with paired endpoints preserved.
- Native block fragments capture supported body, furniture, note, comment, table-cell, text-box and nested control owners. Imports close their styles, numbering, notes, review metadata and reachable supported package graphs. Glossary, AutoText, building-block and content-control template operations use this staged ownership model.
- Native field authoring supports simple and complex fields, nested operands, rich cached results and locks. Checked refresh covers pagination, captions, sequences, cross-references and typed note copies. INDEX, figure TOC and TOA rebuilding retains unsupported caches with diagnostics and uses its measured en-US ASCII sorting boundary.
- Native citation and bibliography APIs preserve all seventeen source-kind identities, sixteen contributor-role identities and twelve style identities as authoring metadata. Formatting admits 210 of 223 measured dense APA locale configurations, separately earned sparse and rich controls, and lean numeric1033 Book inputs for the eleven other bibliography styles. Citation admissions remain separate. Recognized unfinished formatting branches abort the complete staged refresh. Unrecognized paths retain their caches with diagnostics. Source/style metadata coverage does not promise complete catalogue formatting, arbitrary Unicode ordering or dedicated Python/WASM bibliography APIs.

### Fixed

- Multi-paragraph comment threads retain parent and resolved metadata and authored paragraph boundaries, incorporating the core of [PR 271](https://github.com/tensorbee/rdocx/pull/271) with additional closure checks for [Issue 270](https://github.com/tensorbee/rdocx/issues/270).
- Explicit cell `nil` borders suppress inherited borders, while omitted or `none` borders retain inheritance, for [Issue 272](https://github.com/tensorbee/rdocx/issues/272) and [PR 274](https://github.com/tensorbee/rdocx/pull/274). Page and column breaks inside cells no longer introduce unwanted line breaks or spaces, for [Issue 273](https://github.com/tensorbee/rdocx/issues/273) and [PR 275](https://github.com/tensorbee/rdocx/pull/275).
- Direct table alignment survives style resolution for [Issue 277](https://github.com/tensorbee/rdocx/issues/277) and [PR 279](https://github.com/tensorbee/rdocx/pull/279). Absent side margins resolve to zero while cell padding remains authoritative. Measured legacy table placement uses the first accepted cell's left margin, resolving [Issue 276](https://github.com/tensorbee/rdocx/issues/276) and [Issue 278](https://github.com/tensorbee/rdocx/issues/278) through a hardened equivalent of [PR 280](https://github.com/tensorbee/rdocx/pull/280).
- Comment removal closes its thread and companion metadata atomically for [Issue 282](https://github.com/tensorbee/rdocx/issues/282). Typed accepted comment anchors and identity-preserving anchor moves address [Issue 283](https://github.com/tensorbee/rdocx/issues/283) and [Issue 284](https://github.com/tensorbee/rdocx/issues/284). Exact direct-control mapping and cached sibling/nested-field story text address [Issue 289](https://github.com/tensorbee/rdocx/issues/289) and [Issue 291](https://github.com/tensorbee/rdocx/issues/291), using selected hardened equivalents from [PR 287](https://github.com/tensorbee/rdocx/pull/287).
- Scoped paragraph, cell and document replacement preserves selected formatting and review state for [Issue 285](https://github.com/tensorbee/rdocx/issues/285), using a hardened equivalent of [PR 286](https://github.com/tensorbee/rdocx/pull/286). Complete story replacement and building-block removal reconcile review ownership for [Issue 288](https://github.com/tensorbee/rdocx/issues/288). Qualified marker omission happens before selected fragment dependencies are captured for [Issue 292](https://github.com/tensorbee/rdocx/issues/292), using the relevant contribution from [PR 290](https://github.com/tensorbee/rdocx/pull/290).
- Namespace declarations alone no longer mark an imported numbering level as unmodeled. Real unknown attributes and children retain the existing preservation and admission rules, incorporating the core of [PR 293](https://github.com/tensorbee/rdocx/pull/293).
- The shared font loader accepts source variants enabled by downstream Cargo feature unification, incorporating [PR 269](https://github.com/tensorbee/rdocx/pull/269). Existing ordinary file and binary caches retain their behavior.

### Compatibility

This pre-1.0 release adds required members to existing public Rust types. Update exhaustive struct literals for Word `LayoutInput` with `sequence_snapshot`, `legacy_table_positioning`, `modern_footnote_layout`, `footnote_layout_like_word8`, `note_defaults`, `story_part_names` and `story_bodies`. Add `text_body` and `source_text_box_owner` to `CT_Shape`, `special_references` to `CT_NoteProperties`, and `draw_marker` to `NoteLayout`. Use default/update construction where the type provides it.

Shared `TextSegment`, `GlyphRun` and `MultilingualGlyphRun` literals now need `note_reference_source`. Exhaustive `FieldKind` matches must handle `Section`, `SectionPages`, `SequenceContext` and `SequenceRepeat`. Exhaustive `WordStory` matches must also handle `TextBox`. `FieldSource.index` now counts nested instruction and result fields in preorder.

Native fallible header/footer setters validate staged content before publication. Existing infallible wrappers retain their panic contract and can reject invalid XML text earlier. Python setters propagate failures and revise document state only on success. Zero-result scoped edits preserve existing handles.

Table border, placement and padding corrections intentionally change some rendered output. Note placement follows the authored OOXML policy, including the documented Word-for-Mac `beneathText` discrepancy. Unsupported, locked or ambiguous field owners retain their documented diagnostic/error boundaries. Full bibliography catalogue parity belongs to F-X192 in S91, and navigation F-283 is carried to S91. These changes do not promise universal native Word pixel parity.

### Contributors

- Hadrien Mary [@hadim](https://github.com/hadim) supplied the reports and reference implementations behind comment threads ([PR 271](https://github.com/tensorbee/rdocx/pull/271)), cell borders and breaks ([PR 274](https://github.com/tensorbee/rdocx/pull/274), [PR 275](https://github.com/tensorbee/rdocx/pull/275)), table alignment and margins ([PR 279](https://github.com/tensorbee/rdocx/pull/279), [PR 280](https://github.com/tensorbee/rdocx/pull/280)), comment cleanup ([Issue 282](https://github.com/tensorbee/rdocx/issues/282)), scoped replacement ([PR 286](https://github.com/tensorbee/rdocx/pull/286)), anchors and field snapshots ([PR 287](https://github.com/tensorbee/rdocx/pull/287)), and whole-story review ownership and fragment omission ([PR 290](https://github.com/tensorbee/rdocx/pull/290)). Accepted portions landed directly with additional controls or through the hardened equivalents described above, rather than wholesale adoption of the stacked PRs.
- Pedro Assumpcao [@pedroassumpcao](https://github.com/pedroassumpcao) contributed the namespace-only numbering reader correction in [PR 293](https://github.com/tensorbee/rdocx/pull/293). The accepted core has expanded local preservation tests.
- Changjoon [@changjoon-park](https://github.com/changjoon-park) contributed the shared fontdb feature-unification repair in [PR 269](https://github.com/tensorbee/rdocx/pull/269), also consumed by Word.

## rpptx-v0.14.0

### Highlights

The shared OOXML and PowerPoint Rust crates, PowerPoint CLI and Python distribution share version 0.14.0. The shared font loader accepts additional source variants enabled by downstream Cargo feature unification.

### Added

Shared layout records carry explicit note-reference provenance. The PowerPoint, chart and PDF consumers initialize or propagate that provenance through their existing rendering paths. No separate PowerPoint editing feature is introduced by the Word-specific changes in this release range.

### Fixed

The shared font loader uses fontdb's face-data callback for source variants enabled by downstream feature unification, incorporating the core of [PR 269](https://github.com/tensorbee/rdocx/pull/269). It retains ordinary file and binary caches, the original face index and existing defaults.

### Compatibility

This pre-1.0 release adds `note_reference_source` to the existing public `TextSegment`, `GlyphRun` and `MultilingualGlyphRun` types. Update exhaustive Rust struct literals, or use default/update construction where provided. Exhaustive `FieldKind` matches must handle `Section`, `SectionPages`, `SequenceContext` and `SequenceRepeat`. `FieldSource.index` counts nested instruction and result fields in preorder. The font repair leaves public loader APIs and bundled font assets unchanged.

### Contributors

Changjoon [@changjoon-park](https://github.com/changjoon-park) contributed the fontdb feature-unification fix in [PR 269](https://github.com/tensorbee/rdocx/pull/269). Word-specific contributions are credited in v0.16.0.

## v0.15.0

### Highlights

The seven stable Word Rust crates, the `rdocx` CLI, and the `rdocx` Python
distribution move together to 0.15.0 from the reviewed S88 source. One
release carries six CLI archives, six `cp39-abi3` wheels, one source
distribution, and a `SHA256SUMS` covering all thirteen payloads. Each
payload has build provenance verifiable with `gh attestation verify FILE -R
tensorbee/rdocx`. This answers
[@hadim](https://github.com/hadim)'s
[Issue 266](https://github.com/tensorbee/rdocx/issues/266) request for one
current-main release with matching CLI and Python versions.

### Added

- Word authoring now covers rich content and comments in all six header and
  footer variants, together with rich footnotes and endnotes whose references,
  relationships, pictures, hyperlinks, and unmodelled XML survive save and
  reopen. High-level style formatting covers alignment, spacing, indentation,
  font, size, bold, and colour for
  [Issue 264](https://github.com/tensorbee/rdocx/issues/264).
- Python editing exposes tables, sections, styles, bookmarks, fields, story
  content, pictures, hyperlinks, tracked revisions, and counted replacement.
  The CLI can inspect and compare related stories, anchor comments on text,
  select accepted or tracked revision views, and emit structured results.
  These surfaces incorporate
  [PR 176](https://github.com/tensorbee/rdocx/pull/176),
  [PR 186](https://github.com/tensorbee/rdocx/pull/186),
  [PR 187](https://github.com/tensorbee/rdocx/pull/187),
  [PR 194](https://github.com/tensorbee/rdocx/pull/194),
  [PR 198](https://github.com/tensorbee/rdocx/pull/198),
  [PR 201](https://github.com/tensorbee/rdocx/pull/201),
  [PR 203](https://github.com/tensorbee/rdocx/pull/203),
  [PR 204](https://github.com/tensorbee/rdocx/pull/204),
  [PR 212](https://github.com/tensorbee/rdocx/pull/212),
  [PR 220](https://github.com/tensorbee/rdocx/pull/220),
  [PR 236](https://github.com/tensorbee/rdocx/pull/236), and
  [PR 256](https://github.com/tensorbee/rdocx/pull/256).
- Native Word story traversal and replacement reach content controls, text
  boxes, footnotes, endnotes, tracked insertions, simple fields, smart tags,
  and custom XML wrappers. This incorporates
  [PR 177](https://github.com/tensorbee/rdocx/pull/177),
  [PR 179](https://github.com/tensorbee/rdocx/pull/179),
  [PR 180](https://github.com/tensorbee/rdocx/pull/180),
  [PR 191](https://github.com/tensorbee/rdocx/pull/191),
  [PR 195](https://github.com/tensorbee/rdocx/pull/195),
  [PR 202](https://github.com/tensorbee/rdocx/pull/202),
  [PR 210](https://github.com/tensorbee/rdocx/pull/210), and
  [PR 211](https://github.com/tensorbee/rdocx/pull/211).

### Fixed

- The `rdocx` CLI runs comparison and other document commands on an eight MiB
  stack on Windows, avoiding a stack overflow in the complete Word workflow.
- Editing Word parts writes compact XML without repeating canonical namespace
  declarations on every retained paragraph. It preserves producer aliases,
  root attributes, unknown content, and unchanged package parts, building on
  [PR 154](https://github.com/tensorbee/rdocx/pull/154) and
  [PR 251](https://github.com/tensorbee/rdocx/pull/251) for
  [Issue 245](https://github.com/tensorbee/rdocx/issues/245).
  Newly created documents now carry Word-compatible application metadata, and
  nested tables retain their required trailing paragraph.
- Word layout uses the governing section for each paragraph and table. Line
  height, inline picture height, tab stops, TOC entries, page fields,
  keep-with-next chains, row minimum heights, and terminal spaces follow the
  reviewed Word cases. This incorporates
  [PR 148](https://github.com/tensorbee/rdocx/pull/148),
  [PR 199](https://github.com/tensorbee/rdocx/pull/199),
  [PR 200](https://github.com/tensorbee/rdocx/pull/200),
  [PR 225](https://github.com/tensorbee/rdocx/pull/225),
  [PR 237](https://github.com/tensorbee/rdocx/pull/237),
  [PR 241](https://github.com/tensorbee/rdocx/pull/241), and
  [PR 242](https://github.com/tensorbee/rdocx/pull/242).
  [PR 265](https://github.com/tensorbee/rdocx/pull/265) supplied independent
  mixed-section regression coverage.
- Comparison now records changed picture bytes and final table or paragraph
  changes as revisions that survive acceptance and rejection. Accepted text,
  HTML, Markdown, JSON, PDF layout, and media traversal omit deleted table
  rows and join paragraphs whose marks were deleted. This incorporates
  [PR 228](https://github.com/tensorbee/rdocx/pull/228),
  [PR 229](https://github.com/tensorbee/rdocx/pull/229),
  [PR 257](https://github.com/tensorbee/rdocx/pull/257),
  [PR 258](https://github.com/tensorbee/rdocx/pull/258),
  [PR 259](https://github.com/tensorbee/rdocx/pull/259),
  [PR 260](https://github.com/tensorbee/rdocx/pull/260),
  [PR 261](https://github.com/tensorbee/rdocx/pull/261),
  [PR 262](https://github.com/tensorbee/rdocx/pull/262), and
  [PR 263](https://github.com/tensorbee/rdocx/pull/263).
- CLI and package saves refuse unintended overwrite and stage replacement
  atomically. Edited border and namespace declarations, shared-run field
  text, and XML-valid characters survive their supported read and write paths.
  This incorporates
  [PR 174](https://github.com/tensorbee/rdocx/pull/174),
  [PR 178](https://github.com/tensorbee/rdocx/pull/178),
  [PR 185](https://github.com/tensorbee/rdocx/pull/185),
  [PR 197](https://github.com/tensorbee/rdocx/pull/197),
  [PR 232](https://github.com/tensorbee/rdocx/pull/232),
  [PR 233](https://github.com/tensorbee/rdocx/pull/233), and
  [PR 239](https://github.com/tensorbee/rdocx/pull/239).

### Compatibility

The exact seven publishable stable crates move from 0.14.0 to 0.15.0:
`rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, and `rdocx-cli`. Update their pins together.
They use the separately prepared shared OOXML 0.13.1 family. The Python
distribution and import name stay `rdocx`, and its 0.15.0 wheels require
Python 3.9 or newer. Historical `v0.14.0` and `py-rdocx-v0.14.0`
remain unchanged. The unpublished `rdocx-wasm` and npm package are outside
this crates.io and PyPI release.

This pre-1.0 minor release has source-incompatible Rust model changes.
`rdocx_layout::PageGeometry` was `Copy` in 0.14.0 and is now `Clone`
because it owns column tracks and optional page metadata. Pass it by
reference or call `.clone()` where a copy was previously implicit. Full
`PageGeometry` struct literals must supply `columns`,
`column_separator`, `page_borders`, `line_numbers`,
`vertical_alignment`, `mirror_margins`,
`do_not_use_html_paragraph_auto_spacing`, and `body_rotation`.
Public `GlyphRun`, `LineBreakParams`, and `LayoutInput` also gained
fields. Update full literals and exact-output fixtures when upgrading.
`rdocx replace` now refuses input-as-output and existing output paths.
Scripts that overwrote a source file must choose a new output path.

### Contributors

[@hadim](https://github.com/hadim), Hadrien Mary, reported
[Issue 266](https://github.com/tensorbee/rdocx/issues/266) and
[Issue 245](https://github.com/tensorbee/rdocx/issues/245). The linked Word,
Python, CLI, layout, comparison, and preservation PRs above supplied
reference behavior that was reconciled or hardened in maintainer commits.
Those GitHub PRs were closed without direct merges.
[@set2374](https://github.com/set2374) supplied the independent
[PR 265](https://github.com/tensorbee/rdocx/pull/265) mixed-section regression.
[@Andrewkoro105](https://github.com/Andrewkoro105) requested the
[Issue 264](https://github.com/tensorbee/rdocx/issues/264) high-level style
API. Atul Sharma reviewed and integrated the release family.

## rpptx-v0.13.1

### Highlights

The shared OOXML and PowerPoint Rust crates, `rpptx` CLI, and `rpptx` Python
distribution move together to 0.13.1. One family release carries six CLI
archives, six `cp39-abi3` wheels, one source distribution, and a `SHA256SUMS`
covering all thirteen payloads. Each payload has build provenance that can be
checked with `gh attestation verify FILE -R tensorbee/rdocx`. This answers
[@hadim](https://github.com/hadim)'s
[Issue 266](https://github.com/tensorbee/rdocx/issues/266) request for a
current release from the reviewed main merge with matching CLI and Python
versions.
The immutable `rpptx-v0.13.0` tag stopped at its package inventory gate
before either registry publication or GitHub release creation.

### Added

- Native and Python presentation editing now covers retained run handles,
  inherited shape geometry, table rows and cells, picture crop,
  grouping, shape hyperlinks, and text-range comment anchors.
- Shape authoring covers slide jumps, line dash and ends, outer shadows,
  preset geometry replacement, and theme effect selection. Cross-deck slide
  import carries supported media and notes. Counted replacement can target a
  single slide or text frame, and built-in table styles render their
  backgrounds.
- The CLI can add comments and report speaker notes. Text, outline, and
  inspection commands have JSON forms for automation.

### Fixed

- The `rpptx` CLI runs validation and other presentation commands on an eight
  MiB stack on Windows, avoiding a stack overflow in the complete deck chain.
- DrawingML text edits retain unmodelled body properties and turn assigned
  line feeds into paragraphs while preserving line separators in layout.
  Newly authored shapes receive theme style references in schema order.
- Presentation PDF output paints slide backgrounds, preserves supported
  producer paragraph properties, and keeps searchable text paired with its
  glyphs. Rich and plain line layout follow the reviewed spacing and line
  break cases.

### Compatibility

The exact 15 publishable shared OOXML and PowerPoint crates move from 0.12.1
to 0.13.1 together. They are `oxml-core`, `oxml-opc`, `oxml-media`,
`oxml-layout`, `oxml-drawing`, `oxml-pdf`, `oxml-sml`, `oxml-cli-support`,
`oxml-chart`, `rpptx-oxml`, `rpptx-chart`, `rpptx-layout`, `rpptx-render`,
`rpptx`, and `rpptx-cli`. Rust callers should update every shared internal
crate pin together. This pre-1.0 minor release adds authored-shape and
presentation APIs and changes the output of affected shape, text, and PDF
operations. `oxml-layout::LineBreakParams` gained required public fields, so
Rust callers constructing it with a struct literal must update those
literals. Review exact-output fixtures and exhaustive use of public models.

The Python distribution and import name remain `rpptx`. The 0.13.1 wheels
require Python 3.9 or newer. Stable Word crates, `rdocx` Python, and the
unpublished `rpptx-wasm` and npm packages are outside this tag. Historical
`rpptx-v0.12.1` and `py-rpptx-v0.12.1` remain unchanged.

### Contributors

[@hadim](https://github.com/hadim), Hadrien Mary, reported
[Issue 266](https://github.com/tensorbee/rdocx/issues/266) and contributed the
presentation changes in
[PR 173](https://github.com/tensorbee/rdocx/pull/173),
[PR 181](https://github.com/tensorbee/rdocx/pull/181),
[PR 189](https://github.com/tensorbee/rdocx/pull/189),
[PR 192](https://github.com/tensorbee/rdocx/pull/192),
[PR 208](https://github.com/tensorbee/rdocx/pull/208),
[PR 209](https://github.com/tensorbee/rdocx/pull/209),
[PR 218](https://github.com/tensorbee/rdocx/pull/218),
[PR 219](https://github.com/tensorbee/rdocx/pull/219),
[PR 221](https://github.com/tensorbee/rdocx/pull/221),
[PR 223](https://github.com/tensorbee/rdocx/pull/223),
[PR 224](https://github.com/tensorbee/rdocx/pull/224),
[PR 230](https://github.com/tensorbee/rdocx/pull/230),
[PR 231](https://github.com/tensorbee/rdocx/pull/231),
[PR 234](https://github.com/tensorbee/rdocx/pull/234),
[PR 235](https://github.com/tensorbee/rdocx/pull/235),
[PR 238](https://github.com/tensorbee/rdocx/pull/238), and
[PR 252](https://github.com/tensorbee/rdocx/pull/252). The shared rendering
and layout work came from
[PR 175](https://github.com/tensorbee/rdocx/pull/175),
[PR 188](https://github.com/tensorbee/rdocx/pull/188),
[PR 196](https://github.com/tensorbee/rdocx/pull/196),
[PR 206](https://github.com/tensorbee/rdocx/pull/206),
[PR 207](https://github.com/tensorbee/rdocx/pull/207),
[PR 222](https://github.com/tensorbee/rdocx/pull/222), and
[PR 242](https://github.com/tensorbee/rdocx/pull/242). These PRs were closed
after their behavior was integrated or hardened in the repository. Atul
Sharma reviewed and integrated the release family.

## rpptx-v0.12.1

### Highlights

This patch recovery replaces the failed immutable `rpptx-v0.12.0` workflow
attempt. That tag published no crates and created no GitHub release. The
package content below is the reviewed S73 incubating family prepared at the
new coherent version.

The shared OOXML and PowerPoint family moves to 0.12.1 with correct picture
transparency, safe rendering of large and transparent raster pictures, notes
that render and update reliably, portable authored charts, and searchable PDF
output in logical reading order. The `rpptx` CLI now ships as prebuilt
binaries attached to its GitHub release.

### Added

- Picture transparency from `a:alphaModFix` is applied across slide, layout,
  master, background, preview, SVG, PDF, and raster output, based on
  [PR 105](https://github.com/tensorbee/rdocx/pull/105) for
  [Issue 91](https://github.com/tensorbee/rdocx/issues/91).
- Speaker notes text can be replaced while keeping the notes shape identity
  and run formatting. Presentation replacement counts slide and notes
  matches, and `rpptx replace` refuses to overwrite its input or an existing
  output, validates an expected match count, and publishes atomically. This
  resolves the replacement half of
  [Issue 120](https://github.com/tensorbee/rdocx/issues/120).
- Native presentation APIs expose shape geometry and identity, run font
  details, autofit mode, and formatting-preserving run text replacement for
  the requests in [Issue 121](https://github.com/tensorbee/rdocx/issues/121).
- Rust release tags attach checksummed `rpptx` CLI archives for six Linux,
  macOS, and Windows targets, and `cargo binstall rpptx-cli` resolves them, as
  requested in [Issue 100](https://github.com/tensorbee/rdocx/issues/100).
- Authored charts emit explicit title layout, overlay, marker, and smoothing
  defaults that viewers such as Word and Pages read consistently. This
  builds on [PR 71](https://github.com/tensorbee/rdocx/pull/71) and
  [PR 123](https://github.com/tensorbee/rdocx/pull/123).
- `Presentation::slide_png_deterministic` and `slide_pngs_deterministic` render
  slides with bundled fonts only, supporting the presentation rendering
  automation requested in
  [Issue 76](https://github.com/tensorbee/rdocx/issues/76).
- Layout output exposes `FieldSource` provenance on field glyph runs so
  page-dependent field caches can be written from one deterministic
  pagination pass, as planned in
  [Issue 93](https://github.com/tensorbee/rdocx/issues/93).

### Fixed

- Raster output now premultiplies straight-alpha pictures before compositing,
  and the decoded-pixel render limit rises from 16 MiB to 64 MiB, so larger
  pictures render instead of being skipped silently. The 16 MiB encoded-file
  limit is unchanged. This resolves
  [Issue 119](https://github.com/tensorbee/rdocx/issues/119).
- Notes render when a producer-valid notes slide omits its reverse
  relationship to the slide, while conflicting owners still fail closed. This
  resolves the rendering half of
  [Issue 120](https://github.com/tensorbee/rdocx/issues/120).
- Date, footer, and slide-number placeholders placed on the slide itself stay
  visible regardless of master `p:hf` flags, which only govern inherited
  placeholders. This answers
  [Issue 92](https://github.com/tensorbee/rdocx/issues/92).
- Generated PDFs keep complete logical text lines for search and extraction,
  which resolves the shared renderer part of
  [Issue 74](https://github.com/tensorbee/rdocx/issues/74).
- Shared line layout honors run-level page breaks for page-dependent Word
  output, from [Issue 88](https://github.com/tensorbee/rdocx/issues/88) and
  [PR 102](https://github.com/tensorbee/rdocx/pull/102).

### Compatibility

The exact 15-package shared OOXML and PowerPoint crates.io family moves
together from 0.11.0 to 0.12.1. The selected set is `oxml-core`, `oxml-opc`,
`oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`, `oxml-sml`,
`oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`, `rpptx-layout`,
`rpptx-render`, `rpptx`, and `rpptx-cli`.

This pre-1.0 minor release contains source-incompatible Rust changes. Public
structs gained fields, including `ChartData` (axis titles and a palette, now
with `Default`), `LayoutLine` (`forced_break_after`), field glyph runs
(`field_source`), and `ResolvedImage` (`opacity`). Exhaustive enums gained
variants, including `OpcError::DuplicatePartName`. Code that builds these
structs with full struct literals or matches these enums exhaustively must be
updated. `Presentation::replace_text` now also replaces matches in speaker notes
and includes them in its returned count. Callers that must leave notes untouched
need to account for that change.

`rpptx replace` now fails when no match is found unless an expected count of
zero is given, and it refuses to overwrite an existing output file. Scripts
that relied on silent zero-match success or in-place overwrite must pass an
explicit expected count and a new output path.

Stable Word crates, Python distributions, WASM, and npm packages are outside
this release. `rpptx-wasm@0.12.1` is not a crates.io package.

### Contributors

[@hadim](https://github.com/hadim) reported the rendering, notes, header and
footer flag, automation, field, and binding gaps in Issues
[74](https://github.com/tensorbee/rdocx/issues/74),
[76](https://github.com/tensorbee/rdocx/issues/76),
[88](https://github.com/tensorbee/rdocx/issues/88),
[91](https://github.com/tensorbee/rdocx/issues/91),
[92](https://github.com/tensorbee/rdocx/issues/92),
[93](https://github.com/tensorbee/rdocx/issues/93),
[100](https://github.com/tensorbee/rdocx/issues/100),
[119](https://github.com/tensorbee/rdocx/issues/119),
[120](https://github.com/tensorbee/rdocx/issues/120), and
[121](https://github.com/tensorbee/rdocx/issues/121). They also contributed the
picture transparency implementation in
[PR 105](https://github.com/tensorbee/rdocx/pull/105) and the run page-break
layout in [PR 102](https://github.com/tensorbee/rdocx/pull/102).
[@chevinbrown](https://github.com/chevinbrown) contributed portable authored
charts in [PR 71](https://github.com/tensorbee/rdocx/pull/71) and the line-chart
viewer defaults in [PR 123](https://github.com/tensorbee/rdocx/pull/123).
Atul Sharma integrated, hardened, and released the family.

## rpptx-v0.12.0

### Highlights

The shared OOXML and PowerPoint family moves to 0.12.0 with correct picture
transparency, safe rendering of large and transparent raster pictures, notes
that render and update reliably, portable authored charts, and searchable PDF
output in logical reading order. The `rpptx` CLI now ships as prebuilt
binaries attached to its GitHub release.

### Added

- Picture transparency from `a:alphaModFix` is applied across slide, layout,
  master, background, preview, SVG, PDF, and raster output, based on
  [PR 105](https://github.com/tensorbee/rdocx/pull/105) for
  [Issue 91](https://github.com/tensorbee/rdocx/issues/91).
- Speaker notes text can be replaced while keeping the notes shape identity
  and run formatting. Presentation replacement counts slide and notes
  matches, and `rpptx replace` refuses to overwrite its input or an existing
  output, validates an expected match count, and publishes atomically. This
  resolves the replacement half of
  [Issue 120](https://github.com/tensorbee/rdocx/issues/120).
- Native presentation APIs expose shape geometry and identity, run font
  details, autofit mode, and formatting-preserving run text replacement for
  the requests in [Issue 121](https://github.com/tensorbee/rdocx/issues/121).
- Rust release tags attach checksummed `rpptx` CLI archives for six Linux,
  macOS, and Windows targets, and `cargo binstall rpptx-cli` resolves them, as
  requested in [Issue 100](https://github.com/tensorbee/rdocx/issues/100).
- Authored charts emit explicit title layout, overlay, marker, and smoothing
  defaults that viewers such as Word and Pages read consistently. This
  builds on [PR 71](https://github.com/tensorbee/rdocx/pull/71) and
  [PR 123](https://github.com/tensorbee/rdocx/pull/123).
- `Presentation::slide_png_deterministic` and `slide_pngs_deterministic` render
  slides with bundled fonts only, supporting the presentation rendering
  automation requested in
  [Issue 76](https://github.com/tensorbee/rdocx/issues/76).
- Layout output exposes `FieldSource` provenance on field glyph runs so
  page-dependent field caches can be written from one deterministic
  pagination pass, as planned in
  [Issue 93](https://github.com/tensorbee/rdocx/issues/93).

### Fixed

- Raster output now premultiplies straight-alpha pictures before compositing,
  and the decoded-pixel render limit rises from 16 MiB to 64 MiB, so larger
  pictures render instead of being skipped silently. The 16 MiB encoded-file
  limit is unchanged. This resolves
  [Issue 119](https://github.com/tensorbee/rdocx/issues/119).
- Notes render when a producer-valid notes slide omits its reverse
  relationship to the slide, while conflicting owners still fail closed. This
  resolves the rendering half of
  [Issue 120](https://github.com/tensorbee/rdocx/issues/120).
- Date, footer, and slide-number placeholders placed on the slide itself stay
  visible regardless of master `p:hf` flags, which only govern inherited
  placeholders. This answers
  [Issue 92](https://github.com/tensorbee/rdocx/issues/92).
- Generated PDFs keep complete logical text lines for search and extraction,
  which resolves the shared renderer part of
  [Issue 74](https://github.com/tensorbee/rdocx/issues/74).
- Shared line layout honors run-level page breaks for page-dependent Word
  output, from [Issue 88](https://github.com/tensorbee/rdocx/issues/88) and
  [PR 102](https://github.com/tensorbee/rdocx/pull/102).

### Compatibility

The exact 15-package shared OOXML and PowerPoint crates.io family moves
together from 0.11.0 to 0.12.0. The selected set is `oxml-core`, `oxml-opc`,
`oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`, `oxml-sml`,
`oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`, `rpptx-layout`,
`rpptx-render`, `rpptx`, and `rpptx-cli`.

This pre-1.0 minor release contains source-incompatible Rust changes. Public
structs gained fields, including `ChartData` (axis titles and a palette, now
with `Default`), `LayoutLine` (`forced_break_after`), field glyph runs
(`field_source`), and `ResolvedImage` (`opacity`). Exhaustive enums gained
variants, including `OpcError::DuplicatePartName`. Code that builds these
structs with full struct literals or matches these enums exhaustively must be
updated. `Presentation::replace_text` now also replaces matches in speaker notes
and includes them in its returned count. Callers that must leave notes untouched
need to account for that change.

`rpptx replace` now fails when no match is found unless an expected count of
zero is given, and it refuses to overwrite an existing output file. Scripts
that relied on silent zero-match success or in-place overwrite must pass an
explicit expected count and a new output path.

Stable Word crates, Python distributions, WASM, and npm packages are outside
this release. `rpptx-wasm@0.12.0` is not a crates.io package.

### Contributors

[@hadim](https://github.com/hadim) reported the rendering, notes, header and
footer flag, automation, field, and binding gaps in Issues
[74](https://github.com/tensorbee/rdocx/issues/74),
[76](https://github.com/tensorbee/rdocx/issues/76),
[88](https://github.com/tensorbee/rdocx/issues/88),
[91](https://github.com/tensorbee/rdocx/issues/91),
[92](https://github.com/tensorbee/rdocx/issues/92),
[93](https://github.com/tensorbee/rdocx/issues/93),
[100](https://github.com/tensorbee/rdocx/issues/100),
[119](https://github.com/tensorbee/rdocx/issues/119),
[120](https://github.com/tensorbee/rdocx/issues/120), and
[121](https://github.com/tensorbee/rdocx/issues/121). They also contributed the
picture transparency implementation in
[PR 105](https://github.com/tensorbee/rdocx/pull/105) and the run page-break
layout in [PR 102](https://github.com/tensorbee/rdocx/pull/102).
[@chevinbrown](https://github.com/chevinbrown) contributed portable authored
charts in [PR 71](https://github.com/tensorbee/rdocx/pull/71) and the line-chart
viewer defaults in [PR 123](https://github.com/tensorbee/rdocx/pull/123).
Atul Sharma integrated, hardened, and released the family.

## v0.14.0

### Highlights

The complete stable Word family publishes the M23 from-scratch business
document boundary at 0.14.0. You can now author complete business documents
from a blank package in Rust, including tables, rich headers and footers,
drawings, text boxes, watermarks, styles, numbering, and page-accurate fields.
This release also brings every Word fix reported and contributed since 0.13.1
to crates.io, together with prebuilt `rdocx` CLI binaries.

### Added

- Create Word-compatible DOCX, DOCM, DOTX, and DOTM packages with complete
  settings, properties, themes, font tables, embedded fonts, style graphs, and
  numbering definitions.
- Edit ordered sections, page geometry, rich per-section headers and footers,
  story text in every container, and generic content through transactional
  insert, move, clone, remove, and cross-document fragment import operations.
- Author complete tables, rows, and cells, ordered run content, rich HTML
  fragments in any story, pictures with crop and floating placement, text
  boxes, and section-aware watermarks. Measure paragraphs and tables at a
  caller width for equal-height layouts.
- `update_layout_backed_fields` and `update_page_fields` write PAGE, NUMPAGES,
  PAGEREF, and TOC caches from one deterministic layout, as proposed in
  [Issue 93](https://github.com/tensorbee/rdocx/issues/93) and
  [PR 114](https://github.com/tensorbee/rdocx/pull/114).
- Clone and remove table rows
  ([Issue 95](https://github.com/tensorbee/rdocx/issues/95),
  [PR 113](https://github.com/tensorbee/rdocx/pull/113)), replace an existing
  picture atomically
  ([Issue 96](https://github.com/tensorbee/rdocx/issues/96),
  [PR 107](https://github.com/tensorbee/rdocx/pull/107)), split runs at Unicode
  character offsets
  ([Issue 97](https://github.com/tensorbee/rdocx/issues/97),
  [PR 112](https://github.com/tensorbee/rdocx/pull/112)), and control
  `w:updateFields`
  ([Issue 98](https://github.com/tensorbee/rdocx/issues/98),
  [PR 104](https://github.com/tensorbee/rdocx/pull/104)).
- Insert pictures into any story with `insert_picture_to_story` and anchor
  comments inside table cells with `add_story_comment`, from
  [Issue 121](https://github.com/tensorbee/rdocx/issues/121). Native content
  index lookup, paragraph search, named highlight, shading, and style helpers
  back the Python round-two surface from
  [Issue 94](https://github.com/tensorbee/rdocx/issues/94) and PRs
  [108](https://github.com/tensorbee/rdocx/pull/108),
  [109](https://github.com/tensorbee/rdocx/pull/109),
  [110](https://github.com/tensorbee/rdocx/pull/110), and
  [111](https://github.com/tensorbee/rdocx/pull/111).
- Story items report their direct main-body owner
  ([Issue 86](https://github.com/tensorbee/rdocx/issues/86)), TOC rebuilding
  returns ordered diagnostics, and comments accept optional RFC 3339 dates
  ([Issue 117](https://github.com/tensorbee/rdocx/issues/117)).
- `rdocx-cli` gains schema-versioned comment, revision, comparison, TOC,
  structured text, layout, and guarded replacement commands from
  [Issue 76](https://github.com/tensorbee/rdocx/issues/76). Rust release tags
  attach checksummed CLI archives for six targets that `cargo binstall
  rdocx-cli` resolves, as requested in
  [Issue 100](https://github.com/tensorbee/rdocx/issues/100).

### Fixed

- Documents with drawing ids reused across parts open
  ([Issue 72](https://github.com/tensorbee/rdocx/issues/72)), unused root
  default namespaces save safely
  ([Issue 73](https://github.com/tensorbee/rdocx/issues/73)), and generated PDFs
  keep logical reading order
  ([Issue 74](https://github.com/tensorbee/rdocx/issues/74)).
- Document comparison preserves drawings
  ([Issue 75](https://github.com/tensorbee/rdocx/issues/75)), paragraphs with
  complex fields
  ([Issue 85](https://github.com/tensorbee/rdocx/issues/85),
  [PR 106](https://github.com/tensorbee/rdocx/pull/106)), and two or more
  appended paragraphs
  ([Issue 115](https://github.com/tensorbee/rdocx/issues/115)).
- Explicit false table toggles stay false
  ([Issue 83](https://github.com/tensorbee/rdocx/issues/83),
  [PR 101](https://github.com/tensorbee/rdocx/pull/101)), content-control type
  payloads survive save
  ([Issue 84](https://github.com/tensorbee/rdocx/issues/84),
  [PR 103](https://github.com/tensorbee/rdocx/pull/103)), and fractional line
  spacing values are accepted
  ([PR 122](https://github.com/tensorbee/rdocx/pull/122)).
- Word PNG output premultiplies straight-alpha pictures before compositing
  ([Issue 119](https://github.com/tensorbee/rdocx/issues/119)).
- Run-level page breaks start a new page
  ([Issue 88](https://github.com/tensorbee/rdocx/issues/88),
  [PR 102](https://github.com/tensorbee/rdocx/pull/102)), and header and footer
  pictures render from their own story part
  ([Issue 89](https://github.com/tensorbee/rdocx/issues/89),
  [PR 102](https://github.com/tensorbee/rdocx/pull/102)).
- TOC rebuilding accepts Word's default `\z` switch
  ([Issue 90](https://github.com/tensorbee/rdocx/issues/90),
  [PR 101](https://github.com/tensorbee/rdocx/pull/101)) and writes entries with
  localized styles, section-derived tab stops, and structural numbering tabs
  ([Issue 116](https://github.com/tensorbee/rdocx/issues/116)).
- Comment content types, identities, thread links, and dates survive save and
  reopen ([Issue 117](https://github.com/tensorbee/rdocx/issues/117)), and story
  reads are linear and include tracked insertions and inline content controls
  ([Issue 118](https://github.com/tensorbee/rdocx/issues/118)).
- Note-only edits keep paragraph cache reuse, and restart checkpoints survive
  body-length changes, restoring the editing performance reported in
  [Issue 69](https://github.com/tensorbee/rdocx/issues/69).
- Authored Word charts are portable across Word and Pages
  ([PR 71](https://github.com/tensorbee/rdocx/pull/71),
  [PR 123](https://github.com/tensorbee/rdocx/pull/123)), and reader and
  serializer behavior incorporates hardened equivalents of
  [PR 77](https://github.com/tensorbee/rdocx/pull/77),
  [PR 78](https://github.com/tensorbee/rdocx/pull/78),
  [PR 79](https://github.com/tensorbee/rdocx/pull/79), and
  [PR 80](https://github.com/tensorbee/rdocx/pull/80).

### Compatibility

The exact seven-package stable crates.io family moves together from 0.13.1 to
0.14.0. The selected set is `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`,
`rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli`. It depends on the
separately published shared OOXML 0.12.1 family. This release also satisfies
the crates.io publication request in
[Issue 99](https://github.com/tensorbee/rdocx/issues/99). The unpublished 0.13.2
crates.io train is superseded rather than backfilled.

This pre-1.0 minor release contains source-incompatible Rust changes, including
the following.

- `TocRebuildReport` replaces `diagnostic_count` with ordered `diagnostics` and
  is no longer `Copy`. Read `diagnostics.len()` for the count.
- `ParagraphItemRef::CommentRangeStart` and `CommentRangeEnd` are struct
  variants, and `BookmarkStart` and `BookmarkEnd` gained `has_child_content`.
  Update patterns that match these variants.
- `ListLevel` is no longer `Copy` or `Eq` and now has private fields, so
  `ListLevel { format, start: Some(n) }` literals become
  `ListLevel::new(format).start(n)`.
  `ListNumberFormat` is no longer `Copy` and now has 61 variants, so exhaustive
  matches must handle the new formats.
- `rdocx::Error` gained a `Story` variant, and exhaustive matches must handle
  it.
- `ChartData` and several `rdocx-oxml` `CT_*` model structs gained public
  fields, so full struct literals must add them. `CT_Anchor::from_xml` and
  `CT_Inline::from_xml` now take a namespace-aware `NsReader`.

`rdocx replace` now stages its result
and refuses an output path that already exists, including the input file. It
also checks an exact expected match count before publishing when one is given.
Scripts that replaced a document in place must write to a new output path.
Python, WASM, npm, and PyPI publication authority is unchanged, and
`rdocx-wasm@0.14.0` is not a crates.io package.

### Contributors

[@hadim](https://github.com/hadim) reported the real-document failures and
requests in Issues [72](https://github.com/tensorbee/rdocx/issues/72),
[73](https://github.com/tensorbee/rdocx/issues/73),
[74](https://github.com/tensorbee/rdocx/issues/74),
[75](https://github.com/tensorbee/rdocx/issues/75),
[76](https://github.com/tensorbee/rdocx/issues/76),
[83](https://github.com/tensorbee/rdocx/issues/83),
[84](https://github.com/tensorbee/rdocx/issues/84),
[85](https://github.com/tensorbee/rdocx/issues/85),
[86](https://github.com/tensorbee/rdocx/issues/86),
[88](https://github.com/tensorbee/rdocx/issues/88),
[89](https://github.com/tensorbee/rdocx/issues/89),
[90](https://github.com/tensorbee/rdocx/issues/90),
[93](https://github.com/tensorbee/rdocx/issues/93),
[95](https://github.com/tensorbee/rdocx/issues/95),
[96](https://github.com/tensorbee/rdocx/issues/96),
[97](https://github.com/tensorbee/rdocx/issues/97),
[98](https://github.com/tensorbee/rdocx/issues/98),
[99](https://github.com/tensorbee/rdocx/issues/99),
[100](https://github.com/tensorbee/rdocx/issues/100),
[115](https://github.com/tensorbee/rdocx/issues/115),
[116](https://github.com/tensorbee/rdocx/issues/116),
[117](https://github.com/tensorbee/rdocx/issues/117),
[118](https://github.com/tensorbee/rdocx/issues/118),
[119](https://github.com/tensorbee/rdocx/issues/119), and
[121](https://github.com/tensorbee/rdocx/issues/121), and requested the Python
editing surface those native helpers support in
[Issue 94](https://github.com/tensorbee/rdocx/issues/94).
They contributed the implementations reconciled from PRs
[101](https://github.com/tensorbee/rdocx/pull/101),
[102](https://github.com/tensorbee/rdocx/pull/102),
[103](https://github.com/tensorbee/rdocx/pull/103),
[104](https://github.com/tensorbee/rdocx/pull/104),
[106](https://github.com/tensorbee/rdocx/pull/106),
[107](https://github.com/tensorbee/rdocx/pull/107),
[108](https://github.com/tensorbee/rdocx/pull/108),
[109](https://github.com/tensorbee/rdocx/pull/109),
[110](https://github.com/tensorbee/rdocx/pull/110),
[111](https://github.com/tensorbee/rdocx/pull/111),
[112](https://github.com/tensorbee/rdocx/pull/112),
[113](https://github.com/tensorbee/rdocx/pull/113), and
[114](https://github.com/tensorbee/rdocx/pull/114).
[@pedroassumpcao](https://github.com/pedroassumpcao) contributed the reader and
serializer contracts in PRs [77](https://github.com/tensorbee/rdocx/pull/77),
[78](https://github.com/tensorbee/rdocx/pull/78),
[79](https://github.com/tensorbee/rdocx/pull/79), and
[80](https://github.com/tensorbee/rdocx/pull/80), and fractional line spacing in
[PR 122](https://github.com/tensorbee/rdocx/pull/122).
[@chevinbrown](https://github.com/chevinbrown) contributed portable authored
Word charts in [PR 71](https://github.com/tensorbee/rdocx/pull/71) and the
line-chart viewer defaults in [PR 123](https://github.com/tensorbee/rdocx/pull/123).
[@emptinessform](https://github.com/emptinessform) diagnosed and measured the
note-edit and restart performance regression in
[Issue 69](https://github.com/tensorbee/rdocx/issues/69). Atul Sharma
integrated, hardened, and released the family.

## py-rdocx-v0.14.0

### Highlights

`rdocx 0.14.0` for Python matches the native 0.14.0 Word family and makes the
binding a practical automation surface. Python can now insert, clone, move,
and remove content, format paragraphs and runs, edit stories and hyperlinks,
resolve revisions, update page fields from layout, and insert pictures, all
with typed, owner-checked operations.

### Added

- Indexed content insertion, cloning, movement, removal, and counted
  replacement, based on [PR 109](https://github.com/tensorbee/rdocx/pull/109)
  and [PR 111](https://github.com/tensorbee/rdocx/pull/111).
- Paragraph style and numbering plus run style, named Word highlights, and
  separate shading, based on
  [PR 108](https://github.com/tensorbee/rdocx/pull/108).
- Header, footer, and story text mutation, hyperlink creation, revision
  resolution, field updates, and exact story item XML, based on
  [PR 109](https://github.com/tensorbee/rdocx/pull/109) and
  [PR 110](https://github.com/tensorbee/rdocx/pull/110). Together these complete
  the round-two requests in
  [Issue 94](https://github.com/tensorbee/rdocx/issues/94).
- `Table.clone_row` and row removal, including Word files whose root declares
  an unused default namespace
  ([Issue 95](https://github.com/tensorbee/rdocx/issues/95),
  [PR 113](https://github.com/tensorbee/rdocx/pull/113)), picture replacement
  ([Issue 96](https://github.com/tensorbee/rdocx/issues/96),
  [PR 107](https://github.com/tensorbee/rdocx/pull/107)), run splitting
  ([Issue 97](https://github.com/tensorbee/rdocx/issues/97),
  [PR 112](https://github.com/tensorbee/rdocx/pull/112)), and the
  `update_fields_on_open` setting
  ([Issue 98](https://github.com/tensorbee/rdocx/issues/98),
  [PR 104](https://github.com/tensorbee/rdocx/pull/104)).
- `update_page_fields` and `update_layout_backed_fields` with an owned report
  ([Issue 93](https://github.com/tensorbee/rdocx/issues/93),
  [PR 114](https://github.com/tensorbee/rdocx/pull/114)), ordered TOC rebuild
  diagnostics ([Issue 90](https://github.com/tensorbee/rdocx/issues/90)), and
  direct body ownership on story items
  ([Issue 86](https://github.com/tensorbee/rdocx/issues/86)).
- Picture insertion from bytes at story positions, comment ranges inside table
  cells, and optional comment dates, from
  [Issue 121](https://github.com/tensorbee/rdocx/issues/121) and
  [Issue 117](https://github.com/tensorbee/rdocx/issues/117).

### Fixed

- `story_items` and `hyperlinks` run in linear time, and `Paragraph.text` and
  `Paragraph.runs` include tracked insertions and inline content controls
  ([Issue 118](https://github.com/tensorbee/rdocx/issues/118)).
- The package includes the native 0.14.0 fixes for comparison with drawing
  namespaces declared on ancestor elements and with complex fields
  ([Issue 75](https://github.com/tensorbee/rdocx/issues/75),
  [Issue 85](https://github.com/tensorbee/rdocx/issues/85)), appended
  paragraphs ([Issue 115](https://github.com/tensorbee/rdocx/issues/115)),
  straight-alpha pictures in PNG output
  ([Issue 119](https://github.com/tensorbee/rdocx/issues/119)), table toggles
  ([Issue 83](https://github.com/tensorbee/rdocx/issues/83),
  [PR 101](https://github.com/tensorbee/rdocx/pull/101)), content controls
  ([Issue 84](https://github.com/tensorbee/rdocx/issues/84),
  [PR 103](https://github.com/tensorbee/rdocx/pull/103)), run page breaks and
  header pictures
  ([Issue 88](https://github.com/tensorbee/rdocx/issues/88),
  [Issue 89](https://github.com/tensorbee/rdocx/issues/89),
  [PR 102](https://github.com/tensorbee/rdocx/pull/102)), TOC entries
  ([Issue 116](https://github.com/tensorbee/rdocx/issues/116)), comment identity
  ([Issue 117](https://github.com/tensorbee/rdocx/issues/117)), comparison with
  complex fields ([PR 106](https://github.com/tensorbee/rdocx/pull/106)), and
  fractional line spacing
  ([PR 122](https://github.com/tensorbee/rdocx/pull/122)).

### Compatibility

The distribution and import name remains `rdocx`, now at 0.14.0 to match the
native crate. It requires Python 3.9 or newer through six `cp39-abi3` platform
wheels and one source distribution. Existing 0.13.2 document calls remain source
compatible, and the direct-body `RunPosition` call shape is unchanged. The one
incompatible signature change is the `TocRebuildReport` constructor, which now
takes `diagnostics` instead of `diagnostic_count`. Its `diagnostic_count`
property remains. `Paragraph.runs` now includes runs inside tracked insertions
and inline content controls, so code that stored run indexes from 0.13.2 should
look them up again. This release does not publish `rpptx`, any crates.io
package, WASM package, or npm package.

### Contributors

[@hadim](https://github.com/hadim) drove the round-two and round-three Python
surface through Issues [86](https://github.com/tensorbee/rdocx/issues/86),
[90](https://github.com/tensorbee/rdocx/issues/90),
[93](https://github.com/tensorbee/rdocx/issues/93),
[94](https://github.com/tensorbee/rdocx/issues/94),
[95](https://github.com/tensorbee/rdocx/issues/95),
[96](https://github.com/tensorbee/rdocx/issues/96),
[97](https://github.com/tensorbee/rdocx/issues/97),
[98](https://github.com/tensorbee/rdocx/issues/98),
[117](https://github.com/tensorbee/rdocx/issues/117),
[118](https://github.com/tensorbee/rdocx/issues/118), and
[121](https://github.com/tensorbee/rdocx/issues/121), and reported the native
defects in Issues [75](https://github.com/tensorbee/rdocx/issues/75),
[83](https://github.com/tensorbee/rdocx/issues/83),
[84](https://github.com/tensorbee/rdocx/issues/84),
[85](https://github.com/tensorbee/rdocx/issues/85),
[88](https://github.com/tensorbee/rdocx/issues/88),
[89](https://github.com/tensorbee/rdocx/issues/89),
[115](https://github.com/tensorbee/rdocx/issues/115),
[116](https://github.com/tensorbee/rdocx/issues/116), and
[119](https://github.com/tensorbee/rdocx/issues/119). They contributed the
implementations reconciled from PRs
[101](https://github.com/tensorbee/rdocx/pull/101),
[102](https://github.com/tensorbee/rdocx/pull/102),
[103](https://github.com/tensorbee/rdocx/pull/103),
[104](https://github.com/tensorbee/rdocx/pull/104),
[106](https://github.com/tensorbee/rdocx/pull/106),
[107](https://github.com/tensorbee/rdocx/pull/107),
[108](https://github.com/tensorbee/rdocx/pull/108),
[109](https://github.com/tensorbee/rdocx/pull/109),
[110](https://github.com/tensorbee/rdocx/pull/110),
[111](https://github.com/tensorbee/rdocx/pull/111),
[112](https://github.com/tensorbee/rdocx/pull/112),
[113](https://github.com/tensorbee/rdocx/pull/113), and
[114](https://github.com/tensorbee/rdocx/pull/114).
[@pedroassumpcao](https://github.com/pedroassumpcao) contributed fractional
line spacing in [PR 122](https://github.com/tensorbee/rdocx/pull/122). Atul
Sharma integrated, hardened, and released the distribution.

## py-rpptx-v0.12.1

### Highlights

This patch recovery follows the failed immutable `rpptx-v0.12.0` Rust tag,
which published no crates and created no GitHub release. It keeps the Python
distribution aligned with the recovered native family.

`rpptx 0.12.1` for Python matches the native 0.12.1 presentation family and
adds notes editing, shape inspection, and formatting-preserving text
replacement, together with correct transparency and notes rendering.

### Added

- Set speaker notes text through `Slide.notes_text` while keeping the notes
  shape identity and run formatting, requested in
  [Issue 121](https://github.com/tensorbee/rdocx/issues/121) and
  [Issue 120](https://github.com/tensorbee/rdocx/issues/120).
- Inspect shape geometry and identity, run font details, and autofit mode, and
  replace run text without losing formatting, from
  [Issue 121](https://github.com/tensorbee/rdocx/issues/121).

### Fixed

- Notes render when the notes slide omits its reverse relationship
  ([Issue 120](https://github.com/tensorbee/rdocx/issues/120)).
- Picture transparency is applied in every output
  ([Issue 91](https://github.com/tensorbee/rdocx/issues/91),
  [PR 105](https://github.com/tensorbee/rdocx/pull/105)), straight-alpha
  pictures composite correctly, and the decoded-pixel render limit rises from
  16 MiB to 64 MiB, so larger pictures render instead of being skipped
  ([Issue 119](https://github.com/tensorbee/rdocx/issues/119)).
- Slide-owned date, footer, and slide-number placeholders stay visible under
  master header and footer flags
  ([Issue 92](https://github.com/tensorbee/rdocx/issues/92)).

### Compatibility

The distribution and import name remains `rpptx`, now at 0.12.1 to match the
native crate. It requires Python 3.9 or newer through six `cp39-abi3` platform
wheels and one source distribution. Existing 0.11.0 calls remain source
compatible. This release does not publish `rdocx`, any crates.io package, WASM
package, or npm package.

### Contributors

[@hadim](https://github.com/hadim) reported the transparency, header and footer
flag, raster, notes, and binding gaps in Issues
[91](https://github.com/tensorbee/rdocx/issues/91),
[92](https://github.com/tensorbee/rdocx/issues/92),
[119](https://github.com/tensorbee/rdocx/issues/119),
[120](https://github.com/tensorbee/rdocx/issues/120), and
[121](https://github.com/tensorbee/rdocx/issues/121), and contributed the
transparency implementation in
[PR 105](https://github.com/tensorbee/rdocx/pull/105). Atul Sharma integrated,
hardened, and released the distribution.

## py-rpptx-v0.12.0

### Highlights

`rpptx 0.12.0` for Python matches the native 0.12.0 presentation family and
adds notes editing, shape inspection, and formatting-preserving text
replacement, together with correct transparency and notes rendering.

### Added

- Set speaker notes text through `Slide.notes_text` while keeping the notes
  shape identity and run formatting, requested in
  [Issue 121](https://github.com/tensorbee/rdocx/issues/121) and
  [Issue 120](https://github.com/tensorbee/rdocx/issues/120).
- Inspect shape geometry and identity, run font details, and autofit mode, and
  replace run text without losing formatting, from
  [Issue 121](https://github.com/tensorbee/rdocx/issues/121).

### Fixed

- Notes render when the notes slide omits its reverse relationship
  ([Issue 120](https://github.com/tensorbee/rdocx/issues/120)).
- Picture transparency is applied in every output
  ([Issue 91](https://github.com/tensorbee/rdocx/issues/91),
  [PR 105](https://github.com/tensorbee/rdocx/pull/105)), straight-alpha
  pictures composite correctly, and the decoded-pixel render limit rises from
  16 MiB to 64 MiB, so larger pictures render instead of being skipped
  ([Issue 119](https://github.com/tensorbee/rdocx/issues/119)).
- Slide-owned date, footer, and slide-number placeholders stay visible under
  master header and footer flags
  ([Issue 92](https://github.com/tensorbee/rdocx/issues/92)).

### Compatibility

The distribution and import name remains `rpptx`, now at 0.12.0 to match the
native crate. It requires Python 3.9 or newer through six `cp39-abi3` platform
wheels and one source distribution. Existing 0.11.0 calls remain source
compatible. This release does not publish `rdocx`, any crates.io package, WASM
package, or npm package.

### Contributors

[@hadim](https://github.com/hadim) reported the transparency, header and footer
flag, raster, notes, and binding gaps in Issues
[91](https://github.com/tensorbee/rdocx/issues/91),
[92](https://github.com/tensorbee/rdocx/issues/92),
[119](https://github.com/tensorbee/rdocx/issues/119),
[120](https://github.com/tensorbee/rdocx/issues/120), and
[121](https://github.com/tensorbee/rdocx/issues/121), and contributed the
transparency implementation in
[PR 105](https://github.com/tensorbee/rdocx/pull/105). Atul Sharma integrated,
hardened, and released the distribution.

## py-rdocx-v0.13.2

### Highlights

`rdocx 0.13.2` adds the complete user-facing PyPI project page that was absent
from the immutable first Python release. The distribution now publishes its
README, installation and quick-start guidance, capability summary, typing
contract, classifiers, and project links alongside the native Word automation
surface.

### Added

- The PyPI long description documents installation, DOCX creation and editing,
  comparison, comments, stories, deterministic layout, PDF and image output,
  type checking, compatibility boundaries, support links, and licensing.
- Package metadata now identifies the project author, focused search keywords,
  Python and Rust classifiers, and direct homepage, repository, issue tracker,
  and changelog links.
- The release contains six platform wheels and one source distribution for
  CPython 3.9 and newer through `cp39-abi3`.

### Fixed

- Artifact validation now rejects a wheel or source distribution whose
  embedded metadata lacks the reviewed Markdown description, project summary,
  author, keywords, classifiers, or project links.
- The latest package retains the direct fixes for
  [Issue 72](https://github.com/tensorbee/rdocx/issues/72),
  [Issue 73](https://github.com/tensorbee/rdocx/issues/73),
  [Issue 74](https://github.com/tensorbee/rdocx/issues/74), and
  [Issue 75](https://github.com/tensorbee/rdocx/issues/75), plus the Python
  publication outcome requested in
  [Issue 76](https://github.com/tensorbee/rdocx/issues/76).
- Reader and serializer behavior retains hardened equivalents of
  [PR 77](https://github.com/tensorbee/rdocx/pull/77),
  [PR 78](https://github.com/tensorbee/rdocx/pull/78),
  [PR 79](https://github.com/tensorbee/rdocx/pull/79), and
  [PR 80](https://github.com/tensorbee/rdocx/pull/80).

### Compatibility

The distribution and import name remains `rdocx`. It requires Python 3.9 or
newer and uses the stable ABI from Python 3.9. The native `rdocx`, binding
crate, built artifacts, release tag, and PyPI version are aligned at 0.13.2.
This packaging correction changes no Python runtime API. The immutable
`rdocx 0.13.1` release remains available.

### Contributors

[@hadim](https://github.com/hadim) reported the real-document failures in
Issues [72](https://github.com/tensorbee/rdocx/issues/72),
[73](https://github.com/tensorbee/rdocx/issues/73),
[74](https://github.com/tensorbee/rdocx/issues/74), and
[75](https://github.com/tensorbee/rdocx/issues/75), and requested the Python
publication surface in
[Issue 76](https://github.com/tensorbee/rdocx/issues/76).
[@pedroassumpcao](https://github.com/pedroassumpcao) contributed the reader and
serializer contracts retained through PRs
[77](https://github.com/tensorbee/rdocx/pull/77),
[78](https://github.com/tensorbee/rdocx/pull/78),
[79](https://github.com/tensorbee/rdocx/pull/79), and
[80](https://github.com/tensorbee/rdocx/pull/80). Atul Sharma integrated the
native, binding, metadata-validation, and release paths.

## py-rdocx-v0.13.1

### Highlights

The first `rdocx` Python release brings the distinctive Word automation
surface to PyPI at the same 0.13.1 version as the native `rdocx` crate. It
supports CPython 3.9 and newer through `cp39-abi3` wheels.

### Added

- Comparison, comments and replies, revision filtering and resolution,
  deterministic layout snapshots, TOC rebuilding, ordered section and style
  inspection, rich header and footer stories, and relationship-safe
  hyperlinks are available from Python.
- Installed type declarations are checked with strict mypy and stubtest. The
  release contains six platform wheels and one source distribution.

### Fixed

- Documents now accept drawing identifiers scoped per package part, retain
  unused Word default namespaces safely, preserve drawings through comparison,
  and emit searchable PDF text in logical reading order. These changes resolve
  [Issue 72](https://github.com/tensorbee/rdocx/issues/72),
  [Issue 73](https://github.com/tensorbee/rdocx/issues/73),
  [Issue 74](https://github.com/tensorbee/rdocx/issues/74), and
  [Issue 75](https://github.com/tensorbee/rdocx/issues/75).
- Reader and serializer behavior incorporates hardened equivalents of
  [PR 77](https://github.com/tensorbee/rdocx/pull/77),
  [PR 78](https://github.com/tensorbee/rdocx/pull/78),
  [PR 79](https://github.com/tensorbee/rdocx/pull/79), and
  [PR 80](https://github.com/tensorbee/rdocx/pull/80).
- Publication rejects a partial or mixed project artifact set, mismatched name
  or version metadata, a non-ABI3 wheel, manual-dispatch publication, and any
  release that lacks fresh approval at the reviewed commit.

### Compatibility

The distribution and import name is `rdocx`. It requires Python 3.9 or newer
and uses the stable ABI from Python 3.9. Existing locally built binding APIs
remain source compatible. This release does not publish or change `rpptx`, any
crates.io package, WASM package, or npm package.

### Contributors

[@hadim](https://github.com/hadim) reported the four real-document failures in
Issues [72](https://github.com/tensorbee/rdocx/issues/72),
[73](https://github.com/tensorbee/rdocx/issues/73),
[74](https://github.com/tensorbee/rdocx/issues/74), and
[75](https://github.com/tensorbee/rdocx/issues/75). They also identified the
missing high-value Python surface and made the case for PyPI publication in
[Issue 76](https://github.com/tensorbee/rdocx/issues/76).
[@pedroassumpcao](https://github.com/pedroassumpcao) contributed the reader and
serializer contracts in PRs [77](https://github.com/tensorbee/rdocx/pull/77),
[78](https://github.com/tensorbee/rdocx/pull/78),
[79](https://github.com/tensorbee/rdocx/pull/79), and
[80](https://github.com/tensorbee/rdocx/pull/80). Atul Sharma integrated and
hardened the native, CLI, binding, and release paths.

## py-rpptx-v0.11.0

### Highlights

The first `rpptx` Python release brings presentation automation to PyPI at the
same 0.11.0 version as the native `rpptx` crate. It supports CPython 3.9 and
newer through `cp39-abi3` wheels.

### Added

- Deterministic presentation PDF and slide PNG output, speaker notes as text,
  PDF, and PNG, plus modern comment authors, threads, replies, and ordered
  comment movement are available from Python.
- The PyPI long description documents installation, quick-start examples,
  presentation editing, rendering, notes, comments, type checking,
  compatibility boundaries, support links, and licensing. Project metadata
  also supplies the author, keywords, classifiers, and direct project URLs.
- Installed type declarations are checked with strict mypy and stubtest. The
  release contains six platform wheels and one source distribution.

### Fixed

- The Python binding crate, project metadata, built artifacts, release tag,
  and PyPI version now agree with the native `rpptx` 0.11.0 version.
- Artifact validation rejects a wheel or source distribution whose embedded
  metadata lacks the reviewed Markdown description or project links.
- Publication rejects a partial or mixed project artifact set, mismatched name
  or version metadata, a non-ABI3 wheel, manual-dispatch publication, and any
  release that lacks fresh approval at the reviewed commit.

### Compatibility

The distribution and import name is `rpptx`. It requires Python 3.9 or newer
and uses the stable ABI from Python 3.9. Existing locally built binding APIs
remain source compatible. This release does not publish or change `rdocx`, any
crates.io package, WASM package, or npm package.

### Contributors

[@hadim](https://github.com/hadim) identified the missing presentation binding
surface and made the case for publishing `rpptx` on PyPI in
[Issue 76](https://github.com/tensorbee/rdocx/issues/76). Atul Sharma
implemented and hardened the native binding and release paths.

## v0.13.1

### Highlights

The complete stable Word family publishes the M22 Word-depth boundary at
0.13.1. This patch release recovers from the immutable v0.13.0 attempt, which
published only five low-level stable packages before registry verification
stopped the workflow.

### Added

- Publish the complete DOCX, DOCM, DOTX, and DOTM package identity and bounded
  Flat OPC interchange surface.
- Publish bounded MHTML import and export with contained image and link
  resolution, deterministic output, and stable loss diagnostics.
- Publish the M22 OfficeMath, field, table-of-contents, merge, comparison,
  executable-content, glossary, and package-story outcomes in one coherent
  stable family.

### Fixed

- Pin the stable family to the separately published shared OOXML 0.11.0 family,
  which contains the Word main content-type constants required by `rdocx`.
- Replace the incomplete v0.13.0 publication boundary with a new coherent patch
  family without moving its tag or altering its five immutable registry entries.

### Compatibility

The exact seven-package stable crates.io family moves together from the last
complete 0.12.0 release to 0.13.1. The selected set is `rdocx-opc`,
`rdocx-oxml`, `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and
`rdocx-cli`. It depends on the separately published shared OOXML 0.11.0 family.

The native Rust APIs remain additive pre-1.0 surfaces. The immutable v0.13.0
attempt and its five published packages remain available as historical release
evidence. Python, WASM, npm, and PyPI publication authority is unchanged, and
`rdocx-wasm@0.13.1` is not a crates.io package.

### Contributors

Atul Sharma maintained the release. No external issue or pull request belongs
to the selected stable-family changes since `v0.12.0`, so this release has no
external contribution notification.

## rpptx-v0.11.0

### Highlights

The shared OOXML and PowerPoint family moves to 0.11.0 with the Word main
content-type constants required by modern stable package-class handling.
`oxml-opc` now names DOCX, DOCM, DOTX, and DOTM main parts through one shared
vocabulary.

### Added

- Add `WORD_DOCUMENT`, `WORD_DOCUMENT_MACRO_ENABLED`, `WORD_TEMPLATE`, and
  `WORD_TEMPLATE_MACRO_ENABLED` to `oxml-opc`.
- Cover the four Word main content-type constants in the shared relationship
  vocabulary regression.

### Fixed

- Publish the shared contract required by packaged stable `rdocx`, which the
  immutable shared 0.10.0 archive does not contain.

### Compatibility

The exact 15-package shared OOXML and PowerPoint crates.io family moves
together from 0.10.0 to 0.11.0. The constants are additive pre-1.0 Rust API.
The stable Word family remains at 0.13.0 during this selected-family release.
Its immutable v0.13.0 attempt published five low-level packages before
registry verification exposed the missing shared constants.

Stable Word, Python, WASM, npm, and PyPI packages are outside this release's
publication authority. `rpptx-wasm@0.11.0` is not a crates.io package.

### Contributors

Atul Sharma maintained the release. No external issue or pull request belongs
to the selected family changes since `rpptx-v0.10.0`, so this release has no
external contribution notification.

## v0.13.0

### Highlights

The stable Word family completes the M22 Word-depth boundary at 0.13.0.
Documents can author and render OfficeMath, rebuild fields and tables of
contents, perform advanced mail merge and comparison, inspect executable and
building-block content, preserve modern package identity, and exchange bounded
Flat OPC and MHTML through the native Rust facade.

### Added

- Add OfficeMath modeling, authoring, deterministic rendering, and MathML and
  LaTeX conversion for the supported equation subset.
- Add field evaluation and updates, dynamic table-of-contents rebuilding,
  sectioned mail merge, and body, header, footer, footnote, and endnote
  comparison with explicit revision policy.
- Add exact VBA, OLE, and ActiveX inventory and mutation, plus glossary,
  building-block, and package-story access without executing embedded content.
- Add DOCX, DOCM, DOTX, and DOTM package identity, output-only class
  conversion, and bounded Flat OPC import and export with payload retention.
- Add bounded MHTML import and export with contained image and link resolution,
  deterministic MIME output, and location-aware loss diagnostics.

### Fixed

- Reject malformed XML lexical forms through the shared `oxml-core` validator
  while preserving each Word owner's established error and rollback surface.
- Preserve unsupported XML, namespace bindings, relationship ownership,
  executable bytes, template semantics, and package-signature invalidation
  evidence through the new editing and interchange paths.
- Keep unsafe, external, ambiguous, malformed, or over-limit MHTML and Flat OPC
  resources from publishing a partial document.

### Compatibility

The exact seven-package stable crates.io family moves together from 0.12.0 to
0.13.0. The selected set is `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`,
`rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli`. It depends on the
separately published shared OOXML 0.10.0 family.

The native Rust APIs are additive pre-1.0 surfaces. Package-class conversion
changes only output identity and does not remove macros. Flat OPC output uses a
deterministic container representation while preserving unsupported inner XML
and binary payload bytes. MHTML supports the documented contained PNG and JPEG
resource subset and never fetches external content. Python, WASM, npm, and PyPI
publication authority is unchanged, and `rdocx-wasm@0.13.0` is not a crates.io
package.

### Contributors

Atul Sharma maintained the release. No external issue or pull request belongs
to the selected stable-family changes since `v0.12.0`, so this release has no
external contribution notification.

## rpptx-v0.10.0

### Highlights

The shared OOXML and PowerPoint family moves to 0.10.0 with one strict XML
1.0 lexical validator owned by `oxml-core`. Format-specific readers can share
declaration, character, name, namespace, reference, comment, and processing
instruction checks while keeping their schema rules and public errors local.

### Added

- Add `XmlLexicalError` and `validate_strict_xml_1_0` to `oxml-core` for
  bounded format-neutral lexical validation.
- Add baseline-aware inline groups to `oxml-layout`, allowing shared layout
  consumers to carry exact ascent and descent without teaching a backend
  document grammar.
- Add the Word glossary content type and relationship constants to `oxml-opc`
  so package owners can resolve glossary parts through shared vocabulary.

### Fixed

- Reject malformed declarations, forbidden XML 1.0 characters, invalid names
  and namespace bindings, duplicate expanded attributes, invalid references,
  malformed comments, and reserved processing instruction targets through one
  reviewed policy.

### Compatibility

The exact 15-package shared OOXML and PowerPoint crates.io family moves
together from 0.9.0 to 0.10.0. The selected set is `oxml-core`, `oxml-opc`,
`oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`, `oxml-sml`,
`oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`, `rpptx-layout`,
`rpptx-render`, `rpptx`, and `rpptx-cli`.

The lexical validator and glossary constants are additive pre-1.0 APIs.
Callers constructing `InlineItem::Group` or `LineItem::Group` literals must
initialize the new `baseline` field. `None` preserves the established
top-aligned group behavior.

The stable Word family remains at 0.12.0 and current source now pins shared
dependencies to 0.10.0. It is outside this release's publication authority.
Python, WASM, npm, and PyPI publication authority is unchanged, and
`rpptx-wasm@0.10.0` is not a crates.io package.

### Contributors

Atul Sharma maintained the release. No external issue or pull request belongs
to the selected family changes since `rpptx-v0.9.0`, so this release has no
external contribution notification.

## v0.12.0

### Highlights

The stable Word family moves to 0.12.0 with richer relationship-safe reader
facts and bounded restart pagination for ordinary prose. Warm layouts now keep
their complete prefix across note references and page-spanning paragraphs
without the document-wide second pagination pass reported after v0.11.1.

### Added

- Expose namespace-aware hyperlink targets, external-image relationships, and
  drawing safety facts shaped by [PR
  61](https://github.com/tensorbee/rdocx/pull/61).
- Expose document, table, row-grid, border, formatting, and retained-property
  completeness facts shaped by [PR
  62](https://github.com/tensorbee/rdocx/pull/62).
- Expose numbering identity, level metadata, and effective paragraph and run
  formatting shaped by [PR
  63](https://github.com/tensorbee/rdocx/pull/63).
- Expose bounded nested-revision projection, preserved insertion facts, and
  ordered complex-field display segments shaped by [PR
  64](https://github.com/tensorbee/rdocx/pull/64).

### Fixed

- Keep direct body paragraphs with footnote or endnote references eligible for
  exact cache reuse while retaining note-part invalidation, addressing [Issue
  65](https://github.com/tensorbee/rdocx/issues/65).
- Admit ordinary multi-line prose and complete block-boundary restart records
  under the shared aggregate cache budget, addressing [Issue
  66](https://github.com/tensorbee/rdocx/issues/66).
- Keep the completed recorded pagination pass when a paragraph spans a page,
  then publish only later complete-boundary checkpoints, addressing [Issue
  67](https://github.com/tensorbee/rdocx/issues/67).

### Compatibility

The exact seven-package stable crates.io family moves together from 0.11.1 to
0.12.0. The selected set is `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`,
`rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli`. It depends on the
separately published shared OOXML 0.9.0 family.

Native reader additions use existing non-exhaustive or additive pre-1.0
surfaces. Full low-level `rdocx-oxml` struct literals written against 0.11.1
must initialize the new preservation fields or use the existing constructors
and `Default` implementations. The cache and pagination fixes require no
migration and leave rendered output unchanged. Python, WASM, npm, and PyPI
publication authority is unchanged, and `rdocx-wasm@0.12.0` is not a crates.io
package.

### Contributors

Thanks to `@pedroassumpcao` for the relationship-safe hyperlink and drawing
reader design in [PR 61](https://github.com/tensorbee/rdocx/pull/61), document
and table completeness design in [PR
62](https://github.com/tensorbee/rdocx/pull/62), numbering and effective
formatting design in [PR 63](https://github.com/tensorbee/rdocx/pull/63), and
tracked insertion and field safety design in [PR
64](https://github.com/tensorbee/rdocx/pull/64).

Thanks to `@emptinessform` for the note-reference cache report in [Issue
65](https://github.com/tensorbee/rdocx/issues/65), ordinary-prose restart report
in [Issue 66](https://github.com/tensorbee/rdocx/issues/66), and page-spanning
paragraph regression report in [Issue
67](https://github.com/tensorbee/rdocx/issues/67).

No named external patch landed directly. Each contribution landed through a
reviewed hardened equivalent that retains namespace identity, raw XML,
bounded work, exact warm and fresh equality, and compatibility contracts. The
four pull requests and Issues 65 and 66 remain closed after their release-bound
thank-yous. Issue 67 remains open after its release-bound thank-you.

## rpptx-v0.9.0

### Highlights

The shared OOXML and PowerPoint family completes the M21 presentation-depth
boundary at 0.9.0. Native Rust callers can inspect and edit modern
collaboration, timing, media, SmartArt, embedded-content, security, and package
variant state. They can also exchange presentations through bounded ODP, HTML,
and PDF workflows and export deterministic animations, notes, and handouts.

### Added

- Read and author modern comments, threaded replies, sections, slide numbers,
  dates, footers, and notes-master and handout-master settings.
- Inspect, create, and invalidate presentation signatures, and read or write
  password-protected packages through explicit optional security features.
- Model animation timing, transitions, morph metadata, audio and video,
  posters, playback settings, and deterministic timeline state.
- Export animated GIF or Motion JPEG AVI output with bounded frame rate,
  dimensions, duration, and media fallback policy.
- Inspect and edit the supported SmartArt data and layout subset, retain
  unsupported diagram content, and render six pinned layout families.
- Inventory, extract, replace, and remove relationship-owned OLE, ActiveX, and
  VBA payloads without executing them.
- Read and write the declared ODP subset, preserve PPTX, PPTM, POTX, POTM,
  PPSX, and PPSM package identity, and export deterministic notes pages and all
  six handout grids.
- Import a bounded HTML5 and CSS subset as editable slide content with pinned
  Chrome structure and render comparisons.
- Import PDF pages as preserved graphics or as the declared editable text,
  image, path, and URI-link subset with strict resource limits and pinned
  Poppler render comparisons.

### Fixed

- Preserve relationship ownership, raw unsupported XML, executable payloads,
  package signatures, and source package class across the new mutation and
  conversion paths.
- Keep unsupported HTML, PDF, SmartArt, media, animation, and interchange
  content explicit through stable diagnostics or retained opaque content
  instead of silently approximating it.
- Keep ordinary static rendering and the 49-entry deterministic output harness
  unchanged while adding the new presentation-depth paths.

### Compatibility

The exact 15-package shared OOXML and PowerPoint crates.io family moves
together from 0.8.0 to 0.9.0. The selected set is `oxml-core`, `oxml-opc`,
`oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`, `oxml-sml`,
`oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`, `rpptx-layout`,
`rpptx-render`, `rpptx`, and `rpptx-cli`.

The native Rust facade and model additions are additive pre-1.0 APIs. Existing
callers need no migration unless they opt into the new methods. HTML import
uses the existing default-template boundary and PDF import uses the existing
render boundary. The stable Word family remains at 0.11.1. Python, WASM, npm,
and PyPI publication authority is unchanged, and `rpptx-wasm@0.9.0` is not a
crates.io package.

### Contributors

Atul Sharma maintained the release. No external issue or pull request belongs
to the selected PowerPoint family for this release.

## v0.11.1

### Highlights

The stable Word family completes the S58 release at 0.11.1 after the immutable
partial v0.11.0 attempt published only `rdocx-opc@0.11.0` and
`rdocx-oxml@0.11.0`. It adds language-aware conditional hyphenation,
multi-script shaping, bidirectional paragraph and run layout, logical text
extraction over visually ordered output, and bounded restart pagination for
ordinary note, header, and footer workloads. The complete stable family now
uses the published shared 0.8.0 family.

### Added

- Apply Word automatic-hyphenation settings and run languages for English,
  French, German, and Spanish without assigning source ranges to generated
  hyphens.
- Shape Arabic, Devanagari, Thai, and Simplified Chinese with deterministic
  font selection, cluster-safe breaking, complete glyph offsets, and stable
  source mapping.
- Carry Word paragraph and run direction into paragraph-wide UAX 9 resolution,
  line-local visual ordering, and logical PDF and SVG extraction.
- Reuse bounded restart pagination for unchanged notes, headers, and footers.
  This hardened equivalent addresses the 700-paragraph note and header or
  footer workloads reported in
  [Issue 53](https://github.com/tensorbee/rdocx/issues/53).
- Avoid the redundant retained-context font byte comparison after the font
  manager has already accepted the exact ordered font set. This hardened
  equivalent addresses the 22 MiB caller-font workload reported in
  [Issue 54](https://github.com/tensorbee/rdocx/issues/54).
- Accept exact whole-valued decimal table measurements while rejecting
  fractional, exponent, overflow, unit-bearing, and malformed forms. This
  hardened equivalent includes the outcome proposed in
  [PR 55](https://github.com/tensorbee/rdocx/pull/55).
- Preserve tracked table-grid history as inert revision metadata while keeping
  the active grid as the only layout input. This hardened equivalent includes
  the outcome proposed in
  [PR 56](https://github.com/tensorbee/rdocx/pull/56).
- Classify the narrow enabled legacy VML horizontal-rule form for native
  inspection while retaining its exact raw XML. This hardened equivalent
  includes the outcome proposed in
  [PR 57](https://github.com/tensorbee/rdocx/pull/57).
- Prime locked Cargo dependencies before the intentionally offline Word
  fidelity harness. This hardened equivalent includes the locked Word fidelity
  dependency preparation proposed in
  [PR 58](https://github.com/tensorbee/rdocx/pull/58).

### Fixed

- Preserve exact warm and fresh layout equality when related stories, note
  references, language, hyphenation, direction, source-less generated content,
  or caller fonts participate in cache identity.
- Preserve logical searchable text while visually positioning mixed-direction
  rich runs, inline objects, numbering markers, stored fields, tab leaders,
  and generated conditional hyphens.
- Preserve namespace-aware unsupported table and run XML, including ancestor
  bindings, without changing schema child order or treating historical grids
  as active layout data.
- Recover the complete stable registry family without moving or deleting the
  immutable v0.11.0 tag. The five stable packages absent from that partial
  attempt are published only as part of the complete 0.11.1 family.

### Compatibility

The exact seven-package stable crates.io family moves together from 0.10.1 to
0.11.1 and requires the separately published shared 0.8.0 family. The stable
set remains `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, and `rdocx-cli`. Python, WASM, npm, and PyPI publication
authority is unchanged.

Low-level Rust callers that use full `CT_RPr`, `LayoutInput`, `CT_PPr`,
`CT_TblGrid`, `TextSegment`, and positioned-layout struct literals must
initialize the new language, automatic-hyphenation, direction, preservation,
and source-mapping fields or use existing defaults and constructors. In
particular, callers constructing full `TextSegment` literals must initialize
the `direction` field. These are intentional pre-1.0 source changes. The native
facade additions are additive, and existing binding method names remain
unchanged. The legacy VML horizontal rule is classified for inspection and
preservation, not rendered.

The immutable partial v0.11.0 attempt contains exactly `rdocx-opc@0.11.0` and
`rdocx-oxml@0.11.0`. It has no GitHub release and receives no contribution
notification. Recovery cleanup may yank those two incomplete entries only
after all seven 0.11.1 packages and the release body verify independently.

### Contributors

Atul Sharma maintained the release. Thanks to authenticated reporter
`@emptinessform` for the note and header or footer restart evidence in
[Issue 53](https://github.com/tensorbee/rdocx/issues/53), including the
corrected attribution and independent page-count ceiling, and for the isolated
caller-font byte-comparison evidence and unsound shallow-comparison caveat in
[Issue 54](https://github.com/tensorbee/rdocx/issues/54). Both final fixes are
hardened equivalents. Both issues remain open after their release-bound
thank-yous.

Thanks to authenticated contributor `@pedroassumpcao` for the whole-valued
decimal table-measurement case in
[PR 55](https://github.com/tensorbee/rdocx/pull/55) at
`056d48fdf23f35e3538ef3d6ff78cf9e3863e3a5`, tracked table-grid history in
[PR 56](https://github.com/tensorbee/rdocx/pull/56) at
`8b79c4cd0452defafe0a58e86b332c98e7fe52d7`, the legacy VML reader
classification in [PR 57](https://github.com/tensorbee/rdocx/pull/57) at
`44498f042a2290ef40c7a6c26025f38e38e9ce2a`, and locked Word fidelity
dependency preparation in
[PR 58](https://github.com/tensorbee/rdocx/pull/58) at
`c8fed1d1268fd765d602bac2da6524900c1c1cfd`. All four outcomes are hardened
equivalents. All four pull requests remain open after their release-bound
thank-yous.

No named external patch landed directly. Each named report or proposal landed
through the hardened equivalent described above so that namespace, exact
lexical parsing, raw preservation, bounded cache identity, and offline oracle
contracts remain intact.

## rpptx-v0.8.0

### Highlights

The shared OOXML and PowerPoint family publishes the complete text-direction
contract required by current stable Word source. This release supplies the
registry boundary needed for the stable 0.11.1 recovery after the immutable
partial v0.11.0 attempt stopped during `rdocx-layout` verification.

### Added

- Carry resolved text direction in the shared `TextSegment.direction` field so
  Word, PowerPoint, PDF, raster, and SVG paths use one reviewed direction
  contract.
- Preserve paragraph-wide bidirectional levels, line-local reordering, logical
  searchable text, inline objects, tab leaders, numbering markers, stored
  fields, and conditional hyphens across the shared layout path.
- Retain the deterministic Arabic, Devanagari, Thai, and Simplified Chinese
  font and shaping substrate published in the previous shared family.

### Fixed

- Make the current stable layout source compile against a complete published
  shared family instead of relying on an API added after 0.7.0.
- Keep malformed rich-run validation, source mapping, and logical extraction
  behavior aligned across PDF, raster, and SVG backends.

### Compatibility

The exact 15-package incubating crates.io family moves together from 0.7.0 to
0.8.0. The additive `TextSegment.direction` field is an intentional pre-1.0
Rust source change for callers that construct full `TextSegment` literals.
Callers using existing shaping and layout entry points receive the resolved
direction from those APIs.

The stable workspace remains prepared at 0.11.0 during this shared release.
The later stable 0.11.1 recovery pins shared dependencies to 0.8.0. The
unpublished `rpptx-wasm` preparation carrier moves to 0.8.0 without gaining
crates.io, npm, or other publication authority.

### Contributors

Atul Sharma maintained the release with the rdocx maintainers. This
shared-family carrier release adds no authenticated external issue or pull
request to its selected contribution inventory, so it prepares no external
notification.

## v0.11.0

### Highlights

The stable Word family adds language-aware conditional hyphenation,
multi-script shaping, bidirectional paragraph and run layout, and logical text
extraction over visually ordered output. Bounded restart pagination now covers
ordinary note, header, and footer workloads, while unchanged caller fonts no
longer pay a second full byte comparison on the warm layout path.

### Added

- Apply Word automatic-hyphenation settings and run languages for English,
  French, German, and Spanish without assigning source ranges to generated
  hyphens.
- Shape Arabic, Devanagari, Thai, and Simplified Chinese with deterministic
  font selection, cluster-safe breaking, complete glyph offsets, and stable
  source mapping.
- Carry Word paragraph and run direction into paragraph-wide UAX 9 resolution,
  line-local visual ordering, and logical PDF and SVG extraction.
- Reuse bounded restart pagination for unchanged notes, headers, and footers.
  This hardened equivalent addresses the 700-paragraph note and header or
  footer workloads reported in
  [Issue 53](https://github.com/tensorbee/rdocx/issues/53).
- Avoid the redundant retained-context font byte comparison after the font
  manager has already accepted the exact ordered font set. This hardened
  equivalent addresses the 22 MiB caller-font workload reported in
  [Issue 54](https://github.com/tensorbee/rdocx/issues/54).
- Accept exact whole-valued decimal table measurements while rejecting
  fractional, exponent, overflow, unit-bearing, and malformed forms. This
  hardened equivalent includes the outcome proposed in
  [PR 55](https://github.com/tensorbee/rdocx/pull/55).
- Preserve tracked table-grid history as inert revision metadata while keeping
  the active grid as the only layout input. This hardened equivalent includes
  the outcome proposed in
  [PR 56](https://github.com/tensorbee/rdocx/pull/56).
- Classify the narrow enabled legacy VML horizontal-rule form for native
  inspection while retaining its exact raw XML. This hardened equivalent
  includes the outcome proposed in
  [PR 57](https://github.com/tensorbee/rdocx/pull/57).
- Prime locked Cargo dependencies before the intentionally offline Word
  fidelity harness. This hardened equivalent includes the locked Word fidelity
  dependency preparation proposed in
  [PR 58](https://github.com/tensorbee/rdocx/pull/58).

### Fixed

- Preserve exact warm and fresh layout equality when related stories, note
  references, language, hyphenation, direction, source-less generated content,
  or caller fonts participate in cache identity.
- Preserve logical searchable text while visually positioning mixed-direction
  rich runs, inline objects, numbering markers, stored fields, tab leaders,
  and generated conditional hyphens.
- Preserve namespace-aware unsupported table and run XML, including ancestor
  bindings, without changing schema child order or treating historical grids
  as active layout data.

### Compatibility

The exact seven-package stable crates.io family moves together from 0.10.1 to
0.11.0 and requires the separately published shared 0.7.0 family. The stable
set remains `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, and `rdocx-cli`. Python, WASM, npm, and PyPI publication
authority is unchanged.

Low-level Rust callers that use full `CT_RPr`, `LayoutInput`, `CT_PPr`,
`CT_TblGrid`, and low-level positioned-layout struct literals must initialize
the new language, automatic-hyphenation, direction, preservation, and source
mapping fields or use existing defaults and constructors. These are
intentional pre-1.0 source changes. The native facade additions are additive,
and existing binding method names remain unchanged. The legacy VML horizontal
rule is classified for inspection and preservation, not rendered.

### Contributors

Atul Sharma maintained the release. Thanks to authenticated reporter
`@emptinessform` for the note and header or footer restart evidence in
[Issue 53](https://github.com/tensorbee/rdocx/issues/53), including the
corrected attribution and independent page-count ceiling, and for the isolated
caller-font byte-comparison evidence and unsound shallow-comparison caveat in
[Issue 54](https://github.com/tensorbee/rdocx/issues/54). Both final fixes are
hardened equivalents, and both issues remain open for their release-bound
notifications.

Thanks to authenticated contributor `@pedroassumpcao` for the whole-valued
decimal table-measurement case in
[PR 55](https://github.com/tensorbee/rdocx/pull/55) at
`056d48fdf23f35e3538ef3d6ff78cf9e3863e3a5`, tracked table-grid history in
[PR 56](https://github.com/tensorbee/rdocx/pull/56) at
`8b79c4cd0452defafe0a58e86b332c98e7fe52d7`, the legacy VML reader
classification in [PR 57](https://github.com/tensorbee/rdocx/pull/57) at
`44498f042a2290ef40c7a6c26025f38e38e9ce2a`, and locked Word fidelity
dependency preparation in
[PR 58](https://github.com/tensorbee/rdocx/pull/58) at
`c8fed1d1268fd765d602bac2da6524900c1c1cfd`. All four outcomes are hardened
equivalents, and all four pull requests remain open for their release-bound
thank-yous.

No named external patch landed directly. Each named report or proposal landed
through the hardened equivalent described above so that namespace, exact
lexical parsing, raw preservation, bounded cache identity, and offline oracle
contracts remain intact.

## rpptx-v0.7.0

### Highlights

The shared OOXML and PowerPoint family adds one complete multilingual text
substrate for the later Word hyphenation, complex-script, and bidirectional
layout stories. Conditional hyphenation, script-aware shaping, cluster-safe
breaking, paragraph direction, line-local visual ordering, and deterministic
complex-script fonts now share one format-neutral contract.

### Added

- Offer conditional hyphenation for English, French, German, and Spanish while
  retaining contiguous source spans and omitting a source for generated
  hyphens.
- Shape Arabic, Devanagari, Thai, and Simplified Chinese with deterministic
  font fallback, explicit script and language, logical clusters, and complete
  two-dimensional glyph advances and offsets.
- Apply ICU complex-script boundaries, CJK prohibited-punctuation rules, and
  UAX 9 bidirectional levels and line-local visual ordering without rewriting
  logical searchable text.
- Carry typed DrawingML paragraph direction through an additive PowerPoint
  sidecar into PDF, raster, and SVG output.
- Bundle licensed Noto Sans Arabic, Devanagari, and Thai fonts plus a
  reproducible Noto Sans Simplified Chinese subset for the approved fixture
  repertoire.

### Fixed

- Apply explicit right-to-left direction to numeric and Latin text across
  styled runs and forced line breaks using one paragraph-wide bidi context.
- Reject malformed rich glyph positioning safely across PDF, raster, and SVG
  backends. SVG retains logical searchable text with an explicit positioning
  approximation diagnostic.

### Compatibility

The exact 15-package incubating crates.io family moves together from 0.6.0 to
0.7.0. This is an intentional pre-1.0 minor boundary for new additive shared
text types, non-exhaustive variants, and sibling resolver and renderer entry
points. Existing legacy Latin struct and entrypoint shapes remain valid, and
their deterministic output remains byte-identical.

The seven-package stable family remains at 0.10.1 and does not opt into the new
shared path in this release. Word property parsing, facade authoring, and final
Word oracle acceptance remain in the later product stories. `rpptx-wasm` is
prepared at 0.7.0 for binding checks but remains unpublished on crates.io, npm,
and every other registry.

### Contributors

Atul Sharma maintained the release with the rdocx maintainers. The selected
F-X058 substrate adds no new authenticated external issue or pull-request
record after rpptx-v0.6.0, so this release prepares no external notification.

## v0.10.1

### Highlights

The stable Word family adds native RTF, HTML, and OpenDocument Text input,
native RTF and OpenDocument Text output, deterministic EPUB and SVG export,
and ordered compatibility readers that preserve unsupported XML. Layout also
honors caller font aliases and restores bounded editor-scale reuse.

This patch release recovers the complete stable family after v0.10.0 published
only `rdocx-opc` and `rdocx-oxml`, then stopped during `rdocx-layout` package
verification. Version 0.10.1 is the first complete stable family carrying the
S56 outcome.

### Added

- Read RTF, HTML5, and OpenDocument Text into the native `Document` facade with
  bounded inputs, ordered diagnostics, and no network access.
- Write deterministic RTF byte streams with stable lossy diagnostics and
  failure-atomic path replacement.
- Write deterministic OpenDocument Text archives with stable lossy diagnostics
  and failure-atomic path replacement.
- Export deterministic EPUB publications that retain supported links, images,
  headings, lists, tables, and accessibility structure while reporting
  unsupported source content.
- Export searchable SVG pages with deterministic fixed-page geometry, images,
  and safe links while reporting unsupported visual content.
- Export exact selected pages as opaque or transparent PNG, quality-controlled
  JPEG, or one deterministic multi-page TIFF through native Word, Python, and
  the general Word and PowerPoint CLI paths.
- Inspect direct table-cell children through `CellRef::items`. This hardened
  equivalent includes the outcome proposed in
  [PR 47](https://github.com/tensorbee/rdocx/pull/47).
- Inspect direct run content through `RunRef::items`, including text, breaks,
  drawings, fields, notes, comments, and preserved XML. This hardened
  equivalent includes the outcome proposed in
  [PR 48](https://github.com/tensorbee/rdocx/pull/48).
- Inspect direct paragraph and hyperlink content without changing established
  flattened accessors. This hardened equivalent includes the outcome proposed
  in [PR 49](https://github.com/tensorbee/rdocx/pull/49).
- Classify retained unsupported body XML through borrowed qualified-name,
  namespace, and child-content facts without inventing source bytes. This
  hardened equivalent includes the outcome proposed in
  [PR 50](https://github.com/tensorbee/rdocx/pull/50).

### Fixed

- Honor caller-supplied font family aliases without duplicating font bytes,
  while retaining deterministic fallback and bounded caches. This hardened
  equivalent addresses
  [Issue 44](https://github.com/tensorbee/rdocx/issues/44) and the reference
  implementation in [PR 45](https://github.com/tensorbee/rdocx/pull/45).
- Restore bounded reusable layout performance for document load, typing, undo,
  and table mutation while retaining exact shaping, source mappings, and page
  structure. This hardened equivalent addresses
  [Issue 46](https://github.com/tensorbee/rdocx/issues/46).
- Reject undecodable ordinary and deleted Word text instead of publishing an
  empty lossy value. This hardened equivalent includes the correction proposed
  in [PR 52](https://github.com/tensorbee/rdocx/pull/52).

### Compatibility

The exact seven-package stable crates.io family moves together from 0.9.0 to
0.10.1. This is an intentional pre-1.0 Rust source boundary. The shared OOXML
and PowerPoint family is published at 0.6.0. Python, WASM, npm, and PyPI
publication authority is unchanged.

The immutable v0.10.0 attempt contains only `rdocx-opc` and `rdocx-oxml`.
Callers should select 0.10.1 for a coherent seven-package stable graph. No
v0.10.0 tag or registry entry was moved, replaced, or reused.

`ST_NumberFormat` now preserves producer-defined values in `Other(String)`.
The enum no longer implements `Copy`, and exhaustive matches must handle the
new value-bearing variant. Callers should borrow or clone numbering formats as
needed and retain unknown values rather than substituting a modeled marker.
This hardened equivalent includes the preservation outcome proposed in
[PR 51](https://github.com/tensorbee/rdocx/pull/51).

External layout backends must continue to recurse through
`PositionedElement::MarkedContent` or use `oxml_layout::walk`, as documented for
v0.9.0. No migration is required for callers that use the high-level document
facade and do not exhaustively match `ST_NumberFormat`.

### Contributors

Atul Sharma maintained the release. Thanks to `@emptinessform` for the caller
font-alias report and reference implementation in
[Issue 44](https://github.com/tensorbee/rdocx/issues/44) and
[PR 45](https://github.com/tensorbee/rdocx/pull/45), and for the editor
performance measurements and migration evidence in
[Issue 46](https://github.com/tensorbee/rdocx/issues/46).

Thanks to `@pedroassumpcao` for the ordered cell, run, paragraph, and hyperlink
reader designs in [PR 47](https://github.com/tensorbee/rdocx/pull/47),
[PR 48](https://github.com/tensorbee/rdocx/pull/48), and
[PR 49](https://github.com/tensorbee/rdocx/pull/49), the unsupported XML facts
in [PR 50](https://github.com/tensorbee/rdocx/pull/50), producer-defined
numbering preservation in [PR 51](https://github.com/tensorbee/rdocx/pull/51),
and fail-closed text decoding in
[PR 52](https://github.com/tensorbee/rdocx/pull/52).

No named external patch landed directly. Each named report or proposal landed
through the hardened equivalent described above so that current namespace,
non-exhaustive API, bounded-allocation, diagnostic, and compatibility contracts
remain intact.

## rpptx-v0.6.0

### Highlights

The shared OOXML and PowerPoint family adds caller-controlled font aliases,
bounded reusable layout work, deterministic page image output, and common CLI
page selection. This release publishes the shared APIs required by the stable
Word family before its separate v0.10.1 recovery release.

### Added

- Configure bounded caller font family aliases through the shared font manager
  without repeating font bytes or changing deterministic fallback.
- Render exact selected pages as transparent or opaque PNG, quality-controlled
  JPEG, or deterministic multi-page TIFF through the shared raster backend.
- Select page ranges and image output options through the common CLI support
  layer and the PowerPoint CLI.

### Fixed

- Honor caller-supplied family aliases through deterministic fallback and
  bounded caches. This hardened equivalent addresses
  [Issue 44](https://github.com/tensorbee/rdocx/issues/44) and the reference
  implementation in [PR 45](https://github.com/tensorbee/rdocx/pull/45).
- Restore bounded reusable layout performance for document load, typing, undo,
  and table mutation while retaining exact shaping, source mappings, and page
  structure. This hardened equivalent addresses
  [Issue 46](https://github.com/tensorbee/rdocx/issues/46).

### Compatibility

All 15 crates.io packages in the shared OOXML and PowerPoint family move
together from 0.5.0 to 0.6.0. This is an intentional pre-1.0 minor boundary.
Callers that build a `FontManager` directly can use the new caller-alias
configuration. Existing callers that do not configure aliases retain the same
deterministic fallback behavior.

The stable Word family remains at its prepared 0.10.0 source boundary while
this incubating family publishes. `rpptx-wasm` is prepared at 0.6.0 but remains
unpublished on crates.io. Python, WASM, npm, and PyPI publication authority is
unchanged.

### Contributors

Atul Sharma maintained the release. Thanks to `@emptinessform` for the caller
font-alias report and reference implementation in
[Issue 44](https://github.com/tensorbee/rdocx/issues/44) and
[PR 45](https://github.com/tensorbee/rdocx/pull/45), and for the editor
performance measurements and migration evidence in
[Issue 46](https://github.com/tensorbee/rdocx/issues/46).

No named external patch landed directly. Each named report or proposal landed
through the hardened equivalent described above so that the current bounded
cache, deterministic fallback, and reusable layout contracts remain intact.

## v0.10.0

### Highlights

The stable Word family adds native RTF, HTML, and OpenDocument Text input,
native RTF and OpenDocument Text output, deterministic EPUB and SVG export,
and ordered compatibility readers that preserve unsupported XML. Layout also
honors caller font aliases and restores bounded editor-scale reuse.

### Added

- Read RTF, HTML5, and OpenDocument Text into the native `Document` facade with
  bounded inputs, ordered diagnostics, and no network access.
- Write deterministic RTF byte streams with stable lossy diagnostics and
  failure-atomic path replacement.
- Write deterministic OpenDocument Text archives with stable lossy diagnostics
  and failure-atomic path replacement.
- Export deterministic EPUB publications that retain supported links, images,
  headings, lists, tables, and accessibility structure while reporting
  unsupported source content.
- Export searchable SVG pages with deterministic fixed-page geometry, images,
  and safe links while reporting unsupported visual content.
- Export exact selected pages as opaque or transparent PNG, quality-controlled
  JPEG, or one deterministic multi-page TIFF through native Word, Python, and
  the general Word and PowerPoint CLI paths.
- Inspect direct table-cell children through `CellRef::items`. This hardened
  equivalent includes the outcome proposed in
  [PR 47](https://github.com/tensorbee/rdocx/pull/47).
- Inspect direct run content through `RunRef::items`, including text, breaks,
  drawings, fields, notes, comments, and preserved XML. This hardened
  equivalent includes the outcome proposed in
  [PR 48](https://github.com/tensorbee/rdocx/pull/48).
- Inspect direct paragraph and hyperlink content without changing established
  flattened accessors. This hardened equivalent includes the outcome proposed
  in [PR 49](https://github.com/tensorbee/rdocx/pull/49).
- Classify retained unsupported body XML through borrowed qualified-name,
  namespace, and child-content facts without inventing source bytes. This
  hardened equivalent includes the outcome proposed in
  [PR 50](https://github.com/tensorbee/rdocx/pull/50).

### Fixed

- Honor caller-supplied font family aliases without duplicating font bytes,
  while retaining deterministic fallback and bounded caches. This hardened
  equivalent addresses
  [Issue 44](https://github.com/tensorbee/rdocx/issues/44) and the reference
  implementation in [PR 45](https://github.com/tensorbee/rdocx/pull/45).
- Restore bounded reusable layout performance for document load, typing, undo,
  and table mutation while retaining exact shaping, source mappings, and page
  structure. This hardened equivalent addresses
  [Issue 46](https://github.com/tensorbee/rdocx/issues/46).
- Reject undecodable ordinary and deleted Word text instead of publishing an
  empty lossy value. This hardened equivalent includes the correction proposed
  in [PR 52](https://github.com/tensorbee/rdocx/pull/52).

### Compatibility

The exact seven-package stable crates.io family moves together from 0.9.0 to
0.10.0. This is an intentional pre-1.0 Rust source boundary. The shared OOXML
and PowerPoint family remains at its published 0.5.0 boundary. Python, WASM,
CLI, npm, and PyPI publication authority is unchanged.

`ST_NumberFormat` now preserves producer-defined values in `Other(String)`.
The enum no longer implements `Copy`, and exhaustive matches must handle the
new value-bearing variant. Callers should borrow or clone numbering formats as
needed and retain unknown values rather than substituting a modeled marker.
This hardened equivalent includes the preservation outcome proposed in
[PR 51](https://github.com/tensorbee/rdocx/pull/51).

External layout backends must continue to recurse through
`PositionedElement::MarkedContent` or use `oxml_layout::walk`, as documented for
v0.9.0. No migration is required for callers that use the high-level document
facade and do not exhaustively match `ST_NumberFormat`.

### Contributors

Atul Sharma maintained the release. Thanks to `@emptinessform` for the caller
font-alias report and reference implementation in
[Issue 44](https://github.com/tensorbee/rdocx/issues/44) and
[PR 45](https://github.com/tensorbee/rdocx/pull/45), and for the editor
performance measurements and migration evidence in
[Issue 46](https://github.com/tensorbee/rdocx/issues/46).

Thanks to `@pedroassumpcao` for the ordered cell, run, paragraph, and hyperlink
reader designs in [PR 47](https://github.com/tensorbee/rdocx/pull/47),
[PR 48](https://github.com/tensorbee/rdocx/pull/48), and
[PR 49](https://github.com/tensorbee/rdocx/pull/49), the unsupported XML facts
in [PR 50](https://github.com/tensorbee/rdocx/pull/50), producer-defined
numbering preservation in [PR 51](https://github.com/tensorbee/rdocx/pull/51),
and fail-closed text decoding in
[PR 52](https://github.com/tensorbee/rdocx/pull/52).

No named external patch landed directly. Each named report or proposal landed
through the hardened equivalent described above so that current namespace,
non-exhaustive API, bounded-allocation, diagnostic, and compatibility contracts
remain intact.

## v0.9.0

### Highlights

The stable Word family adds native package encryption and signing, accessible
and archival PDF output, exact redaction, and editor-scale layout reuse. It
also corrects duplicated shaped text, header and footer delivery, and dense
form table layout while retaining unsupported OOXML.

### Added

- Open and write Microsoft Agile encrypted OOXML packages with authenticated
  AES-256 and SHA-512 processing, bounded inputs, and failure-atomic output.
- Verify and create RSA-SHA256 OPC digital signatures with exact declared part
  and relationship coverage. Certificate-chain trust remains caller policy.
- Emit tagged PDF structure with deterministic marked-content ownership, and
  emit PDF/A-2b or PDF/A-3b with an output intent and conformance metadata.
- Remove exact non-empty literals from Word stories, metadata, chart caches,
  and embedded workbooks through a transactional native redaction API.
- Share immutable font bytes and page frames, transfer reusable layout work
  only across an exact checked context, restart pagination at safe boundaries,
  and retain bounded transactional caches. The hardened equivalent was shaped
  by [Issue 39](https://github.com/tensorbee/rdocx/issues/39),
  [PR 40](https://github.com/tensorbee/rdocx/pull/40), and
  [PR 41](https://github.com/tensorbee/rdocx/pull/41).

### Fixed

- Preserve default, first-page, even-page, inherited, and multi-section header
  and footer text through reopened layout and deterministic PDF output, closing
  [Issue 15](https://github.com/tensorbee/rdocx/issues/15).
- Reshape final Unicode break segments exactly, so spaces, hyphens, ligatures,
  combining text, and CJK do not duplicate source text or glyphs. This resolves
  [Issue 23](https://github.com/tensorbee/rdocx/issues/23).
- Keep nested tables recursive and honor grid-span-aware vertical merges,
  exact and minimum row rules, table-style cascades, paragraph-mark metrics,
  outer border fallbacks, and cell-relative anchors. This hardened equivalent
  addresses [Issue 42](https://github.com/tensorbee/rdocx/issues/42) and
  [PR 43](https://github.com/tensorbee/rdocx/pull/43).

### Compatibility

The exact seven-package stable crates.io family moves together from 0.8.0 to
0.9.0. This is an intentional pre-1.0 Rust source boundary. `FontData.data`
now uses `Arc<[u8]>`, `LayoutResult.pages` now uses
`Vec<Arc<PageFrame>>`, table cells retain ordered `CellBlock` values, and
typed table styles expose additional preserved and conditional properties.
Callers that construct these low-level values must update their literals or
use the provided constructors.

`PositionedElement` remains non-exhaustive. Visible content in
`PageFrame::elements` can be nested under
`PositionedElement::MarkedContent`. External backends must recurse through
`MarkedContent::children` or use `oxml_layout::walk` when visiting page
elements. A wildcard arm that ignores the wrapper can otherwise produce empty
output.

The `agile-encryption` and `digital-signatures` features remain default-off.
The high-level native additions do not expand Python, WASM, or CLI method
surfaces, and those packages remain unpublished on crates.io. The separate
shared OOXML and PowerPoint family remains at its published 0.5.0 boundary.

### Contributors

Atul Sharma maintained the release. Thanks to `@mantissaman` for the
authenticated header and footer report in
[Issue 15](https://github.com/tensorbee/rdocx/issues/15) and the duplicated
text report in [Issue 23](https://github.com/tensorbee/rdocx/issues/23).
Thanks to `@emptinessform` for the break-opportunity diagnosis on Issue 23,
the editor profiling and reference implementations in
[Issue 39](https://github.com/tensorbee/rdocx/issues/39),
[PR 40](https://github.com/tensorbee/rdocx/pull/40), and
[PR 41](https://github.com/tensorbee/rdocx/pull/41), and the dense-form report
and reference implementation in
[Issue 42](https://github.com/tensorbee/rdocx/issues/42) and
[PR 43](https://github.com/tensorbee/rdocx/pull/43). Those contributions
landed directly where noted or through the hardened equivalents described
above.

## rpptx-v0.5.0

### Highlights

The shared OOXML and PowerPoint family adds native package encryption and
digital signatures, accessible PDF structure, and deterministic PDF/A output.
Layout results also retain their largest immutable page and font payloads by
shared ownership, avoiding deep copies when results are cloned or retained.

### Added

- Read and write Microsoft Agile encrypted OOXML packages with authenticated
  AES-256 and SHA-512 output, bounded input processing, and failure-atomic
  publication.
- Verify and create RSA-SHA256 OPC digital signatures with exact declared part
  and relationship coverage. Certificate trust remains caller policy.
- Emit tagged PDF structure with deterministic marked-content ownership,
  document language, titles, outlines, links, and structure destinations.
- Emit deterministic PDF/A-2b and PDF/A-3b files with an output intent,
  conformance metadata, embedded-file relationship rules, and validator-backed
  fixtures.
- Share immutable `FontData` bytes and completed `PageFrame` values through
  `Arc`, so cloning or retaining a layout result keeps those payloads shared.
  This ownership boundary was shaped by
  [Issue 39](https://github.com/tensorbee/rdocx/issues/39),
  [PR 40](https://github.com/tensorbee/rdocx/pull/40), and
  [PR 41](https://github.com/tensorbee/rdocx/pull/41).

### Fixed

There are no user-facing defect corrections unique to the incubating family in
this release. Word-only rendering and document-editing corrections remain on
the separate stable release train.

### Compatibility

All 15 crates.io packages in the shared OOXML and PowerPoint family move
together from 0.4.0 to 0.5.0. This is an intentional pre-1.0 Rust source
boundary. `FontData.data` now uses `Arc<[u8]>`, and `LayoutResult.pages` now
uses `Vec<Arc<PageFrame>>`. Callers that construct those low-level values must
wrap owned data with `Arc::from` or `.into()`. Callers that only inspect values
can continue through deref coercion or `.as_ref()`.

The `agile-encryption` and `digital-signatures` package capabilities remain
default-off. Existing PowerPoint facade behavior requires no migration.
`rpptx-wasm` is prepared at 0.5.0 but remains unpublished on crates.io.

### Contributors

Atul Sharma maintained the release. Thanks to `@emptinessform` for the
authenticated editor profiling and reference implementations in
[Issue 39](https://github.com/tensorbee/rdocx/issues/39),
[PR 40](https://github.com/tensorbee/rdocx/pull/40), and
[PR 41](https://github.com/tensorbee/rdocx/pull/41). Their font-copy and
page-copy measurements informed the shared ownership surface that landed as a
hardened equivalent. The format-specific transfer, pagination, and cache work
remains on the stable release train.

## v0.8.0

### Highlights

The stable Word family now combines native document automation with a complete
layout result that downstream renderers and editors can inspect and reuse.
This release includes structured fields, templates, mail merge, tracked
comparison, watermarks, chart support, source provenance, and bounded relayout
caches while preserving unsupported OOXML.

### Added

- Parse and evaluate Word fields with explicit update policies, including safe
  displayed results for complex fields.
- Create, reply to, resolve, and remove comments and threaded conversations.
  Bind content controls to namespace-aware custom XML without rewriting
  unrelated package data.
- Create bookmarks and resolve `REF` and `PAGEREF` cross-references through
  fields and final pagination.
- Inspect tracked revisions and accept or reject all or a filtered selection
  while preserving unsupported revision XML.
- Render accepted or tracked revision views with visible insertions,
  deletions, and changed paragraphs. Read document-protection intent and its
  recorded enforcement metadata without claiming to enforce the restriction.
- Author Word charts and render them through the shared ChartML model.
- Expand structural templates with conditions and loops, then produce separate
  or sectioned mail-merge documents from flat records.
- Compare documents into deterministic tracked revisions whose accepted and
  rejected views reproduce the edited and original bodies.
- Author and render text or image watermarks through header-scoped VML.
- Expose complete native `WordLayoutResult` bundles with owned font data,
  diagnostics, and result-local source paths for body and related stories.
- Reuse safe paragraph layout, shaping, and font work through bounded caches
  that preserve cold-layout bytes, diagnostics, and current provenance.
- Traverse direct body paragraphs, tables, content controls, and unsupported
  XML in source order through `Document::body_items`.

### Fixed

- Preserve reader-owned unsupported XML, namespace bindings, table facts,
  paragraph borders, hyperlink tooltips, header and footer content, and safe
  field results across opened-document round trips.
- Keep watermark edits, failed relayouts, tracked views, caller fonts, and
  context-sensitive paragraphs from leaking stale cached layout state.

### Compatibility

The seven crates.io packages move together to 0.8.0. The release contains
intentional pre-1.0 Rust source breaks in low-level OOXML and layout structs.
Python, WASM, CLI, and the high-level `rdocx::Document` facade retain their
existing surface contracts. The shared and PowerPoint family remains on its
separate 0.4.0 train.

#### Migration table

| Previous path or crate | Replacement | Compatibility |
|---|---|---|
| `rdocx::Length` | `oxml_core::Length` | `rdocx::Length` remains an exact re-export |
| `rdocx_oxml::{core_properties, error, raw_xml, units}` | The same modules under `oxml_core` | The `rdocx_oxml` paths remain exact re-exports |
| `rdocx_opc` | `oxml_opc` | `rdocx-opc` is a deprecated exact re-export shim, except for the removed Word-only constructors listed below |
| Word-owned image sniffing, sizing, and media naming | `oxml_media::{resolve, probe, ImageFormat, ImageInfo, NativeSize, MediaNamer}` | These shared APIs are available directly from `oxml-media` |
| `rdocx_layout::bundled_fonts` | `oxml_layout::bundled_fonts` | The old module path is removed |
| `rdocx_layout::font::{FontManager, FontMetrics, ShapedText}` | The same types at the `oxml_layout` root | `rdocx_layout::input::FontFile` and `rdocx_layout::FontFile` remain exact re-exports of `oxml_layout::FontFile` |
| `rdocx_layout::error::{LayoutError, Result}` | `oxml_layout::{LayoutError, Result}` | The types also remain exact re-exports at the `rdocx_layout` root |
| `rdocx_layout::line::{InlineItem, LayoutLine, LineBreakParams, LineItem, TextSegment, break_into_lines}` | The same names at the `oxml_layout` root | The old `rdocx_layout::line` module is removed |
| `rdocx_layout::output::{Color, DocumentMetadata, FieldKind, FontData, FontId, GlyphRun, LayoutResult, OutlineEntry, PageFrame, Point, PositionedElement, Rect}` | The same names at the `oxml_layout` root | Types previously exported at the `rdocx_layout` root remain exact re-exports there |
| Exhaustive `TextSegment` and `GlyphRun` literals | Add `source: Option<SourceSpan>` | Use `None` for generated or unattributed text. Word provenance results supply exact result-local node ids and Unicode-scalar ranges |
| `rdocx_pdf` | `oxml_pdf` | `rdocx-pdf` is a deprecated exact re-export shim |
| `rdocx_pdf::raster::{render_page_to_png, render_all_pages}` | `oxml_pdf::{render_page_to_png, render_all_pages}` | The old nested `raster` path is removed. The functions remain available at the `rdocx_pdf` root through the shim |

`rdocx-oxml` and `rdocx-layout` are retained format-specific crates, not
deprecated shims. `rdocx-oxml` continues to own WordprocessingML types.
`rdocx-layout` continues to own the Word flow engine, paginator, blocks,
tables, style resolver, and Word-to-shared conversion boundary. The `rdocx`,
`rdocx-cli`, and `rdocx-html` crate names are unchanged.

### Shared dependencies

New direct users can select the format-neutral crate that owns each surface:

```toml
[dependencies]
oxml-core = "0.4.0"   # Length, units, XML helpers, document properties
oxml-opc = "0.4.0"    # OPC package, relationships, and content types
oxml-media = "0.4.0"  # Image detection, dimensions, and media naming
oxml-layout = "0.4.0" # Layout output, fonts, and line breaking
oxml-pdf = "0.4.0"    # PDF and PNG rendering backends
```

### Breaking API changes

- `rdocx_opc::OpcPackage::new_docx()` and
  `rdocx_opc::ContentTypes::new_docx()` are removed. Use
  `oxml_opc::OpcPackage::new()` or `OpcPackage::with_main_part(...)`, plus
  `oxml_opc::ContentTypes::minimal()`, and add Word-specific defaults and
  overrides at the application boundary.
- `rdocx::Error::Opc` now contains `oxml_opc::OpcError`, and
  `rdocx::Error::Layout` now contains `oxml_layout::LayoutError`. The
  deprecated OPC shim and retained layout facade re-export those exact shared
  types, but code that spells payload paths in exhaustive matches should use
  the shared paths.
- The public `rdocx_layout::line` module is removed. Its shared replacement
  uses `MediaId` instead of relationship-scoped `embed_id` strings for image
  items. `TextSegment` uses `oxml_layout::Underline` and adds `line_gap`.
  `LayoutLine` adds `line_gap`. `LineBreakParams` replaces Word tab stops,
  alignment, and stringly typed line rules with `TabStop`, `Align`, and
  `LineSpacing`, and adds `wrap`.
- `rdocx_layout::engine::layout_paragraph(...)` and
  `rdocx_layout::table::layout_table(...)`, plus
  `rdocx_layout::paginator::paginate(...)` and
  `rdocx_layout::paginator::paginate_sections(...)`, now take a shared
  `MediaRegistry`. Construct it once from `LayoutInput::images` so relationship
  lookup and pagination use the same collision-resolved IDs, bytes, and
  content types.
- `rdocx_layout::AnchoredContent::Image` replaces `embed_id: String` with
  `media_id: MediaId`.
- `rdocx_layout::ParagraphBlock::jc` replaces `Option<ST_Jc>` with
  `Option<oxml_layout::Align>`.
- `PositionedElement` is non-exhaustive, replaces the optional image
  `embed_id` with `MediaId`, and adds `Path` and `Group` variants. External
  matches must include a wildcard arm.
- `PageFrame` is non-exhaustive and adds `background`. Construct it with
  `PageFrame::new(...)` when a default background is wanted.
- `LayoutResult` is non-exhaustive and adds `diagnostics`. Construct it with
  `LayoutResult::new(...)` when an empty diagnostics list is wanted.
- `oxml_layout::TextSegment` and `oxml_layout::GlyphRun` add the required
  `source: Option<SourceSpan>` field. External exhaustive literals must set it
  to `None` unless they own an exact source range. This source change ships in
  the incubating 0.4.0 family and the stable 0.8.0 family. Word callers can use
  `rdocx_layout::layout_document_with_provenance` or its deterministic variant
  to receive `WordLayoutResult`, resolve result-local nodes to
  `WordSourcePath`, and interpret exclusive character ranges as Unicode scalar
  indices in the recorded revision view.
- The nested `rdocx_pdf::raster` module is removed. Import its two rendering
  functions from the `oxml_pdf` root or from the compatible `rdocx_pdf` root.

### Media behavior and additive API

Word media insertion now detects the image format from its bytes before using
the filename extension. It allocates the next numeric media suffix after the
greatest occupied suffix, so gaps do not overwrite an existing part.

`rdocx::Document::add_picture_auto(image_data, image_filename)` adds an image
at its intrinsic size. It uses declared per-axis DPI when valid and a 72 DPI
fallback otherwise. If dimensions cannot be determined, it returns
`rdocx::Error::UnavailableImageDimensions` before changing the document.

### Contributors

Thanks to Pedro Assumpcao for the ordered-body contribution in PR 36 and the
reader compatibility work included in this release. Thanks to `@emptinessform`
for the Issue 37 complete-layout report and the Issue 39 relayout measurements
and cache proposal.

## rpptx-v0.4.0

### Highlights

The complete shared OOXML and PowerPoint family moves together to 0.4.0. This
is the first release to publish `oxml-chart`, making the typed ChartML model,
authoring surface, and renderer available from its format-neutral home.

### Added

- `oxml-chart` now owns shared ChartML parsing, editing, authoring, and render
  geometry. `rpptx-chart` remains an exact compatibility re-export.
- `oxml-layout` glyph runs can carry exact `SourceSpan` provenance through
  shaping and line splitting, with generated or transformed text left
  truthfully unattributed.
- Normal host-font layout reuses a bounded process font snapshot, file-backed
  bytes, and exact-key shaping results. Deterministic and caller-font paths
  remain isolated from that state.

### Fixed

- Bounded OPC reads reject oversized declared ZIP entry counts before the ZIP
  index is constructed, and retain the configured byte and entry ceilings
  throughout package access.
- Deterministic PDF output now writes font, Unicode-map, and image resources in
  stable order, so identical inputs produce identical bytes.

### Compatibility

This is an intentional pre-1.0 source boundary. External exhaustive literals
for `TextSegment` and `GlyphRun` must add `source: None` unless they own an
exact `SourceSpan`. Existing `rpptx-chart` imports remain valid through the
exact re-export, while new direct users should depend on `oxml-chart`.

Normal system-font discovery is now a process-lifetime snapshot. Installing,
removing, or replacing host fonts requires a process restart. Deterministic and
caller-provided font behavior is unchanged. `rpptx-wasm` is prepared at 0.4.0
but remains unpublished on crates.io.

### Contributors

Atul Sharma maintained the release. `@emptinessform` supplied the provenance
and cache reports behind Issues 38 and 39. Pedro Assumpcao
(`@pedroassumpcao`) contributed bounded OPC reads in PR 33 and carried the
entry-limit hardening through PR 34. Jon Stokes (`@jonstokes`) authored the
ZIP entry-admission hardening commit integrated by PR 34.
