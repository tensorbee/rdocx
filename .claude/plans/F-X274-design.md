# F-X274, Honor section and field page-number formats

**Status**: approved
**Sprint**: S96
**Size**: M
**Depends on**: F-X273

## Problem

Issue 423 reports PAGE rendering ignores section `w:fmt` and explicit `\* roman`. Current main is newer than its reproduction: `crates/rdocx/src/field.rs:16705` already implements section-sensitive cache formatting, while `engine.rs:8418` substitutes layout fields using separate context. `CT_PageNumberType` retains the original pgNumType XML without a modeled format. Fix the cross-consumer gap while preserving the current dynamic header/footer saved-cache contract.

## Spec reference

- `docs/hld/08-rendering-spec.md`, "Word section geometry and page numbering" and "Word bookmark field pagination".
- `docs/hld/04-opc-and-packaging.md`, "Producer read and writer authority" and retained XML rules.
- `docs/hld/10-bindings-spec.md`, layout-backed field caches and combined `fields_update` contract.
- `docs/hld/12-testing-strategy.md`, Word live-story/page-authoring and installed compatibility controls.
- `docs/hld/14-development-backlog.md`, F-X274.

## Approach

Read namespace-qualified page format from the already retained section source. `parse_page_number` retains external bindings through `raw_with_external_bindings`, so use those bindings rather than matching any attribute's local name. Preserve raw bytes, chapter settings and unknown extensions unchanged. Absence selects existing decimal fallback. An imported unknown format or chapter composition retains stored field display and reports an unsupported diagnostic as the existing bounded policy. An unknown-format explicit override is retained with a diagnostic in this boundary until native qualification proves its exact interaction. No unsupported section state is guessed as successful parity. Do not add a page-format setter or authoring enum.

Move the existing cache result policy to the existing layout engine provider, with a concrete doc-hidden cross-crate function `pub fn format_page_field_result(instruction: &FieldInstruction, value: usize, section_format: Option<&str>) -> std::result::Result<String, String>`. Resolve section source via a second concrete doc-hidden helper `pub fn section_page_number_format(section: &CT_SectPr) -> Option<String>` in that same file. These are additive implementation-facing native symbols, not Python authoring additions. No new trait, generic, crate, module or file. Cache updates and layout substitution call the same result policy with their respective physical field owner or PAGEREF target section. NUMPAGES/SECTION/SECTIONPAGES keep decimal defaults. Keep physical page, displayed integer and formatted text distinct.

Namespace-aware retained-source reading requires the existing rdocx-layout manifest to declare `quick-xml = { workspace = true }`. The workspace already uses quick-xml0.41 and its locked resolution. This adds a direct dependency edge for the approved engine helper, not a new repository crate, version, source file or public lower-layer API. Keep Cargo.lock changes limited to that edge, without unrelated upgrades. The existing final WASM, archive/size, supply-chain and hash riders cover this wiring. The implementing owner acknowledges this clarification while retaining the immutable claim Base.

An explicit numeric picture or numeral format switch overrides section numeral style. Character-only switches and MERGEFORMAT/CHARFORMAT must not accidentally erase the section default. Preserve existing ordered numeric/character/general formatting stages and unsupported-switch diagnostics. Reconcile this distinction against the current `page_field_result`, which treats any non-MERGEFORMAT/CHARFORMAT general switch as numeric. PAGE letters use repeated-letter numbering, with boundary cases 26=Z,27=AA,28=BB,52=ZZ,53=AAA, not spreadsheet column labels. Reuse the current bounded page formatter, not the different ordinary alphabetic formatter.

Keep existing Roman 1 to 3999 and alphabetic 4096 character implementation limits as explicit bounded support, with retained display and diagnostics outside the supported range. These limits are existing library behavior, not claimed native maxima. Native PAGE documentation states a 32767 count limit, so do not silently claim unrestricted Word parity. Unknown/zero/hostile restarts retain existing numeric handling and diagnostic boundaries, with no truncation/rounding change.

Rust `Document::update_page_fields(&mut self) -> Result<usize>` and `update_layout_backed_fields(&mut self) -> Result<LayoutBackedFieldUpdateReport>` retain signatures. Python exposes the corresponding no-argument methods returning `int` and report. Dynamic header/footer PAGE/NUMPAGES saved caches stay unchanged because one physical cache cannot encode every placement of a reused footer. Verify render-time text separately on all pages. Eligible placed body/note/text-box caches update through the existing staged path. Read-only `evaluate_field` continues to defer pagination-dependent values. CLI rendering and existing `fields update` use native policy. No new WASM method or binding default.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Qualified section format | Retained `CT_PageNumberType.raw_xml`, section inventory | `section_page_number_format`, source namespace bindings | Parse, read, untouched save, restart edit, reopen |
| PAGE formatting | `substitute_fields`, `page_field_result` successor, explicit cache updates | Shared concrete result helper, actual source section and displayed integer | Layout glyph text, eligible saved cache, PDF/PNG/text |
| PAGEREF/default fields | Existing target section context, NUMPAGES/SECTION/SECTIONPAGES | Target section for PAGEREF, decimal unrelated totals | Layout, cache update, TOC labels/targets |
| Dynamic furniture cache | `update_page_fields`, `update_layout_backed_fields`, `fields_update`, Python/CLI | Existing retained-cache policy, rendering remains dynamic per placement | Before/after update, save/reopen, multi-page footer render |
| Cached reuse | Header/footer cache and substitution/restart snapshots | Include format and instruction in existing equality/input context | Same-owner format/switch mutation, warm versus cold |

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| Section scope | pgNumType governs page numbers in its owning section, start/restart remains integer | ISO29500-1 first-edition clause reproduced in Microsoft OpenXML3.0.1 `PageNumberType`, https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.pagenumbertype?view=openxml-3.0.1 | Section 1 roman 1/2, section 2 decimal 1/2 |
| Explicit switch | PAGE numeral switch overrides section Number format | Microsoft Support "Field codes: PAGE field", Syntax/Switches, retrieved 2026-10-28, https://support.microsoft.com/en-us/word/field-codes-page-field | Decimal section with roman switch, Roman section with Arabic switch |
| Native count bound | Microsoft documents PAGE count does not exceed 32767 | Same Microsoft PAGE source, note before Syntax | Native qualification separate from existing library bounded-format limits |
| Format kinds | decimal, lowerRoman, upperRoman, lowerLetter, upperLetter | OpenXML3.0.1 NumberFormatValues lists exact serialized tokens, https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.numberformatvalues?view=openxml-3.0.1 | All five formats, upper/lower and rollover controls |
| Imported state | Missing format decimal, unknown/chapter composition preserved with diagnostic, qualified aliases distinguished from extension `fmt` | Current raw source parser and `section_page_format`, source Base `24af95d3` | Namespace shadow/alias and unknown-extension exact-byte controls |
| Saved caches | Dynamic header/footer PAGE/NUMPAGES retain source cache while displayed per-page text changes | Existing `field.rs:16615` contract and native saved-cache qualification | Footer placed twice with one physical cache, body eligible cache and retained footer both asserted |

## Dependency readiness

F-X273 passed its approved genuine LOCAL provider gate and clean independent microscope pass2. Its reviewed source is integrated at `fca298cc1a90dd77682ebbe6b652c56dd74bd228`. The existing provider run state and BACKLOG/CURRENT_SPRINT record implementation completion at the bounded dependency-prefix checkpoint. Its plan remains approved with all final differential/native debt retained. Claim this consumer only from a Base descended from that integrated source, and authenticate the ancestry before implementation. Final native differential evidence for both stories remains due at sprint closure.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Exact Issue 423 four-page input | Footer render reads i,ii,1,2, physical pages 1 to 4 remain distinct | Rust/Python/CLI PDF/text/PNG | Current-Base exact repro then repaired text assertions | Native Word input/output plus pinned LO |
| A2 | Five formats across restart/continue sections and default/first/even variants | Correct case/text for actual displayed integer and selected source section | Layout/cache/PAGEREF | Section and owner-context matrix | Native page-format controls and installed consumers |
| A3 | Explicit roman/ROMAN/Arabic/ALPHABETIC, numeric picture, character-only and MERGEFORMAT | Numeral override beats section style, nonnumeral formatting retains intended section base, ordering unchanged | Shared formatter plus actual callers | Same-source cache/render switch-precedence cases | Native switches and strict installed usage |
| A4 |26/27/28/52/53 letter values, Roman 3999/4000, very large restart | Repeated letters within bound, unsupported result retains cache with diagnostic, no unbounded allocation | Layout/cache/Python | Boundary and hostile-value discriminators | Native letter rollover, explicit bounded divergence record |
| A5 | Missing/unknown format, chapter settings, namespace alias/shadow, unknown siblings | Decimal fallback only for absence, unknown supported-source states preserved and diagnosed, no extension collision | Parse/save/render/update | Exact raw-byte retention and namespace control cases | Native source/reopen and package qualification |
| A6 | Body eligible cache and one physical dynamic footer used on multiple pages | Body cache equals formatted placement, footer cache remains unchanged while per-page render agrees | Update methods, combined fields update, save/reopen | Cache-retention and per-placement render paired assertions | Installed 3.9/3.12 and source-qualified native saved-cache proof |
| A7 | Same-owner format/switch change, held/fresh layout and run reads | Warm/cold values equal, published owner/cache invalidation follows existing contract | Native/Python/CLI | Fact-changing publication and neighboring-owner controls | Full current installed/runtime/typing union |
| A8 | F-X273 tall furniture/merged headings plus F-X272 prior redline | TOC target, PAGE label, physical geometry and saved source stay coherent | All three stories and CLI | Focused integrated composition after provider | Actual final native/PDF/PNG/text/hash union |

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | Roman restart then decimal restart, continue third section | Same displayed integer may have different text/physical page | Formatting by physical page or first section globally |
| C2 | Explicit numeral versus LOWER/MERGEFORMAT on a roman section | Override versus character operation applied to correct base | Treating every general switch as decimal reset |
| C3 |28/53 letter labels,3999/4000 Roman boundary | BB/AAAand bounded refusal | Reusing spreadsheet alphabetic function or silently claiming native unbounded support |
| C4 | Qualified Word alias versus extension `fmt`, unknown token and chapter data | Correct namespace or retained diagnostic, raw bytes survive | Matching by local name or rewriting pgNumType |
| C5 | Footer reused across pages, body field, same-owner switch edit | Dynamic text varies by placement, saved footer cache retained, warm/cold agree | Writing one footer cache as every page or cache hit with stale format |

## Rejected alternatives

- Add a page-format authoring API. The issue asks to render preserved inputs.
- Write every footer PAGE cache. A shared physical field has several placement values and existing native-compatible behavior retains it.
- Implement another independent formatter. Cache/render caller context already differs and needs one policy.

## Test plan

**Test gate**: differential, verbatim from backlog.

Local: add sentence-named `issue423_...` cases to existing engine tests, `regression_test.rs`, CLI integration and selected Python carrier. Current-Base RED/fixed GREEN for actual rendering and missing caller context, paired with current passing cache controls. Use `cargo check -p rdocx-layout -p rdocx --all-targets` while iterating, then only selected meaningful tests/scoped lint/fmt. No new integration binary.

Final: source-qualified Word page formatting including original input, pinned LO26.2.5.2/Poppler 26.01.0, deterministic library PNG/PDF/text, installed 3.9/3.12 runtime and strict typing/stubs, affected archive verification/size and README examples, WASM compile and full S96 risk union. Preserve every failed/old producer and unsupported row explicitly. No standalone native/oracle gate per worker.

## HLD impact

- `docs/hld/08-rendering-spec.md`
- `docs/hld/10-bindings-spec.md`

## Risk routing

Parser/serializer adds namespace-qualified imported variants and exact opaque-byte retention, unchanged schema order. Layout/pagination/units adds deterministic baselines and physical/displayed/text separation with unchanged unit arithmetic. Shared semantic/cache projection adds same-input cache/render/CLI/Python and warm/cold/held/fresh/reopen controls. Public implementation-facing additions add semver disposition, current archive dry-run/size. Bindings add final WASM compile/installed typing and runtime. Native compatibility requires pinned primary clauses and actual Word artifact provenance. Contribution issue mapping retains the full reproduction.

## Hash harness

Expected existing decimal cases unchanged. Roman/letter rendering is an intentional separately labelled behavior change. Namespace-capture corrections may intentionally add only the inherited alias/default or producer binding closure required to preserve expanded page-number and opaque-attribute identity. Ordinary canonical-w source stays unchanged. Any resulting source/output delta must trace to the qualified capture case and receive the same final combined review, never automatic baseline replacement. Any 49 hash/7 pixel delta must map to inspected section/field inputs and reviewed expected text/geometry before a successor baseline is recorded. F-X273 owns the baseline resource in this serial sequence, with one final combined account.

## Implementation checklist

- [ ] Close primary format/absence/rollover facts and shared provider semantics.
- [ ] Approve provider design, then require actual F-X273 local completion/ancestry.
- [ ] Implement concrete shared formatting/read policy in existing files.
- [ ] Keep raw XML/namespace authority and dynamic saved-cache behavior.
- [ ] Cover all A-IDs with discriminating source-qualified checks.
- [ ] Align HLD/docs and local scoped gate, independent microscope, handoff.
- [ ] Retain actual native/full-union debt until final acceptance.

## Open questions

No new user product choice. Current dynamic header/footer cache retention is the existing approved behavior, not an omitted issue fix. Absent format uses the existing application decimal fallback, not an invented schema default. Chapter composition and unknown formats retain display plus diagnostics, including explicit overrides, until qualified support is added. Existing Roman and alphabetic resource bounds stay in force. Native observations remain final acceptance obligations. These are evidence obligations rather than a request to weaken acceptance.

## Contract closure

Proposed concrete helper signatures are defined once above, no authoring API. All existing public operations retain signatures. Imported/absent/unknown states and read/preserve/update behavior are mapped. Provider design is approved before this dependent design. Primary numeric and namespace source authority is recorded, and the concrete helper/caller map defines implementation. Actual provider completion and final native observations remain required.

## Intake identity

Read-only GitHub refresh on 2026-10-28 found no body or contribution-head changes from the S96 intake. Original payloads and comments are retained under `/private/tmp/rdocx-s96-run-refresh`.

- Record 423, updated `2026-10-27T17:56:04Z`, body SHA-256 `a38f05c33ac3ec614982970cdf9dc08db53308deff619145f3fc326ce5cd56ca`.

## Primary specification pin

Official ECMA376 Part 1, fifth edition December 2016, retrieved 2026-10-28 from https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip. ZIP SHA-256 `9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129`,43501721 bytes. PDF and extracted text retained in `/private/tmp/rdocx-s96-run-refresh`.

Clause 17.6.12, printed pages 583 to 585, defines section format and omitted-start continuation. Clause 17.18.59 defines repeated-letter formatting and language-sensitive alphabets. Local letter examples above use explicit en-US. Preserve imported language and record unsupported/non-en-US cases rather than claiming all-locale parity. This closes the repeated-letter source fact. Native output and format-absence/chapter behavior still require qualification.

Clause 17.16.5.44, printed page 1234, defines PAGE and general formatting switches. Clause 17.16.4.3.1, printed page 1185, directly defines repeated-letter ALPHABETIC/alphabetic formatting. Decimal on omitted `fmt` is an explicitly retained application fallback. Unknown section/chapter composition refusal is the existing conservative policy, not an unsupported claim about native fallback. Explicit supported numeral switches override supported section numeral styles. No new compatibility claim outside the five declared formats and existing numeric bounds is made.

Design closure on 2026-10-28: consumer map, observable A-IDs, discriminating cases, public signatures, imported states, HLD impact and local/final evidence were checked against current source. Q-S96-01 is answered in F-X272. Approvals follow provider contract order. No material user decision remains. Pending native/hosted/installed observations are retained final acceptance debt, not recorded passes.

## Qualified TOC labels and helper absence semantics

Read-only consumer tracing found `field.rs:12066`, deterministic_toc_page_values, parsing rendered PAGEREF text as usize and constraining that displayed integer to the physical page count. Roman/letter text cannot satisfy that predicate, and a valid displayed restart can exceed the document's physical count. This consumer belongs to the existing TOC label/target and physical/displayed/formatted contract, not a new feature.

Use the existing `WordLayoutResult::bookmark_page_section(&self, name: &str) -> Option<WordPageSection>` qualification to distinguish actual physical target, displayed integer and owning section. Validate physical placement with physical page facts. Preserve the shared formatted PAGEREF label as text and prove its current supported formatting rather than parsing it as a physical index. Missing/stale/ambiguous targets, unsupported format and unresolved placeholder text must still refuse atomically. Do not turn a retained unsupported placeholder into a successful rebuilt cache. No new public method or model is needed for this consumer.

Local A2/A4/A8 controls include rebuilt TOC labels with supported roman/letter target sections, a supported displayed restart greater than the physical page count, and unknown/unsupported or stale-target negatives. Retain actual Base wrong-label/refusal outcomes. After the shared render fix, keep any intermediate label-parsing failure separately until qualified substitution succeeds. Check saved/reopened cache text, same-input render text and physical targets, not only successful rebuild status. Native and final composition evidence remain due at the final union.

The proposed `section_page_number_format` keeps the existing cache policy's Option meaning: absent page-number source or absent qualified format selects Some(decimal), while None signals unsupported or unqualified section state. An unknown qualified format can remain Some(unknown token) for the shared formatter to diagnose conservatively. The formatting helper does not treat None as an omitted-format decimal fallback. Unknown/chapter state remains unsupported even with an explicit override under the approved boundary. Keep both concrete signatures already defined and actual namespace-qualified reads. This closes an internal representation ambiguity without changing the approved five-format support, authoring surface, provider barrier, HLD impact or final gates.

## Original reproduction input qualification

Read-only preparation of the exact Issue 423 builder using python-docx 1.2.0 found an input distinction before implementation. The original pg helper appends a numbering element. Adding the second section first copies its inherited lowerRoman element, so the second section contains lowerRoman/start1 followed by decimal/start1. Both sections append numbering after cols/docGrid. ECMA376 Part 1 Annex A, EG_SectPrContents schema lines 1534 to 1555, defines a sequence with singleton pgNumType before cols and docGrid. The pinned extracted text at 188234 to 188255 confirms both the order and default maxOccurs 1. This is an input fact, not product RED or an observed native recovery rule.

Retain the exact generated original at `/private/tmp/rdocx-s96-run-refresh/issue423-original-input/roman.docx`, SHA-256 `f68b81a3d2ecdd64d789d24f377b591f7b456e1c8c870c694411d7b4f53a726a`. The first post-generation inventory assertion incorrectly expected the first second-section element to be decimal and failed. Its failure and complete actual inventory remain in preparation-qualified-inventory.json. No rdocx import, rendering or native observation ran. The original input bytes remain unchanged.

Use the separately labelled roman-schema-valid-intent.docx control, SHA-256 `cc422fa4211ad04b40c61a4c1b9a631146ebfa78074e23a807b587e9a9e3f566`, to discriminate the intended two-section formatting contract. It has exactly one correctly ordered lowerRoman/start1 and decimal/start1 element respectively. Body text and all other package part bytes are unchanged. The control never substitutes for the exact original A1 witness, whose actual Word recovery, output and saved source remain pending at the final native gate.

Existing document.rs:1565 to 1576 and the paired-element branch at 1921 model the first numbering element and preserve later duplicates opaquely. The existing duplicate_page_number_elements_preserve_source_order regression at 6034 explicitly expects the first start value. The shared format helper must not silently select a later duplicate or broaden that parser/authoring ownership policy. Preserve exact source and distinguish the original invalid input from the valid intended control in local and final evidence. If actual final native observations contradict the existing first-owner policy, reconcile that fact explicitly against this design before changing selection or claiming A1 complete. No original criterion or final gate is waived by this factual clarification. Independent reviewer fx272_scope_review confirmed the builder, XML, schema and current parser facts without product execution.

## Retained page-number namespace capture

Two namespace-helper repair attempts exposed an existing source-capture defect. Root designer and independent reviewer investigated before further edits. Root diagnostics against the unchanged existing rdocx-oxml artifact show that direct checked section-fragment construction can lose inherited q=Word and producer-w bindings from the retained pgNumType. Full document import with section-local declarations closes them correctly, while part-root inheritance can lose them from the isolated section carrier. The complete writer can then expose an originally foreign w:fmt as Word-qualified after publishing its canonical w binding. This is actual parser/writer evidence, not a native Word observation or a harmless formatting fallback. Exact commands, artifact/source/output hashes and all six source routes are retained in fx274-root-writer-probe.json and fx274-root-writer-probe-v2.json under `/private/tmp/rdocx-s96-run-refresh`.

Approve a narrow correction in the existing `crates/rdocx-oxml/src/document.rs` page-number capture paths. Preserve the actual effective URI map from namespace_bindings(word_prefixes), merged with current owner overrides, into the existing retained raw carrier before that authority is lost. Apply existing binding-closure helpers to both Empty and Start forms and both modeled and opaque duplicate page-number occurrences. Keep foreign-w bindings child-local when the section writer aliases producer attributes. Preserve ordinary canonical-w bytes where the established fixed writer already supplies the verified Word binding. Do not change unrelated section children, general fragment construction, model fields, serializer policy, first-owner selection or public API. This fulfills existing HLD04 source-preservation intent and changes no additional HLD intent file.

The approved section-format helper may read the existing section writer's temporary output under its canonical Word envelope, honoring all retained owner and child-local scopes. It must positively identify the modeled direct owner, excluding nested sectPrChange and later opaque duplicates. A modeled but unqualified owner returns None and cannot fall through to another qualified child. An absent page-number model is qualified absence. Writer/parse errors remain unsupported. Reuse computed section-format facts per section and operation where practical instead of serializing an owner per field.

Retain actual pure lower old-source RED before this correction, in addition to the earlier helper failures. Local controls cover direct fragment and full document import, part-root and section-root declarations, alias/default namespace, producer-w shadow and foreign attributes, Empty/Start, duplicate/nested exclusion, start mutation and source nonmutation. Verify preserved expanded names and opaque values after serialization/reopen, with ordinary canonical source unchanged. The existing parser/serializer, archive, WASM, source-qualified native, hash and full-risk union riders remain due. This factual capture repair introduces no new product choice or public surface. The owner acknowledges the revised contract while retaining the immutable claim Base and preserving the failed attempts.
