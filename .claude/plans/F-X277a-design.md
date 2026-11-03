# F-X277a, Emit authored radical degrees and preserve imported absence

**Status**: approved
**Sprint**: S96
**Size**: S
**Depends on**: none
**Test gate**: regression

## Problem

MathRadical::new in crates/rdocx-oxml/src/math.rs:1026 initializes degree_present false. The writer at:1096 consequently omits the degree child for an ordinary authored square root. Both native conversion routes and Python paragraph authoring publish this structure. Issue432 A277-1 requires a degree element after radical properties, empty for a square root. Imported absence must remain absence under the existing preservation contract.

This child owns exactly the original A277-1, including its import and public-consumer controls. A277-2 through A277-8 remain F-X277b. F-X277 is the completion umbrella. No acceptance row is dropped or counted as delivered by splitting it.

## Spec reference

- docs/hld/03-architecture.md, OfficeMath ownership and raw-boundary paragraph publication, around:837-845.
- docs/hld/10-bindings-spec.md, native equation conversion and Python Paragraph.add_equation, around:1942-1951.
- docs/hld/04-opc-and-packaging.md, preservation of unmodelled XML and schema child order. Mandatory parser/serializer reading, not a child implementation edit.
- docs/hld/06-presentationml-model.md, preservation and fixed-prefix serialization. Mandatory risk-route reading, no Presentation implementation work.
- docs/hld/14-development-backlog.md, F-X277a registration, regression gate and independent scope.
- Official Microsoft OpenXML SDK3.0.1, generated schemas_openxmlformats_org_officeDocument_2006_math.g.cs, Radical.ConfigureMetadata:1641-1654. URL and exact source digest are recorded below. Its particle is Sequence(radPr0..1, deg1..1, e1..1). The prior ECMA22.1.2.88 citation is historical context only. This preparation did not inspect the ECMA PDF.

## Source and schema pins

Product Base before metadata is ed588356bc50982516fff571eeba21bde8d76868. Current math sources, native/Python paragraph publishers and Rust integration carrier match that Base byte for byte. The actual claim must pin a reviewed metadata-only descendant whose product tree is identical to that Base. Any later provider assumption must be reviewed rather than silently changing the Base.

Issue432 body SHA256 is 43b11e97e001b5dcec2a11ae10b83f5cd2295b7764ac1bb3e5f72b4664d47a67. Its captured PR441 head is aca195c4412f65cfc497a955b34a3e52efd3e2e7, unchanged in authenticated07:35UTC and10:45UTC refreshes on2026-11-03. The latter is the latest retained observation, recorded in /private/tmp/rdocx-s96-github-refresh-20261103-1045/COMPARISON_RECEIPT.json. Reauthenticate mutable issue/PR identities before claim. The worker implements the contract rather than copying donor source.

Primary schema source: https://raw.githubusercontent.com/dotnet/Open-XML-SDK/v3.0.1/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_officeDocument_2006_math.g.cs. Version3.0.1, complete-source SHA256 4e182e11492f4dc80fab470438e8bd1d5683fe69dbde591f9a8fe1d49c6ecdf2. Source receipt and exact metadata excerpt are retained in the preparation packet's primary directory. Official package XML documentation separately confirms qualified child types, but the generated metadata is the sequence/cardinality pin. No schema validation or native result has been executed for this child.

## Approach

Change only authored constructor initialization so MathRadical::new(base: MathArgument) -> Self has an explicitly present empty degree, with hide_degree true. Keep MathRadical::with_degree(degree: MathArgument, base: MathArgument) -> Self unchanged. Keep MathRadical::from_raw's independent degree_present tracking and existing preservation slots unchanged. The current writer already writes a present degree in the proper slot.

No new source file, module, crate, feature flag, trait, generic, method, public signature or enum is needed. Existing unit, integration and Python test carriers hold the regressions. Do not rewrite all imported radicals to satisfy an authoring contract. Do not alter the math converter, bounds, raw parser, paragraph sidecar or bindings when a constructor correction suffices.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Authored empty degree | MathRadical::new, math.rs:1026 | Grammar constructor owns presence and hide_degree | Model, serialized XML and reopen |
| Explicit nth-root degree | MathRadical::with_degree, math.rs:1036 | Existing degree content and visibility stay intact | Model and serialized child order |
| LaTeX authored roots | equation_from_latex, native math.rs:2201-2215 | Optional index chooses with_degree, absent index chooses new | Public conversion, add_equation and package XML |
| MathML authored roots | equation_from_mathml, native math.rs:747-756 | msqrt chooses new, mroot chooses with_degree | Public conversion and package XML |
| Native paragraph publication | Paragraph::add_equation(OfficeMath) -> Result<()>, paragraph.rs:1041 | Serializes equation before adding existing raw-boundary and typed sidecars | Inline/display save and reopen |
| Python paragraph publication | Paragraph.add_equation(source, *, format, display=False), Py paragraph.rs:836 | Existing latex/mathml routes reuse native conversion | Actual installed call, saved document.xml and reopen |
| Imported missing/present degree and extensions | CT_OMath::from_xml/to_xml, MathRadical::from_raw, math.rs:1045 | Parser owns original presence and unknown-child slots | Import, serialize, reopen, serialize |

No dedicated CLI or WASM equation-insertion API is introduced. Existing downstream export/layout consumers receive the same typed base and hide_degree values. No pagination or glyph-position behavior is promised by this markup repair.

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| m:rad content | Written sequence is optional m:radPr, exactly one m:deg and exactly one m:e for newly authored roots. The writer emits radPr with degHide, so deg directly follows that element on an ordinary authored root | SDK3.0.1 Radical.ConfigureMetadata:1648-1652 and current writer:1087-1110 | C1/C2 sibling expanded-name assertions |
| Square root | Empty degree, degHide1, base unchanged. Empty base argument remains empty and gains no invented degree text | new:1026-1033 and Issue432 A277-1 | Zero/one/multiple authored roots |
| Nth root | Caller degree text and hide_degree false unchanged | with_degree:1036-1043, LaTeX:2201-2215, MathML:751-756 | Degree3 remains3 |
| Imported absence | Missing degree remains missing. Missing radPr still normalizes to radPr/degHide0 as the old writer does. This preservation promise does not claim the imported missing-degree radical is schema-valid | from_raw:1045-1084, writer:1087-1110, existing test:3647 | C3, independent of schema-positive authored documents |
| Opaque imported child | Retain exact raw bytes and existing schema slot. Do not byte-compare modeled degree content because MathArgument rewrites it | Preservation/from_raw:1075 and existing slot test:3680 | C4, raw substring and sibling order |
| Public API/errors | No signature, default, error type, accepted format, numeric conversion, ownership or atomicity policy changes. Existing invalid converter inputs still refuse before insertion | Paragraph::add_equation:1041, Python add_equation:836 | Existing conversion and publication controls retained |

Numeric quantities are presence, boolean visibility and existing expression content. There is no unit conversion, rounding or new range. Empty degree is meaningful. Existing caller errors and parser bounds remain unchanged.

## Dependency readiness

F-X272 is not complete and is not claimed as a provider. This leaf does not read or change its note identity, comparison, custody, resource-budget or checked-mutator mechanisms. MathRadical's constructor/writer and the exact native/Python equation routes are unchanged at the reviewed product Base. Final save/reopen on the integrated result remains required, since later C1 changes may affect a surrounding document parser.

F-X277b retains F-X272. F-X277 closes only after both children and the existing F-X272 edge satisfy their real gates. F-X280's existing F-X277 and F-X272 dependencies remain intact. Child approval/completion alone does not unblock F-X280.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A277-1a | Direct new(empty/text base), LaTeX square root and MathML msqrt | One empty m:deg after m:radPr, degHide1 and unchanged base | Grammar constructor and both converters | issue432_authored_radicals_emit_deg_and_imports_keep_absence | Actual generated DOCX SDK validation |
| A277-1b | LaTeX indexed root and MathML mroot with degree3 | Present degree3, unchanged visibility and order | with_degree and converters | Same bounded matrix | Integrated schema-positive output |
| A277-1c | Imported absent degree, imported absent radPr, present degree and an opaque child | Absence retained, old degHide0 normalization retained, present degree retained, raw bytes/slot retained | CT_OMath parse/write/reopen | Same matrix and existing preservation tests | Final raw-child preservation rider. Imported invalidity is not relabelled valid |
| A277-1d | Add the public converted roots as inline/display equations, save and reopen | Package XML has the same authored structure, equations and surrounding text remain present | Native Paragraph::add_equation | issue432_public_equation_radicals_keep_authored_degree | Final integrated Rust acceptance |
| A277-1e | Python add_equation with latex/mathml and display False/True | Actual saved XML agrees with native authoring and Python reopen retains equations | Installed Python binding | test_issue432_authored_radicals_emit_degree_through_python | Final installed Python3.9/3.12 suites and typing |

These subrows preserve original A277-1, not additional issue scope. Issue432's other seven rows remain with F-X277b and the umbrella.

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 authored | Empty/text base, one square root, two square roots, direct/LaTeX/MathML | Present empty degree for each root, exact expanded sibling order | Fixing one converter only, nonempty invented degree, duplicate degree |
| C2 indexed | LaTeX index3 and MathML mroot, inline and display paragraph | Existing degree3/visibility retained | Setting presence by replacing degree or always hiding degree |
| C3 imports | Missing degree, missing radPr, present degree | Exact documented distinction across parse/write/reopen | Always emitting degree on imported serialization or falsely claiming whole import byte equality |
| C4 preservation | Unknown child and unknown attributes under imported radical, explicit alias | Exact opaque bytes and retained slot | A serializer rewrite that drops or moves an unsupported child |
| C5 public publication | Native and installed Python, save/reopen, surrounding before/after text | Same radical structure without losing equations/text | Testing only constructor output while public package still wrong |

Do not cross-product every case. Existing conversion-error and preservation tests remain controls. No note, numbering, TOC, chart, table-cell or compare test belongs to this child.

## Rejected alternatives

- Always emit a degree in write_xml. That changes imported absence.
- Patch only LaTeX or Python. That leaves other constructor consumers inconsistent.
- Restore arbitrary donor math source. The current shared constructor already supplies the correction point.
- Add a field-presence setter, schema framework or new test binary. No current consumer requires them.

## Test plan

The sole story gate is regression, inherited from the registered A277-1 contract.

| Timing | Category | Test | Asserts |
|---|---|---|---|
| local scoped | regression | Existing math.rs unit carrier, new issue432_authored_radicals_emit_deg_and_imports_keep_absence and relevant existing optional/preservation controls | Actual old-source omission RED, then fixed-source order/presence GREEN. Imported controls must remain accepted and preserve their documented output |
| local scoped | regression/integration | Existing crates/rdocx/tests/integration_test.rs, new issue432_public_equation_radicals_keep_authored_degree | Actual public conversion, add_equation, save and reopen. Batch the related cases once |
| local scoped | check | cargo check -p rdocx-oxml -p rdocx --all-targets, affected warnings-denied Clippy and fmt | Exact source builds. Default inner loop is check rather than build |
| local scoped | binding discriminator | Existing test_core.py function test_issue432_authored_radicals_emit_degree_through_python | Actual local installed wheel from reviewed child source when an appropriate producer is available. If unavailable, retain pending status and do not invent a run |
| final sprint | schema | /private/tmp/rdocx-s96-openxml-sdk-validator/run-validator.sh --type docx --version Microsoft365 --input ABSOLUTE_GENERATED_DOCX --max-errors 0 | Actual final authored square/nth-root DOCX documents pass the independently pinned SDK. No imported missing-degree fixture is asserted schema-valid |
| final sprint | raw preservation | Integrated C3/C4 source-preservation controls | Opaque bytes/slots and import absence survive the final union |
| final sprint | normal full gate and riders | /verify --full, final installed Python3.9/3.12 plus typing, WASM check, applicable publish dry-runs/archive bounds, normal hosted CI and hash harness | Shared final integrated qualification once. Worker focused checks do not replace these |
| final sprint | native drawing residual | Existing Word/LibreOffice radical drawing observation | Pending residual under the parent design. It is not a local regression gate or a native parity claim |

Only actual executed results may be recorded. The worker creates the failing cases against its real Base before the production edit. Baseline setup/compile failures are not semantic RED. No workspace suite, external oracle or hash record runs per worker.

## HLD impact

- docs/hld/03-architecture.md
- docs/hld/10-bindings-spec.md

In HLD03's OfficeMath paragraph describe authored presence versus preserved imported absence. In HLD10's native equation conversion/Python paragraph describe the same public markup result and unchanged signatures. Do not edit the parent's notes/numbering/TOC sections or claim native rendering qualification. HLD14 and sprint metadata are integrator-owned split work, not worker product edits. HLD04 is mandatory reading but no child mechanism change is needed there.

## Risk routing

| Matched trigger | Required concern | Exact additional qualification |
|---|---|---|
| Parser/serializer, risk-routing row25 | Schema sequence and raw preservation | C3/C4 plus final SDK authored-document check |
| Public published API behavior, row28 | Strict SemVer and package compatibility | No signature change. Intentional authored XML delta, prepared Word0.17.0 minor release, final publish dry-run/archive riders |
| Binding consumer, row29 | Actual Python use, not similarly named native tests | A277-1e and final installed Python/typing/WASM riders |
| Compatibility/collection order, row33 | Pin source and imported distinctions | Primary SDK particle and sibling-order/import matrix |
| Multiple XML contexts, row35 | Inline/display and square/nth publishers | C1/C2/C5, no new placements |
| Shared semantic consumers, row39 | Same structure through each real entrypoint | Native and Python package comparisons |
| Grammar/empty variants, row40 | Empty degree and imported absence differ | C1/C3 |
| Completion/evidence metadata, row41 | Preserve counters and final rider debt | Root audits split/claim and applicable metadata gates |
| Contributed issue/PR, row42 | Whole issue accounting retained | A277-1 delegated only, other criteria stay with F-X277b/umbrella |

No layout/font algorithm, dependency graph, new public method or raw-input acceptance rule changes. Whole union checks remain once with the integrated owner.

## Hash harness

Authored roots intentionally gain one empty degree element. Samples without an authored radical are expected unchanged. Actual affected harness keys must be established at the final integrated hash wave, rather than guessed from the donor's old five-key umbrella estimate. The worker does not write baselines. The integrator records any actual intentional delta in its own reviewed labelled commit, deterministic font mode, with no unrelated source move.

## Review lineage and workflow

The parent has used all three design peers. This extracted unchanged criterion inherits that exhausted design cycle and its relevant accepted findings, including imported absence and no-radPr normalization. No fourth design peer or reset follows from the split. Root completed the affected audit of this child, the remainder and metadata contract before approving this child. The original implementation phase has used no passes for this criterion. Its maximum remains three. The split and these read-only engineering preparations do not reset a used cap or constitute design peers, implementation peers or product acceptance. Any later review starts from the genuine recorded source and preserves existing evidence.

The worker uses /start-feature F-X277a --claimed, /implement-feature F-X277a and scoped verification, then performs a clearly identified self-review and stops. A GLM self-review in the same implementation session is not independent microscope evidence. Root performs the actual independent microscope as a distinct read-only pass, or invokes an actual separate reviewer. The author repairs findings, with at most three implementation passes under the original cap, and stops again for the required independent assessment. Only after a clean actual independent pass does Root authorize /complete-feature F-X277a --prepare. Root final audit may supply the required independent pass and is not duplicated merely as a confirmation pass. Keep the original F-ID and all numbered A277-1 subrows in the handoff. Worker does not append sprint totals, push, dispatch, comment, tag or publish. Preparation is not final acceptance.

## Implementation checklist

- [x] Root approves the split contract and exact metadata-only claim Base, with source and workflow pins checked.
- [ ] Worker acknowledges its source/test/HLD lease and canonical plan revision.
- [x] Reauthenticate mutable issue/PR identity before claim, retaining exact existing pins. Fresh10:45UTC snapshot is unchanged.
- [ ] Prepare the complete bounded old-source regression batch and preserve real outcomes.
- [ ] Correct authored degree presence without changing import parsing, converters or public API.
- [ ] Complete native public and Python discriminator coverage in existing carriers.
- [ ] Update precisely HLD03/HLD10 sections before source freeze.
- [ ] Obtain scoped checks and identified self-review, stop for actual independent Root microscope, repair any findings within the original maximum of three implementation passes, and obtain a clean independent result before preparation.
- [ ] Prepare a validated handoff with all final union riders explicitly pending.

## Open questions

| ID | Affected rows | Decision or action | Status |
|---|---|---|---|
| Q-F277a-split | All | Adopt this independent extraction while preserving remainder dependencies, parent gates and review lineage | Approved by Root final1, unchanged A277-1 scope |
| Q-F277a-claim | All | Reconcile the existing dirty-canonical coordination exception to this exact child and product Base | Approved by Root final1, exact ed588 product Base and metadata-only descendant, temporary path /private/tmp/rdocx-s96-f277a-opencode |
| Q-F277a-python | A277-1e | Lease test_core.py to the worker while Root preserves its unrelated57-line carry outside Base | Reserved by Root, canonical F272 writer checkpointed and paused for adoption, unrelated Python carry preserved |

No unresolved product choice is introduced. These operational readiness rows must close before claim. Root approved this child after its affected final1 audit. Actual claim/Base authentication and worker acknowledgment remain required before implementation.

## Root approval and source lease

Root approved the unchanged child contract in .claude/reviews/F-X277-split-design-codex-final-1.md. F-X277b and the umbrella remain draft. The canonical F272 writer has paused at CHECKPOINT_01 for exact metadata-only staging. Its complete dirty source is guarded and excluded from the Base. Root reserves the stated math/test/HLD sections and approves the exact temporary worker path. F272 may resume disjoint implementation after Root returns its lease. Worker self-review stops before preparation for actual independent Root microscope. Final acceptance and all original rider debt remain pending.
