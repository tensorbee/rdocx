# F-X272, Preserve existing revisions during document comparison

**Status**: approved
**Sprint**: S96
**Size**: L
**Depends on**: F-X229, F-X230, F-X245

## Problem

`crates/rdocx/src/comparison.rs` rejects compared stories containing modeled revisions. Issue 425 requires a second reviewer to retain a prior redline while tracking changes elsewhere, or a documented limitation and usable workaround. PR427 head `c1e261ca63ebda5b7da1adec886ee40364a4e3de` supplies bounded owner preservation, with material exclusions. Its stale donor measurements and 116 missing main commits cannot be adopted.

## Spec reference

- `docs/hld/03-architecture.md`, checked comparison and physical story ownership.
- `docs/hld/04-opc-and-packaging.md`, "Producer read and writer authority", "The package" and "Checked Word operation source ownership".
- `docs/hld/10-bindings-spec.md`, native tracked comparison and `ComparisonOptions` contracts around lines 2857 to 2961.
- `docs/hld/12-testing-strategy.md`, "Word document-operation parity controls".
- `docs/hld/14-development-backlog.md`, F-X272.

## Approach

Approved decision Q-S96-01: adopt bounded source-preserving comparison. Unchanged revision-bearing owners retain source bytes. Changing, inserting, deleting or moving such an owner refuses with its physical story/owner location. Whole-table and block-control ownership boundaries remain explicit. The preceding paragraph mark is guarded when a final paragraph change would modify it. Changed revised section properties and whole comment-package replacement over revised comments also refuse.

Adapt PR427's own behavioral and binding/count commits onto current Base. Preserve current source authority, forms, live owners, dirty caches, tests and versions. Keep the existing id allocator's compared-story scope, including its ignored-story exclusions. Resolve postcondition copies by newly allocated ids only. Verify the original revisions survive byte for byte and in order. CLI counts the newly generated revision elements rather than total revisions. Do not introduce the donor generic helper unless two actual concrete instantiations justify it. Existing concrete count code is preferred.

Keep signatures unchanged: Rust `Document::compare(&mut self, edited: &Document, author: &str, timestamp: &str) -> Result<Vec<ComparisonDiagnostic>>` and `compare_with_options(..., options: &ComparisonOptions)` return the same result. Python keeps `Document.compare(self, edited: Document, author: str, timestamp: str, *, granularity: Literal["run", "word", "character"] = "run", ignore_formatting: bool = False, ignore_whitespace: bool = False, ignore_fields: bool = False, ignore_comments: bool = False, ignored_stories: Sequence[Literal["body", "header", "footer", "comment", "text_box", "footnote", "endnote"]] | None = None) -> tuple[ComparisonDiagnostic, ...]`, as inspected in `_rdocx.pyi:1215`. Empty ignored stories and `None` retain existing behavior. Invalid options map to `RdocxError`. CLI `rdocx compare` keeps existing options and atomic output behavior. No WASM authoring API is added.

For the refused cases, document a checked workaround in existing `README.md`, "Honest boundaries", and CLI README comparison guidance: work on independent copies, resolve revisions with existing checked all-story author/id APIs on both copies, then compare the resulting baselines. State explicitly that this discards the resolved tracked markup, retains no false promise of preserving prior revisions, and may require retaining the original redline separately. Use existing Rust `accept_revisions_by_author(&mut self, author: &str) -> Result<usize>` or `accept_revision_id(&mut self, id: i32) -> Result<usize>`, and Python `accept_revisions_by_author(self, author: str) -> int` or `accept_revision_id(self, id: int) -> int`. Resolve each relevant author/id on both copies before their existing `compare` call. CLI uses `revision accept --author` on both independent inputs before `compare`. Unknown/malformed selected changes retain current refusal policy. Test the exact public calls and counts, including a repeated author name where ids are needed to distinguish reviews. The user selected this recommendation on 2026-10-28. Broader revision-bearing-owner editing is outside this approved boundary.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Prior revision preservation | `compare`, `compare_with_options`, Python `compare`, CLI `compare` | Comparison source owner and revision scanner | No-op, staged compare, retained XML, accept/reject new ids, reopen |
| Id scope | Compared main/header/footer/comment/note/text-box stories | Existing allocator plus ignored-story filtering | Both input seeds, new allocations, same-id ignored controls |
| Live owner/cache publication | Native staged commit and Python candidate/revision handling | Existing checked publication boundary | Held versus fresh reads after no-op, success and late refusal |
| New-only counts | CLI human and JSON output | Existing all-story revision listing, difference by identity | Before/after compare, saved output |
| Unsupported edits | Existing checked resolution plus comparison | Explicit limitation/workaround, no raw XML editing | Independent copies, resolution, compare, original retained separately |

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| Compare versus Combine | Preservation is bounded Combine-like behavior, not legal-blackline parity | Microsoft Support "Compare document differences using the legal blackline option", step 8 and 9, retrieved 2026-10-28, https://support.microsoft.com/en-gb/word/compare-document-differences-using-the-legal-blackline-option | Native artifact observation still due |
| Multiple reviewers | Combine is Microsoft's route for retaining multiple authors' edits | Microsoft Support "Combine document revisions", retrieved 2026-10-28, https://support.microsoft.com/en-gb/word/combine-document-revisions | Supports distinction, not unrestricted algorithm parity |
| Owner preservation | Same revision-bearing paragraph/table/block control copied byte for byte, otherwise atomic refusal | PR427 pinned head, `HeldRevisions` and `compare_body`, candidate implementation not authority | A1 toA5 |
| New resolution | Resolve only ids allocated by this comparison for postconditions | Current revision scanner/allocator and staged accept/reject infrastructure | Prior author/count/bytes remain while new author's acceptance reconstructs edited input |
| Unsupported source | Unknown revision XML remains opaque, namespace aliasing preserved | Existing OPC/raw source contract | Alias and unmodelled-subtree controls, no general XML validator added |

## Dependency readiness

F-X229, F-X230 and F-X245 are done in BACKLOG. F-X271 is an execution-order tooling prerequisite, not a new formal product dependency. Final native/installed evidence stays pending. Q-S96-01 is answered and all S96 designs are approved. F-X272 has passed local qualification and independent full microscope pass4 and has a validated final-acceptance handoff. F-X271 is integrated with hosted and final acceptance pending.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Exact Issue 425 two-reviewer reproduction | First author's two revisions retained, new author deletion/insertion added elsewhere | Rust/Python/CLI | Adapt `issue425_comparison_keeps_existing_revisions_and_tracks_changes_elsewhere` and exact Python repro | Installed 3.9/3.12 and native retained revision inspection |
| A2 | Prior revisions in body, table, control, header/footer, footnote/endnote, text box, section properties | Unchanged owners retain XML/authors/dates/ids/location, noncolliding new ids in compared scope | All compared stories | Extend existing regression carrier with per-owner cases | Combined source-qualified native save/reopen |
| A3 | New author accepted or rejected | Reproduce edited or original input respectively while preserving first review | Native all-story resolution, Python, CLI | Author-filtered resolution and independent byte/order assertions | Installed suite and native author-filtered controls |
| A4 | Changed/deleted/moved/inserted revised owner, final paragraph mark, changed revised section | Under the approved bounded policy, location-qualified refusal and no source/live-owner/output change | Rust/Python/CLI | Adapt `issue425_comparison_refuses_changes_that_touch_an_existing_revision`, add moved/table/control cases | Limitation and executable checked workaround, Q-S96-01 |
| A5 | Existing revised comment plus comment graph change, ignored comments/stories | Whole comment replacement refuses if it would overwrite revisions, ignored sources follow current exclusions | Comparison/comment graph | Existing comment-package refusal plus ignored-story controls | Installed and combined graph acceptance |
| A6 | Identical redline, no-op and late invalid input | Zero new counts, exact no-op, late refusal leaves original and output intact | CLI human/JSON, native, held Python handles | New-only count and late-failure controls | Current installed runtime and strict typing/stubs |
| A7 | Unsupported edited revised owner | Documented limitation plus tested checked workaround, no preservation claim after resolving baseline | Existing README boundaries and CLI comparison guidance, public APIs and help/stub docs | Execute public workaround on same input | Issue 425 alternative acceptance, Q-S96-01 approved |

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | No revisions, unchanged prior revisions, changed revised owner | Existing clean comparison unchanged, preservation succeeds, unsupported edit refuses | Blanket refusal or accepting all old revisions in postconditions |
| C2 | Body and related owners, alias prefixes, table/control nesting | Exact prior revision bytes and qualified locations | Main-only scanner or local-name-only revision identity |
| C3 | Included/ignored stories and comments, existing same numeric ids | Noncollision within compared set, ignored contract unchanged | False global-id uniqueness claim or silently seeding ignored stories |
| C4 | Same new/old author name, no-op and late failure | Counts use new identities, not author name or final total | Counting all revisions by the requested author |
| C5 | Held/fresh Python owner handles before save | No-op/refusal retains identity and values, success follows established invalidation policy | Reopen-only assertion hiding stale live handles |

## Rejected alternatives

- Compare accepted views silently. It loses the prior markup the issue asks to keep.
- Treat donor owner preservation as full Word Compare parity. Microsoft's legal-blackline behavior differs.
- Hand-edit XML as the workaround. It bypasses checked publication.

## Test plan

**Test gate**: regression, verbatim from the backlog.

Local scoped: current-Base RED and fixed GREEN for adapted Issue 425 cases in `cargo test -p rdocx --test regression_test issue425`, focused comparison lib cases, `cargo test -p rdocx-cli --test integration issue425`, and selected installed Python Issue 425 nodes only after authenticating a successful fresh extension build. Batch related cases in existing binaries. Existing unchanged behavior controls are not falsely recorded as RED.

Final sprint: native all-story revision and author-filtered resolution, installed Python 3.9/3.12 complete suites, strict mypy/stubtest, CLI composition and `/verify --full` union. Genuine affected archive verification and size under 10 MiB, actual README examples and current measurement. No donor archive claims.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`

## Risk routing

Contributed acceptance adds pinned full issue/PR rows and omission mapping. Parser/serializer and XML preservation add alias/schema-order and opaque-byte controls. Compatibility, multi-consumer and cached projections add same-input native/binding/CLI and held/fresh/reopen comparisons. Public behavior/binding docs add final WASM compile, installed typing/stubs and affected archive dry-runs/size. External native oracle adds actual producer/tool/input/output identities. Metadata adoption adds affected final policy checks. No new trait, module, crate or source file. The existing main-key and all-story revision count paths instantiate the concrete multiset helper with tuple and StoryRevision values today. The accept-new-id and reject-new-id closures are the two existing callback instantiations for captured postcondition resolution.

## Hash harness

Expected unchanged 49 hashes/7 pixels. New comparison behavior is a separately labelled commit. No baseline change is predicted from this story.

## Implementation checklist

- [x] Resolve Q-S96-01 and close concrete workaround signatures before approval.
- [x] Reconcile own PR427 commits against current Base without donor measurements.
- [x] Guard all revised-owner mutations, preserve exact old bytes.
- [x] Prove newly allocated id postconditions and new-only CLI counts.
- [x] Cover all A-IDs and held/fresh/late-refusal controls locally.
- [x] Align public docs/stub/help and existing README boundaries.
- [x] Update listed HLD intent, scoped gate and independent microscope.
- [x] Prepare the validated final-acceptance handoff.
- [ ] Retain final native/installed/archives debt until actual integrated acceptance.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision and recommendation | Status and answer/source |
|---|---|---|---|
| Q-S96-01 | F-X272 A4/A7, combined S96 acceptance | Recommend bounded unchanged-owner preservation plus documented/tested checked resolve-and-compare workaround. Changes within a revised paragraph, whole table/control or revised section remain unsupported. Alternative: expand implementation to carry revisions through such edits, increasing scope and requiring a revised design/estimate | Answered 2026-10-28 by the user: "Receommeneded", selecting recommended bounded preservation plus tested checked workaround. The issue permits this alternative |
| Q-S96-02 | F-X272 A2/A3, ordinary simple-field instruction edits and all-story resolution | Permit a standard complex-field comparison representation for clean admitted instruction edits, explicitly retaining the observable form change after accept/reject. Neither source guards nor opaque metadata may silently weaken. Result-only edits retain simple form | Answered 2026-10-29 by the user: "Complex fields (recommended)". Exact design closure and schema/native/lifecycle proof remain required before repair and acceptance |
| Q-A430-28 | F-X272 Issue 430, G430-02 shared-owner native positive | One-release exception solely for the unobserved positive native case with two physical footnoteReference elements targeting one normal definition. All remaining library, SDK, binding/CLI and actual comparison-output native obligations stay mandatory | Answered 2026-11-03 by the user: "Approve the narrow exception (Recommended)". Status: unobserved under user-approved acceptance exception, never passed. Exact decision and evidence: /private/tmp/rdocx-cross-review/Q-A430-28-native-evidence-decision.json. Canonical amendment below governs later adoption. No implementation or publication approval |

## Q-A430-28 one-release native acceptance amendment

The user explicitly answered "Approve the narrow exception (Recommended)" on
2026-11-03. The exact decision is
[/private/tmp/rdocx-cross-review/Q-A430-28-native-evidence-decision.json](/private/tmp/rdocx-cross-review/Q-A430-28-native-evidence-decision.json),
question call_caca9c39f5cc465892c11a39f1e82bc4. This amendment governs canonical
later adoption. Frozen V13 proposal and review artifacts stay historical and
unchanged.

For this release only, the positive native observation with two physical
footnoteReference elements targeting one normal definition is **unobserved under
user-approved acceptance exception**. It is never recorded passed. The bounded
probes establish no admitted fixture, not a universal Word limitation. The
[V3 alternative-revision negative](/private/tmp/rdocx-cross-review/native430-alternative-reference-v3/NATIVE_SCALAR_FACTS.json)
remains negative. The
[V4 NOTEREF positive](/private/tmp/rdocx-cross-review/native430-noteref-v4/NATIVE_SCALAR_FACTS.json)
uses a different graph. The
[V5 actual tracked move](/private/tmp/rdocx-cross-review/native430-native-move-v5/NATIVE_SCALAR_FACTS.json)
uses separate owners. Neither positive satisfies the waived observation.

A430-6, A430-7 and Q-A430-22 remain unchanged, including every physical reference
occurrence, retained-definition cleanup, complete custody-unit selection, atomic
partial-selection refusal, original unused definitions and destination reference
inventories. Library lifecycle, reopen and selective resolution, applicable
bindings/CLI and full SDK validation including expected-invalid controls remain
mandatory for shared-owner imported fixtures. Retain their native admission
failures. G430-02 selected-reference cleanup and native checks of actual library
comparison outputs from Word-admitted inputs remain mandatory in pending,
selected and resolved states. A Word-refused comparison output is not exempt.
If an admitted fixture emerges before final acceptance, qualify actual library
outputs against it.

Stable follow-up R-A430-28 in the existing backlog requires re-evaluation on a
later Word release. This is not a permanent silent waiver. The answer changes no
feature status, totals, provider readiness or Issue430implementation authority.
C1 remains incomplete. No reporter comment, source integration, sprint completion
or release publication is approved by this answer.

## Contract closure

Q-S96-01 is answered. No new public comparison method or option is proposed. The additive doc-hidden lower native residual bridge is the explicit technical exception recorded below. Exact Python defaults and checked all-story author/id resolution signatures are transcribed above from current binding/stub and `revision.rs:251to278`. The checked workaround has actual local Rust, CLI and fresh authenticated Python evidence, including independent author-resolution copies followed by comparison and retention of the original redline. Historical and current source-qualified receipts are recorded below and in progress. Compatibility source distinction is verified. Final integrated and native saved-artifact oracle riders remain pending.

## Intake identity

Read-only GitHub refresh on 2026-10-28 found no body or contribution-head changes from the S96 intake. Original payloads and comments are retained under `/private/tmp/rdocx-s96-run-refresh`.

- Record 425, updated `2026-10-27T20:08:05Z`, body SHA-256 `3f4f9b5577506d36ec13cdb09f145d551cce7701e93f4fc53419f0a74f78d26b`.
- Record 427, updated `2026-10-27T20:59:51Z`, body SHA-256 `5cef6bc8e0440076c8da5c46778ee22595c208f6698f367d2fbe9635208a46f4`.
  Contribution head `cc3473d0c05659167847a941e0c7eb2bbd7217f5`.

The draft initially assumed a `gaps.md` file existed. Repository inventory shows none. Put the limitation and checked workaround in existing README and CLI comparison documentation instead. This is a documentation-location correction, without changing the accepted bounded policy or adding an unapproved source file.

Design closure on 2026-10-28: consumer map, observable A-IDs, discriminating cases, public signatures, imported states, HLD impact and local/final evidence were checked against current source. Q-S96-01 is answered in F-X272. Approvals follow provider contract order. No material user decision remains. Pending native/hosted/installed observations are retained final acceptance debt, not recorded passes.

## Source-qualified namespace clarification

The actual added aliased-main-root regression exposed the existing literal `w:body` extraction assumption after bounded revision preservation was implemented. Preserve that full input. Select the direct Word body by expanded namespace identity, carry document/body declarations to the existing owner-span wrapper and postcondition readers, and use the existing fragment namespace closure to add a missing canonical Word binding for generated markup. Existing producer bindings must never be overwritten. A foreign namespace body lookalike is a separate discriminator. Imported self-closing bodies retain their existing refusal boundary here. This concrete source-context correction stays in the existing comparison file and the approved A2/raw-preservation scope, with actual failure and successor results retained.

## Implementation-time contribution refresh

Read-only refresh on 2026-10-28 after the user scope answer found PR 427 rebased onto released main `24af95d3d94c25e95f61a75e7e21f18fa95b6363`, with current head `c1e261ca63ebda5b7da1adec886ee40364a4e3de`. Its current body SHA-256 is `dc490c52d857c06751586bf2a945ee1ecb6cc81922b03c21cf40b0b3544c643d`, updated `2026-10-28T06:43:32Z`. Issue 425 body is unchanged. Historical initial pins above remain attributed to that intake.

The current four own commits are `56523be00a3c58b4b730717af99ce66e57911f88`, `e2796f0567d26c47688fdc8057d8cde21417ca06`, `9d0744301b604fdac003686d47690a67028a5100` and `c1e261ca63ebda5b7da1adec886ee40364a4e3de`. Comparing per-file added/removed payloads confirms all product code, tests, stub and HLD contribution changes are identical to the original pin. Changed README/archive measurements remain excluded and will be remeasured on the final union. The issue acceptance and Q-S96-01 decision are unchanged.

Raw refreshed bodies/comments/PR metadata, commits and exact current diff are retained under `/private/tmp/rdocx-s96-run-refresh`. The running worker keeps Base `c8acc502b74f29ed7e9546939f17007709a9f22b` and its dirty implementation/evidence. Its owner must acknowledge this canonical pin revision before further implementation. No reset/rebase or source replacement is authorized by this factual refresh.

## Namespace descendant implementation clarification

The N1/N2 investigation is consumed for private source context, without changing bounded revision ownership. Actual effective namespace declarations accompany source owner fragments. Unchanged revision-bearing owners are emitted from their exact original slices. Source-capable descendant scanners resolve expanded child and parent names, and source closing boundaries come from the source reader. Child slices receive their parent scope through the existing namespace closure without rebinding producer declarations. Generated canonical Word markup is bound locally when an existing producer w binding prevents a missing-only part-root binding. No new public API, source file, module, crate or trait. The original plain aliased-root reproduction and foreign-body control remain separate. New discriminators cover mixed inherited aliases, an empty aliased paragraph, opaque foreign parents, changed unrevisioned aliased tables/controls and a producer w binding. All new cases require actual pre-next-repair RED before GREEN.

## Source relocation contract clarification

Independent namespace investigations identified lost source scope across comparison descendants and revision paragraph merging. This is an A2/A3 preservation repair within the approved bounded policy. It does not authorize edits inside prior revision-bearing owners. The existing `revision.rs` renderer is the private provider for main and related-story accept/reject operations, including all, author, id, id-set and date selection. Its paragraph merging, kept-wrapper promotion and property replacement must preserve the expanded names of transported elements and attributes.

Use actual source ancestry and destination context at relocation. Preserve the first merged paragraph shell and attributes, the last paragraph's property precedence and each child's own source scope. Carry incoming wrapper promotion into a merged shell, preserve child-local shadowing and an absent default namespace, and never globally rebind producer w. Ordinary unchanged-context traversal retains its bytes. Add only namespace closure needed for actual relocation. Keep selection, resolution counts, schema order, atomicity and both comparison reconstruction postconditions intact. Existing parent links and local declarations provide source ancestry. No new public signature, parser, trait, module, crate or source carrier is authorized or needed.

Local evidence adds the actual final-mark failure with inherited q properties and foreign producer w, plus direct accept/reject resolution with conflicting source paragraph scopes, a child-local shadow, opaque element/attribute identity, save/reopen and a related story. Preserve genuine pre-repair outcomes for newly added controls. Batch the affected cases with `unwrapped_revisions_keep_wrapper_namespace_bindings`, `rejected_property_changes_keep_owner_namespace_bindings`, `contextual_paragraph_markers_merge_the_adjacent_paragraphs`, `contextual_cleanup_preserves_foreign_property_shells`, `paragraph_mark_removed_before_a_table_removes_an_emptied_paragraph` and `rejecting_paragraph_property_change_keeps_mark_and_section_properties`. Retain prior comparison, ignored-source, workaround and revised-owner refusal controls. These are focused local regressions, not complete suites or native oracle acceptance.

The existing HLD03/HLD10 work list remains unchanged and must describe this shared source authority. Namespace investigation pass 2 records the observed 12-test result, 11 passed and 1 failed, separately from source-traced challenges and prior setup/semantic failures. Full feature microscope and every final rider remain pending. Before implementation resumes, the running owner must acknowledge the canonical revision containing this clarification, preserve Base and dirty work, and reconcile this section into its annotated worker plan without resetting it.

## Relocation context qualification

The shared implementation transports changed inherited source bindings only at actual relocation and preserves child-local overrides. It does not infer namespace use only from element or attribute names, because compatibility attribute values also require scope. ECMA-376 Part3, fifth edition December2015, clauses7.2and7.4 require in-scope bindings for Ignorable and MustUnderstand prefixes. The primary archive is [ECMA Part3](https://ecma-international.org/wp-content/uploads/ECMA-376-3_5th_edition_december_2015.zip),741741bytes, SHA256 `42294159fbbbe9393ccadac95b859d7729cc68d908898bcbe31034dda059daa8`. Authenticated text and pin receipt are retained in the existing temp intake. Ordinary same-context output stays byte-identical.

Direct Word-property destination selection and insertion both use actual direct expanded Word identity and the rendered fragment's external namespace scope, rather than the first lexical QName anywhere in opaque descendants. Local evidence and final readiness are recorded separately below when observed.

## Revision identity clarification

Independent investigation pass 3 establishes N3, separate from N1 transport: an unrelated public comment edit removes redundant paragraph namespace declarations while retaining identical old revision bytes and effective bindings. `CT_Revision` equality currently mistakes private parser-context enumeration order for an owner change. The fresh authenticated Python observation is three passing cases and one semantic failure, retained with its original wheel/module identities. No actual revised-owner edit or new user decision is involved.

Repair only identity of `CT_Revision`'s complete private parser context in the existing `crates/rdocx-oxml/src/revision.rs`. Retain the original parser vector and its order for parsing, editing and writer prefix selection. Existing PartialEq may compare those complete entries without depending on enumeration order, preserving multiplicity and every binding value. Every other field remains compared, including raw XML, kind/id/author/date, typed and projected content, nested revisions and parse policy. Do not sort the shared namespace scanner, discard context, add a semantic XML comparator, normalize published original bytes or weaken revised-owner refusal.

Add actual pre-repair low-level controls in existing carriers for empty, paired and nested/projected revision identity under reordered equivalent context, with unchanged raw bytes and parser input. Meaningful foreign/Word/default/undeclaration differences remain unequal. Preserve observable raw and projection differences. Add the matching native public comment reproduction, including original paragraph/revision bytes and comment graph after save/reopen and accept/reject of the new author. Rebuild and authenticate a successor wheel before rerunning the same four selected Python nodes. Preserve the failed predecessor and all historical observations.

Scoped checks now include the affected `rdocx-oxml` crate, in addition to the existing changed native, CLI and binding crates. Refresh the affected bounded local tests, including all 22 prior namespace, relocation, ignore, refusal and workaround controls. HLD03/HLD10 remain the complete HLD work list. The running owner must acknowledge this canonical clarification and reconcile it into the annotated worker plan before implementation, without changing Base or resetting dirty work. Full microscope and every named final gate remain due.

## Historical identity-only implementation handback

The dirty worker remains on Base `c8acc502b74f29ed7e9546939f17007709a9f22b`. Canonical factual pin `59dc1f48`, relocation clarification `10d45b11` and identity clarification `0987372c` are acknowledged without reset/rebase or replacement of worker annotations. That identity-only predecessor was locally qualified for independent full microscope. Its listed evidence remains attributed to that historical producer. Subsequent full review and property-source remediation supersede its readiness status. Namespace investigations are design escalations, not a clean feature microscope.

| Local observation | Actual result | Retained evidence under `/private/tmp/rdocx-s96-run-refresh` |
|---|---|---|
| Focused lower revision module | 20 passed, including actual pre-repair identity RED and namespace/projection/edit/policy controls | `fx272-identity-unit-{red,green}.log` |
| Native regressions | 23 passed, preserving all prior 22 controls and public comment reproduction | `fx272-identity-corrected-frozen-native.log` |
| CLI regression and documented workaround | 1 passed | `fx272-identity-corrected-frozen-cli.log` |
| Changed-crate compile and lint | rdocx-oxml, rdocx, CLI and Python all-target checks and all-feature clippy passed | `fx272-identity-corrected-frozen-native-cli.json`, `fx272-identity-clippy.log` |
| Formatting | Passed on frozen product source | `fx272-identity-corrected-frozen-fmt.log` |
| Fresh installed Python 3.12 | Same four selected nodes passed, 643 deselected, 5.04 seconds | `fx272-identity-successor-python.json`, `fx272-identity-successor-python-green.log` |

The successor dev wheel is semantic local evidence, not release-size or performance qualification. Complete product source manifests match before build, after build and after tests. Installed native bytes match the successor wheel member. The failed authenticated predecessor and all intermediate setup, compile, semantic and test-expectation failures remain preserved. Detailed artifact identities and original paragraph/comment graph assertions are in the existing progress record.

Pending before completion: a different non-implementing agent's full feature microscope and structured handoff/integration. Final integrated acceptance still owes full workspace and complete Python 3.9/3.12 suites, strict mypy/stubtest, native all-story/author resolution, actual producer oracles and hosted jobs, 49 hashes/7 pixels, WASM, actual README examples, refreshed measurements, affected release archives and their size under 10 MiB, plus the plan's risk and metadata-policy riders. No ledger, commit, push, publication, native-parity or final acceptance claim is made by this worker.

## Full microscope pass1 remediation checkpoint

The earlier local handback is historical evidence for its named authenticated producer. Full microscope pass1 reports D1 and D2, with zero smells/nitpicks. D1 has actual CLI RED and fixed2/2GREEN using stable story coordinates. D2 has actual four-context token-boundary RED and a first repaired24-test result,23passed1failed: token parsing now succeeds, but an opaque rPr root attribute is lost after final property resolution. Root additionally observed a generated-row namespace failure on the authenticated predecessor, not yet qualified on the new candidate. Product writing stops for focused independent reassessment of property source authority and generated publishers. Current implementation is not locally ready for integration or acceptance. Exact receipts, source trace and write-stop are in the existing progress record.

## Property source and publication closure

Independent investigation pass 4 clarifies D2 and reopened N1 without changing the bounded user policy. Parsed opening boundaries fix the quoted-attribute parse failures, but the actual successor loses an opaque rPr attribute after property reconstruction. Typed CT_RPr deliberately does not retain opaque root attributes, so comparison must carry private ORIGINAL and EDITED raw property-source snapshots beside typed formatting precedence. No new public model field/type, parser for unknown attributes, public API, trait, module, crate or source carrier is authorized or needed.

Retain each side's direct pPr/rPr source shell, root QName, effective namespace context, attributes and raw children before adding comparison markers. Carry the selected snapshot through emitted paragraphs, trailing and moved predecessor marks, final insertion/deletion and block-control/related-story callers. Preserve existing own/merged accepted/rejected precedence and the final paragraph's own section break. Patch only intended direct Word property/change slots, respecting schema order. An absent property owner differs from an empty owner carrying opaque metadata. Unknown-formatting policy remains left-biased where it already was. Audit ordinary paragraph and run-property regeneration under the same raw-source authority instead of leaving typed-only equivalents beside the final-paragraph repair.

Every fresh canonical Word container must close at its actual publication boundary, including generated rows, row properties, cells, paragraphs, tables, controls, fields and property early returns for equal or ignored formatting. Reuse the existing qualifier policy. Retained raw fragments keep their actual source context. Do not globally rebind producer w or put a source producer-w attribute onto a fresh Word-w root with changed meaning. A marker's local Word binding does not qualify its ancestors. Preserve old revision bytes, atomic refusal and both reconstruction postconditions.

The existing consumer map is extended by pass 4's cited publisher inventory. Local checks retain all original quoted-opening cases and add distinct own/merged property values, absent versus metadata-bearing empty owners, a source Word q:rPr with opaque producer-w attributes and a supported block-control or related-story final path. Assert expanded identities, opaque bytes, intended formatting and both author-resolution sides after save/reopen. Obtain current-source RED for the canonical/alias/foreign-w row diagnostic, then a bounded inserted/deleted row, nonempty property and existing move/control matrix that separates fresh containers from retained fragments. Do not add a new move capability or use the predecessor wheel as current-candidate qualification.

D1's stable kind/part/owner count repair has actual local CLI evidence. N3's equality-only repair remains unchanged. Refresh only affected scoped evidence and build/authenticate a fresh successor wheel for changed native source before the same selected Python checks. HLD03/HLD10 remain the work list. Full microscope pass 2 and every final sprint rider remain pending. The implementing owner must acknowledge this canonical revision and reconcile this section into its annotated plan before resuming, preserving Base, dirty source and every historical observation.

## Private live-source ownership invariant

Raw lookup, when needed, is confined to an operation-local stack frame whose keys identify only the immutable original and edited typed parse owners borrowed by one compare_body call. Each frame is created from those actual raw bodies/spans/bindings and popped on every Result exit, before those owners can drop. Nested control/related body parsing pushes a separate frame and restores the still-live parent frame. No global registry, historical address entry or temporary cloned property is registered. Lookup occurs at the live-source boundary and copies owned source snapshots into accepted/rejected property sides before adding markers. Those owned snapshots, not later pointer lookups or tracked output, carry precedence through final/move/control publication.

## Property projection and resolution boundary

Independent investigation pass 6 separates three actual boundaries. Rejecting a complete prior rPr duplicates a current opaque child inside the same run owner. Granular text comparison can legitimately copy formatting onto several resulting runs. A canonical paired-pPr unrelated-edit probe loses metadata before comparison is invoked. Preserve the failed receipts and these distinctions. No new public scope or user decision follows.

Repair complete prior rPr replacement in the existing shared resolver for ordinary runs as well as paragraph marks. Retain independently required revision markers and selection accounting, without appending current formatting to an already complete prior. Keep pPr base, mark and section roles separate. Intentional identical opaque siblings remain distinct. Never deduplicate XML to hide the defect. Make the text-changing fixture owner-aware: every resulting run retains its exact opaque property payload once and intended modeled formatting, with ordered expected text. Preserve exact global counts for owners that did not split and the same-owner duplicate discriminator.

A private source-backed property projection may reconcile the parser's Empty versus Start representation only when the actual direct Word property source is positively checked. Retain root QName, opaque attribute lexical bytes, expanded identity, encoded and decoded values, effective bindings and opaque child identity. Exclude a generic raw Empty token only when this positive property identity represents that exact source owner. Account for all reader event classes, including text, CDATA, comments and processing instructions. Either retain their exact bytes or refuse the special empty-container reconciliation. Declaration-only generated empty rPr can qualify, while nonempty or attributed rPr cannot. Keep positive Empty/paired controls and negative removed/changed attributes, meaningful binding changes, encoded values, opaque children and non-element content. This is not a generic semantic XML comparator or a waiver of physical preservation.

Actual comparison, selective accept/reject, save and reopen must preserve paired and Empty property metadata, source namespaces and raw children. The additional later unrelated authoring-edit probe is exploratory differential evidence, not a new F-X272 capability. Its canonical paired-pPr baseline fails before comparison at the existing typed main-part serialization boundary in document.rs. Preserve and attribute that inherited limitation in the existing plan/progress and final report, without claiming it passed, pinning the bug as desirable behavior, or expanding this feature into a public CT_PPr model or general serializer repair. Compare any later-edit observation with the identical pre-comparison baseline. Newly introduced loss in the authorized comparison lifecycle remains blocking. Keep all required comparison lifecycle cases and separate the exploratory baseline from their gate.

No new source file, public model/API, trait or parser for unknown attributes is authorized. Refresh affected scoped tests and fresh Python artifact qualification after correction, then require independent full microscope. HLD03/HLD10 and all final sprint obligations remain unchanged. The worker must acknowledge this canonical revision before resuming, preserving Base, dirty work and historical observations.

## Historical property-publication remediation handback

Base remains c8acc502 and canonical acknowledgments59dc1f48,10d45b11,0987372c,ade6458f and0e9ae23a are reconciled without reset/rebase. Full microscope pass1 findings D1/D2 have locally qualified remediation. N1 source/publication and N3 equality retain their separate invariants. The six independent investigations are historical escalations, not full microscope pass2.

Current source uses bounded live input frames and owned original/edited property snapshots before markers. Complete source child slots preserve opaque order and namespace authority. Actual retained run children remain exact under equivalent source context, and genuinely fresh Word containers qualify at publication, including default-Run inline controls. Complete prior rPr replaces ordinary or paragraph-mark formatting while preserving scoped markers and intentional opaque sibling multiplicity. Empty/paired default-property projection positively checks source metadata and all reader event classes without changing public parser/model fields or generic XML equality. Paired and Empty comparison, selective resolution, save and reopen cases pass with actual source identity assertions.

| Current focused observation | Actual result | Evidence under `/private/tmp/rdocx-s96-run-refresh` |
|---|---|---|
| Frozen native provider and comparison controls | 28 passed,2468filtered,8.90seconds | `fx272-pass6-publisher-final-native.log` |
| Positive/negative source projection unit | 1 passed | `fx272-pass6-publisher-final-projection-unit.log` |
| CLI counts, no-op, refusal and documented workaround | 2 passed | `fx272-pass6-publisher-final-cli.log` |
| Four changed crates all-target check and all-target/all-feature clippy | Passed, denied warnings | `fx272-pass6-publisher-final-{check,clippy}.log` |
| Formatting and source applicability | Passed, before/after source hashes identical | `fx272-pass6-publisher-final-native-cli.json` |
| Unchanged lower revision identity unit evidence | 20 passed, exact lower source unchanged | `fx272-identity-unit-green.log` |
| Fresh dev-wheel Python3.12, same four selected nodes | 4 passed,643deselected,5.95seconds | `fx272-property-successor-python.json` |

The new wheel/native/import and complete before/build/after-test source identities are recorded in the existing progress file and artifact receipt. This semantic dev artifact does not qualify release size or performance. Earlier artifacts, actual semantic failures, setup/compiler failures and the intermediate test-only expectation failure remain immutable.

The canonical paired-property unrelated authoring-edit loss remains separately attributed inherited exploratory evidence. The exact probe fails before comparison is invoked, and no general serializer or public-model repair was authorized. Its source and actual before/after XML remain under the evidence directory. It is not a failing normal feature gate or desirable pinned behavior. Required comparison lifecycle preservation remains blocking and has passed the current controls.

Pending before integration readiness: a different non-implementing agent's full microscope pass2. Pending final integrated acceptance: complete workspace/native/hosted/installed Python3.9/3.12, strict typing/stubs,49hashes/7pixels,WASM,real README examples and measurements, affected release archives/size, producer oracles and all approved risk/metadata-policy riders. No commit, push, ledger update, self-microscope, release or final acceptance is claimed by this worker.

## Protected source identity clarification

Full microscope pass 2 records actual D2 bypasses in five contexts: nonempty paragraph pPr, paragraph-mark rPr, ordinary run rPr, a whole protected table and a whole protected block control. Changing a producer property attribute value is a meaningful edit inside the protected owner even when the typed model omits that attribute. The approved unchanged-owner refusal takes precedence over ordinary non-held unsupported-formatting left bias. This clarifies the existing Approach, A4 and owner-preservation row without adding a capability or user decision.

Establish complete protected source identity before typed equality can authorize raw passthrough. Reuse physical owner spans, both actual source contexts and existing private source inventories. Cover paragraph, whole table and block-control boundaries and their equivalent main/related callers. Do not register temporary owners or introduce public model fields, a parser for unknown metadata, a new source file or generic semantic XML equality. A targeted value-only patch is insufficient if equivalent unmodeled metadata on a modeled property child still bypasses the same guard. Preserve changed attribute values, bindings and opaque payloads as meaningful input differences. Benign namespace declaration ordering, redundant closure and compatible public comment edits remain supported. Keep the existing compared-story exclusions and ordinary non-held left bias.

Obtain actual current-source RED for the five recorded contexts and a discriminating opaque attribute on a modeled property child. GREEN must show location-qualified refusal, exact original bytes and earlier revision identity, no output publication and applicable held-handle preservation. Retain positive reordered/redundant namespace and compatible-comment cases, the ordinary non-held left-bias control, and ignored-story policy controls. Keep previously earned property/publication/resolution successes applicable where their product inputs are unchanged. Refresh affected scoped checks and authenticate a fresh successor Python artifact after the guard repair, then require independent full microscope pass 3. Full sprint obligations and HLD03/HLD10 remain unchanged. The worker must acknowledge this canonical revision before resuming and preserve Base, dirty work and all earlier evidence.

## Modeled meaning and source residual authority

Independent investigation pass 7 records actual regression in the proposed all-attribute guard. The authenticated predecessor accepts a public adjacent-comment/save route that normalizes bold `val=true` to bare `b` and spacing `00120` to `120` while keeping old revision 7/A unchanged. The current guard refuses that route. A separate valid predecessor table case normalizes tblW, gridCol, trHeight and tcW values `00120` to `120` and succeeds with its outside comment. Preserve these actual positives and the current failure. The earlier table probe with content after terminal sectPr is invalid setup, not semantic compatibility evidence.

A protected owner requires its existing typed modeled identity AND positively checked omitted source residual. The supplement must not repeat known modeled identity through lexical attribute spelling or count known-only nodes that the writer may omit. Preserve already supported typed normalization without adding a general semantic normalizer. Retain unknown attributes including unknown Word attributes, opaque root metadata and events, association, sibling order/multiplicity and meaningful namespace context. A successful comparison still emits the exact original protected owner and old revision bytes.

Explicit technical approval: add one concrete read-only implementation-facing native bridge in the existing `crates/rdocx-oxml/src/properties.rs`, exported as `rdocx_oxml::properties::protected_source_residual`, with signature `pub fn protected_source_residual(xml: &[u8], inherited_bindings: &[(String, String)]) -> Result<Vec<String>>`. Use the existing lower-crate Result type. Mark it doc-hidden and document its output as an opaque internal identity representation, not an authoring or persistent interchange format. This additive exported symbol is a narrow exception to this plan's prior no-new-public-API wording. It adds no public model type/field, Python/WASM method or authoring operation. The semver disposition is additive, without changing existing signatures.

Keep modeled-source classification at the owning lower layer. Reuse existing property attribute/overlay authorities where they match the actual parser, adjusting crate-private access as needed. Numbering's Paragraph/Run subset and text-trimming overlay are not complete authorities for every parser or event. Reconcile the actual paragraph/run/section, table/row/cell/grid, control-property and modeled-content consumers. Do not copy a second schema into comparison.rs, invent a blanket namespace exclusion, classify every accepted validator attribute as modeled, or alter numbering serialization incidentally. The complete-event residual scanner must preserve source data that typed readers omit. Existing parser/writer behavior remains the authority for supported attributes and partial raw carriers.

Associate residual data with actual owner ancestry and recognized property/content slots and occurrences, retaining opaque multiplicity and order. Known-only node/default omission must not shift another owner's residual identity. A source node with genuinely opaque attributes or nonempty metadata cannot disappear through the default reconciliation. Foreign lookalikes and value-only namespace dependencies remain opaque and source-qualified. The existing HeldRevisions physical scopes, section guard and ignored-story policy consume this one bridge before typed equality authorizes raw copy.

Add actual current-source GREEN for both public predecessor compatibility routes, keeping exact original lexical spellings in emitted protected owners, new comment graphs and selective resolution/save/reopen. Add independent numeric/toggle and applicable already-supported default representation controls. Mixed-node tests combine known formatting normalization with unchanged opaque data, then change producer or unknown Word attributes to require refusal. Keep all17 refusal contexts, text-node metadata, opaque encoded values/events/bindings, whole-owner boundaries, benign namespace, ordinary-left-bias and ignored-story controls. New helper units must challenge modeled ownership and complete residual content. Refresh affected scoped lower/native/CLI/Python checks and authenticate a fresh successor wheel, then require full independent microscope pass 3. HLD03/HLD10 and final sprint obligations remain unchanged. The implementing owner must acknowledge this canonical revision before work resumes, preserving Base, dirty work and every prior observation.

## Residual-source local qualification before full pass3

Canonical d3a028e1 is acknowledged and implemented through the sole additive doc-hidden lower native bridge. Existing typed equality establishes modeled meaning, while parser-owned source residuals establish omitted metadata identity before protected raw passthrough. Owning value checks retain unsupported known-key values, unknown Word/producer attributes, opaque events, source association and effective bindings. No numbering behavior, public model, authoring API or Python/WASM method changed.

Actual frozen local evidence is recorded in `fx272-residual-final-native-cli.json`: native30/30, lower88/88 including numbering/revision/residual units, positive source projection1/1, CLI2/2, four-crate all-target check, denied-warning all-target/all-feature clippy and fmt. Both exact public comment canonicalization routes pass while emitted protected owners retain original raw spellings and old7/A bytes. The17original source-refusal contexts, unsupported-known-value controls, opaque content metadata, namespace/ignored-source positives and ordinary nonheld left bias remain in the bounded batch. Genuine predecessor/current-source RED and failed setup/compile logs remain independently attributed in progress.

Fresh dev successor qualification is actual4/4 on the same Python nodes, including the protected metadata refusal extension. Receipt `fx272-residual-successor-python.json` authenticates wheel SHA256 `541224310d3d7d2bcbc255f3e2d04a797ccc421ccdacbd2a4e5868df61f3ce0b`, wheel/imported native SHA256 `32560b83cdfeea04e23910b7b40fa31abe4f2c4ff91ca15f3e5190ccccef105a` and unchanged product manifest `ee557b98818464ca1fb6806e299e0ffdce250a8047042d4de8d12c7ad42d51a2` before build, after build and after tests. This dev artifact establishes focused semantics, without claiming final release/archive measurements.

Local implementation is ready for a different agent's full independent microscope pass3. D2 resolution requires that review and is not self-certified. F-X272 remains in-progress, Base c8acc502 and dirty source/history preserved. All final integrated full-suite, native/hosted, installed Python3.9/3.12, typing/stub, WASM,49hash,7pixel, README/measurement, archive-size, oracle and risk riders remain pending. The inherited pre-comparison paired-property authoring limitation remains explicitly attributed and outside the bounded serializer scope. No final acceptance, integration, commit or ledger update is claimed.


## Opaque-root namespace closure remediation

Full microscope pass3 retains D2 at the residual unknown-root literal boundary. The existing canonical d3a028e1 benign namespace contract already admits public raw-carrier closure. This is a narrow implementation correction, without another API or scope change. Actual current native RED covers producer children in pPr and mark rPr, an unknown Word child and a whole-table caller. The same public adjacent-comment/save/reopen route fails all three Python shapes on authenticated native32560b83, while the lower opaque-root unit records the literal-blob difference. Distinct `fx272-opaque-root-*red*` logs preserve those actual observations.

The lower scanner now positively records opaque-root expanded name, every non-declaration attribute identity and encoded value, effective scope and shape independently of root namespace declaration plumbing. All bytes after that opening remain untouched residual source, including inner events, descendant-local declarations and closing markup. Sibling association/order/multiplicity remains checked. Known lower ownership/value/default logic is unchanged. Revision subtrees keep entire literal source identity, and successful publication still emits the original protected raw owner. Updated local qualification is recorded below. Independent full microscope pass4 remains required before scoped integration readiness.


## Opaque-root successor local qualification

Actual repaired candidate passes all30prior native controls in `fx272-opaque-root-final-native.log`. The new test initially overconstrained a selectively rendered outer known paragraph shell. Its captured diagnostic proves exact properties, opaque child and old revision bytes survived while only redundant outer namespace declarations were added. The corrected new native case passes1/1 across all4contexts and both resolution directions in `fx272-opaque-root-qualified-native.log`. Immediate comparison output still requires the exact original whole held owner. Selective saved/reopened outputs require exact original properties/old revision bytes, independently read expanded opaque metadata and multiplicity, text and comment graph.

`fx272-opaque-root-test-only-applicability.json` reconstructs the three historical source/test files by reversing only the later test assertions. All reconstructed hashes match the earlier receipt and all other inputs agree. No non-test implementation changed after that candidate. Prior30passes remain applicable without replay. Qualified receipt `fx272-opaque-root-qualified-native-cli.json` records lower89/89, projection1/1, CLI2/2, four-crate all-target check, denied-warning all-target/all-feature clippy and fmt, with final source/test inputs unchanged and actual CLI application authenticated.

NEW dev successor wheel SHA256 `5f9dfbfc3c95156141ffb28217aa0e1661d5376cf466f11199f14abeae0d50de` and wheel/imported native SHA256 `ba936fc0610d59b71ec31b0d878ec8a136de73f441f84af146c2be77170049ad` are recorded in `fx272-opaque-root-successor-python.json`. Product manifest `db09bb8db0f0d27ea67de5ce9a63be2fcc7c99de61e67d7ff4fd20fbdae08ffd` agrees before build, after build and after tests. The SAME4Python nodes pass, including all3actual public namespace-move inputs, prior-redline/comment preservation, typed snapshots and atomic refusal/live-handle/checked workaround controls. Build and selected tests use a dev artifact, without release measurements or final installed qualification claims.

Local implementation is ready for a different agent's full independent microscope pass4. D2 remains subject to that review. Base c8acc502, dirty source, old artifacts and every failed attempt remain preserved. All final integrated/native/hosted/full-installed/typing/WASM/hash/pixel/archive/README/measurement/oracle/risk riders and the attributed inherited authoring limitation remain explicit. No final acceptance, commits, push or ledger update is claimed.


## Independent pass4 and preparation evidence

Independent full microscope pass4 reports0defects,0smells,0nitpicks and approves scoped integration readiness. Root authenticated all241current source/test inputs,227product files, executed Rust binaries and embedded CLI application, shipped Python stub/init, fresh wheel and isolated native import in `fx272-opaque-root-root-auth.json`. Root then actually executed the current full named native selection31/31,2468filtered,7.73seconds, preserving source and executable identities in `fx272-opaque-root-root-current-native.json`. Binary SHA256 is `135efe500377b06fb878e827a8f50668be5ab859e6147835a75d96f3d132040c`. This supersedes historical shared-executable applicability ambiguity, without relabelling the earlier failed command.

The durable independent review `.claude/reviews/F-X272-all-pass-4.md` records current source/artifact identities, all acceptance rows, local outcomes and final debt. Qualified lower89/projection1/CLI2/check/clippy/fmt and fresh SAME4Python observations remain applicable and are consumed without reruns. HLD03/HLD10 intent and all product inputs stay frozen after review. Preparation retains approved status, in-progress delivery records, the inherited unrelated-authoring limitation and every named final acceptance rider. The worker does not complete delivery or edit sprint ledgers.


Prepared handoff `.claude/handoffs/F-X272-ready.md` names source producer `daba1a47c71e9fd1607a7b1fc5c1515514420092`, an ancestor of the metadata tip, with immutable Base c8acc502. Validator passes the sprint-final-acceptance checkpoint. No crates/HLD intent changes follow that producer. The plan stays approved and final delivery checks stay unticked pending actual integrated acceptance.


## Full07 physical field-span ownership clarification

Root approved one narrow additive, doc-hidden native ownership method on2026-10-29 after independent source investigation and before remediation. The exact signature is `CT_P::field_source_run_fragments_at(&self, run_idx: usize) -> Result<Option<(usize, Vec<Option<Vec<u8>>>)>>`. This is an implementation-facing native bridge, not an authoring operation, persistent exchange format or Python/WASM API. No new public model, type, field, module or source file is authorized. The earlier no-new-bridge restriction remains binding for every other addition.

Use the existing lower `field_span_at` and `field_span_parts` authority. A successful result belongs only to the actual still-eligible shared physical field span beginning at run_idx. The first value is its positive modeled-run coverage. The vector has exactly that length and physical source order, with Some containing the exact materialized outside-field run fragment, including its actual source QName, properties and namespace shadows, and None at field entries. Retain the existing materializer output before parsing the outside run, rather than serializing a typed CT_R to invent source metadata. No intact eligible span, or a shape whose correspondence is not positively proved, returns None. Reader failures remain errors. Do not publish a partial vector or fabricate a one-to-one physical shell correspondence. Fragments remain operation-local and no global owner registry is introduced. Existing lower callers keep their behavior.

Comparison source snapshots use this bridge for ordinary property/run sources inside shared spans. Whole-paragraph publication and run replacement correlate actual retained intervals and modeled coverage through existing field-source ownership instead of greedily counting detached field runs. Direct and nested controls retain full source events beside modeled payloads. Preserve supported text around fields, sibling fields in one run, overlapping field owners and their actual distinct first, intermediate and last property roots. Discriminators include inherited aliases and producer-w bindings as well as canonical Word input, unchanged spans and changed or detached no-span controls.

The existing HLD03 and HLD10 impact list remains exact. Describe the bounded bridge and source authority there before freeze. Every existing no-revision comparison and both accept/reject reconstruction checks remain binding. The full07 and hosted failures remain actual RED evidence. This technical ownership amendment does not broaden the unchanged-revision-owner editing policy or claim final native/installed acceptance.


## Physical raw-slot occurrence identity through comparison fragments

The second actual preservation repair passes the retained70native regressions but a
new current-CLI sibling discriminator refuses acceptance. Its original/edited inputs
contain an earlier opaque child and a later closure-equivalent child with authored
declarations. Both raw spellings and their original order must survive each selected
input role. This is an existing preservation requirement, not a new public capability.
The atomic refusal, full source identity and prior green checks stay separately
retained. Fresh wheel/lint and final acceptance remain pending.

After the mandatory independent investigation, root approves this bounded private
repair on2026-10-29. The read-only report is
`/private/tmp/rdocx-s96-run-refresh/fx272-authored-closure-sibling-probe/independent-investigation.md`,
SHA256`82d9d3c459f4be2874112f0bed64ff8e77ccd9618713773257003bd8dd6cf29d`. The prior implementing owner stopped all
writes/builds before reassignment. No third repair precedes this investigation.

A source-backed child is identified by input role, physical owner and original source
occurrence/boundary, not by XML bytes, hashes, a normalized value or namespace-closure
spelling. Closure is a relocation representation of an already identified occurrence.
Keep actual raw slots and typed-child roles distinct. Run-root metadata is not a child.
Comment records, alternate-content projections and private text-box markers retain
their specialized existing roles. Qualified unknown Word children remain opaque too.

Extend only the existing operation-local private RunSource/AttributedRunUnit carriers
and helper arguments. Capture the ordered relation between actual lexical spans and
the parsed owner's existing raw-slot ordinals before attribution. Carry those original
ordinals, owner roles and needed original boundaries through attribute_raw_children,
word/character units, coalescing and merge_unit_runs. Retain the current same-owner-pair
coalescing restriction. Whole-run publication explicitly selects complete surviving
content occurrences. Fragment publication explicitly selects just its actual units,
including an empty selection when no opaque content survives. Never restart an
unqualified whole-owner search for each fragment.

At publication, validate generated role/boundary coverage against the selected source
occurrences, then emit each identified occurrence's exact raw bytes under a positively
equal destination scope. At a real relocation, transport that occurrence from its
actual scope, preserving authored declarations. Generated/changed modeled text,
delText, drawings with remapped relationships and other typed children retain generated
authority. Any unchanged typed-child source restoration also needs explicit source
association and transformation checks. Generated XML spelling is a consistency check
after identity is established, never an identity search key.

Keep left formatting/shell and right inserted content authority separate. Check
ordered multiplicity and assignment across all fragments of each surviving input
role. The same source occurrence may appear once in deletion and once in insertion
only under their separately selected roles. It cannot be reused twice among sibling
fragments of one surviving role. Do not deduplicate equal payloads, prioritize a global
exact-byte match, use a per-call cursor over the whole source or normalize opaque XML.
Both reconstruction postconditions and atomic refusal remain unchanged.

The publisher work list is shared publish_run_source/preserve_run_source, whole and
granular ordinary runs, whole paragraph children, controlled runs and unmatched inline
controls, including related/table/note/text-box callers. Preserve marked_direct_runs_xml
source-span publication. Property-slot current/prior source precedence and the complete
prior-rPr resolver remain unchanged. No additional public method, model, type, field,
module, source file, trait, generic, registry or serializer metadata API is approved.
The earlier exact field_source_run_fragments_at exception remains the only new lower
bridge for this remediation.

Retain the actual CLI RED. Add ordered[A,B],[B,A],[A,A,B],[A,B,B] closure-equivalent
sibling controls, plus raw children at distinct typed boundaries that yield a later
B-only word/character fragment. Assert exact per-position bytes, multiplicity and
current/prior formatting after accept/reject/save/reopen. Include differing original
and edited payload lists, ignore_formatting, aliases, foreign/default binding shadows,
QName-valued opaque metadata, root metadata and specialized raw records. Route decisive
cases through ordinary/controlled fragments, whole paragraph changes and an applicable
related/table/note/text-box boundary. New semantic REDs require actual pre-repair source
execution. Source tracing alone stays labelled as such.

Refresh affected scoped native/lower/CLI checks, strict check/clippy/fmt and a fresh
authenticated opt2debug wheel for the exact five hosted failures and established four
Python controls. Require a different actual microscope reviewer afterward. Existing
HLD03/HLD10 must state the source-occurrence invariant before source freeze. All final
workspace/native/hosted/installed/typing/hash/pixel/archive/risk gates remain due. The
reassigned implementing owner must read and acknowledge the revised plan content hash
without changing Base or discarding any previous source/evidence.

## Paired textbox formatting policy without changing host revision granularity

The actual occurrence-inventory controls expose original bold becoming edited italic
under ignore_formatting=true in four unrevised main-story textboxes. Preserve the
exact REDs and source recipes. Actual authenticated predecessors refuse these exact
operations, so they prove neither successful formatting compatibility nor an inherited
passing behavior. Actual pre-slot-fix CLI bdd4260f accepts five clean host-axis controls.
Those current positives remain binding, including changed geometry and inserted,
deleted and moved hosts. Broad main-story recursive dispatch would introduce different
host-shell/count refusal behavior and is rejected.

Root's bounded technical decision is to retain existing outer-host revision granularity
and existing main dispatch. At a genuinely paired outer replacement, qualify physical
source raw-slot and qualified pict/txbx path ownership before applying original-formatting
policy to the edited inner body. Use existing comparison alignment and private source
carriers. Text, drawing names, byte equality and namespace-closure equality are not
identity. A changed-formatting pair whose ownership cannot be established unambiguously
refuses atomically with its physical location. Unpaired inserted/deleted hosts retain
existing behavior. Option false is unchanged. No material new user decision was
identified by the independent closure investigation.

Reuse existing inner paragraph/table/run comparison policy rather than inventing an
rPr-only formatter. The admitted projection compares actual inner owners with the same
options, snapshots allocator state, captures only the genuinely newly allocated IDs,
and selectively accepts those temporary revisions in isolated namespace-qualified inner
XML. Every input-held revision span remains exact and ordered, and no selected temporary
ID may escape the projection. Temporary allocator gaps must preserve uniqueness and
all existing output/count controls. Descent is strictly into a qualified source child,
with progress checked to prevent parent re-entry or duplicate projection. Global held
and reconstruction checks remain unchanged.

Splice the projected inner content only into the admitted edited txbxContent interval.
Preserve edited host geometry, source attributes and opaque host events outside it.
Within the interval, the existing source-authority comparison policy governs original
formatting and edited content. Carry that derived edited-content authority explicitly
through raw-slot publication, with exact interval validation. Do not claim a policy-
projected subtree is the unchanged original raw source, or weaken source coverage to
accept arbitrary generated opaque XML.

All five paired publishers must reach the same source-qualified policy: compare_paragraph,
compare_granular_paragraph, compare_complex_paragraph, compare_control_content and
compare_control_runs_with_options. Complex paragraphs directly publish policy output,
so a publish_run_source-only hook is insufficient. Existing preserve/equal/insert/delete,
structural row/control markers, related recursion and ignored owners keep their actual
policies. Include a discriminator for every applicable publisher or a source-backed
exclusion, not an assumption from one simple textbox.

No new public method, type, field, file, module, trait, generic or flag is authorized.
The existing private selective resolver may receive minimum crate-private visibility.
Alternatively an internal narrow selective helper must own substantive validation of
held-span byte/order preservation and complete removal of selected temporary IDs. A
wrapper that only forwards is forbidden. No public accept_all, temporary document
package, author-name filtering or second independent formatter is permitted.

Keep the four original ordered/multiplicity cases and execute accept/reject independently
so the first failure cannot hide the second observation. Add a minimal no-opaque same-
host bold Old to italic New, equal-text formatting-only, geometry-plus-formatting,
new inserted italic, multiple/nested hosts, aliases/local Word shadows, unchanged held
revisions and ambiguous changed-formatting atomic refusal. Preserve the five actual
clean host positives, direct/related/control/complex publisher controls, ignored Main
and TextBox behavior, exact opaque occurrence order and both global postconditions.
Capture actual pre-patch outcomes before implementation. Test assertions do not change
merely because the old postcondition excludes ignored formatting.

Independent reports are fx272-textbox-policy-flow-v1/diagnosis.txt, SHA256
413c1d4c9caced641c9e12346505f926508e26cbb2253ade58873ccccd6c0cf1,
and fx272-textbox-host-compatibility-v1/final-caller-and-invariant-addendum.txt,
SHA256902d7139149c4463e1ff212f07bdc704b0b6fb06cb9ea6bd2fc48d9ea65998c2.
The independent closure verdict and all actual five-host/predecessor outcomes remain
retained alongside them. After the occurrence implementer explicitly stopped all writes
and builds, root approves this bounded technical amendment on 2026-10-29. It closes
existing clean comparison policy C1 and A2/A3/A6 lifecycle checks without relaxing the
user-approved bounded existing-revision decision or adding public surface.
HLD03/HLD10 remain the exact HLD work list. Fresh scoped checks, opt2 wheel, independent
microscope and all final native/full/installed/typing/hash/risk acceptance remain due.

The assigned implementing owner must read and acknowledge this full canonical plan
content identity before affected work. No implementation lease follows from approval
alone. Root assigns each implementation separately and preserves one writer/build owner.

### Paired formatting-only no-op publication

Root clarified the existing A6/C1/C5 contract on 2026-10-29 before implementation.
When a positively source-qualified paired projection removes every change in a
formatting-only textbox edit, retain the original source owner and emit zero new
records. Identical outer deletion/insertion records would manufacture a content edit.
Require actual projection and source-role proof, not equal text or namespace closure
as occurrence identity. Genuine remaining geometry, content and opaque changes keep
the established outer revision pair. Preserve the no-op lifecycle, including original
source bytes, held owners, public values and held handles. An equivalent staged clone
must not replace the receiver merely because it has zero records. Keep both global
reconstruction checks and add only the minimal source-proved no-op publication guard.
Capture equal-text formatting-only and text-changing controls before the policy patch,
with independently observed acceptance and rejection.

## Independent nested-body guard and supported textbox representation closure

The first paired-policy successor actually passed71of74 tests. Preserve that failure,
the two diagnostic preparation compile errors, actual accepted/rejected vectors and
all earlier REDs. The successful diagnostic shows original inner bold/intended text,
but adds extraction declarations on pict/rPr and fails the acceptance policy identity.
Rejection independently matches original. The equal-text case still emits an outer pair.
The qualified direct anchored WPS diagnostic separately succeeds at comparison but
incorrectly accepts edited italic. Neither cause is closed by the previous source patch.

Root approves two distinct bounded corrections after reading the independent
fx272-textbox-postcondition-investigation-v1/diagnosis.md in full and confirming the
implementer's stop checkpoint. The original and edited reconstruction comparisons,
prior-held raw checks, atomic publication and all no-op/live-owner requirements remain.
This changes the defective internal policy representation, not the required equalities.

First restore extraction-only lexical closure at the actual admitted host/property
source boundary. Preserve actual lexical templates and source/destination scopes.
Remove only a positively proved added redundant root declaration at that exact owned
root. Preserve authored declarations and lexical values, unknown attributes, descendant
bytes and all opaque event order. The guard must not conceal this publication error.
Do not weaken exact host equality/no-op suppression, globally strip namespaces, or
manufacture equality by copying original inner XML before proving nonformatting content
unchanged. A zero temporary-ID set alone is insufficient for that proof.

Second add a substantive guard-specific nested Word-body identity provider, used by
normalized_package and normalized_story_part for both reconstruction comparisons.
It independently projects each actual input. It never invokes comparison, allocates or
resolves revisions, borrows an expected document or substitutes expected content.
Do not change global run_signature, drawing_signature, semantic_run_raw or option-run
signatures in ways that alter matching, source pairing or outer revision granularity.
Do not dispatch by function-pointer identity or a global mask. Shared move normalization
keeps its current behavior. If a shared source projector must change instead, stop and
close its main/related move asymmetry before implementation.

Represent each admitted raw host by its ordered exact outside-body source segments and
ordered recursive Word-body policy identities. Use a collision-proof private structured
identity, never a marker serialized into XML or a raw string used as an ownership key.
Actual physical paths/ordinals associate the identity with its real source/model owner.
They do not become source-specific equality payload: byte offsets, original/accepted
input role and run partitions may differ legitimately after resolution. Preserve the
existing canonical ordered content boundary semantics when comparing the independently
projected values. Text/kind/control/table/cell/field/opaque differences remain significant.
Only the existing typed Word-body formatting policy is excluded when requested.

Keep the already relationship-normalized outer model. An outer story reparse must not
discard canonical docPr or relationship payload identities. Each admitted inner body is
parsed with actual inherited/local/default/undeclared namespace context, normalized with
the same actual Document and physical OPC part owner, then passed through the existing
source-property/postcondition provider recursively. Main/header/footer/note/comment
owners are distinct. Images/charts/hyperlinks with different IDs but equivalent targets
retain existing graph policy, and a changed target remains meaningful. No raw-ID-only
shortcut or caller-default main-part substitution is permitted.

Close both supported representations in the publisher and the guard:

- VML pict in CT_R.extra_xml uses captured raw-slot occurrence/boundary authority.
- Direct anchored WPS in RunContent::Drawing uses actual typed-slot authority and
  CT_Anchor.shape.text_body, including tables/controls. Keep canonical non-text drawing
  fields and edited/generated geometry/relationship authority, while replacing only
  the admitted inner policy identity. The actual source/raw_xml serialization authority
  must be updated coherently and reparsed where needed. A typed-only mutation whose old
  raw_xml still wins is invalid. Use authoritative text_body once. A supported legacy
  text fallback without body requires real source association or the approved qualified
  refusal, never fabricated source intervals.
- AlternateContent uses existing selected Choice/copy authority. Associate the selected
  modeled drawing with its actual physical source interval, project it once, and retain
  unmodeled Choice/Fallback custody exactly. Do not infer correspondence from counts,
  text, glyph names or equality. Do not project the typed view and the same raw body twice.

CT_Inline currently has image/chart/raw custody and no typed shape-body projection.
The actual revised probe proves that boundary. No inline shape modeling is added here.
Retain its existing unchanged opaque source behavior and source preservation. It cannot
be used as a passing modeled-textbox policy test. A qualified-looking textbox under an
unsupported opaque wrapper remains opaque and significant. Namespace==Word alone is
not admission. Reuse existing upper document/story-owner grammar and compatibility
branch authority, with minimum crate-private visibility where necessary. No new lower
public export, public model, authoring method, file/module/trait/generic/flag or general
XML framework is approved.

At each recursion select only outermost admitted bodies, then descend strictly into
qualified source child intervals with checked owner/progress. Nested boxes recurse in
the inner pass. Foreign w/aliases/default shadows and undeclarations cannot acquire Word
identity through a synthetic wrapper. The same concrete ownership grammar must serve
publication and the guard, without copying a second permissive tag allowlist. Existing
held-revision checks stay earlier and exact. All five paired publishers remain covered,
including complex direct publication and both controlled granular routes.

Required actual evidence includes the original four raw cases, both minimal no-op/text
controls and independently executed accept/reject observations, all existing71passes,
and the genuine anchored WPS RED. Add independent guard discriminators where only
formatting differs versus changed text/kind/table/control/field/opaque order/host geometry.
Perturb actual accepted/rejected content to show both guards reject. Test raw/typed/MCE
physical authority, selected/unselected copies, foreign/opaque decoys and one projection
per owner. Retain source-exact pict/rPr spelling and live no-op owners. Batch aliases,
namespace shadows, related part relationships, held revisions and all five publisher
controls in existing carriers. Do not borrow parser/layout passes as comparison proof.
The initial mixed inline probe is recorded as contaminated by fixture overreach, while
the qualified anchored probe is an actual semantic RED.

This remains the existing A2/A3/A6/C1/C2 source/policy/no-op contract. HLD03/HLD10 remain
the exact HLD work list. Independent microscope, fresh opt2 installed binding proof,
scoped lower/comparison/CLI/lint checks and all final native/full/typing/hash/risk gates
remain due. No baseline update, commit, push or final acceptance follows from approval.
The assigned owner must acknowledge the full new canonical content hash before a second
policy repair. A further unsuccessful repair of the same invariant requires fresh
independent escalation under implement-feature discipline.

## Diagnostic-only escalation after the second policy failure

The second policy attempt executes six library passes and three failures. Its74
regression selections are built but not executed. Preserve the exact checkpoint
fx272-policy2-stop-identity.json, diff12dc5d3b and product mappingc6789535. Root has
independently authenticated every227product input and all tracked stop identities.
The different investigator's diagnosis is retained at
fx272-policy3-independent-investigation-v1/diagnosis.md, SHA256
de1c516fb369d6144739ea1560cf0b826c2353694cf6a87b53573cc67587ffd0.
Root has read that report in full. It is an investigation, not a clean microscope.

No third product repair is approved. Root authorizes a separate diagnostic-only
phase in the existing comparison test carriers. Keep every production result,
control-flow decision, source association check, reconstruction equality and test
assertion unchanged. Any added production-site observation is cfg(test) only.
Do not correct namespace closure, pairing, publication, canonicalization or no-op
behavior during this phase. No new tracked file, public surface, residual/schema
framework, parser capability, baseline, commit, push or completion claim is added.

Batch the actual failing cases and their existing original controls. Collect the
physical stage, OPC owner, qualified source path/slot/boundary, original versus
model raw bytes and first difference at each raw-role refusal. Retain the actual
pre-wrapper source and post-drawing-closure source separately. Do not identify a
physical occurrence by matching these byte strings. For canonical/alias/default
and direct/selected equal-text cases, record revision IDs/counts, both halves of
the exact no-op predicate, package/part first differences, temporary IDs, lexical
source templates, source/destination scopes and held owner identities. Record
accepted versus edited and rejected versus original vectors independently, even
when one fails. Include typed non-text fields, source namespace bindings and
selected/fallback custody without replacing the guard's actual input.

The source trace identifies descendant drawing closure before physical capture
as a plausible shared cause. It does not supply the missing actual stage trace,
first differing bytes, alias no-op counts or expanded WPS signature difference.
Do not report those unmeasured values as observations. Capture them before a
bounded third implementation contract is approved.

Add diagnostic-only guard and public comparison discriminators for unknown
metadata attributes on inner modeled w:t, with unchanged text, changed text and
protected-owner controls. Preserve existing ordinary unsupported-formatting left
bias. This phase approves neither silently excluding previously significant raw
metadata nor importing the entire protected residual policy into unheld content.
Any required classification decision returns to design with the actual evidence.

Use one batched native opt1 test execution with assertions, overflow checks and
8MiB stacks after compile preparation. Retain individual outcomes and all aborted
or unexecuted cases. Existing failed outputs remain immutable. Stop afterward,
authenticate all diagnostic source and evidence, update the existing progress
record, and return the sole writer/build lease. Root reviews the observations and
reconciles the canonical contract before any third product patch. Final scoped,
independent microscope, binding, native and integrated union debt remains due.

## Third policy correction after measured independent reconciliation

Root has read the actual diagnostic handback and the different investigator's
contract-proposal.md in fx272-policy3-contract-reconciliation-v1, SHA256
7137dd915440739ee8490bcda5b1bf7fde0a725d6f0c36c584e7d4a67897c16b.
All227product inputs and eleven tracked diagnostic checkpoint identities were
independently authenticated. The diagnostic batch is9pass/3fail, not a repair.
Root approves the following bounded technical correction after this escalation.
The prior observations and synthetic-alias contamination remain attributed exactly.

Actual public alias no-ops produce two records, changed bytes and replaced owners.
Their public guards agree. Actual expanded WPS no-op produces zero records and
equal faithful paired publications, but later candidate closure changes source
and replaces its owner. Expanded WPS acceptance adds an observed mc binding while
rejection matches original. Related-header and public MCE failures are measured
at descendant inline/anchor closure before strict physical capture. Unknown inner
text metadata is actually omitted by the new guard. These are distinct obligations.

Establish physical source association before descendant drawing closure or graph
normalization. A source-faithful parse or equivalent bounded association phase must
retain actual qualified ancestry, raw/typed slots, boundaries and lexical templates.
Carry explicit correspondence into later analysis clones. No discarded model pointer,
byte/text/hash equality, normalized drawing value or first match identifies an owner.
Keep strict raw-role validation substantive, complete source/run-count validation,
field fragment correlation, actual control event ownership and selected MCE authority.
Use the same existing upper story-owner grammar in ordinary comparison and guards.
Do not relax descendant declaration equality or discover owners inside opaque payloads.

Retain original qualified lexical children for actual associated content slots whose
existing typed policy proves them unchanged. q:t/default t must not become generated
w:t merely because the surrounding original formatting is retained. Keep actual left
formatting shell and right content source authority separate. Fresh generated Word
markup still needs qualification under producer-w shadows. That mode must not bypass
restoration for retained source. Restore only positively proved extraction additions
using actual pre-extraction templates and equal destination scope. Preserve authored
namespace declarations, encoded/QName metadata, descendants, opaque events and order.
Real relocations close against the actual destination instead of removing required
bindings. Empty temporary IDs or equal text alone never justify copying an entire body.

Keep edited source-qualified canonical non-text drawing fields when replacing inner
content. Reparse to establish coherent typed/body/raw serialization authority, then
reconcile only analysis-added context against actual associated source and deliberate
generated dependencies. No global signature change, wholesale binding removal or broad
namespace filter is approved. A real binding used by opaque QName data remains
significant. Unprovable scope plumbing refuses at its physical location. Direct and
selected MCE views must use the same existing docPr and relationship normalization
with the real OPC owner, including nested drawing bodies. No raw-ID shortcut or main
part default is allowed. Preserve raw selected authority and typed view coherently,
project each physical selected body once, and retain fallback/unselected bytes exactly.
Direct WPS text_body remains authoritative once. Legacy fallback needs actual source
proof or qualified refusal. Unsupported inline shape modeling remains excluded.

Keep lexical package source separate from parse-only drawing-closed models throughout
comparison_input, tracked candidate staging and final publication. Restoring package
bytes while leaving polluted raw_xml is insufficient because a later save can leak it.
A global early return based only on zero IDs is also insufficient. Preserve the final
package equality guard and both independently computed reconstruction equalities.
A genuine whole-operation no-op retains original receiver, owners, parts, relationships
and content types. With an independent meaningful edit elsewhere, publish that edit
while keeping the unchanged textbox's exact lexical source authority. Existing handle
invalidation after genuine mutation stays unchanged. No-op/refusal identities remain.

Unknown nonformatting metadata on admitted modeled inner content remains significant.
ignore_formatting retains existing typed formatting exclusion and ordinary unsupported
property left bias. It does not discard unknown Word/producer attributes on w:t. Reuse
the existing lower protected_source_residual owning classification only on actual
source-proved nonformatting content slots with real parent grammar and bindings. A
private read-only r envelope may supply the existing content classifier's context.
It never becomes persisted XML or occurrence identity. Exclude unrelated property
subtrees. Do not import the complete protected residual into unheld formatting policy,
copy classifier tables, add a schema framework or introduce another public export.

Associate this metadata with existing canonical content boundaries, preserving ordered
multiplicity across supported split/coalesced runs. Physical ordinals/run partitions
are association evidence, not semantic equality payload. Unchanged positively mapped
metadata must survive text/formatting edits. Accepted success contains edited metadata
and rejected success original metadata under the supported source policy. A metadata
edit that cannot be represented or mapped with existing authority refuses atomically
with a physical location. It never becomes a false no-op. This is not blanket refusal
for all metadata-bearing source. Earlier protected-owner refusal stays exact. If the
existing bridge cannot express a slot without new classification/API scope, stop that
portion for design rather than inventing or waiving its policy.

Required actual successor evidence retains every old assertion and includes:

- Strict association controls and all ten source/closed owner contexts, then real header,
  footer, footnote, endnote and comment graph cases with equivalent IDs/targets, changed
  image/chart/link targets, main decoys, direct/selected views and save/reopen.
- Exact alias/default and expanded WPS no-ops with zero records, exact parts/source and
  held native owner identity. Text-changing controls independently accept New/original
  formatting and reject Old/original formatting. No extraction-only binding survives.
- Selected MCE one-owner projection, canonical graph identity, exact fallback/unselected
  custody, real namespace-dependency differences and unchanged original WPS controls.
- A formatting-only textbox beside an unrelated body edit, related-story edit or changed
  box. Mixed inner bodies retain unchanged source beside real text/table/control edits.
  These controls reject whole-body copying and zero-ID shortcuts.
- Unknown Word/producer metadata with equal/changed text and metadata, aliases/default
  shadows, repeated same-text owners and word/character split/coalescing. Assert unequal
  guard metadata, exact saved accepted/rejected inventories on supported success, or
  location-qualified atomic refusal when mapping/publication is unprovable. Keep ordinary
  property left-bias positives and protected counterparts.
- All five paired publishers, applicable granular modes, tables/controls/related recursion,
  nested boxes with strict progress, prior held spans, new-only IDs and all raw occurrence
  order/multiplicity/B-only cases. Keep unpaired italic, optionfalse, ignored stories,
  clean host insert/delete/move/geometry/opaque positives and injected late refusal.

Deliberately perturb actual candidates to prove each reconstruction guard independently
rejects. The guard never borrows an expected model or calls comparison. Synthetic helper
aliases without the public local_word_binding setup are not public RED provenance.
The exact lower lost-binding intermediates and candidate closure insertion site remain
source tracing until observed during this bounded repair and its discriminating checks.

Implementation uses existing files/private carriers with no new public surface, trait,
generic, module, crate or feature flag. HLD03/HLD10 remain the complete HLD work list.
After one coherent repair, batch actual local tests, affected lower/comparison/CLI and
scoped checks/lints, fresh authenticated opt2 binding proof, then a different independent
microscope. A continuing core invariant failure stops for independent reassessment.
No wheel/readiness, baseline update, final acceptance, commit or push follows merely
from this approval. All final integrated/native/installed/typing/hash/risk debt remains.
The new owner must acknowledge this full canonical plan identity before implementation.

## Destination and option diagnostics before another correction

The third coherent repair actually executes12pass/3fail library and72pass/2fail
regression selections. The root TOC passes. Its namespace failure remains real.
Root read the different investigator's reassessment.md in
fx272-namespace-reassessment-v1, SHA256
e641867a6222f23dec67506301a5515020b6aa833beba965b84d2f51e9247d5e,
and authenticated all227product/eleven tracked/progress/diff stop identities.
No fourth product repair is approved by this section.

Authorize a separate diagnostic-only phase in existing carriers. Production behavior,
namespace decisions, source association, guards and publication remain unchanged.
Observe actual source raw/scoped/parent frames, actual enclosing source shell, fresh
wrapper bindings, pre/post-wrapper XML, first changed binding, both resolved outputs
and save/reopen. Cover retained and fresh roles across all five publishers and direct
preserve/equal, insert/delete/replace and structural wrapping callers. Include canonical,
producer-w, default Word namespace, local producer-w shadows and actual relocation.
Do not use the source shell's scope as a stand-in for the future wrapper's scope.
No allocation or resolver call may be added solely to measure a frame if it changes
IDs, source or control flow. Production-site observations are cfg(test) only.

Root explicitly approves two test-harness corrections after independent evidence.
The mixed-header fixture must add its missing real section headerReference, verify
the namespace-qualified reference and relationship/part selection before comparison,
then observe the mixed operation. The former unselected-header outcome stays preserved
as setup contamination. Later operation/representation cases remain unrun until they
actually execute. Do not infer a selected-header product failure from the old fixture.

Replace only the overconstrained single-w:t predicate in the publisher regression.
The independent parser examined all14printed failures: ordered text is exactly the
expected New./Old., every text-bearing inner run is bold and none italic. Word and
character granularity legitimately split text nodes. The new observer must independently
parse actual qualified textbox content, assert exact ordered text per paragraph/content
boundary and original bold/no edited italic on every text-bearing run, while retaining
all existing host geometry, opaque source, counts and resolution checks. Keep Run-mode
exact source assertions where promised. Add wrong-text, missing-bold and edited-italic
negative observer controls. Do not call the production guard/projection to decide the
observer result. Retain the old predicate as a labelled historical diagnostic if useful,
but its single-node assumption is no longer the acceptance oracle. No other assertion
is waived. Old failing XML, source and receipts remain immutable.

Measure equal unchanged content metadata with ignore_whitespace false/true and unchanged,
whitespace-only and real nonwhitespace text changes. Exercise both metadata-bearing leaf
length and metadata-free whitespace before it. Record canonical vectors, raw/source
boundaries, IDs, no-op bytes and held owners, actual publication and both resolutions.
Metadata values, expanded names, scope and order do not become whitespace-normalized.
These are observations of the current candidate, not a preapproved cursor repair.

Use real simple and complex field/result source for ignore_fields false/true, with
unchanged/changed instruction or result, equal metadata and a later metadata-bearing
ordinary leaf. Record actual parsed Field/result/source-fragment ownership. Separately
measure unknown field metadata changes. Do not replace a physical field with a t merely
labelled as one. Existing modeled field exclusion and unknown metadata significance
must be reconciled explicitly after the observations, without inventing a new field API.

Measure actual nonformatting root/container attributes using the existing source roles:
r and p roots with Word/producer attributes, equal/changed metadata crossed with text,
a metadata-free neighbor, repeated equal text and formatting-induced split/coalescing.
Include inherited/shadowed bindings and ordinary compatible typed-value normalization.
Boundedly audit table/row/cell/control/hyperlink source roles against their existing
signatures and classifier. Do not claim all containers fail. A run-root record remains
a root fact, never a fake raw-child slot. Reuse only existing owning classification for
observations, with no new residual/schema/public bridge or whole protected-policy import.

Batch these discriminators and the corrected fixture/observer controls once with native
opt1, assertions, overflow checks and8MiB stacks. Keep all remaining genuine namespace
assertions active and per-case aborted/unrun distinctions explicit. Stop after evidence,
progress and frozen identities and return the sole writer/build lease. Root reconciles
the actual destination, option composition and metadata owner boundaries before another
product correction. New passing observers do not close the remaining namespace failure.
HLD03/HLD10, fresh bindings, independent microscope and every final union gate remain due.

## Actual destination and metadata correction after independent reconciliation

Root authenticated the diagnostic235file/30artifact freeze and read findings SHA256
78be6e4c6264e20b55b24485b52a7cb70320fae90d4290d970c1262777eb7b31.
A different nonimplementing investigator then reconciled the actual source and
measurements in fx272-destination-contract-reconciliation-v1/reconciliation.md,
SHA256359b1c98096d6ddcaa0766a8c027f148a4dc8b364af0acdeee036bb7011cff2e.
Root read that report in full and checked its concrete existing APIs. The diagnostic
results remain13/16library and74/75regression, not164semantic passes. No user/product
decision changes. Root approves this bounded technical correction and the proving
sequence below. An explicit sole-writer assignment is still required.

Use actual immediate publication scope. Add a private destination_parent parameter
using existing NamespaceBindings to preserve_run_source, publish_run_source,
paired_run_replacement and affected compared_run_xml callers. Derive retained-run
scope from its actual selected opening under that parent, and content scope from
the actual content source parent/root. Revision payloads use the planned fresh
ins/del parent scope through the same private qualification decision as
IdAllocator.generated. Do not duplicate local_word_binding predicates or allocate
dummy revisions to discover scope. Preserve allocation order/counts. Faithful
unwrapped publication proves the no-op first. A required outer pair republishes
against each actual wrapper before allocation/wrapping in the existing order.

Cover ordinary, granular, complex direct-splice, control-content and controlled
granular publishers, plus marked_control_content, marked_control_owned_rows and
marked_direct_runs_xml. Complex nested hyperlink/control shells and synthetic
inlineControl read wrappers are distinct from actual output parents. Same-parent
unchanged source stays exact. Actual relocation retains required producer/default/
local bindings and qualifies fresh Word children locally. Source/property/content
authorities remain separate. No global namespace closure, generation-mode bypass,
registry, source lookup by bytes/text or global matching/move change is approved.

Retain significant actual root, content and field metadata using the existing
protected_source_residual owning classifier. Its returned strings remain opaque.
For a root-only private read view, use that actual source owner's opening/closing
and inherited bindings with descendants omitted. Never invent a raw child or move
root attributes onto a fabricated leaf. For content/field facts, retain real parent
grammar and positively associated source roles, excluding only separately governed
actual Word property roles. Do not double-count a physical fact, copy classifiers,
import the complete protected policy or add another lower/public API.

This correction remains on the admitted paired textbox inner-source path and its
independent guard. Equal positively associated metadata beside supported text and
formatting changes must remain supported, including all8required Run positives.
Changed significant r/p/row/control/field facts without an existing tracking
publisher refuse atomically with a physical location. They must not false-no-op or
silently retain original metadata. Preserve actual hyperlink edit positives,
table/cell refusals, held-source stricter policy and ordinary unsupported-property
left bias. No generic unknown-attribute tracking capability is added.

Fix metadata coordinates and publication independently. Equality coordinates use
the existing option-normalized content/boundary policy, including ignored fields
and token barriers. Values, encoded spelling, expanded identity, scope, order and
multiplicity remain exact. Whitespace-only/empty metadata stays a boundary fact.
Physical partitions identify source occurrences but are not semantic equality keys.
Ambiguous collapsed correspondence refuses specifically, without discarding facts.
Extend the existing private typed selection with original ordinal and optional
Unicode-scalar subrange, captured during source-owned splitting and carried through
merge/publication. Validate ordered nonoverlapping complete role coverage beside
existing raw attribution. Whole unchanged leaves retain bytes. Proved pure-text
fragments keep their actual qualified shell/attributes and serialize only their
selected text with existing escaping/xml:space policy. Opaque mixed-event shapes
without mapping refuse. Never recover correspondence by text search or duplicate
a whole old leaf across fragments.

For fields retain ignore_fields exclusion first, then compare actual modeled field
pairs through existing compare_field_xml before ordinary ignore_formatting return.
Keep field alignment/writer refusal and ordinary property precedence. Unknown field
or result metadata is separately significant even when modeled fields are ignored.
Use existing Field source_replacement/source_owner_id and CT_P source-span methods
only after physical owner/span/coverage proof. field_source_run_fragments_at returns
None at field entries and is not a field-source provider. Owner IDs are local
association witnesses, never cross-document equality. If existing APIs cannot
prove the needed role/scope, stop that portion for design rather than invent an API.

Reconcile direct/selected mixed-body mc only from positively proved extraction
additions at exact owned roots, including when temporary IDs are nonempty. Carry
source/root correspondence through temporary resolution, splice admitted intervals,
reparse and keep persisted raw/typed selected authority coherent. Required generated
bindings, authored mc/QName dependencies, opaque shadows and fallback/unselected
bytes remain significant. No binding filter, drawing signature change, expected
model substitution or whole-original-body copy is approved.

Root separately authorizes the measured header count observer correction. Replace
only the main-only revisions() predicate with saved/reopened story_revisions()
inventory for actual Policy author and /word/header1.xml owner, expected inserted/
deleted kinds and distinct IDs. This fixture retains zero main revisions. Execute
both later author-resolution/save branches, require no remaining Policy inventory,
HeaderNew/HeaderOld respectively and unchanged textbox custody/formatting. Global
guard IDs are not a substitute for saved all-story inventory. Preserve all other
source, geometry, fallback, count and observer-negative assertions.

Before production mutation, batch actual stopped-source discriminators and missing
measurements in existing carriers: clean canonical simple/complex field instruction/
result baselines, exact INNER bridge/source/owner vectors, result metadata edits,
field multiplicity/repeated owners and later metadata leaf, saved selected-header
inventory/resolution after the approved observer correction, and the first actual
mc extraction introduction through nonempty-ID projection. Keep setup/errors,
counterfactual source tracing and semantic RED distinct. Required newly introduced
controls run against the stopped implementation before a successor. If observations
contradict this contract or existing APIs cannot represent source authority, stop
the affected implementation and return concrete evidence for reconciliation.

Then make one coherent correction. Retain all20paired and20direct/structural
namespace cases, six property controls, required Run positives,12whitespace,
32field,48root,6normalization,20container and12mixed controls. Add actual destination
shadow/inverse relocation, pure-text split/empty/ignored boundaries, independent
missing-metadata perturbation, real mc dependency and absent-headerReference setup
negatives. Both reconstruction guards, held byte/order checks, raw occurrence
authority, no-op/live-owner identity and late atomic refusal remain unchanged.

Only existing private carriers/files are authorized. HLD03/HLD10 remain the exact
work list, implemented before freeze. Existing risk routing still adds qualified XML
relocation/source preservation, same-input lifecycle/cache checks, fresh authenticated
opt2 binding proof and separate installed/native/final riders. After local actual
RED/GREEN, run affected scoped/lower/CLI/check/lint/fmt evidence and obtain a different
independent microscope. A continuing core failure stops for reassessment. No final
full gate, baseline, ledger, commit, push, readiness or release is implied.

## Source occurrence restoration and field owner correction

The coherent destination successor actually executes18/20library and30/30Issue425
regression, exit101. Root authenticated all235sources,156artifacts, both executed
binaries and frozen diff2183bc6b before further work. The report42a9acb9 and192
individual diagnostic outcomes remain historical evidence, not readiness. A different
nonimplementing investigator's metadata-reassessment-v1/report.md SHA256
bf9a388781b9a8e0f148bad2b0a01f12088d61bb041944afef971a51b600fcea and
contract-addendum.md SHA256
4edd551955593ddbdec6e9fa129a3134fb842fc3b2711aa21081ea85919f3066
were read in full. Root checked the actual private resolver and deletion paths and
authenticated3834pinned inputs, allowing only its disclosed progress update.

Root approves one bounded correction after this independent escalation. No public
authoring contract changes. The new writer must first capture the missing exact
diagnostics and genuine stopped-source REDs below. Contradictory evidence returns
affected work for reconciliation before production mutation. Existing failed outputs,
assertions and unrun-branch limitations remain preserved.

Restore source occurrences in publication, not by weakening guard equality. The
paired-host inner revisions are temporary and accepted before outer replacement.
Capture each relevant actual tracked pure-text leaf interval at emission, while
input role, physical source owner, original leaf ordinal, scalar subrange and
shell/property/namespace authorities still exist. Carry that association through
the actual selective resolver. After resolution, reconstruct only complete ordered
partitions of that same proved source occurrence, using actual resolved scalar
payloads and its actual selected source shell. Do this before splicing the projected
body. Never substitute an expected input body or infer origin from equal text,
attribute values, adjacency, counts, coordinates or a newly parsed ordinal.

Compatible source shell/properties, actual parent grammar and namespace authority
are required. Different authored occurrences, differing properties, opaque events,
controls, held revisions and other source barriers may not be merged. Explicitly
authored empty/whitespace leaves keep their own provenance even at one normalized
cursor. Keep the genuine one-occurrence mapped observer unchanged. The generated
whitespace partition is retired by restoring its source leaf in actual XML. The
independent guards then inspect that XML without expected-model or provenance access.
Keep all8required Run positives and existing supported clean text. Unproved source
correspondence retains the qualified physical refusal, without a blanket refusal
for metadata-bearing text.

Authorize one private additive crate-visible resolver entry in existing revision.rs:

    pub(crate) fn resolve_story_xml_with_leaf_ranges(
        source: &[u8], resolution: Resolution, scope: RevisionScope<'_>,
        leaves: &[Range<usize>],
    ) -> Result<(Vec<u8>, usize, Vec<Option<Range<usize>>>)>;

Only existing standard types appear. This is internal transport, not a new public
field/source API. Input indices correspond to positively associated actual tracked
leaf intervals. Validate bounds, whole actual qualified leaf ownership and disjoint
coverage before output. Each result entry is that same leaf's actual emitted output
interval, or None if removed. Keep/remove/copy/namespace-opening transformations
transport attribution through the SAME existing resolver operations and actual child
append offsets. Prefix/suffix and wrapper offsets are checked explicitly. Unsupported
transformations decline this narrow proof instead of manufacturing a range. Preserve
the existing resolve_story_xml signature, bytes/count behavior and selection policy.
Use concrete private carriers only where they own substantive checked composition.
No second resolver, generic callback framework, marker attributes, registry, public
model/type/export or new source file is authorized. Retire transport after assembly.

Close the fact-kind contract explicitly. Preserve the existing positive r-root
coverage normalization that accommodates physical run partitions. It is coverage
semantics, not a claim of strict authored positive run-root cardinality. Values,
expanded names, scope and multiplicity INSIDE its opaque fact vector remain exact.
Zero-width r-root facts remain ordered boundary facts and are not absorbed by equal
neighbors. Content leaves, logical fields, p/table/row/cell/control structural roots
and unproved shared-root inventories retain ordered occurrence multiplicity. Their
placement coverage is not permission to merge distinct owners. Remove blanket
cross-kind positive merging. Do not broaden it through zero-width facts. Required
source-proved leaf restoration occurs before either reconstruction guard, preserving
the earlier partition-independent run policy without a false exact-cardinality claim.

Fix simple-field classifier context privately. Existing typed Field parsing owns
ordinary direct Word result runs. After actual field/source/slot proof, collect the
fldSimple root facts with descendants omitted in its real parent context and collect
the actual associated result-run roots/content through the existing owning classifier.
Keep unknown direct children/events opaque and ordered in the remaining field view.
Do not double-count result roots or physical shared runs, invent a replacement t,
parse opaque identity strings, copy classification tables, reparent persisted source
or change the lower classifier globally. Its unchanged owner/value policy remains
authoritative. Stop an unproved shape for design. Modeled instruction/result values
must not leak into nonformatting facts, even when ignore_fields is true. Unknown
field/root/result metadata stays significant under both option values. Keep the
existing field comparison ordering before ordinary ignore_formatting.

The actual CT_P outside-fragment bridge and detached field payloads supply distinct
roles. Shared ownerID identifies a physical span, not a logical field. Restore actual
Before/Middle/After source leaf attributes through any detach/materialize/projection
step using proved slots and coverage. Classify a physical root only once. Keep the
bounded unproved significant shared-r-root refusal and test it separately from rPr
formatting. When ignored modeled field changes leave every other proved meaning
unchanged, preserve the exact genuine no-op source/owners. A zero temporary-ID set
alone is insufficient. Do not blanket-refuse clean simple/complex/shared fields.

Bound deletion instruction conversion to actual owned direct Word instruction
occurrences. Preserve the documented field-spanning-paragraph instruction deletion
behavior, qualified aliases and namespace authority. The current recursive lexical
instrText renaming reaches inside the opaque pict host and violates its raw custody.
Do not rename nested textbox/opaque or foreign lookalike payload. Existing explicit
typed Field publishers retain their own authority. Keep strict raw occurrence Some(0)
validation exact, with no derived-raw exemption or broad refusal used to hide this fault.

Before production mutation, acquire one bounded existing-carrier diagnostic/RED batch:

- Immediately before the existing raw mismatch, record actual selected source and
  generated child, raw ordinal, caller/location, role and source/destination scope.
  Use clean complex result false and shared instruction true, plus a simple control.
  The instrText/delInstrText causal chain remains source tracing until that executes.
- Record actual fixture ZIP/part inputs at construction, without inducing a separate
  pre-comparison authoring rewrite. Print actual accepted/reopened parts before the
  mapped assertion and independently exercise rejection/save/reopen on a diagnostic
  clone. Do not skip, soften or reorder away the genuine semantic assertion. Label
  these actual artifacts separately from earlier assembled diagnostic envelopes.
- Add semantic assertions for the promised clean simple/complex/shared modeled
  instruction/result paths and ignored-field no-op behavior. Diagnostic completion
  alone cannot qualify them. Retain actual field/result/source role observations and
  unknown metadata atomic refusals, with existing unsupported/malformed boundaries.
- Add adjacent identical authored leaves, an authored empty fact between equal
  nonempty facts, two authored empty facts at one cursor, metadata-free whitespace
  between different owners, repeated text and differing attribute values. Preserve
  ordered multiplicity in both resolution directions/reopen and show independent
  removal/value/scope perturbations fail both guards. Preserve pure-text Run/Word/
  Character ignored-whitespace REDs, exact mapped count and authored empty positives.

Then make one coherent correction and run affected actual local successor cases.
Preserve all namespace/destination,20paired/20direct,12mixed/header, physical graph,
held-byte/order, no-op/live-owner, late refusal and ordinary property left-bias controls.
Private resolver changes additionally require existing selective resolution/property/
paragraph-marker/namespace tests and exact unmapped bytes/count compatibility. New
range-map positives/negatives must cover kept/removed leaves, transformed openings,
aliases/default shadows, wrapper offsets, invalid/overlapping/nonleaf input ranges
and source barriers. Only existing comparison.rs, revision.rs and existing diagnostic/
test carriers may change, with no new public/lower bridge. HLD03/HLD10 remain exact.

After actual local GREEN, complete affected scoped/lower/CLI/check/lint/fmt and fresh
authenticated opt2 binding proof, then obtain a different independent microscope.
A continuing core failure stops again for independent reassessment. Final integrated,
installed/native/typing/hash/pixel/archive/README/hosted/risk evidence remains due.
No baseline, ledger, commit, push, publication or readiness follows from approval.

## Fifth correction reassessment and diagnostic-only next phase

The coherent successor under99ad38a1 is not qualified. Its actual execution is
23library passes/3failures and30Issue425regression passes. Root independently
authenticated235sources,2760artifacts, both binaries and full diffbe18265a. Report
4517060a and independent causal report9b50abeb retain all outcomes and gaps.
The implementing owner returned every writer/build lease and stopped. Historical
compile interruption and linker failure remain distinct from the executed result.

The existing paired textbox policy already requires edited content with original
formatting under ignore_formatting=true. Within a positively source-proved existing
leaf pair, inserted scalar portions must therefore use that original policy shell.
Genuinely unpaired inserted owners retain edited formatting. Optionfalse and ordinary
outer comparison policies remain unchanged. This clarifies existing policy rather
than approving a global formatting change or inferring correspondence from text.

Independent technical addendum6fc485a7 proposes separate physical emitted provenance,
paired logical occurrence, policy shell and ignored scalar transformations, using
actual owner/slot/alignment witnesses before fragmentation. It also identifies the
existing whole-field publisher as the narrower simple-result route. It is a proposal
for the next coherent correction, not permission to mutate production now. No global
fldSimple scanner/listing/held/resolve grammar expansion is approved.

Root approves only the following bounded diagnostic/test phase in the existing
comparison.rs and existing diagnostic/test carriers. Production behavior, revision.rs,
HLD03/HLD10 and every existing guard/assertion stay unchanged during this phase.

- Print the actual tracked/resolved XML, wrapper offsets, supplied physical leaf
  intervals and transported ranges, original/edited owner/slot/scalar witnesses,
  pre-group alignment/actions, post-group actions and actual policy properties.
  At every restoration refusal retain the exact group and root/scope/property/
  terminal operands. Measure both failing Character cases rather than guessing.
- At the simple-result held/new-ID check print actual held spans/bytes, tracked and
  remaining spans/IDs, temporary IDs and broad word-ID inventory. Include the
  successful changed-instruction whole-field control for comparison. Keep the
  combined guard and exact original public fixtures active.
- Add inline source discriminators for inverse ignored whitespace, unequal text
  lengths, repeated same-text authored leaves and authored empty boundaries in the
  existing test module. Record each actual outcome separately. No fixture file or
  public API is added. Do not invent a source pair from equal values or expected text.
- Correct the new resolver positive to genuine modeled Word story ancestry, asserting
  actual revision selection/count first. Execute Accept/Reject, aliases/default
  shadows, wrapper offsets and invalid-range cases. Retain the original opaque-root
  setup as a negative with zero selection and unchanged bytes/count/range transport.
  Preserve the historical invalid-positive failure instead of treating it as RED.

Run one bounded actual opt1/assertion/overflow/8MiB semantic diagnostic batch after
source capture, with exact configuration,235before/after guards and executed binary
identity. No stale binary after a failed build. Return the actual missing operands,
source-pair admission limits and all outcomes to root before production correction.
Only explicit separate assignment grants the sole writer/build lease. No sixth
product repair, scoped readiness, binding build, microscope, baseline, ledger,
commit, push or release is approved by this diagnostic phase. All final debt remains.

## Measured paired-source alignment and whole-field publication correction

The diagnostic-only phase completed at report79ab6cf0 and seal08e5ed63. Actual
library25/27 and regression30/30 executed with235unchanged inputs. Root read the
full report and independently verified235sources,1638artifacts, both binaries,
diff6d23d79c and production-prefix identity after reviewed cfg(test) exclusions.
Root authentication is5deba2ce. Corrected resolver controls pass. The exact remaining
property mismatch, crossed source owners, scalar gaps and surviving simple-field
IDs are measured. These are not hypothetical gaps or a successful product repair.

Root read independent measured proposalecaf41f816ec83b86535306e2e1bfc666d8473349cf58d892d6fdfe376c992a6.
The prior writer has returned every lease. Root approves one coherent correction
under the following bounded contract, with implementation assigned separately.

Inside an admitted genuinely paired textbox projection, explicitly select the
existing deterministic ordered whole-run alignment before granular_text or unit
construction. Resolve its compared indices to actual live source owners and captured
qualified slots. Inside each selected run pair, establish ordered content-leaf
correspondence from actual CT_R.content, typed lexical spans and raw boundary roles.
Signatures select alignment but are never identity keys. Whole-run pairing alone,
equal text/attributes/bytes, matching counts or independently parsed ordinal equality
do not prove a leaf pair. Validate ordered role coverage and injective significant
leaf correspondence. Preserve authored empty/whitespace leaves as distinct slots.

Use that selected relation to segment private projection scalar alignment BEFORE
grouping. Align only within an admitted leaf pair, carry unmatched leaves as actual
insertion/deletion owners, and retain pair/leaf/raw/control cuts through grouping and
coalescing. This reaches both ordinary and controlled granular routes and Run when
ignore_whitespace causes fragmentation, as well as Word/Character. Global align,
ordinary outer matching/moves, optionfalse, ignored-story policy and host granularity
remain unchanged. Nested boxes use their own checked descendant source context.

Keep four separate substantive facts in existing private carriers: actual emitted
physical input and scalar interval, selected logical leaf pair, original formatting
shell for that existing pair, and complete actual policy action coverage. Preserve/
Equal that publish left record truthful left provenance. Insert records actual right
provenance. Within the proved pair all scalar pieces use the approved original policy
formatting before resolution. Genuinely unpaired inserted owners retain edited
formatting. Nonformatting root/leaf metadata stays significant. Do not copy a whole
original shell over a changed unknown fact or reinterpret field-owned policy as
ordinary plain text policy. No global registry or new public/lower API is added.

Restore exactly one occurrence using actual resolved emitted payload and source-
proved shell/context. Replace physical edited full-length coverage with stronger
complete transformation accounting: every source interval has its actual Preserve,
Equal, Replace, Insert, Delete or permitted ignored Drop disposition. Every emitted
piece has an actual range/payload or legitimate resolver removal. A Drop has no
fabricated XML leaf. Reject unexplained gaps, duplication, reordering or crossed
owners. Never fill ignored spaces from expected edited text. Compatible properties,
scope, facts, parent grammar and barriers remain required before combining pieces.
Independent guards consume assembled XML/graphs only, without this private relation.

For same-instruction simple-field result changes, use the existing whole-field
deletion/insertion placement. Deletion contains the actual original field. Insertion
uses the SAME original field shell plus its actual edited result interval, retaining
the current field-owned property/source policy. Do not blindly insert all edited_xml
or put temporary revisions under fldSimple. Keep instruction spelling, known flags/
defaults, unknown root/result metadata, opaque children/events and namespace authority
under their existing policies. ignore_fields exact no-op remains first. Complex/shared
behavior remains unchanged. No revision scanner/listing/held/resolve grammar change
or second resolver is approved. A concrete field-policy contradiction returns to root.

Require actual GREEN for all12clean field and9fragment semantic rows, all8required
Run positives and the unchanged mapped one-occurrence lifecycle. Turn the8new inverse/
unequal/repeated/empty observations into substantive semantic checks, adding Word
coverage and source-pair/action/range assertions. Include different repeated-leaf
attributes, two authored empty boundaries and independent corrupt/missing proof
negatives. Boundaries may not cross metadata-free neighbors, opaque events, controls,
fields or held revisions. Retain unpaired italic, optionfalse and mixed unrelated edits.

Field discriminators include bold1toitalic2 under both formatting options, same
effective instruction with different lexical spelling, known flag/default controls,
unchanged opaque/result metadata and local bindings, repeated result siblings, and
unsupported metadata value/order/binding edits with atomic refusal. Assert actual
accepted/rejected property/root/source inventories after save/reopen, not only text.
Both recognized outer temporary IDs must disappear. Preserve every passing complex/
shared, namespace/mixed/header, graph, raw/multiplicity, held-byte/order, no-op/live-
owner, late-refusal and newly green resolver control.

The work list remains existing comparison.rs and existing test/diagnostic carriers,
with revision.rs behavior unchanged and HLD03/HLD10 updated before freeze. Reach all
five paired publishers and both granular composition/capture routes. No new tracked
file/module/crate, public model, generic framework, marker attribute or expected-model
substitution is approved. Unprovable imported shapes need a concrete source/slot
contradiction, without dropping required positives or blanket metadata refusal.

After actual local GREEN run affected scoped/lower/CLI/check/lint/fmt, fresh opt2
binding evidence and a different independent microscope. A continuing core failure
stops this correction for independent reassessment. Every final union/native/install/
typing/hash/pixel/archive/README/hosted/risk gate remains due. No readiness, baseline,
ledger, commit, push, merge, tag or publication follows from this approval.

## Two-input action closure and mode-specific field evidence

The source-pair successor at69b084ea remains unqualified:27library passes/2failures
and30regression passes, with235unchanged execution inputs. Original12clean fields
and9fragment semantic rows pass. New inverse Character fails before later added
rows, and the new field test aborts its first false-mode spelling assertion. Root
authenticated235current sources,1910artifacts, both binaries and diff8ab8b4f6 at
079b8ae2. The writer returned all leases and stopped. Every original outcome remains.

Root read independent causal report2da50c6526a4cb35457b9ac4c9317b8c2df6ce3ad33533671d73545a6f346a2a.
Its exact first-guard derivation uses complete actual groups and deterministic source
order, explicitly distinguished from a runtime-printed operand. Physical original
coverage0to3 is complete. Edited coverage0to1 is complete independently. The actual
three emitted spaces are Preserve, not Drop. The incomplete proof splits original-
only Preserve pieces away from their already selected logical pair.

Root approves one bounded completion of the existing two-input transformation rule.
Each selected logical occurrence owns separate original and edited source domains.
Every action retains optional original/edited interval witnesses independently, plus
actual emitted/removal disposition. Look up pair membership from either actual source
role/owner/slot, including original-only Preserve and edited-only Drop. Missing a
simultaneous counterpart unit must not erase the preselected relation. Never relabel
original intervals as invented edited coordinates or mark emitted Preserve as Drop.

Validate ordered complete coverage for EACH input domain separately, with distinct
missing/repeated-original and missing/repeated-edited negative controls. Validate
physical emitted ranges/payloads independently. Inverse3to1 must assemble one proved
occurrence from three actual original spaces, with original coverage0to3 and edited
coverage0to1. Do not fill expected text, skip gaps, merge authored owners or change
ignored-whitespace policy. Both ordinary and controlled capture/composition routes
must retain these witnesses. Existing scalar/barrier/scope/property/raw guards remain.

The field failure is an invalid source-spelling oracle for the actual executed mode.
ignore_formatting=false bypasses inner projection and accepts the ordinary edited raw
host, so accepted spelling is edited and rejected spelling original. Do not change
that outer dispatch. The lower parser trims instruction payload for effective field
identity, so raw source inequality is not an established cause. True-mode inner field
publication has not yet executed in this added test.

Correct the test to distinguish those callable boundaries while retaining exact
historical invalid-oracle evidence. For admitted inner comparison assert actual
effective-instruction equality and branch admission, then original field shell with
actual edited result/property authority after accept, and original after reject.
Retain exact root/result metadata and lexical source assertions per mode. Include
the existing direct typed-field caller where needed to test field-owned policy under
both options. Do not trim field source globally, expand grammar or weaken assertions.

Execute added rows and both resolution branches independently or aggregate failures
so an early assertion cannot hide later inverse/unequal/repeated/empty or true-mode
observations. Preserve all12clean/9fragment and30regression gains, all8Run positives,
both width directions in all granularities, authored empty/multiplicity controls,
unpaired italic, optionfalse and both independent reconstruction guards.

Only existing comparison/test carriers and HLD03/HLD10 may change under a separate
sole-writer assignment. revision.rs/F273/CURRENT/rootTOC stay frozen. No new tracked
file, public/lower API, generic framework or global matching/default policy is added.
One actual opt1/assertion/overflow/8MiB coherent semantic batch follows meaningful
scoped checks. A continuing core failure stops for independent reassessment. Only
actual local GREEN permits remaining scoped/lower/CLI/lint/fmt/fresh opt2 binding
proof and a different independent microscope. All final union/native/install/risk/
typing/hash/pixel/archive/README/hosted gates and release approvals remain due.

## Token-boundary publication and terminal-child reconstruction

The stopped two-input successor is actual RED, not acceptance:28library passes,
1aggregated failure and30regression passes. Root independently authenticated all
235current sources,1949artifacts,both actual binaries and diffa779a7e5 in temporary
receipt26023abc. All original12clean/9fragment rows, inverse3modes and4field
resolution branches remain successful. Seven new comparisons refuse atomically.

Root read both independent full causal reports: token boundaries5e2eab9b and
authored empty custody in report92f227c9/seal7df8efe9. Root authenticated both reports and their
241/2192source-and-artifact entries with zero mismatch in receipt10ed9fe0.
These are independent causal investigations, not clean microscopes or GREEN.
The previous writer stopped and returned every canonical/source/build lease.

Root approves one bounded reconciliation of publication and reconstruction.
The independent guard's existing Run/Word Unit boundaries remain authoritative.
Complete source interval coverage alone does not justify word joining when
fragments are restored into one leaf. In the private source-proved paired text
projection, publish a necessary actual source-backed ignored separator BEFORE
assembly when omitting it would join the surviving granularity-specific Units.
Use the captured source interval, actual selected payload, existing source-aware
publisher and original paired property policy. Record truthful emitted Insert
for a right-only separator that is actually published. Never retain a Drop label
for emitted text, invent original coordinates, inject expected edited text during
restoration or change global matching/token normalization/metadata coordinates.
Matched original whitespace and original-only Preserve retain existing width
policy. Leading/trailing, whitespace-only and width-only ignored changes do not
become blanket insertions. Character retains its scalar equality policy. Both
ordinary and controlled granular capture/composition routes must compose under
this same rule. Explain its precise predicate and source-pair confinement before
building, and return a contradiction to root rather than weaken the guard.

For the measured authored-empty terminal-suffix shape, preserve the actual
terminal run opening, rPr, effective scope, raw suffix and closing tag. Replace
only its selected child budget with the single reconstructed source-associated
leaf containing actual admitted payloads. Retire preceding fragment runs ONLY
when a complete event scan proves all their nonproperty children belong to the
same proved group and the existing root/scope/property/adjacency proofs agree.
Multiple member children in one preceding run need complete coverage. Terminal
member children may be replaced or removed only through ordered disjoint child
budgets with every independent suffix byte untouched. Preserve source raw ordinal,
preceding typed boundary and ordered multiplicity. A new empty-only carrier is
not equivalent. Independent events between member children, before the selected
terminal child, in a retired run, or under another parent still block this
reconstruction. Different shells/properties/scope, held/field/control ancestry,
unaccounted bytes and independent neighboring groups remain barriers. Do not
turn opaque empties into typed text or infer their identity from equal bytes.

Keep the actual seven failed source cases and all previous assertions/history.
Execute all12new shape/granularity comparisons with independent or aggregated
outcomes, both public resolutions and save/reopen where admitted. Assert ordered
content Units and source metadata occurrences, actual emitted separator intervals,
both domain ledgers, final payload identity and untouched suffix bytes/ownership.
Retain deliberately joined-token refusal under Run/Word. Cover leading/trailing
ignored units, punctuation boundaries, width-only/inverse controls, same-text
separate leaves, different attributes and metadata-free neighbors. Add two
distinct and two identical authored empties at one boundary and genuine internal
barrier controls. Missing, changed, reordered or duplicate empty facts must still
fail independent guards. If a new discriminator was not executed on the prior
source, label source tracing as counterfactual and retain that limitation.

All original12clean/9fragment/30regression gains, all8Run positives, four field
resolution branches, both independent reconstruction guards, action-domain and
physical-range negatives, raw/namespace/multiplicity/atomic/no-op/held controls
stay required. Source barriers are narrowed by proved edit budgets, not waived.
No expected-body substitution, second resolver, scanner grammar change, new
public/lower API, generic framework or tracked file/module/crate is approved.

Only existing comparison/test carriers and HLD03/HLD10 may change under a new
explicit sole-writer assignment. revision.rs/F273/CURRENT/rootTOC stay frozen.
Update HLD current intent before freezing. After meaningful scoped checks, run
one native opt1/assertion/overflow/8MiB coherent semantic batch. A continuing core
failure stops source and builds for independent reassessment. Only actual local
GREEN permits remaining scoped/lower/CLI/lint/fmt/fresh opt2 binding evidence and
a different independent microscope. Final union/native/install/typing/hash/
pixel/archive/README/hosted gates and release approvals remain due. No baseline,
ledger, commit, push, merge, tag, publication or readiness follows this amendment.

## Ordinary field schema and lifecycle diagnostic boundary

The focused separator/terminal successor is actual32library/30Issue425regression
GREEN, with88independent resolution observations and27additional library passes.
The approved wider union is actual83of86regressions passing. The three failures
are the ordinary simple-field result-owner matrix and two full-story record
observers. Root authenticated235current inputs,2358artifacts,both actual/copied
binaries and diff0206dafd in ee8de18b. All source/build leases returned to root.
No lower/CLI/lint/binding/final gate followed that failed prerequisite.

Root read independent report0c697acc and authenticated its nine artifacts. It
establishes an ordinary placement regression from the earlier unconditional
simple-result publisher, separate from the latest separator correction. It also
establishes that the two first-field-prefix observers cannot prove insertion is
missing from the complete part. The retained official schema permits result
revisions inside fldSimple but forbids a direct fldSimple under run del/ins.
Both same-instruction and changed-instruction outer field routes conflict with
that parent grammar. Frozen shared revision ancestry excludes fldSimple, and
owning lower result projections also skip its revision children. Restoring the
old placement alone is not accept/reject closure. Actual strict field guards
retain instruction, form and cached result. Clean simple-to-complex conversion
would require an explicit form/source policy decision, not a hidden guard change.

Earlier whole-field publication approval is suspended wherever it conflicts
with this schema/lifecycle evidence. Root approves DIAGNOSTIC-ONLY acquisition
of the missing current operands before another product repair or representation
decision. No conversion, scanner extension, accepted-body shortcut, assertion
waiver, field refusal expansion or new public policy is approved in this phase.
All existing semantic/record assertions stay exact. A later product amendment
must reconcile ordinary/result/instruction/form/source/schema contracts and any
material public choice before implementation.

One existing cfg(test) carrier may add bounded diagnostic cases. Production
comparison prefix, revision behavior, lower parser behavior, document/rootTOC,
F273 and CURRENT remain byte-exact. No source file/module/crate/API is added.
Retain exact original/edited ZIPs and complete main/header parts for ordinary
DATE result-only and the existing two-field MERGEFIELD result/instruction fixture.
Capture default and ignore_formatting modes, actual compared parts, qualified
revision/field parents, typed/all-story listings, source/field form/instruction/
cache/flags/root/result/property/namespace facts and actual public result/error.
Do not use a first-field prefix as complete ownership evidence.

On independent copies of actual current candidates record accept/reject via
all, author, ID and applicable date APIs, with real counts, modeled/broad ID sets,
complete resulting parts and save/reopen field/source inventories. Capture actual
independent accepted/edited and rejected/original operands where the existing
owning guard is callable. Failed/refused/unreached branches stay explicit. Beside
these, construct a separately labelled schema-shaped nested-result revision
INPUT from actual declared source slices. Its current public listing/resolution/
read observations are generated-input diagnostics, not current comparison output
or observed historical production. Do not execute an alternate product patch.

Retain the actual paired simple-result and instruction controls beside ordinary
cases using their existing qualified host/body/run/field source admission. Capture
the real leaf_capture lifetime, field payload, temporary IDs and final body. The
existing focused88 observations remain separately credited. Mode alone supplies
no source pairing proof. Direct helper observations must name their context.

After scoped compile checks, execute one native opt1/assertion/overflow/8MiB
coherent diagnostic batch with235before/after sources and actual Running binary.
Its exit or case count is diagnostic execution, not semantic GREEN or readiness.
No follow-on product patch or next build follows its findings without root review.
Use the existing pinned schema/MCE tool on actual retained complete Word parts,
record original as well as candidate/resolved outcomes and reduction limits.
Opaque-extension input failures do not establish a new placement violation, and
must not be hidden by an invented reduced expected file. No general XML framework
or altered schema is permitted.

One sole writer receives only this cfg(test)/progress/Cargo diagnostic lease.
HLD intent and existing source behavior stay frozen during acquisition. Preserve
all prior failures, green operands, original12clean/9fragment/30regression cases,
88rows,27library controls and86scoped members. Stop/seal/return all leases after
the diagnostic packet. A different independent reassessment and explicit contract
closure follow before any product repair. Every scoped/lower/CLI/lint/binding/
microscope/final union/native/install/typing/hash/pixel/archive/README/hosted/risk
obligation remains due. No baseline, ledger, commit, push or release is authorized.

## Ordinary field repair contract after Q-S96-02

This section is ROOT-APPROVED for implementation on2026-10-29 after independent
contract closure review b171a43a at reviewed plan c54d9014 found zero remaining
concrete contradictions. It supersedes schema-conflicting whole-simple-field
publication instructions and the diagnostic-only grammar freeze. Implementation
still requires a separate sole-writer/build assignment. All historical receipts
and assertions remain at their recorded source identities. This approval supplies
no runtime, native or final acceptance and changes no product, test or HLD source.

### Authority, actual RED and representation decision

The designer resumed branch sprint/s96 at Base6506fa66 with the existing dirty
work intact. BACKLOG843 and CURRENT_SPRINT22 still say in-progress. F-X229,
F-X230 and F-X245 are done. Incoming plan18e45448 records the user's exact
Q-S96-02 answer, "Complex fields (recommended)". This permits a persistent
complex comparison representation for clean ordinary instruction edits. Root
also explicitly applies the existing HLD03 complete-owner FORM-change intent
to that same clean standard-complex carrier, with an actual form-delta witness.
The form-only application is ROOT'S technical decision, not another user quote.
Unchanged, ignored and result-only fields cannot borrow either exception.
Typed field construction, FieldForm and Run.add_field remain unchanged.

The independent reports consumed in full are ordinary-field-design-reassessment-v2
reportd6270484, complex-candidate-v1 addendumf49b933a, ordinary-field-contract-inventory-v1
inventory6f3ac297, field-form-contract-inventory-v1 addendum42dccc10 and
result-owner-consumer-map-v1 addendumd38f2579, under the existing temporary evidence
root. Root's73d5854e authentication binds the actual16case/128branch/312ZIP operands.
The diagnostic exit0 is not semantic GREEN. The actual schema result is270valid
and42invalid packages: FOUR ordinary current comparison candidates, NINETEEN
foreign-input variants and NINETEEN opaque-input variants. The first four have
fldSimple directly under run del/ins. The other38 retain invalid input ancestry.
Clean generated nested-result INPUTs validate but resolve0 selected IDs40/41,
retain those IDs after reopen, and read an empty lower cache. They are generated
inputs, never current comparison output. Default guard equality on them is a
measured blind spot. Native report6413b310's14footer simple-to-complex changes
occur during Word open/F9/save only. They prove neither comparison nor lifecycle.

Relevant source anchors are `crates/rdocx/src/comparison.rs:13395` and
`:13446/:13459/:13479`, `crates/rdocx-oxml/src/text.rs:905` and
`:1291/:4533/:4569`, `crates/rdocx/src/revision.rs:1665`,
`crates/rdocx-oxml/src/revision.rs:542`, and
`crates/rdocx-layout/src/engine.rs:2372` and `:11030`. Numeric citations in
this section refer to those frozen files: comparison1b36562b, production prefix
adc603df, revision0a2cfa45 and lower text84b1bc93. The implementation must retain
actual old-source RED, diagnostic operands and any later unexecuted branches.

### Dispatch and physical XML contract

The following order is mandatory in compare_field_xml_impl and every equivalent
publisher, before any Field PartialEq shortcut:

| Operation witness | Physical publication and source authority | Form and guard disposition |
|---|---|---|
| ignore_fields, or source-proved unchanged field | Existing original exact source passthrough. Significant protected/source facts are checked before passthrough | No conversion, no new revisions, existing live-owner/no-op contract |
| Actual qualified paired-host admission | Existing outer host owns the change. Inside it publish policy-resolved actual field payload. Same-instruction uses original field shell and actual edited result/property authority. Instruction/form changes use the existing selected-side complete field policy | Keep simple sides simple. Actual source pair, physical part/path/slot and checked inner interval are required. leaf_capture alone, an option or equal text is insufficient |
| Ordinary same effective instruction AND same simple form, changed result | Retain actual original fldSimple root QName, lexical instruction, flags, namespace scope, unknown root facts and schema-ordered fldData. Publish nested result revisions AFTER fldData at actual result owners | Simple remains simple in both directions. Accept actual new result, reject actual old result. Field-owned property policy precedes plain ignore_formatting |
| Ordinary same effective instruction AND same complex form, changed result | Existing complete complex source span and marker shells stay authoritative. Track actual result owners only | Keep complex/shared source and all outside-run fragments. No whole-span rebuild |
| Ordinary actual instruction delta OR actual simple/complex form delta | Publish complete old/new STANDARD COMPLEX spans under ordinary del/ins. Convert only source-proved clean simple sides. Existing complex side keeps its actual source/span | Persistent complex after both resolutions. Only the exact conversion witness permits that side's source-simple versus actual-complex mismatch |

Field PartialEq at text.rs2293 omits form. Therefore equality cannot precede
the form-delta decision. Keep its global semantics unchanged. A comparison-local
no-op needs equal effective instruction, equal actual form, existing value/cache
policy and independently checked source facts. Same-effective lexical differences
do not create an instruction witness. A real form-only edit still allocates and
publishes the complete old/new pair even if normalization yields equal complex
payloads. It is not suppressed as a no-op. Public all/author/date resolves2 new
revisions, each new ID resolves1. This explicitly records the form-delta review
while accepting the chosen persistent complex representation. It does not promise
that native acceptance restores a simple side.

Do not wrap arbitrary EG_PContent in run revisions. The retained Transitional
wml.xsd CT_SimpleField1231 has fldData then EG_PContent2228. CT_RunTrackChange995
uses EG_ContentRunContent2082, which excludes direct fldSimple AND hyperlink.
Apply this at compare_field_xml_impl, tracked_field_result, complete-field
instruction/form replacement, simple-to-complex result carriage and the five
paired publishers, including direct complex-paragraph publication. For a matched
nested field or hyperlink, preserve its real enclosing source shell and recurse
through the existing qualified field/hyperlink policy to revise permitted run
children. Existing controls use their actual supported content model. Do not add
an invented wrapper, strip authored controls or recursively relabel unknown XML.
A whole-field outer run revision admits only an actually proved legal run-content
payload, including each transparent wrapper's real inherited grammar. A nested
simple field or hyperlink cannot pass by name alone or by being one level deeper.
Unchanged raw descendants remain exact. No blanket exclusion of metadata-bearing
result runs is allowed. An unpaired incompatible nested owner without an existing
legal publication route is an explicit unproved shape, not permission to emit
invalid XML or change a required positive to refusal.

### Clean simple-to-complex source mapping

A conversion witness is created only after ordinary instruction/form dispatch and
actual source correspondence. Its concrete private data belongs in existing
comparison carriers: physical OPC part/story, source input role, owner path and
ordinal, logical field/span slot, exact source bounds/bindings, reason Instruction
or Form with the actual before/after facts, generated span bounds and allocated
revision IDs. Source IDs, addresses, signatures, text or bytes alone are not
logical occurrence identity. Repeated and shared physical field owners require
actual ordered logical slots. There is no persisted marker or global registry.

For EACH simple side, inventory every root attribute and child event. Admit known
qualified instr, optional dirty/fldLock and at most one correctly ordered fldData,
plus namespace declarations whose dependencies can be retained without rebinding
opaque descendants. Map flags onto the generated begin fldChar with their actual
presence, lexical on/off spelling and expanded identity, rather than regenerating
from Option<bool>. Carry the exact fldData subtree as that begin fldChar's one
choice child. Do not invent ffData or numberingChange. Emit one instruction run
with instrText and xml:space=preserve containing the actual decoded XML-normalized
instr attribute value. Preserve significant whitespace and effective instruction.
Attribute quoting/entity spelling cannot survive attribute-to-text conversion,
which is the explicit encoding exception. Do not use trimmed instruction.raw or
Field::from_instruction. Emit separate even for an empty result, carry actual
ordered result source, then end. Marker/instruction runs acquire no invented rPr.
Result run/root/rPr/text/opaque facts retain their existing owners and formatting.
Close only actual namespace dependencies at the true destination, never overwrite
producer w/default/local shadows. A simple empty spelling is allowed to expand.

Deletion renames only source-proved qualified owned t/instrText to delText/
delInstrText. Use the existing qualified direct-run machinery, not the lexical
whole-subtree deleted_text_xml replacement. Stop at nested fields, drawings,
textboxes, foreign lookalikes and opaque custody boundaries, each of which retains
its own publication authority. Already-complex sides keep complete source and
outside-span materialization. Mixed forms work in BOTH directions. Original simple
and edited complex means reject also stays complex. The reverse means accept stays
complex. Form-only follows the same explicit witness, never global kind equality.

Unknown simple-root attributes, unqualified producer attributes, root xml/MCE
semantics without a proved destination, malformed known flags, duplicate expanded
attributes, multiple/late fldData, opaque direct events or ambiguous physical bounds
are not clean conversion witnesses. Unknown content INSIDE retained result runs or
fldData is not automatically unclean because it retains its owner. Unchanged and
ignored opaque-root fields remain exact. Result-only retains their real simple
root. The inventory identifies no required ordinary opaque-root instruction-edit
positive. For a newly encountered unmappable ordinary edit, return the exact
unconsumed source fact/owner for design disposition before claiming that class
supported. Do not introduce an early blanket typed-field refusal, alter public
field updates, or invent a generic conversion fallback. Existing required clean,
paired-host and metadata-preservation positives cannot be narrowed by this rule.

### One canonical ordered COW result owner

Root adopted the explicit PENDING low-level mutation contract after reading the
independent borrowing reports f3a41fda and66a98a1b. This is a NEW technical state
contract for the newly retained ordered result, not a redefinition of checked
RevisionRef as a historical snapshot. The previous amendment1d451e78 remains
immutable in fx272-field-contract-design-v1/plan.md and its seal. Its mutable
accepted Vec plus immutable wrapper-listing proposal is superseded in place here.
That scheme would expose stale insertion_paragraph content even with correct IDs.

Field keeps ONE current canonical ordered result CT_P behind concrete Arc<CT_P>
COW storage. Immutable original raw XML, source bounds/bindings and original field
facts remain the provenance baseline. An original parsed CT_P baseline may share
that initial allocation, but must never supply current borrowed revision wrappers
after mutation. No independently mutable flattened accepted Vec or cached_fields
clone mirrors this result. Existing no-retained-owner Field cases keep their
present cache implementation and all existing signatures. Cloning a Field shares
current ordered ownership and immutable baselines. Real mutation detaches its
actual owner through Arc::make_mut after resource preflight. Nested result owners
also share their allocations until mutation.

Parse the actual result interval AFTER fldData in its genuine field scope through
existing CT_P ownership parsing. A parse-only paragraph envelope never becomes
persisted source or identity. Cache and accepted paths come from actual accepted
owners, including supported revisions/controls. Deleted-only source can have an
empty accepted cache with a nonempty stored owner. Original raw bytes remain exact
while unchanged. Root/fldData/unknown events stay outside accepted mutation budgets.
Namespace-qualified foreign/shadowed names and unknown-wrapper descendants remain
opaque. Prior property snapshots are nonlive. Move/property/nested states keep
existing CT_Revision admission. No fldSimple-under-run-revision admission follows.

cached_fields_in_source_order walks actual accepted paths in this current owner.
cached_field_mut follows the SAME actual path/run/Field content slot, detaches
COW ownership and returns that actual stored Field, not a flattened clone. A
private mutable path walker checks every bound/source slot and marks the traversed
result/revision ancestors pending BEFORE returning &mut Field. There is no completion
hook when that borrow ends. Do not use physical_content_paragraph_mut's layout-only
contract as an authoring hook. A no-op mutable borrow may conservatively leave a
pending ancestor until normalization proves no change. It must not manufacture
revisions, lose exact source or replace the receiver on a genuine no-op.

cached_result_runs retains Option<&[CT_R]>. Some requires a genuinely contiguous
actual accepted slice in this canonical owner AND the existing cache consistency
check. Direct result runs remain borrowable. One insertion paragraph's actual run
slice may qualify when complete accepted coverage and absence of interleaved or
excluded owners are proved. Accepted runs spanning several revision/control Vecs
are not one slice and return None. No temporary reference, mutable accepted mirror
or fabricated empty slice is allowed. Every structured consumer first checks for
the retained owner and uses its accepted paths/prepared projection when no slice
is borrowable. None never means no notes/drawings/content. Existing ordinary
no-owner paths and their borrowed positives remain unchanged.

### Concrete projection and normalization operations

Three additive doc-hidden native lower bridges on the existing Field implementation
are approved, including the bounded physical-fact clarification below. They use the existing lower Result and add no Python/WASM API,
public model, trait, generic, file or module:

```rust
#[doc(hidden)]
pub fn cached_result_owner_projection(&self) -> Result<Option<CT_P>>;
#[doc(hidden)]
pub fn normalize_cached_result_owner(&mut self) -> Result<()>;
#[doc(hidden)]
pub fn cached_result_physical_runs(&self) -> Option<&[CT_R]>;
```

The read bridge clone-stages the ACTUAL current COW owner and pending field facts,
applies the checked source-bound budgets below and normalizes the staged result.
It returns an owned current ordered paragraph, excluding fldData/field root. None
means no retained owner, not empty text. It accepts no expected input, view flag or
fallback text. Existing layout projection selects RevisionView. The owned bridge
cannot update existing borrowed references or establish that self was normalized.
It is prepared once per operation, not per glyph/reference.

The mutable bridge is a no-op for no retained owner. Otherwise it stages current
COW state, pending parent cache override and nested actual child changes, normalizes
nested field owners and refreshes affected CT_Revision wrappers BOTTOM-UP. Reuse
edit_accepted_paragraph/refresh_accepted_content's retained wrapper/source-context
mechanism. For a previously escaped mutable child, an explicit pending-chain refresh
must invoke that checked mechanism even though the edit already occurred. Reparse
actual refreshed source and validate content(), content_paragraph(), nested_revisions,
accepted cache/runs, root/source/graph/namespace/grammar and held-source obligations
before assignment. Adopt the actual normalized owner and current source/cache
baselines, while retaining original immutable provenance separately. Associate
reparsed owners with the actual old logical source slots and published ranges
BEFORE assignment. Do not infer correspondence from new parse ordinals/source IDs
or matching text. Checked replacement may invalidate existing handles only under
the existing mutation policy. Do not reparse only serialization output and leave
the live model unchanged.

The existing lower CT_Revision public signatures remain unchanged. Its current
paragraph storage can use Arc<CT_P> privately, sharing it across COW clones and
compatible secondary caches. Existing checked revision edits must detach the actual
paragraph before mutation and commit all projections together. The physical layout-
only mutator retains its existing documented source behavior. Source write_xml
continues to write normalized raw_xml, not an uncommitted layout projection.

set_cached_runs, set_cached_runs_with_comment_ranges, actual facade field update
at field.rs17925 and every checked publication/update path MUST normalize the staged
model BEFORE successful live commit. Then fresh borrowed RevisionRef readers,
insertion_paragraph text/items/nested Field cache/properties, lower content()/nested
references, actual raw source and save/reopen must agree. Atomic facade transactions
keep the original receiver, graph and live-owner policy on failure. An &self source
helper may normalize and serialize a staged copy but cannot count that as normalization
of the live model. Document publication with mutable staged ownership must normalize
that staged model before it becomes the receiver. No binding-only completion rule.

The outer parsed Field source fragment is also a replacement LOCATOR in its
containing paragraph, not merely a cache baseline. source_replacement at lower
text.rs1996 returns that original fragment, and field_source_replacements_at10806
uses it to replace the actual parent bytes. Normalizing an inner current result
must not rebase that outer locator to new bytes while the containing source still
has the old fragment. Retain original outer provenance/locator separately from
current inner normalized source. Only a checked containing-owner publication may
advance their correspondence together. is_unchanged at1809 and shared-span source
fast paths must include actual current ordered-owner deltas against original
source. A normalized metadata-only nested edit cannot disappear merely because
its reparsed child now compares unchanged to its own new local baseline. Test
that edit through current borrowed readers, repeated normalization, parent source
replacement and save/reopen, while a no-op mutable borrow remains byte-exact.

### Honest pending low-level mutation state

An escaped low-level &mut Field or direct public cached_result/instruction/dirty
assignment can leave PENDING state. Immediate raw user values and actual mutable
child object exist, but a public String assignment cannot synchronously rewrite
an enclosing paragraph or refresh raw/content caches through &self. Until explicit
normalization, stored revision subprojections/raw and structural owner facts can
remain last-normalized source. effective_display honors its existing parent override
policy, and the fallible owned read bridge can prepare a current projection.
Borrowed structured readers are not a promise of synchronous publication of those
pending raw assignments. A canonical current insertion paragraph contains the actual
mutated child object, never a stale child from an immutable wrapper baseline, but
its separately stored raw/legacy subprojections still require completion.

This pending phase is explicitly LOW-LEVEL. It does not weaken existing checked
Document mutation or RevisionRef contracts. A fresh borrowed reader AFTER successful
checked setter/update/publication must expose the normalized current owner. Existing
Rust borrowing prevents mutation while a RevisionRef is held. After raw lower mutation,
call normalize_cached_result_owner before requiring normalized borrowed structural,
raw or revision readers. Document this requirement on the existing mutable child
accessor and both bridges, with a direct assignment example. The public raw String
remains immediately readable as pending input. Failed normalization reports the
actual OxmlError, leaves that user's pending input in place and publishes no
normalized candidate. It does not promise rollback of an earlier raw assignment.

No pending-query API is proposed without a named necessary consumer. Tests can
inspect pending raw values, the owned projection and explicit normalization without
one. A conservative private pending-chain flag plus direct fact/baseline divergence
checks drive preparation. Invalid new raw facts never become silent stale successful
fallback. No unsafe, interior mutability, destructor hook or borrowed temporary is
permitted. Existing public signatures and checked handle invalidation remain intact.

### Current simple-field paragraph text

ParagraphRef::text at paragraph.rs2768 calls CT_P::accepted_text. The current
append_accepted_text at text.rs7228 reparses unchanged simple-field source, while
a changed simple Field falls back to CT_R::text, whose effective_projected_text
policy intentionally exposes only complex fields. COW ownership does not fix this.

At the affected admitted Field reader path, explicitly select a current retained
simple Field's accepted/effective result through the same actual owner provider.
For already-normalized current source, unchanged-source preservation still works.
For a pending child, its actual Field cache/item remains readable under the pending
contract, and current effective text uses that child's established display policy
where this paragraph reader promises simple-field text. Unsupported/foreign raw
lookalikes do not enter this typed path. Do not change global CT_R::text or its
complex-only semantics, invent a raw fallback or concatenate deleted/inserted text.
ParagraphRef text, insertion items and nested Field cache after checked completion
must agree with actual normalized source and the accepted projection. This typed
text reader uses the existing nonfallible effective display, including a pending
raw cache value, without serializing or validating that pending input. An invalid
pending instruction can still have a readable cache String. Only fallible checked
normalization/publication certifies valid source. Never silently substitute old
source text when current admitted Field display is available.

### Construction cost before COW cloning

Arc cloning alone does not close parse amplification. CT_Revision::try_from_raw
currently runs parse_content and then parse_accepted_revision_content independently
at lower revision.rs175. A recursively retained field-result owner inside both
routes could be reconstructed twice at every level before any Arc is shared.
Replace that duplicate recursive construction for the new retained owner path with
ONE canonical scoped paragraph parse and derive compatible secondary caches from
that parse's physical source facts. This applies to EVERY enclosing revision
constructor whose actual admitted subtree contains a newly retained result owner,
not only the innermost Field. An old dual-parse ancestor cannot recreate that owner
twice. Unaffected old input admission and read behavior remain unchanged. Keep public RevisionContent and nested_revisions
signatures and their existing source/admission behavior.

Concretely, the existing private paragraph/revision parser records operation-local
physical direct-run and nested-wrapper facts at their actual source paths/ordinals,
including hyperlink source scope, BEFORE complex/shared-field grouping. Reuse those
already parsed actual run/child objects to derive legacy RevisionContent::Runs and
nested_revisions in their existing ordering. Legacy Runs is not the canonical
paragraph's flattened accepted view: it must retain the same direct-run/hyperlink
selection, deleted/prior content, marker representation and opaque exclusions as
parse_content currently promises. The simple-field preflight must also select
the retained-owner parse route BEFORE its old direct-run/cache pass, rather than
parse those owners and then recursively reparse them in a synthetic CT_P fallback.
Never admit fldSimple under revisions or promote
unknown wrappers while deriving the cache. If canonical field grouping changes run
partitions, use the retained physical facts, not text/source_id matching or a
second parse of each entire result subtree.

Store canonical revision paragraphs as private Arc<CT_P> so deriving/cloning a
secondary CT_Revision shares its paragraph and nested Field result owners. Otherwise
deriving a legacy child by deep-cloning its whole paragraph can recreate exponential
construction. Secondary nested-revision Vec copies remain source-derived read caches,
not independent mutation targets. Charge the complete composite-source-fact
allowance before allocation, in the explicit A16 resource units below. The immutable original Field baseline is not reparsed
or deep-cloned per ancestor. This is one bounded provider in existing parser files,
not a general parser registry or a new XML framework.

Preflight actual qualified source ancestry/events iteratively before recursive
construction or COW detachment. Reuse current revision-depth32 and weighted owner256
checks, charge newly retained result-owner edges within that owner traversal, and
use checked arithmetic for node/byte/work bounds. For this path, each source run/
field/revision subtree has ONE recursive construction. Compatible secondary-cache
copies may revisit stored facts only along their actual bounded ancestor paths.
Instrument actual recursive construction visits separately from secondary-copy
visits. The former must not have a two-children-of-the-same-source recurrence.
The latter must be bounded by the preflight sum of actual source fact sizes times
actual enclosing secondary-cache ownership occurrences, calculated before copying.
Initial canonical construction and checked refresh are distinct measured phases.
Bottom-up normalization may reparse an affected ancestor once at its own completion
boundary, but must charge the sum of actual affected subtree sizes before those
refreshes. It must not recursively run both legacy and canonical parsers inside
each refresh. Bound total normalization construction visits by that preflight sum,
and secondary-copy visits by the separately computed ownership-occurrence sum.
Overflow or an existing owner/depth/resource violation refuses before allocation
and leaves source unpublished. Do not invent a timing pass or silently lower an
existing accepted limit. If a required old admission cannot satisfy single canonical
construction plus compatible legacy derivation, stop that implementation boundary
for root rather than keep duplicate recursive parsing with an Arc-only claim.

### A16 resource metric units

For this bounded provider, secondary-copy resource units are complete qualified physical source facts and their actual enclosing legacy-cache ownership occurrences. A fact is the complete physical run or revision-cache owner, including its private modeled and preserved payload. Charge a conservative checked allowance from that complete source interval, current authored payload and namespace context before copying it. Internal owned model payload covered by that allowance does not require a generic recursive heap profiler. Directly visited revision/raw-property buffer lengths are charged before their visible copy and may be reported separately with that explicit scope. No allowance or partial buffer sum is an exact heap-byte or exact internal raw-copy measurement.

Observe canonical Field/revision constructor entries, actual physical-run/revision-cache fact-copy calls including recursive legacy-cache occurrences, whole and own normalization constructor entries, and owned CT_P root-copy calls separately from the preflight allowance. A CT_P root-copy call is not a count of every internal model clone. A recorded source-byte allowance is not an observed copy-byte count. Composite-fact coverage must include every actual supported source/current payload route, including payload carried privately by the modeled owner. Missing complete fact coverage remains blocking rather than disappearing behind an Arc or an allowance label.

Initial construction must still construct each actual retained source owner once. Structural current-owner and source-fact derivation evidence, actual constructor counts and complete wrapper identities must exclude duplicate recursive reconstruction. Before checked normalization, charge the sum of actual affected qualified subtree work across its enclosing refresh boundaries. Observe the whole normalization phase, including child revision refreshes, as well as each owner's own parse. Bound those constructor entries and the actual source-fact copy visits by their separately computed checked source/ownership allowances. Preserve existing depth/resource admission, no-op source/current semantics and atomic refusal. Do not lower accepted limits, invent a timing pass, add a public API or require unrelated model-wide clone instrumentation.

Root approves this bounded technical metric amendment after reading the complete
independent closure report54f0f4fff406421043ad92d07f31ddb80b930a8f430aaa031f5b5144c6335ad9.
The exact recommended paragraphs are retained at recommended-plan-wording.md,
SHAfd7b02e112de5f9073dbf0b155514bd04c15d81b196c2c4edbe27a62cec70a0d,
in fx272-a16-metric-contract-closure-v1. Complete supported private/current/context
payload coverage remains blocking. This approval adds no heap profiler, public
API, unrelated model or native campaign and establishes no runtime acceptance.

### Accepted mutation and deterministic replacement budgets

Direct assignments and nested mutations are detected from current field facts,
original/current baselines and pending owner chains. Preserve the distinction
between observable parent cache override and child-derived updates. Same-slot child
mutation targets the actual canonical path. Root-only lock/dirty/instruction edits
retain stored result source and held bytes. An immutable source baseline is never
chosen as an alternate current mutation authority.

Text override extends the existing first actual accepted text/tab/br replacement
rule into admitted recursive owners. Write the new display at that actual selected
child with existing escaping/whitespace/tab/break semantics and its actual formatting.
Retire other selected accepted text/tab/br budgets, preserving shells/nontext/opaque/
prior/deleted data. Never edit deleted text as accepted. If no accepted text-like
child exists, preserve nontext content and append the existing ordinary result-run
fallback at the actual direct result END before field close. Use existing first-
accepted-property policy, not deleted formatting. A truly empty result has this
boundary after fldData. Empty override invents no visible text. Nested Fields keep
their actual source-owner policy rather than becoming ordinary t children. A NEW
inseparable imported held owner which prevents complete accepted-display coverage
is the explicit checked budget limitation, not stale successful cache publication.
Existing no-owner legacy parent-override positive text.rs21493 remains unchanged.

For checked typed replacement the selected budget is ALL actual accepted run slots
identified by recursive paths and ordered child bounds. Deleted/prior/opaque events
and wrapper roots are excluded. Place the ENTIRE requested validated run sequence
at the FIRST selected accepted run boundary under its actual immediate owner, then
retire only old selected accepted budgets at all paths. With zero existing accepted
slots, anchor at the direct result boundary after fldData before stored result events.
Zero replacement runs adds no payload. More/fewer never zip unequal counts or pick
anchors by text. Explicit typed replacement properties are authoritative, unlike
text override's original formatting. No arbitrary source movement is inferred.

A selected old run with excluded opaque metadata/events or held prior properties
remains as its actual shell/residual owner in place, with only selected modeled
accepted content removed. Insert replacement before the first residual anchor run.
Do not spread its metadata onto new runs. A wholly selected clean run may retire.
The actual normalized accepted projection may therefore contain retained empty
residual runs beyond the requested run count. No promise that getter slice length
matches input length overrides source preservation. Comment-range replacement maps
requested boundaries through the actual insertion/residual positions and validates
that actual annotation graph. No stale flattened offsets or hidden held-marker moves.
Existing no-owner setters and lock/annotation validation remain unchanged.

An inseparable shared field span, unmodeled wrapper, unsupported graph relocation
or nested held-data retirement has a precise checked failed budget/owner, with atomic
publication. This applies to unsupported NEW imported mutation shapes, not blanket
refusal of revisions/metadata or a dropped required positive. Existing setters,
single-insertion results, direct multiple owners, fewer/more/zero, same-slot nested
updates and existing annotations remain required. A contradictory required positive
returns to root. Validate actual source budgets, namespaces/grammar/graph, serialize/
reparse and normalize all current raw/typed revision projections before checked commit.

### Current borrowed listings, identity and read consumers

Typed collect_field borrows CT_Revision wrappers from the CURRENT canonical ordered
owner and traverses its actual stored children once. Original baseline wrappers
are never returned as current accepted insertions. After normalization the wrapper's
content_paragraph, RevisionContent, nested references and raw XML are coherent.
Deleted/prior owners retain their stored source and ordering. No temporary owned
projection can supply a borrowed RevisionRef. Existing CT_Document::revisions and
Document::revisions lifetimes/signatures stay intact. All-story physical listing
and resolution remain the one existing shared scanner/transaction.

Layout/evaluation/source operations own each required prepared projection once,
then freeze its concrete operation-local storage before borrowing shaping/reference
runs. Use actual story/paragraph/outer field path/slot and recursive result owner
path/ordinal. Accepted nested fields map to actual canonical current accepted paths
and Field content slots. Deleted/prior fields map to stored current source paths.
COW/shared source_id, equal bytes/text, cloned pointers or unchecked ordinal zips
are not occurrence identity. Shared physical fields still have distinct logical slots.
Validate complete injective association. Pointer-indexed layout maps may use prepared
clone addresses only AFTER that source/path association is established. Accepted
clones reuse canonical evaluation/source indices, while deleted/prior clones are
stored occurrences and are not reevaluated as live fields. No glyph-time repreparation.
Cold/warm caches include current prepared facts and RevisionView, not old raw only.

Accepted cache/REF/reference readers use actual accepted content. Stored Tracked
painting uses existing project_paragraph_runs/project_revision_runs with inserted
underline/deleted strike and prior property policy. Accepted cache never concatenates
old and new. Existing unlocked computed-field substitution and first accepted stored
property authority stay unchanged. Deleted immutable fields are not recomputed as
live accepted fields. Locked/uncomputed fallback and outer revision suppression of
computation keep their existing caller policy. Do not globally replace computed
fields with stored text. Note/drawing/reference traversal selects the same view and
canonical identities, including accepted-empty text with real structured content.

| Consumer or operation | Concrete provider/change owner | Required observation |
|---|---|---|
| Parse, accepted cache/slice and serialization | text.rs parse_simple_field, current COW owner, both bridges, write_field/update_simple_field_source | One current owner, exact unchanged raw, honest pending phase, normalized current mutation |
| Revision construction | lower revision.rs try_from_raw/parse_content and existing paragraph parser | Single recursive parse, compatible legacy read caches, actual preallocation visit bounds |
| Revision ancestry/allocation/held/resolution | facade revision.rs modeled_child, XmlTree and existing transaction | Qualified real owners only, new visibility/collision coverage, no invalid fldSimple-under-revision admission |
| Typed main listing and insertion readers | lower revision.rs collect_run/collect_field plus facade RevisionRef::insertion_paragraph | CURRENT wrapper and actual child identity, normalized content/raw/nested/text/items agree |
| Paragraph simple-field text | text.rs append_accepted_text and paragraph.rs ParagraphRef | Current admitted simple-owner display, no global complex-only CT_R.text change |
| Resource/graph custody | text.rs resource_owner_children and lower document.rs WordOwner walk | Accepted/deleted/opaque references distinguished, actual Arc allocation/parse-copy visits before recursion |
| Material/reference ownership | document.rs2311/14740 | Structured current result even when borrowed slice None, no deleted-only accepted reference |
| Cache extraction/REF annotation/order/update | field.rs13536/14673/15636/17925 | Current paths/indices, no duplicate evaluator, staged normalization before success |
| Layout registries/notes/references/painting | engine.rs686/734/2233/2372/2890/3646/3683/3710/11030 | One prepared current owner, actual associations, computed/Tracked policy and cold/warm agreement |
| Rust/Python/CLI checked comparison/lifecycle | Existing compare and counted mutation APIs | Same normalized live publication, counts/handles/reopen, no alternate binding resolver |
| Paired host policy | compare_paragraph, compare_granular_paragraph, compare_complex_paragraph, compare_control_content, compare_control_runs_with_options | Both capture routes and granularities, genuine VML/WPS/MCE/related source authority, unchanged guards |

Shared facade/lower simple-result admission must agree at real Word paragraph,
supported inline control and nested result owners. Keep foreign/shadowed/opaque
ancestry excluded, prior snapshots nonlive and existing descendant conflict policy.
New revision visibility affects allocation/held checks, so test collisions with
newly visible field IDs and other owners plus ignored-story exclusions. This is
bounded explicit grammar closure, not a second permissive name list.

### Actual controlled field source owner

Root approves existing lower content_control.rs as an additional implementation
file for the ALREADY REQUIRED A12/A13/A16 controlled result path. Actual stopped
D7 construction and independent source report1e8732f5 show that direct fldSimple
currently remains RawXml there. This is a routine concrete owner/caller scope
clarification. It changes no user decision, public model/API or HLD impact file.
Use the existing Field parser and writer with crate-private visibility where
needed. Do not add a parallel parser, new public SdtContent variant or synthetic
paragraph reparse to discover the direct field.

Only an actual qualified INLINE sdtContent child gains the existing virtual
SdtContent::Run(Field) carrier with its same physical content index and source.
Standalone/body/cell/table/row/foreign/opaque cases retain their current admission.
An existing private physical-child-kind fact may distinguish the raw fldSimple
from a real w:r. Both write_content and inline_runs_as_paragraph must emit a
current virtual Field at its actual field boundary using the existing field
source writer. CT_R::to_xml always emits w:r and cannot publish that virtual
simple field. Physical source/admission policy remains unchanged. For actual qualified INLINE sdtContent only, reuse the existing canonical complex projector over already parsed inline facts to collapse a complete isolated source-proved complex span into one existing SdtContent::Run(Field). Shared/partial/cross-boundary/unproved spans retain the existing physical representation and fallback. No synthetic paragraph/model reparse, new public type/API or separate parser is introduced.

The private bridge uses existing CT_P, run/source and boundary types. Keep the complete original top-level physical source extent and its actual occurrence association in InlineRunSource. Replace all and only that span's logical content/source records once. Rebase outside content and revision boundaries without dropping or duplicating raw/control/marker custody. Listing, accepted/mutable selection, cardinality/splitting, legacy value updates and source replacement use the same actual logical owner and physical span. Both write_content and inline_runs_as_paragraph emit current simple AND complex virtual fields through the existing owner-aware writer, with exact unchanged/no-op source and once-only current publication. The legacy updater maps complete top-level physical spans, never descendant result runs or an assumed one-source-fragment-per-logical-run zip.

Run existing qualified preflight before construction/COW and charge the composite source/current payload and copy/refresh work. Canonical field construction remains once. Moving a logical complex owner adds no physical owner-cost edge. Global Field equality stays as approved. Within existing test bodies, discriminate original simple versus resolved/converted complex control ownership, controlled nested result/current mutation and reopen, repeated adjacent fields/source cardinality, exact no-op/metadata-only publication and mixed legacy/physical runs, while retaining alias/default/shadow/block/foreign/partial positives. No new test binary or native campaign follows.
An equality-only original CT_R fast path must not lose normalized metadata-only
field deltas. Current payload shape and original lookup source are distinct.

set_inline_legacy_form_value must map complete TOP-LEVEL physical children,
including a whole virtual field and existing complex/shared spans. Its old
inline_run_fragments/parse_inline_run search can find descendant result runs and
is not valid correspondence for the new field owner. Stage the complete control
and source records, normalize/validate, then assign atomically. Preserve pending
inline field edits and parent lookup bytes. append_control_field_source_replacements
must publish each current direct virtual field's actual replacement once, alongside
the existing complete complex-span edits. Its current unconditional Run skip is
not the new rule. Do not broaden the separate block-control updater's grammar.

Accepted path read/write, splitting/cardinality/comment placement and removal use
the same physical content slot. A field stays atomic at this outer boundary.
Field-interior changes use the approved result budgets. Original/current control
resource copies share COW owners and charge actual secondary-copy work. Existing
legacy, standalone/placement and control/comment positives remain required. Add
qualified alias/default/shadow versus block/foreign/opaque controls, both writer
paths after real/no-op/metadata edits, mixed physical-run/simple/complex legacy
update, annotation/cardinality and source replacement once, plus ACTUALLY ADMITTED
deep/sibling construction visits. Complete-part schema validation must detect
w:r/fldSimple introduced only after mutation. The report's concrete caller map
is consumed, without claiming these unexecuted implementation controls passed.

### Independent guard and conversion verification

Both default and ignore_formatting guards must inspect actual fields for effective
instruction, actual kind, accepted cache and field-source/root/result facts. The
current default placeholder is insufficient. Use actual physical OPC owner/graph
normalization in main and related stories. The guard does not call comparison,
resolve expected inputs, borrow an expected model, or fill actual text/cache.
A default guard must reject unresolved generated nested IDs/cache loss independently
of the stronger source guard and final held/new-ID checks.

Only a conversion witness permits a local kind/source-parent representation
exception. On independent accepted and rejected actual copies, locate exactly the
corresponding complete complex span by actual publication association. Verify
balanced begin/separate/end, actual instruction decoded identity, exact flag presence/
lexical values, fldData subtree and actual result/metadata/order/cache/graph facts.
All generated fragments must be consumed once. Do not compare only a payload built
by the same constructor. A separate checker inspects actual source events against
mapped source facts. Reject neighboring conversion, duplicated/missing markers,
changed flags/fldData/result, borrowed witness or unconsumed source facts. All other
fields retain exact kind equality. Source metadata significance, held byte/order,
matching/move signatures, scalar ledgers, no-op/live-owner and both reconstruction
equalities remain mandatory. A form witness records the actual original/edited
form delta even where both emitted payloads normalize to the same complex form.

### Observable criteria and next coherent evidence

These rows extend A2/A3/A6/C1/C2 and do not replace the original acceptance rows.
Each uses existing source/test carriers. No new test binary or tracked file is needed.
New controls without actual stopped-source execution are source-derived expectations,
not retrospective RED. Execute both directions independently or aggregate failures.

| ID | Discriminating input | Required GREEN observation and wrong implementation rejected |
|---|---|---|
| A8 | Actual DATE result edit and first MERGEFIELD header result, both formatting options | Nested field-owned del/ins2, simple root/flags/fldData/source preserved, accepted new/rejected old result and properties, complete schema-valid parts. Rejects illegal outer simple wrappers |
| A9 | Actual second MERGEFIELD instruction edit and ordinary DATE-to-TIME | Complete complex old/new spans2, instruction/cache correct in both resolutions and reopen, exact local conversion witness, persistent complex. Intentionally update only Q2's two-simple-root representation assertion, retain old oracle/history and all semantic/count checks |
| A10 | Clean simple-to-complex, complex-to-simple instruction and form-only pairs | Actual instruction OR form witness before PartialEq, existing complex source exact, converted simple facts exact, two new records even form-only equal normalized payload, partial selection legal. Rejects equality shortcut and blanket kind normalization |
| A11 | Actual paired simple result/instruction plus unknown-root source-policy positive | Real host admission, simple retained, original shell/actual edited result for same instruction, field-owned properties, source guards and no-op identities. Rejects leaf_capture-only exception and ordinary conversion leakage |
| A12 | Canonical/alias/default/local-shadow generated nested-result INPUTs, real nested prior revisions and controlled/table v2 | Shared scanner/lower/listing/cache closure, correct accepted content, selected all/author/ID/date counts and saved IDs. Foreign/opaque remain opaque. Default wrong-cache/instruction/form/remaining-ID perturbations fail |
| A13 | Pending direct parent cached_result override, nested cached_field_mut real/no-op borrow, dirty/lock root-only edit, checked typed fewer/more/zero and annotations | Explicit pending values/owned projection versus last-normalized stored subprojections, then normalize and require fresh insertion text/items/Field caches/lower content/raw/nested refs/save-reopen agreement. Checked setters normalize before success. First-text and first-anchor budgets retain excluded source. Locked/invalid attachment fails atomically |
| A14 | Repeated identical fields, shared physical span, accepted-empty/deleted-only, note/drawing/reference content, nested revisions | Correct logical indices, accepted versus Tracked painting/decorations and computed-field policy, one physical revision each, graph custody, no duplicated evaluator/source registration |
| A15 | Known flags absent/false/true lexical variants, fldData, encoded instruction whitespace/entities, opaque result facts, root unmappables and transparent nested payloads | Legal mapping or exact source-bound unproved fact, unchanged/ignored/result-only positives retained. Missing/duplicated source budgets, marker/header prefix tricks, illegal nested fldSimple/hyperlink and guard witness theft fail |
| A16 | Bounded deep/sibling field-plus-revision owners, COW clones and cold/warm caches after real mutation | Single recursive construction with separately bounded secondary-copy visits before allocation, no exponential parse or clone recurrence, detached clone/current values and positions agree. No invented performance pass |

The borrowing/normalization discriminators include a legal nested COMPLEX field
inside a result insertion, plus a nested simple field at its actual permitted
result owner with revisions inside its own result. Never fabricate fldSimple
directly under ins to test the new mutable path. Inspect fresh insertion text,
items, nested cache/properties, lower content(), raw XML and nested references
BEFORE save as well as after reopen. COW-clone mutation leaves the original clone
exact. A no-op mutable borrow clears conservatively pending state on normalization
without new records/source change. A raw outer cache assignment explicitly differs
from normalized borrowed structure until completion, while checked setters expose
no such pending success. Assert invalid normalization leaves raw user input pending
and facade publication atomic. These distinguish current live reader correctness
from correct counts or a prepared output copy alone.

Use a bounded matrix rather than multiplying every axis. The ordinary DATE/main
and two-field MERGE/header cases run both options through all/author/each actual
ID/inclusive-date accept and reject on independent copies. Clean single result
pairs count2 all/author/date and1 per selected ID. The mixed header counts4 total,
with one revision per public ID. Narrow/partial resolution retains the unselected
counterpart, held IDs and valid source. Prior7 plus new40/41 generated INPUTs must
now select2 for the new author/date,1 for each new ID and3 for all where all three
are valid modeled revisions. Existing descendant conflicts still refuse rather
than fabricate these counts. Record actual source/listing/spans and both modeled/
broad IDs before resolution, before save and after reopen. No hidden grouped-ID
side effect or author-only substitute for actual IDs is permitted.

Cover aliases/default/local shadows in different routes, not every Cartesian
combination. Put flags/fldData/encoded instruction on ordinary conversion, prior
rPr/deleted-only on result closure, unknown-root metadata on paired/result-only,
and typed cardinalities on direct versus multiple insertion/control owners.
Use independent deliberately wrong cache/instruction/form/root/source/graph/marker
candidates to prove both guards refuse, including their default option. Repair
the two header observers by selecting COMPLETE actual field ownership, not a
first closing-tag prefix, and keep the pinned record distinctions for each field.
A Q2 representation delta is labelled and reviewed separately from field-owned
result records. Do not accept arbitrary new record normalization.

Preserve all original12clean fields,9field-free fragments,30Issue425regressions,
88independent resolution rows,27library follow-ons and the EXACT86member wider
union, including the three current failures. Retain the32library separator/terminal
successes,8Run positives,12new shape/granularity comparisons, both action-domain
ledgers, necessary separators, actual terminal raw suffixes, all five publishers,
related OPC graphs, aliases/default/MCE custody, metadata order/multiplicity,
held-byte/source/new-ID, injected late refusal, genuine no-op owners and all prior
negative controls. Historical83/86 is not a repaired pass. Existing diagnostic16/
128 acquisition remains diagnostics. Newly admitted runtime semantics need real
assertions and a source-authenticated successful binary, not a self-catching exit0.

After root approval and an explicit separate writer/build assignment, first capture
new stopped-source discriminators where missing, then one coherent source repair
and meaningful scoped checks. Prepare exact named semantic membership with root.
Run one native opt1/assertion/overflow/8MiB semantic batch with before/after source
pins and actual Running binaries. Any continuing core failure stops all product
writes/builds for independent reassessment. Actual GREEN permits proposed affected
lower/layout/resolver/CLI, strict affected-crate check/clippy/fmt, exact selected
Python nodes with fresh opt2/assertion/overflow wheel build/install/import identity,
and a DIFFERENT independent microscope. No stale executable after build failure.

Validate actual complete original/candidate/resolved Word parts using unchanged
schema runner62aaf1a1, wrapper1556888d and29pins, with explicit MCE reduction limits
and original invalidity labels. Require clean candidates and resolutions valid,
including nested payload grammar. Native Word comparison/unresolved display and
all/author/ID accept-reject artifacts remain due, with actual producer/input/output
identities. Existing F274 open/F9 normalization is not a substitute. Same-input
positioned accepted/Tracked layout and note/reference controls are required local
composition checks. Final integrated union, installed Python3.9/3.12 complete
runtime/strict typing/stubs, WASM, CLI compositions, native/hosted/oracle,49hashes/
7pixels, real README examples/measurements, affected archives/size and all risk/
metadata-policy riders remain pending. No baseline, readiness, commit, push, merge,
tag or release is authorized by this amendment.

### Exact impact and closure disposition

The existing HLD impact list remains EXACTLY docs/hld/03-architecture.md and
10-bindings-spec.md. Before implementation freeze, replace the now-stale field
publication intent in HLD03 with simple-result ownership/shared lifecycle,
source-qualified paired projection, clean instruction/form complex carrier,
canonical ordered COW ownership, explicit pending/normalized state, bounded
construction, current simple-field readers, ordered Tracked/layout policy and
local witnessed guard exception. HLD10 records persistent comparison representation,
unchanged authoring/binding signatures, counts/partial resolution, handles/cache
lifecycle and both additive doc-hidden native projection/normalization bridges
with no Python/WASM export. They describe implemented intent, without claiming pending native evidence.
No HLD edit occurs during this designer lease. No shared delivery tracker changes.

The intended existing implementation files are comparison.rs, facade revision.rs,
lower text.rs/document.rs/revision.rs/content_control.rs, layout engine.rs, existing field/document
consumer sites where the table requires a change, and existing test carriers.
No new tracked file/module/crate/trait/generic/public model or general XML/schema
framework is proposed. The two additive exported lower methods are specified above.
No speculative pending-query API is included.
Scanner grammar is changed only for the explicit source-qualified result owners,
not to legalize forbidden whole-field XML or discover opaque revisions globally.
F273/CURRENT and unrelated rootTOC/product behavior remain outside this repair.

Closure disposition: Q-S96-02 is answered. Root has supplied the form-only carrier
application. Callable ownership, mutation cardinalities, listing lifetimes,
projection identity, single-parse construction, pending/normalized reader state,
shared consumers, source mapping and verification boundaries
are specified above for review. No new material user question is identified.
Actual native interoperability, bounded resource behavior, new grammar/typed-owner
agreement and independent guard/source preservation remain UNEXECUTED acceptance
facts, not design approvals. A newly discovered required imported mutation or
ordinary opaque-root instruction positive without the specified source mapping
is a concrete contract contradiction to return to root, never an implicit new
blanket refusal. Root read the full current amendment, authenticated its input
and review packets and approved this contract after the independent finite check.
A separate implementation assignment must acknowledge these canonical bytes and
retain all unexecuted runtime/native/final obligations.


### Local lower semantic implementation check

Root authorizes a finite lower semantic implementation check before dependent
comparison and layout implementation is built. The exact52existing lower members
and3already-named new lower carriers require actual source, runner, profile and
artifact review before execution. This is a local parser and ownership check,
not a dependency or provider completion checkpoint. It changes no feature status,
ledger, final coherent union, required positive, HLD impact or independent review.
A continuing core semantic failure stops product edits and builds and returns
sealed actual evidence for independent reassessment. GREEN preserves its exact
lower transitive source and configuration inputs for later applicability. Lower
tests do not repeat merely because unrelated facade or layout source changes.
The complete A8-A16 correction and independent final microscope remain required.


### Complete complex result ownership for legal conversion

Reuse existing private OrderedFieldResult for a complex field only when actual source mapping proves a complete isolated span with pure separate/end marker runs and wholly owned result geometry. Transfer the already parsed result runs, controls, revisions, opaque records and markers, with exact order/boundaries and namespace context, into one current CT_P. Do not recursively parse a second result model. source_range is the actual result interval after the separate marker run and before the end marker run, including all owned result bytes and excluding marker shells. Zero-result placement is explicit. Begin/instruction/separate/end shells, flags, fldData and outside payload stay under existing source authority.

Any partial/shared marker run, cross-boundary owner/range, overlapping outer geometry or unproved span retains its existing fallback/admission and exact source. Ordinary simple conversion emits only this supported complete subset. Independent emitted-source checks prove root field grouping, instruction/cache/result owner facts, balanced markers, flags/fldData and conversion association. They cannot substitute expected facts for a missing actual RootField.

Ordered complex publication composes the exact result interval with existing complex instruction/flag/span logic. Never seek a fldSimple instruction attribute on complex source. Current and original cache/instruction baselines, source lookup locators, child identities, aliases and parent replacement bytes remain coherent after live normalization and reopen. Apply the same current borrowed revision/cache, first-anchor mutation, atomic checked publication and once-per-operation prepared consumer contracts as simple ordered results.

The logical complex CT_P projection adds no new physical owner-cost edge. Preserve old qualified physical admission and all existing limits, while still traversing its actual result owners and charging composite copy/whole-refresh work before allocation. Do not waive limits or silently lower old valid complex/control depth. Global Field PartialEq remains unchanged. Actual form-change dispatch and local conversion witnesses carry representation changes. No new model/public API/files or broad parser registry is introduced.

Root technical approval follows complete independent source closure758229159357adaf99aa3c135cbf5bc84b3db5e97f76f3c4a7268ed3af1e3c6f. The exact reviewed amendment is aadf61a0ad5ca5b4ba92d144199994fe2b7f184898c400215a64bd503a0ae3f7. This is bounded provider completion for already required legal conversion, with no runtime/native or feature acceptance. Global Eq means the current approved form-insensitive semantics with current-owner comparison when both have owners, and no owner-presence proxy.


### Root approval of complete inline control source ownership

Root fully read independent inline-control closure928ee930 and approved the exact
amendment41a508ebcfb63cf05be0d4d3fe6697c2a944a2d2ccc66057ef98de1e29c83065
before affected lower control edits. The report and reviewed amendment are retained
in /private/tmp/rdocx-s96-run-refresh/fx272-inline-control-complex-contract-v1.
The clarification preserves physical admission and source custody while allowing
one complete source-proved complex logical owner at its actual INLINE control
site. It adds no public model/API, parser, native scope or physical owner edge.

### Current coherent scoped membership

The approved coherent batch is256 exact tests: original250 unchanged plus six
explicit existing lower controls approved for the inline-control amendment.
Cohorts are lower61, layout14, facade65 and regression116. Original250 proposals
and frozen lower55 receipts remain historical artifacts with their original counts.
All original241 existing and nine new names remain required. No membership is
removed or silently substituted. The six additional existing names are:

- content_control::tests::block_control_parser_types_only_children_owned_at_its_placement
- content_control::tests::public_standalone_control_parser_keeps_the_union_child_contract
- content_control::tests::run_control_keeps_comment_anchor_and_hyperlink_boundaries
- text::tests::inline_control_keeps_block_children_opaque_to_typed_projections
- text::tests::nested_control_local_word_alias_projects_bookmarks_after_reopen
- text::tests::accepted_inline_owners_project_bookmarks_once_in_exact_order

Root fully read the source bodies and approved the exact coverage clarification in
/private/tmp/rdocx-s96-run-refresh/fx272-inline-control-test-addendum-root-approval-v1.json.
The actual source freeze, binary/list authentication and semantic execution remain
separate required steps. This is coverage clarification, not a new product criterion
or a claim that any of the added tests has executed on current source.

### Root-approved scoped nested source-fact consumption

The existing source classifier can retain a nested field/control/result subtree as
one opaque metadata contribution under an otherwise modeled parent. In particular,
text.rs source_modeled_attributes recognizes fldSimple under p, hyperlink,
sdtContent and r, while a nested carrier beneath fldSimple can remain opaque.
No global classifier change follows.

The existing private comparison and guard carriers may use an independent source
fact tree to consume a nested encoding exception. Mapping the existing opaque parent
metadata row requires a complete source budget first. Child/control and already
canonical revision intervals must be actual, disjoint and contained. The parent
shell and every remaining byte, attribute and opaque fact compare strictly. Every
omitted span is independently checked recursively. Exact role-local child witnesses
are consumed once. Unclaimed children remain strict. Shared, overlapping and unproved
intervals refuse or remain strict opaque leaves.

This is a source inventory over existing prepared owners and qualified source
primitives. It introduces no public API, model parser, module, file or generic
framework. It does not synthesize expected XML/model, flatten GuardField or erase
ancestor facts merely because a child witness is valid. Existing current-value
preflight and work bounds remain required. The existing named bodies include direct,
controlled and revision-held nested controls and an outside-child parent-fact tamper
negative. All old assertions and exact256 membership remain binding.

Root technically approved these bounded implementation criteria after check85 and
the actual classifier/source ownership report. The implementing owner acknowledges
this clarification before dependent parent metadata mapping. Independent source
closure and current-source semantic acceptance remain pending.

The revision-held nested discriminator keeps the canonical prior revision subtree
as an unchanged strict leaf in the independent guard fixture. It does not authorize
changing a prior revision-bearing owner through public comparison. A separate public
attempted-edit control requires atomic refusal with source/model/live owner and
cached layout identity intact. The existing README refusal and checked-copy
workaround remain authoritative. Root explicitly approved this geometry clarification
after the scoped parent-budget implementation, without expanding the product scope.

### Actual publication transport for conversion guards

Root approves a separate crate-private mode in the existing revision resolver's
source-range transport for positively qualified publication roots. Existing text
leaf admission and every public API remain unchanged. A complex field owns several
sibling roots. Transport the complete actual root sequence, including proved member
and gap coverage, then require both the independently observed resolved full part
extent and logical Field target to match. Equal payload or one begin-run offset is
insufficient. Unproved packed/shared multi-field, property replacement and paragraph
merge geometry retains its transport refusal.

Use the same actual candidate resolution/scope and namespace, unwrap and rename
transformations as normal revision resolution. Root ranges are recorded after inner
changes, so modifications within a containing root preserve the final extent. Select
the actual publication payload roots, never its removed revision wrapper. No expected
XML/model, permissive interval fallback or new public model/API follows. The exact
neighboring/duplicated/missing publication and changed-field guards remain required.
Root approved this bounded criterion after full independent resolver reuse closure
3c1e64251b12991c614c0d966a46e3e9d2fcd6be60bc376d490f842cfb9cf786.

## Physical cache facts versus legacy scalar display clarification

Root approves this bounded technical refinement on2026-10-30after the actual
coherent256RED and source-qualified construction repair. It overrides earlier
bridge-count/no-other-bridge wording only for the third exact native method above.
No new public type, authoring operation, Python/WASM binding, parser, file or module
is added. The two existing consumers are comparison result preparation and its
independent Field graph guard. Both need actual already parsed physical facts.

`cached_result_physical_runs` borrows the existing contiguous current cache facts.
It returns no temporary allocation, fabricated text paragraph or second mutable
owner. Ordered owners retain the existing contiguity/current-cache consistency
rule. For legacy parsed/authored caches, a checked current typed cache is usable.
Otherwise original physical facts are usable only while the current scalar still
equals the immutable original scalar. A pending scalar override retires those
original facts. Absence is not an empty cache. Source qualification and actual
child/run ordinals remain the constructor's responsibility.

A paired raw tab/br can contribute its preexisting field scalar display while
remaining opaque in CT_R. An unwrapped deletion-text item can remain typed in the
physical model while the old scalar projection omits it. Thus physical facts need
not have the same text projection as that legacy scalar. The existing
`cached_result_runs` consistency-filtered semantic view and segmented display
remain unchanged. Layout/evaluation keep their existing display policy. Comparison
and graph identity use the borrowed physical view after ordinary owner preparation,
without model reparsing, invented typed Tab/Break facts or relaxed source guards.

Inside the existing new lower body, discriminate unchanged paired raw source with
physical facts present but semantic cache runs absent, exact no-op source retention,
and retirement after a direct scalar override. The original physical Page/tab/1
and packed source cases remain. Include the legacy raw boundary in the existing
graph body to prove target facts are retained under both guard options. All new
assertions require actual successor execution. HLD03andHLD10describe the distinction
before the next freeze. This changes neither resource limits nor physical owner
cost and makes no runtime, native or feature completion claim.

## Qualified hyperlink result and semantic cache eligibility clarification

Root approves this bounded technical exception before implementation on2026-10-30.
The no-retained-owner statement above excludes namespace-qualified hyperlink-owned
simple results: their actual result interval uses the same canonical ordered CT_P
authority so the accepted hyperlink target and child graph remain observable.
The existing qualified ownership scan selects that route before constructing the
result, once. It adds no owner kind, parser, public type, method, file or second
parse/copy. Foreign hyperlink lookalikes remain opaque. Complex event extraction
uses each actual enclosing hyperlink source scope and the run's local overrides.

Physical legacy cache facts are retained independently of legacy scalar equality.
The semantic borrowed cache still requires equality and a significant typed result,
note/comment reference, actual child Field or actual run properties. Plain ordinary
text and SoftHyphen-only results without properties retain their previous None.
The run-property positive is an explicit additive doc-hidden native eligibility
refinement needed by the existing formatted complex-cache layout consumer. It is
not a claim that its former formatted-cache getter returned Some. Existing scalar
display, original raw source, admission and copy limits remain binding. Physical-only
simple cache facts without a child Field cannot act as a nested derived projection
and suppress a direct scalar clear that happens to equal their typed display.

## Current property roles and physical part proof clarification

The existing opaque preservation and conversion postconditions apply to the source
read view itself. Omit rPr only below a proved current Word run and pPr only below
a proved current Word paragraph. Follow the bounded field-result routes through
actual fldSimple, hyperlink and inline controls. Foreign or unknown wrappers,
fldData, prior revisions, drawing and other unproved owners remain opaque through
all descendants. Property lookalikes cannot disappear before residual comparison.
Both paragraph and control field guards use this one read view. Existing guard
test cases additionally require opaque-only edits to remain observable and refused
with either ignore-formatting value. No new model parse, lower bridge or global
classifier change is introduced.

Physical conversion proof inherits actual part coordinates through admitted table,
row, cell and control child roles. Derive each child base from its proved source
boundary, checked root closure growth and existing synthetic opening offset.
Revision and textbox boundaries retain their own admission contracts. A missing
part extent or hash still refuses conversion. An exact unchanged owner can bypass
detachment only with matching required inherited bindings, using the existing
qualified binding scan. Additional unused inherited declarations do not split
unchanged packed runs. Final independent package graph guards still apply.

These are root-approved concrete repairs of the existing preservation, source
identity and publication contracts. They neither add supported conversion owners
nor relax any guard. Actual successor runtime proof and microscope remain due.

## Issue430 amendment intake, contract closure pending

The user authorized [Issue430](https://github.com/tensorbee/rdocx/issues/430) on
2026-10-30. This is an addition to the existing F-X272, not implemented acceptance
or a replacement for the approved prior repair. Reporter lawrenceakka supplied an
exact Python reproduction for PyPI0.16.4onmacOSarm64/Python3.14. Root authenticated
the issue and complete script in the15member TEMP intake packet. The subsequent
published-producer reproduction below is diagnostic RED. No current-source fix,
new native result, contribution adoption or430notification is claimed.

| Row | Required behavior and evidence | Current status |
|---|---|---|
|430-1| Compare a new middle paragraph with a second footnote. Track the new reference/note and preserve the existing note and following paragraph, with schema/native and resolved package proof | Published add/remove failure reproduced, contract pending |
|430-2| Compare removal of the sole normal note. Track the deletion and reproduce either input's note parts, relationships and content types under checked resolution | Published add/remove failure reproduced, contract pending |
|430-3| Retain the existing note text-edit positive, including its reported two revisions and saved accept/reject content | Published producer returns two revisions, resolution still pending |
|430-4| Preserve documented ignored-story left bias unless explicitly amended. Excluded topology must retain coherent references or refuse atomically | Published ignored-note dangling reference reproduced, contract pending |
|430-5| Preserve the following payment paragraph when the seller paragraph is inserted. Compare with a no-note insertion control before diagnosing alignment | Not reproduced in the two structural controls, native/unreduced qualification pending |

Q-S96-04 is answered for representation: pending comparison-owned footnote
lifecycle uses hidden, zero-width STANDARD Word bookmarks, and imported fully
valid convention markers follow the same convention. Preserve original unused
notes. Cleanup of newly owned notes follows the selected owning revision's
accept or reject outcome, retaining notes still referenced by the retained
result. Normal-note identity, schema placement, custody and cross-document
collisions, and selected-resolution package cleanup remain open contract
details. Native Word retention remains unvalidated. Use actual reference/owner
association rather than positional pairing or bare numeric ID equality. Preserve special separator owners,
prior revisions and opaque data. Independently authored IDs, insertion before an
existing note, collisions and genuine ambiguity need a bounded explicit policy.
Do not wrap an entire footnote element in a revision without schema/Word evidence
or blindly prune all unused imported notes. Shared endnote consumers need a
control without silently promising new endnote topology support.

Close this amendment against primary OOXML and actual retained Word output before
implementation. Add its discriminating tests to the existing Rust and Python
carriers after closure. Rust compare/compare_with_options, Python and CLI share
the provider. Revision id/author/date and all/new selection must agree with actual
revision listings and package cleanup. Native/schema and layout composition remain
final riders, not another full-workspace campaign now. The existing HLD03/10impact
list remains, with its topology-rejection wording revised only when the new
contract is closed. F-X275receives the eventual release/credit consequences.

This intake adds material work. The earlier sprint estimate is no longer current.
No question is being used to waive a gate or pause the existing comparison repair.

The exact published producer reproduction is now executed once on an isolated
Python3.14.7/macOSarm64venv with checksum-pinned PyPIrdocx0.16.4. It reproduces both
reported error messages and the2revision note-text-edit control. Exactly two
supplemental compares show that ignored notes produce reference3without owner3,
while both that case and the no-note insertion preserve the payment paragraph
with zero revisions. Keep the secondary rewrite row as not reproduced, with its
narrow positive structural controls and native/unreduced qualification still due.
The root-authenticated report is issue430-public-repro-v1/REPORT.md under the
existing TEMP evidence root. This is published-producer RED, not current acceptance.


Primary-source closure evidence is retained in
issue430-contract-evidence-v1/REPORT.md under the existing TEMP evidence root.
The pinned Microsoft Open XML SDK3.0.1schema confirms that whole footnote roots
cannot be wrapped in w:ins and that omitted type and explicit normal are equivalent.
The current explicit-normal classifier therefore needs a discriminating repair
case. Reference IDs identify owners within one input, not across independently
authored documents. Native observations must settle note-body revision records,
selected-reference cleanup and absent-versus-separator-only part behavior before
this extension is approved. Existing retained Word artifacts do not cover those
facts. Native control currently fails during pipe startup, before any application
action. This is unavailable evidence, not a user permission question.


## Prior repair scoped lint clarification

The scoped lint repair uses private transparent aliases for repeated physical
run sources, inline projections, source selections, publication roles and resolved
leaf ranges. The approved native bridge retains
that exact expanded return type. No public nominal type or exported alias is
added. The private complex-source helper receives one inclusive source span in
place of separate start/end arguments, with the same two callers and bounds.
These are structural changes without a new behavior, owner or lint waiver.


## D3 whole-paragraph publisher remediation, contract clarification pending

Independent microscope all-pass-5 finds a remaining schema-invalid publisher:
wrapped_children places the carried cross-paragraph TOC's nested PAGEREF fldSimple
directly under a run revision. The existing positive must remain supported. The
qualified primary schema separates paragraph children from run-revision children.
Root reproduced the bad shape through four actual public Python comparisons on
the fresh source-qualified wheel, preserving the exact existing insertion/deletion
fixture at Run/Word granularity. This is actual diagnostic RED despite all256
selected modeled tests passing. No native or full schema validation is inferred.

The remediation inventory is moved_paragraph, final_paragraph, inserted_paragraph,
deleted_paragraph and wrapped_children's hyperlink recursion. A correction must
retain source metadata, unchanged nested field form and checked all/selected
resolution unless an explicit existing contract authorizes the representation
change. The paired ordinary instruction/form conversion witness cannot silently
authorize converting an unchanged nested field. Simply moving fldSimple outside
the revision and wrapping its cache is insufficient without proving owner retirement
on resolution. Do not narrow the existing valid whole-field positive into an
unexplained refusal. Root will close this technical lifecycle contract before code,
with a schema discriminator in the existing regression carrier and the saved
public RED as predecessor evidence. The prior repair remains approved, with this
additional remediation and430scope still open.


## D3 approved same-form custody and selected-resolution contract

Root approves this bounded technical repair on2026-10-30 after reading both
independent support reports and authenticating their seals, the actual four-case
public RED and microscope all-pass-5. It preserves the existing supported positive
and unchanged Field form. No new user product choice or public API is introduced.
Root has resumed sole canonical source/build ownership after the F-X273prefix
handoff. The exact source/marker policy below supersedes the preceding pending
technical clarification for D3 only. Issue430 remains separately draft.

## Intended invariant

Preserve the actual nested simple-field shell, instruction, flags, fldData,
namespace closure and opaque metadata. Keep the shell in paragraph-content
grammar, with legal revision wrappers only around eligible result payloads.
The ordinary paired conversion witness remains restricted to actual instruction
or form differences. This repair adds no persistent-form change, public API,
new source file or private persisted snapshot.

Use one actual source-qualified complete enclosing complex-carriage proof.
It is restricted to admitted consecutive paragraph owners inside the same
structural parent and story. A table, block control boundary, separate note,
unknown path or prior/opaque subtree cannot confer this proof. Actual qualified
fldChar roles and fldCharType attributes, complete balanced own begin/separate/end,
packed-run edges, source gaps and exact nested owner intervals are checked.
A carried-membership Boolean is never an ownership witness.

## Publication

All moved/final/inserted/deleted paragraph publishers and hyperlink recursion
use the same concrete owner policy. Keep fldData in its schema-first location.
Preserve existing refusal of captured nested Word hyperlinks, standalone
uncarried simple-field replacement and empty whole hyperlinks. Preserve supported
whole TOC insertion/deletion and unknown simple-root metadata. No cached scalar
reconstruction, blind string renaming, generated-source equality substitute or
opaque descendant descent is admitted. Qualified deletion leaf mapping reuses
the actual Word role and prefix-preserving source policy. Preserve actual source
budgets, disjoint child intervals, residual parent facts and one-use witnesses.

After two actual failed successors, root adopts the independent bounded fragment
audit on2026-10-30 before a third repair. Both failures remain D3: a whole paired
leaf was passed to a one-tag renamer, then a whole paired atomic owner was passed
to an opening-tag namespace injector. The second actual batch is14PASS/2NONPASS,
with required namespace declarations incorrectly placed on closing q:r tags.
The audit covers every new publisher helper call and the existing fragment
helper callers. Its support report is fx272-d3-fragment-contract-escalation-v1.

One-tag helpers receive exactly the actual opening, empty or closing fragment
their existing contract permits. Compose complete source owners from the same
immutable XmlTree extents. Rename each eligible deletion leaf's opening and
closing independently, retaining its inner bytes. Apply disjoint descendant
splices in reverse source order. Those edits cannot change the atomic root's
opening, so its actual open_end-minus-start remains the opening boundary.
Inject namespace closure only into that opening, append the transformed remainder
once, and consume no old descendant coordinates after the opening grows. Preserve
root QName, non-declaration attributes, quotes, child-local bindings and closing
QName. Keep existing helper semantics, opaque barriers and structural-story
isolation. This clarification adds no public contract, parser or model.

Retain the two actual required-closure alias failures. In the existing custody
body, pair them with empty and no-closure atomic owners, paired and empty deletion
leaves, QName-looking attributes, quoted greater-than and comments/CDATA. Inspect
actual expanded names and exact retained payloads rather than searching a whole
subtree for tag delimiters. Existing field-shell, drawing, selection and source
budget controls remain mandatory. Current-source public/schema/installed/CLI and
native observations remain distinct pending evidence.

Root adopts the bounded qualified-balance support report on2026-10-30 after the
fragment successor passes15of16actual bodies. The remaining opaque-marker body
shows an earlier comparison range-growth scan counting a qualified Word end
inside preserved foreign XML by local name alone. Its false Uncarried decision
precedes the complete qualified proof. Both consumers must share physical marker
qualification, not merely apply a stronger guard after the wrong range decision.

Move the existing FieldEvent enum to the existing revision owner with crate-private
visibility. The paragraph-event adapter uses that owner's existing XmlTree and
source envelope. Share actual eligible Word run/marker qualification and structural
paragraph domains with carriage inventory. Keep coarse Owner events restricted to
qualified direct paragraph children. Preserve coarse Begin/End range growth and
the stronger proof's separate/completeness/material predicates. Foreign, unknown,
prior-property and independent textbox paths supply no marker to this paragraph.
Retain actual source order, packed runs, decoded qualified attributes and source
budgets. No second parser, new public API, owner promotion or global validator is
introduced. Retain the actual opaque-marker failure and existing direct-owner,
standalone-refusal, alias and independent-story controls before refreshing the
affected focused source evidence.

Microscope all-pass-7 retains D4 because carriage admission still quantifies
every all-story owner rather than the publisher's actual framed shells. Root
adopts the precise scope correction before remediation. One private XmlTree
publication-frame map is shared by validator and publisher. Its roots are actual
supplied paragraph children of the source envelope. Only qualified modeled
fldSimple/hyperlink children of an existing frame create another frame. Runs,
SDTs and opaque payloads remain atomic. Their descendant shells need no separate
shell-outside-wrapper admission. Nested shells that really are framed keep the
complete-carriage proof. Preserve the independent all-story inventory, source
domains and qualified SDT deletion policy. A body-domain-only filter is
insufficient because an atomic inline control can occupy the same domain.

Reuse the existing boxed source as a validator positive, retain its independent
owner without borrowing a body carrier, and pair it with an atomic inline SDT
outside a complete body carrier. Preserve the standalone direct-field refusal
and empty-input positive. Distinguish the source-established review from any
actual new pre-repair RED. No public owner admission or new file is introduced.

## D5 typed field excerpt context

Microscope all-pass-8 identifies a separate public discovery failure after the
frame-scope correction. Both current Box and atomic-SDT comparisons have valid
input main parts, but typed field discovery reparses a first-run to last-run
slice with unmatched revision-wrapper boundaries. The enclosing owner has already
been scanned. Exact runtime slice offsets were not captured and are not inferred
as observations. The SDT failed comparison preserves its original bytes. Box
atomicity was not measured. Both public packets remain NONPASS.

Root approves this bounded technical correction before code. The existing
enclosing-run scan carries its range and concrete immediate physical parent start
from the current XmlElementFrame stack. A private concrete run-source value serves
both Start and Empty marker callers. ComplexFieldScan retains the begin run's
parent identity. A complete field supplies an isolated typed candidate only when
its ending run has the same physical parent. Equal attributes, IDs or namespace
maps do not establish that identity. The existing typed probe still decides
admission within that context. Cross-parent source remains preserved and untyped
until its actual owning wrappers are resolved.

Keep intrinsic complete-field validity and enclosing-field propagation separate
from this excerpt eligibility. Preserve template field ranges, nested admission
rules, source budgets and real source parse errors. Do not catch arbitrary errors
as false, rescan whole paragraphs or add a cache, public API or persisted context.
Retain same-parent direct and revision-owned fields, pair sibling wrappers with
equal metadata but distinct physical parents, and test accepted/rejected discovery.
Add the actual Box and atomic-SDT public pairs to the existing regression carrier,
including checked story snapshots before resolution. Refresh both public pairs,
selected-author lifecycle and complete-mainpart schema evidence after repair.

Use an iterative source owner/event walk with explicit interval/grammar/namespace
frames and one final output buffer. Each actual child and gap is inventoried a
bounded number of times. Do not copy each nested output back through its ancestors.
Compute resolver payload-removal facts bottom-up once from the existing XML tree,
then apply the per-occurrence marker predicate. Reuse current physical/model and
source-copy admission before construction, without lowering or waiving those limits. The existing complex balance counter is not an XML nesting
budget. The HLD's former paragraph-plus-one-hyperlink bound must be reconciled
with the actual admitted nested field path, without inventing a broad wrapper
framework or incidental unsupported owner expansion.

## Shared selected resolution

Compute an ephemeral set of proved retiring simple-shell XML indices once from
the existing XmlTree. Apply the same private preparation at BOTH actual render
entry boundaries: public main Document::resolve_revisions and
render_with_source_ranges for related stories and private transport. No XML range splicing or
second persistent model is needed. The normal renderer owns wrapper removal,
namespace transport and paragraph-mark merging. A retiring simple shell invokes
validate_selected_descendants before disappearing so counts remain actual selected
revision element counts. Retained source-range targets cannot fabricate a
missing destination.

For each originally complete, positively proved carrier, evaluate actual selected
removing revision ancestors and selected structural owner markers through the
same scope.matches/removes_content policy used by rendering. Watch only originally proved complete D3 carriers containing actual simple
shells, plus the complete nested occurrences needed for that proof. Track their
own begin/separate/end independently from complete nested carriers. Removing a
paragraph mark alone does not remove marker runs. Use actual owner-removal rules.
Permit either:

- All own structural markers survive. Check any affected complete nested carrier
  under the same rule. Cache-only selection may leave an empty simple result,
  preserving its shell and form.
- All own markers are removed and every owned material payload is selected for
  removal. Only this case grants contained simple-shell retirement.

Mixed own-marker survival, or removal of all own markers while any material
payload remains, refuses atomically before publication. Author/time equality does
not select sibling revisions. No cascade, all-orphan sweep or generic empty-cache
pruning is introduced. Already unmodeled/incomplete imported sources are outside
this new complete-carriage proof, not a new global validation target. The exact integration reuses actual source roles and both existing renderer
construction sites. This is a local D3 preflight, not global field validation.

A direct simple field beneath a qualified modeled hyperlink uses its real schema
owner chain. The existing opaque/prior-property/drawing barriers remain. Do not
make arbitrary similarly named descendants modeled.

## Existing-carrier evidence

Retain the exact existing TOC whole-field insertion/deletion, both Run/Word,
with complete candidate and resolved schema checks. Extend the same existing
regression body, without another test binary. Add bounded discriminators for
unknown root metadata/fldData/lexical flags, aliases and foreign w shadowing,
cache-only empty result, boundary-ID subset refusal with unchanged bytes and
counts, unselected/prior custody, hyperlink-contained simple field, final paragraph
merge, fixed-ID move and block-control composition. Use risk-based caller coverage,
not their full Cartesian product. Public predecessor RED is retained at
fx272-d3-public-diagnostic-prepared-v1. Actual native saved-artifact observations
remain due while native control is unavailable. The user chose to wait for its
recovery and declined the AppleScript fallback. Neither modeled round trip nor this
plan is native/schema acceptance.


The finite partial-boundary supplement is source authority for the marker table,
both render entry points and iterative work closure. Its report SHA256is
1ea8f4ff3e034970382605d4018249956e2dfd1909a4edb0dc3aaad76deaac97. Root has read
and authenticated it. Include complete nested-complex removal versus partial
nested-boundary refusal, two nested simple shells including a control, nested
hyperlink across an intervening field refusal, an admitted deeper source and its
existing over-budget control. A bounded visit counter can prove one-pass work.
No arbitrary depth cap or new public model is authorized. Root adopts this technical amendment on2026-10-30 after actual source/schema
closure and the canonical writer handback. Missing runtime/native evidence remains
pending and no fix or completion is inferred from approval.


### D3 source revision projection boundary

A direct fldSimple under hyperlink remains opaque to the existing lower hyperlink
Field model. This repair adds no typed Field API or lower model/admission change.
The known qualified XML revision traversal through its actual result grammar feeds
the existing shared XmlTree source scanner. Source story_revisions, source counts
and source-range resolution use that same metadata and selection policy. The
borrowed modeled main-document revisions projection retains its existing domain.
Hyperlink-opaque controls assert raw shell/instruction/cache XML, source listing
and resolved counts, and unchanged lower opaque disposition. Typed Field form/cache
assertions apply to the already-modeled direct D3 cases. Preserve unknown children,
prior properties and drawing barriers. Existing source-copy/model admission remains
unchanged. No new public method or general arbitrary subtree editing is authorized.


## Current CI imported-program and ordinary identity boundary

Root adopts this technical clarification of existing A3/A6, C1/C5and the retained
raw/current-cache contracts after actual B7Python failures. It does not change
Q-S96-01/Q-S96-02, implement pending Issue430, add a public API or broaden the
held-revision policy. The complete15-failure diagnostic and independent ordinary
projection/imported-instruction assessments remain under the existing TEMP root.

An unchanged imported unsupported instruction is preserved source, not newly
authored program grammar. The existing Field validation owner may distinguish
actual Parsed source with current raw equal to immutable original raw, existing
instruction_structure_eq and instruction_source_identity_eq, matching the writer's
existing source predicate. These helpers alone do not check nested raw edits.
Validate every current nested instruction child through its own qualified boundary.
Always retain XML-character, model/namespace, ownership/resource/work-budget,
current/new typed-cache, result and legacy kind/value checks. Skip only strict
authoring grammar for the exact unchanged imported program. No token-alphabet
expansion, formula evaluation, name-specific exception or catch-all error fallback
is authorized. Preserve original raw XML and all source owners.

Keep New and edited instruction paths strict, including checked from_raw,
from_instruction and FieldInstruction::new. A newly authored parent must not
accidentally inherit its parsed child's preservation permission. Review all
existing authoring callers and add the actual new-parent/parsed-opaque-child
discriminator. A small private borrowed grammar factor is permitted if required
to retain one policy owner, without a new public method/type/flag or cross-crate
API. Preserve admission before clones, finite work and existing charge rules.

Concrete consumers are Field::validate_for_attachment, ordered-result
normalization/publication, facade preflight_current_field/prepared_current_field
and the existing engine field projection. The imported malformed quoted PAGE
reaches existing unsupported display/cache diagnostics rather than failing the
authoring preflight. The legacy FORMTEXT literal setter changes only the requested
result. Its instruction-region calculated child remains exact, its old result
child is removed only where selected, ffData defaults and bookmark/marker order
remain exact, and held/fresh/reopened values retain their established semantics.
F-X274unsupported-source/native requirements remain separately deciding.

Ordinary clean-comparison proofs must use precise namespace-qualified producer
identity policy for their actual owner: the five CT_PWord rsids, two existing
w14paragraph identities, the three CT_RWord rsids and the current tridentity set.
Factor the existing row filter into one private owner-specific root-fragment
projection, preserving all inner bytes. Reuse it across ordinary p/r/trroot
shells and simple/complex field physical-run and shared-root proof paths. Do not
use the unrelated all-attribute run-root deduplication helper as an identity rule.
Actual source/emitted XML, lower protected_source_residual and HeldRevisions
remain unchanged and strict. Preserve genuine unknown Word/foreign attributes,
alias/shadowed namespace distinctions, opaque elements/events and all scopes on
surviving residual facts. No generic unused-binding/default-namespace pruning
algorithm is authorized.

The exact Issue158first replacement preserves first-title paragraph bytes but
changes inherited default namespace declarations. Its measured guard item0diff
is only incidental scope on p/rproducer-identity residuals. This is not a changed
rsidvalue, and it is not permission to ignore namespace context generally. The
lower bridge emits that known-root scope only with surviving residual attributes,
so excluding ordinary identity-only facts can remove their accidental scope
carrier. Validate that actual case and the complete fresh-binding workflow, not
only a synthetic identity matrix. Original input SHA256is
d05f9c753c00eb804c6e345126ef7a1f7a4fc635d2c9b653b829922030cd875e.

Before source repair, capture actual pre-repair semantic RED for the new lower
discriminators. Cover exact source/raw/structure/child identity, nested raw-only
edits, source-owner substitution, strict new parent, XML-character and legacy
refusals, real stored SVG and cache diagnostics, all18identity matrix rows,
qualified aliases/foreign lookalikes and unchanged held revisions. Batch relevant
cases in existing test carriers. Then acquire a fresh successful installed Python
artifact and run every affected existing field/identity row plus the full
Issue158workflow with the authenticated original fixture. No old wheel qualifies
current Rust. Native, schema, full suite, archives, CLI and final risk riders
remain pending for the integrated source. HLD03/10remain the impact list.

Contract closure confirms existing callable APIs, current preserved-source versus
new-authoring roles, ordinary versus held proof consumers, precise source identity,
atomic refusals and local/final evidence. No new public signature, unsupported
XML creation or unresolved user product choice is introduced. Root grants an
implementation lease only after the current canonical writer and review phases
return control.


## Fresh raw XML authoring and retained field source

Actual DAE MSRVjob124423481181fails the existing F211atomic refusal test when
new replace_paragraph_xml input contains an unclosed quoted REF program. This
is fresh raw authoring, not retained imported display. Root adopts independent
designer reassessment REPORT SHA256
759e71d5c5f6457daf2437918363ddffe149550bdf76a6c27e2d53c34401001a
under existing A3/A6and C1/C5. The original F211plan known-profile and identical
owner contract, and existing HLD04known-profile/source-route clauses, decide
this technical distinction. No new user product choice or public API is added.

After its existing exact whole-owner no-op, checked raw replacement supplies a
new whole fragment. Its modeled fields require strict authoring program grammar
through the existing private validate_element_owner_relationships. All5raw
publishers share that owner: paragraph, run, table, cell and optimized plain
paragraph. Its existing nested/control/revision/ruby/result traversal remains
authoritative. Parsing that new input as FieldSource::Parsed is not retained
original-document authority. An unchanged-looking malformed field in a changed
whole fragment, including an unrelated content-control edit, refuses atomically.
An exact captured same-physical-owner replacement remains an unchanged no-op.
Do not infer partial field retention from spelling, position or copied XML.

Opening/read/render and existing appropriate checked typed cache/result edits
of actual retained fields keep the prior imported-program exception and current
selected-program/resource/model/source checks. Do not globally tighten lower
CT_*fragment parsing, which also serves retained template/comparison/Python
preparation. Unknown/foreign/MC opaque placement and held revision custody remain
unchanged. No source instruction, result, namespace or retained byte is rewritten
merely to validate fresh input.

Use existing checked authoring constructors for strict raw spelling and actual
structured nested operands at the one private fresh-owner boundary, after full
model/work admission. Candidate composition of Field::from_raw and
FieldInstruction::new must first qualify an actual valid nested-complex/switch
control whose flat raw spelling omits its nested operands. If composition
refuses a supported valid control, stop for designer reassessment before a
production patch. Do not duplicate quote/token lexing, add public methods/types
or expand grammar. Quoted initial names, invalid raw switch tokens, malformed
PAGE/REF and nested programs must refuse as well as unbalanced quotes.

Preserve exact old F211note/section/refusal assertions and add risk-based valid
complex/simple, alias/default/foreign, related-story, cell/table and inline/result
controls in existing bodies. Every refusal preserves bytes, location/generation,
relationships, prior revisions and held handles. Reuse exact original retained
PAGE/legacy, New/opaque child, retirement/source identity and full Issue158
controls. Acquire fresh normal CLI/wheel and all49prior affected Python cases
plus real raw binding cases after any facade source fix. Update only current
HLD03/10listed intent. HLD04already states the deciding contract and is not edited.
Actual schema/native/full/typing/WASM/archive and final hosted gates remain due.
Q-S96-04 representation is answered for hidden, zero-width STANDARD Word
bookmark custody. Normal-note identity, schema placement, custody and
cross-document collisions, selected-resolution package cleanup and native Word
retention remain open. No Issue430 implementation is authorized before that
closure.


## Qualified fresh paragraph publication custody

CI-RAW-NS is a separate preexisting namespace-publication defect exposed by the
required positive controls. Current lower checked parsing models canonical,
aliased and default-namespace PAGE fields with cached7correctly. Fresh facade
paragraph reconstruction drops incoming root declarations needed by surviving
field QName bytes. The prior isolated wheel reproduces the same loss. Root
adopts independent closure REPORT SHA256
4d4c14a1d5d9d6bf6b9fc461378c97bec1128f301767b9751fcaebfa66f57b00
before a publication patch. This is source-custody closure under existing
A3/A6and F211/HLD04contracts, not new authoring permission or Q04approval.

Keep the parsed paragraph for every existing known-profile/relationship guard.
Its model is not mutated between those checks and the current reconstruction.
Publish the original already-validated supplied qualified paragraph XML through
the existing ContentFragment transaction, matching current table/cell peers.
Retain exact-owner early no-op, model/resource admission, original/incoming
physical context/placement, fresh strict grammar, identity change, section
profile, relationship/story ownership, actual namespace closure, comment
reconciliation, candidate serialization/reopen and atomic commit/error order.
No required model transformation is bypassed. The actual emitted payload must
carry its incoming QName context, not only a temporary profile-comparison buffer.

Do not hoist all incoming bindings onto a generated canonical w:p, since a
qualified Word root can locally bind w to a foreign namespace. Do not rewrite
field prefixes, unknown raw XML, immutable field source IDs or lower parser
admission. Original qualified root declarations preserve default-element versus
qualified-attribute meaning and local shadowing. Existing supported opaque/MC
cases preserve physical source semantics, while existing unsupported routes
retain atomic refusal. No namespace may leak across sibling scopes.

Incoming fields fail the optimized plain paragraph route before assignment,
so retain that exact gate and fallback. A run edit validates its fresh RUN,
not unrelated retained malformed field siblings in the original paragraph.
Keep the distinct changed-whole-PARAGRAPH strictness and exact captured no-op.
Require actual run replacement beside retained alias/default fields plus a
run-local opaque binding case, table/cell and related-story controls before
claiming all publisher coverage. Do not presume those distinct routes broken.

Preserve the current failed qualified-positive assertion. Extend existing
bodies with canonical/q/default simple cache, nested result/instruction/switch,
qualified emitted owner, typed cache update and save/reopen, foreign w/default
shadow, unqualified attribute distinction, MC namespace values, opaque sibling,
prior revision/bookmark/comment/relationship and held generation controls.
If another source-bound case fails after this minimal correction, stop for
independent publisher reassessment rather than general QName normalization.

Current HLD03/10impact remains unchanged. HLD04already states the deciding
contract and is not edited. Rebuild final coherent Rust evidence, normal CLI
and installed wheel after this actual production change. Earlier40/41and old
wheel observations stay attributed to their original producers. Fresh prior49
plus real raw binding cases and all final native/schema/full/hosted/typing/WASM/
archive requirements remain pending. No baseline re-recording is authorized.


## Current writer field namespace custody

The actual qualified-source diagnostic establishes a surviving current model
writer defect. Physical saved XML remains correctly bound and cached7after
the initial facade payload correction, while virtual main serialization already
emits an unbound q field. Its namespace-complete snapshot cannot recover a
binding absent from that current source. Root adopts independent closure
REPORT SHA256b44400594961ebce82d6f5d683b9e41f5d2756ad2ea1aea1467b0ffa8defabd7
before a second namespace production attempt. Do not substitute pristine
package bytes for current source authority or weaken the checked snapshot.

Separate the two real existing writer roles. Inherited-context source-relative
patches retain exact unchanged raw original/current pairs, immutable source and
owner IDs, cached child associations and shared span geometry. Actual current
model publication closes required namespaces on every actual surviving physical
field root under its real containing context. No FieldSource mutation, new
public API/type, persistent field state, parser mode, trait, generic, feature
flag, QName rewrite or arbitrary first-field association is authorized.

Implement in existing text.rs using concrete private publication-frontier
helpers and existing scoped binding-selection logic. Explicitly cover CT_P
untouched shared-span raw output, direct/span fields, ordered/current field
returns and CT_Rmixed/cached-result publication. Preserve write_field source
patch consumers including source_replacement, detached field/source pairs,
value comparisons and nested source updates. Determine each caller's actual
origin rather than changing all calls because their names suggest publication.
New/rebuilt instruction writers keep current canonical authored behavior.

A complex field can span multiple sibling physical run roots. Preflight and
close each actual top-level root, retaining exact nested bytes and inter-root
whitespace/comments/PIs. Use that field's proved source context and original
local shadowing. Namespace closure adds only necessary declarations. Foreign
w and other conflicting bindings belong on the retained physical source root,
never on generated canonical paragraph/run ancestors. Omit canonical w=Word
only where the actual containing publication context supplies it, including
existing foreign override handling. Default elements do not qualify attributes.

After full model admission, reuse the unchanged normalization_work_budget as
a noncircular publication envelope. Borrowed preflight computes raw length and
exact escaped declaration growth over every physical root with checked
arithmetic. Precompute escaped context lengths once. Refuse overflow or
over-envelope expansion before allocating closed roots, cloning long URI
strings per root or growing an accumulated output. Then emit with one bounded
second walk/buffer. Include actual scan/emission work and existing model/XML
limits. Do not derive a larger budget from required expansion or add a new
higher limit. Existing selection logic may be privately factored into borrowed
indices/length facts while keeping other callers' exact behavior.

Keep the actual snapshot unbound-q RED, plus w/q/default simple/complex/multi-
root/shared-span/nested instruction/switch/cache/ordered-result and typed-update
controls. Prove physical package, virtual writer and checked snapshot semantics
separately. Retain strict fresh malformed refusals, imported PAGE/legacy and
source-origin/no-op controls. Add foreign/local/default/MC/opaque scope pairs,
unchanged source_replacement bytes/IDs, canonical no-needed-declaration byte
identity, smaller admitted long-URI/root family and an actual over-envelope
family refused before output growth. Test work observations must establish
preallocation bounds, not only final namespace counts.

Once common writer closure passes the exact original qualified paragraph
controls through the earlier serialized-payload route, revert the intermediate
original-input payload patch if no independent contract requires it. Prefer
the smallest final diff. Preserve every failed/successful source identity and
never claim the facade-only attempt closed this fault class. If another source
consumer violates custody, stop for independent publisher reassessment before
a third namespace patch.

HLD impact now includes03/04/10before any HLD04edit. Describe implemented current
writer versus source-relative custody and output admission there, not chronology.
Reacquire final affected lower/facade Rust, strict lint, current normal CLI and
installed Python49prior plus raw binding cases. Lower/facade archives now need
affected genuine refresh. Actual49hash/7pixel, source-qualified hosted/native,
schema/full/typing/WASM and all final acceptance remain pending. No baseline
update or Q04ownership implementation is authorized by this clarification.


## Initial immutable field parent scope

Before any shared-writer production patch, actual hyperlink control establishes
missing original scope. Projection A binds a to Word and q to urn:outer, while
its original hyperlink B binds q to Word. Parsed complex PAGE/cache7stores A.
Changing only current mutable hyperlink q to urn:changed retargets field roots.
Current attributes therefore cannot recover original authority. Root adopts
initial-capture closure REPORT SHA256
25452be849f5d6092dc908aba6a1923846c464c7029bdc57c591fb3e2d9bcd64
as a prerequisite to the second namespace production attempt.

At the INITIAL complex Field::parsed construction, select original hyperlink
parent B only when its immutable [first,last) run interval contains the whole
[start,end] field. Derive B from original hyperlink_sources root plus projection
A through existing word_prefixes_at, matching the existing event scanner.
Otherwise select A for direct spans or whole wrapper roots included in raw
source. Never select marker/run-local declarations as the common parent.
Those remain local in exact physical source bytes. Initial correct capture
into existing word_prefixes is allowed, not mutation of already parsed source,
new retained state or authority from mutable HyperlinkSpan attributes.

The final grouped shared span must derive the same proved common parent from
its final interval and original wrappers. Use it for field_span_runs and bare/
part source scans. Every actual span-bearing member retains the corresponding
initial immutable context. prepare_complete_complex_result raw scans must use
the captured field source scope while topology checks keep original projection
maps. No arbitrary nested/simple first-field source is accepted as run context.

Existing partial-crossing admission stays unchanged. has_typed_boundary_inside
declines overlap not containing the whole field. Existing owned-result geometry
allows only a whole result wrapper or a wrapper containing the whole field,
with recursive cached-child checks. Do not silently add fallback admission,
refuse an admitted format or invent per-root source state for already declined
partial crossings. Original raw XML, source/owner IDs, span geometry and
source-relative unchanged pairs remain exact.

Reuse existing reader/XML admission before parsing. Select or borrow original
scope, check prospective byte summation and existing accounting before copying
its strings once into the existing initial slot. Existing model/normalization
admission must observe correct B before subsequent field clones, result
normalization or publication buffers. Do not check A then claim B admitted.
Keep the unchanged noncircular publication envelope and early expansion refusal.
No new higher budget, public API, parser mode or persistent field state.

Retain the actual original-B/current-C mutation control. Assert captured B,
source IDs/raw pairs unchanged, current field/cache/expanded identity and
checked source parse after publication or atomic refusal. Pair direct shared
run/marker-local shadow, wholly contained hyperlink/shared two fields, outer
field with whole result hyperlink and nested field, and existing partial-boundary
decline including owned-result geometry. Continue all canonical, default,
source-relative, typed cache, fresh malformed, scope/opaque and allocation
controls from the previous closure. HLD03/04/10impact remains unchanged.

## Source-associated public cache publication

Root adopts independent mapping closure REPORT SHA256
062e303a5dbd24f52f153cc1cb97b619d5c49e038c79f509d6f0958dfa71d922,
following actual six-case public PAGE diagnostics. All six updates return1.
Only canonical input without opaque metadata saves. Alias/default cases and
all three inputs with an unchanged namespace-dependent opaque sibling refuse
later save. These are successful mutations followed by save refusal, not
atomic update refusals. Preserve their actual raw evidence and private cache8
NONPASS. Added field dirty structure and changed cache markers invalidate old
retained namespace-owner correspondence. Recognition-only repair is insufficient.

Extend the existing private shared cache updater in field.rs and its document.rs
source helpers. Keep existing owner replay structure, exact raw markers,
multiplicity, namespace facts, uniqueness and ambiguity guards unchanged.
Acquire current canonical writer coordinates/body spans and the same immutable
model's namespace_restored_main_source. Only restored XML is parsing authority.
Use existing comment_namespace_source_boundaries for exact event endpoints.
Authorized namespace restoration may change expanded namespace facts. Do not
require its returned changed-facts list empty or use pristine package bytes.

Prove every original direct body item's context-correct body_item_xml equals
its canonical child span. Parse restored main XML through existing
legacy_story_paragraphs(Document). Match complete Root/Table/Row/Cell/SDT
container paths and all physical paragraph records to existing typed traversal.
For each admitted nested paragraph, one temporary paragraph clone emitted by
the original document's body_item_xml gives the same root-context writer bytes.
Inverse-map the restored paragraph boundaries and require exact equality with
those canonical bytes. A flat ordinal/count, first matching substring or bare
CT_P writer without root context is insufficient. Preserve raw wrapper and
foreign Word-p decoy exclusions and exact opaque slot positions.

Match accepted field instruction structure, nesting, inherited locks, ordering
and physical span groups before consuming updates exactly once. Derive actual
original/staged pairs from the independently parsed restored source using
paragraph_field_source_replacements and qualified exact-source lookup. Prepare
all nonoverlapping bounded edits in that SAME source domain, patch once and
install/reparse through set_story_source_xml before any dirty whole-model
assignment. Do not overwrite its resulting model afterward or rediscover a
different canonical source for those pairs. Existing related/textbox domains,
copying, relationships, signatures, invalidation and outer transactions remain.

Full existing model/XML admission precedes clones and emission. Bound temporary
work by admitted owners and current source, reuse existing accounting and field
budgets, and preflight ordered ranges and checked output length before growth.
No new public API, persistent state, higher limit, parser mode, file or module.
Also resolve the paused lower writer's prospective actual-B scope admission
before copying and indexed original hyperlink context selection before freeze.

The private attached Document model setter bypasses source transactions and is
not a public cache-authoring route. Preserve its failed diagnostic unchanged.
Replace that fixture route explicitly, retaining every arbitrary cache8 row:
parse the current namespace-complete snapshot into a detached owner, perform
the checked setter, obtain the OUTER field source_replacement pair, qualify and
patch its exact occurrence into the SAME snapshot bytes, then call public
replace_paragraph_xml and save/reopen8. Preserve all surrounding declarations
and opaque bytes. This is strict fresh authoring and may not launder retained
malformed programs. Lower source-pair/cache8 controls remain. Genuine public
update_fields/update_page_fields must prove actual supported value changes,
while PRODUCER KeepStored proves retention only.

Require the closure report's finite matrix, including all six public RED
successors, simple/complex fields, ordinary evaluator updates, shared/nested
fields, controls/tables/revisions/related/textbox ordering, identical paragraphs,
legal versus raw/foreign Word-p, namespace-dependent opaque siblings and decoys,
root binding/drawing context, exact source identity, smaller/over-envelope work
and late atomic refusal. Any actual admitted input contradicting this mapping
stops implementation for independent reassessment, never a new format refusal
or weaker matcher. HLD03/04/10must describe the implemented transaction before
fresh scoped evidence and independent microscope. Final artifacts, archives,
hosted/native/full/schema/hash/pixel/typing/WASM and Q04remain pending.

## Complete retained-owner publication context

Root adopts control closure9a2c3ae92d6b68bfa06a577df4db3f272469d2f8fb02a14cd165c5a8658160cf,
frozen reconciliationd23f7a6082a32349da19045a22f6c8fe2453ebc8fb27c7bc9db3b7d97ac611e1,
Ruby capture amendmente3c1e0b8c9d1d75f5e5bce07602344aa323f4254fd1c0a2f9ddf5b7f6e20548e
and source isolation48992de3f79278c05b307cc7b2de0ebd6a2bdef6bf809255c24ae57c35f01671.
The source-bound152-member probe packet proves alias/default revision loss,
control root/property/content Word identity loss with foreign-w opaque attrs,
and namespace loss in the existing modeled direct Ruby route. Checked XML
parse success alone does not prove preserved expanded names or typed owners.
Canonical controls and math controls pass. Preserve every actual result.

Existing original control root/content/property/type maps and revision parent
maps are authority for their own raw payloads. Do not borrow a Field scope or
mutable current parent to recover them. Separate full publication from source
projection throughout existing control/property/inline/nested methods. Cover
RootSlot raw, PropertySlot raw and retained type metadata, content raw/revisions,
unchanged inline raw shortcuts and changed fields. Use the common borrowed
namespace root selector with original local overrides and actual destination.
Revisions close from their own captured parent at every full-output caller,
including paragraph/hyperlink/control and property-marker tails. Raw getters,
source refresh and immutable Field source pairs remain source-relative.

Extend the EXISTING private RootBindingScope transport with actual emitted
namespace facts for the duration of serialization. This is not model state,
a global URI cache, a source-history registry or public API. Part/body-item
entrypoints replace an unrelated outer root. Actual emitted local declarations
on paragraph/run/hyperlink/table/row/cell/control/property/content/type owners
overlay that root, and guards restore the caller on return/error/unwind. Borrow
URI comparisons, share one root frame within that operation and keep local
overlays bounded by admitted owner depth. Account root-frame work at each
independent serialization, without claiming URI sharing across unrelated calls.
Recompute old r/mc/wp guarantees from the effective emitted map, never OR them
through a local foreign shadow. Late w14 is guaranteed only where the actual
same-part writer supplies its documented on-use declaration. Unknown standalone
and source wrappers cannot borrow that promise.

Actual matching global bindings suppress redundant declaration growth. Changed,
removed or locally shadowed globals require original meaning at the surviving
raw root. Include long-URI global positive with zero added growth, changed/local
shadow negative, standalone and reentrant independent-part controls. Never
manufacture an over-envelope refusal by omitting a known actual parent binding.

At the observed control/property/content root collision, preserve opaque attr
QName bytes and URI. Explicitly allow a bounded deterministic qualifier for a
GENERATED known Word root when canonical w is locked to a foreign opaque attr.
Prefer a nonconflicting original Word alias/default, with bounded fresh known
qualifier only if necessary. Generated modeled attrs use the appropriate Word
qualifier. Generated descendants restore canonical Word at their own known root
or use the selected known Word qualifier. No generic raw QName rewrite, foreign
w hoist onto a canonical generated root or blanket namespace refusal is allowed.
Retained raw descendants close from their own original frames. Keep schema slot
order, unchanged canonical output and existing facade atomic refusals.

Source field write_field/write_run_field None paths, detached mixed spans,
value comparisons, ordered temporary paragraphs, nested source edits,
with_run_wrapper_content and the ENTIRE inline control source projection use
explicit source role with fresh/cleared actual frame and old guarantee bits.
They cannot consume enclosing publication omission facts. Fixed validation
wrappers establish their OWN actual FIELD_CACHE_VALIDATION_SCOPE. Restore the
outer publication context before final publication closure. Require identical
source bytes/IDs/pairs outside versus inside unrelated x=A/x=B/local shadows,
plus generated canonical-r relationships under original foreign-r source,
error/unwind and publication-source-publication nesting.

Ruby has no retained original frame and its public structs remain unchanged.
Root explicitly selects initial context completion of retained raw fragments
while the actual Ruby-root or Ruby-property-root frame still exists. Add ONLY
needed xmlns attrs at each fragment root, preserving every unknown QName,
existing attr byte and descendant payload. This intentionally changes public
raw_xml bytes where inherited context would be lost. It is a namespace semantic
bug fix for the prepared Word0.17minor, not a new struct field or hidden protocol.
Canonical w=Word remains the guaranteed Ruby publication context, so canonical
w fragments and already-local bindings stay byte-identical. The fragments are
context-complete for that contract, not claimed literally standalone without w.
q/default/foreign-w dependencies complete locally regardless of current global
bindings. The existing typed/fallback decision and placement stay unchanged.
No schema/native acceptance or new Ruby placement is inferred from its legacy
modeled direct paragraph route. Repeated post-read output does not mutate the
model. Field original raw bytes/IDs/unchanged pairs remain exact, while changed
fields may serialize the intentionally completed retained Ruby fragments.

Controls have no existing normalization byte envelope. Privately factor/reuse
the UNCHANGED field affected-source work policy for control/revision/Ruby raw
bytes, original scope strings and actual modeled scalar/attr data. Do not create
a dummy Field, increase a limit or derive allowance from required expansion.
Before the first expanded fragment, preflight TOTAL root/declaration/escaping
growth per owner using original bytes and borrowed context. Do not recycle
expanded property bytes as outer Ruby-tail source allowance. Existing raw
capture allocations remain real and are not described as allocation-free.
Smaller and genuinely over-envelope many-root/long-URI cases must observe
construction work, with zero expanded constructions before refusal.

Implementation may edit existing text/content_control/revision/ruby/document/
table/header_footer/footnotes/comments/paragraph_properties/run_properties/
numbering/styles/glossary files
inside rdocx-oxml ONLY for these concrete roles, frames, guards and caller
overlays. Existing facade field/document and current test carriers remain in
scope. No new crate/module/file, public type/signature, persistent model state
or unrelated rendering/grammar change. HLD03/04/10must describe current intent
before freeze. Retain all source mapping, strict fresh-authoring, private-route
fixture reconciliation and final acceptance obligations from preceding sections.

Before review, qualify actual revision kind/count/id/author/cache, control and
raw property/attr expanded identities, all original field source associations,
changed and unchanged field/control/revision roles, alias/default/local foreign
contexts, independent Ruby q-meaning pair, fallback preservation, actual global
context and source isolation, all six public saves and15cache8 authoring rows,
the corrected paragraph getter and source-qualified SVG. Then reacquire the
coherent affected Rust/lint/CLI/installed-wheel evidence. Any further actual
publisher contradiction stops for independent reassessment. Math and Ruby
placement are not new work. No baseline update, status completion, Q04ownership,
native fallback, source push or publication is authorized by this section.

Root also adopts source known-root qualification amendment SHA256
edcb9a444fd017365096455848806ea71713bc60f0b78b73da1da6a522fafb02.
Generated known Word roots must be namespace-correct in BOTH roles. Source
control projection uses its original contractual content context, choosing
nonconflicting Word qualifiers for generated containing nodes when needed to
keep foreign-w raw descendants under their original bindings. It must not
insert a canonical-w boundary that retargets untouched source snippets. This
explicitly supersedes the earlier blanket source-no-qualifier wording only
for namespace correctness of constructed known nodes. New raw-root closure
and actual global omission remain publication-only. Unchanged FieldSource
bytes/IDs/pairs and unaffected canonical source bytes remain exact.

Concrete private role/core arguments select source versus publication. Existing
public/current serialization defaults to publication, including reentrant
calls. The scoped carrier transports namespace facts, not global source intent.
Source resets cannot turn an unrelated public callback into source output.
Whole-part root providers that currently reach field/control/revision writers
also include glossary and styles/numbering properties. Give each its actual
root replacement. No speculative context hooks are added to unrelated writers.

## Writer-time ancestry and complete namespace capture amendment

Root adopts the independently closed ancestor coordination design732edb96dd6dcdaf8ba4d8b7e9a4759738c8c55397b5ee7aa8c0fa728cd868ab
and default-knowledge design317c28cb05ec401dab7a92e869b617d2041f31d764a7eec64cf55714ba5680ac.
Root authenticated all74members of the actual public ancestor diagnostic,
reportdcaf62694643ab2a76bec8e5b4c935f75c8ed9ded44ed32fe4f605863c03fbcb,
and all108members of the actual Ruby default probe,
report5dc7a296651cd0b5e41bc41de936673bc79d56be0d42360ed822fac26bfcce69.
The earlier implementation pause's154members are also authenticated. Preserve
these failures and their original source identities. They are qualification
inputs, not current acceptance or a reason to weaken the exact source guard.

The first55-child batch stopped at its first child on an incorrect new w14
count expectation. Root's earlier claim about old MC-value declaration omission
was wrong. Preserve the old paragraph value_prefixes exception and unchanged
run lexical replay. The specific MC-only paragraph plus run fixture retains
two local declarations without an explicit global, or root plus both locals
with one. Assert actual expanded meanings and per-owner canonical placement.
Do not blindly replace unrelated field raw-root count assertions. Actual
QName-backed late w14 guarantees remain separate from ordinary actual bindings.

The actual public ancestor diagnostic proves a supported table-local x context
updates with a refusal while canonical and matching document-global controls
update, save and reopen. Whole-part emission inherits table x, while the
part-root-only clone adds a declaration. Keep exact canonical paragraph byte
equality, complete paths, body certificates and field inventories intact.

Explicit exception to the preceding no-new-public-signature restriction: add
one compatible doc-hidden coordination method on existing CT_Document:

```rust
pub fn body_item_xml_with_ancestor_declarations(
    &self,
    item: &BodyContent,
    ancestors: &[&[(String, String)]],
) -> Result<Vec<u8>>
```

Pairs contain prefix and decoded URI, ordered outer to inner. Existing public
body_item_xml behavior and signatures remain unchanged. The concrete consumer
is facade updated_main_field_source. Reuse the same item admission, original
model root_start, original omission policy and conditional late-w14 guarantee,
then overlay the supplied actual non-root local declarations before the SAME
body-item emitter. No public type, feature, trait, generic, model state or new
file. This additive Rust coordination surface is compatible with the prepared
Word0.17minor. No Python or WASM authoring export is added.

Collect actual local ancestor declaration frames in one lexical scan of the
CURRENT canonical writer bytes. Reuse/factor existing story namespace decoding,
pool each decoded frame once and retain indices for requested P starts. Snapshot
BEFORE each target P event, excluding its own declarations and document root.
Include actual non-root body/table/row/cell/control/content ancestors. The
helper reconstructs the original writer-time model root, so never flatten or
import the final canonical root. A later finalizer may have added w14 because
of another paragraph. Canonical bytes supply emitter coordinates only. Restored
current source remains typed authority. Keep same-domain patch/reparse before
installation, all update consumption and exact refusal atomicity. Admit and
account actual source/frame/paragraph work before guard copies under the
unchanged private work envelope. No whole-part serialization or rescan per P.

The12frozen Ruby property probes exit0 for typed counts, immutable vectors and
ordinary sibling Field IDs/source pairs, while independent expanded-name checks
find5failures and7preserved controls. No-default raw children consume a future
default. Property-root default clearing and foreign-default declarations are
also lost. These observations do not qualify tails, changed fields containing
Ruby, schema or native placement. Those remain named proving cases.

Two corrections are required together. Missing DESTINATION default knowledge
is unknown, so remove the generic missing-default-equals-empty exception from
retained_binding_plan. Omit an original empty default only when destination
explicitly contains the real empty binding. At a COMPLETE actual root/frame
replacement, absence of xmlns proves empty default and may install internal
('', '') knowledge without emitting a declaration. Audited complete document,
header/footer, notes, comments, numbering, styles, glossary and fixed validation
envelopes qualify. Child overlays inherit absent declarations, and source
reset/frameNone remain unknown. Ruby capture guarantees canonical w only and
therefore closes required original empty-default roots with xmlns="". Whole-part
known-no-default output retains zero unnecessary declaration growth.

Ruby currently imports the alias-only properties::word_prefixes_at, which
does not update encoded foreign/default capture bindings. Use the EXISTING
numbering::word_prefixes_at full-scope helper for Ruby root/property/child/content
capture. It updates both Word aliases and encoded bindings, including local
clear and foreign shadows. Existing properties recognition/value helpers still
consume bare Word aliases. No new helper or grammar is needed. Completing a
fragment from stale inherited metadata is prohibited even if it parses. Keep
the approved own-original-frame capture contract and unchanged public shapes.

Also restore generated Word binding in CT_P's early empty-element path before
emission, matching the nonempty path under a foreign-w parent. Preserve canonical
empty-P bytes. No unrelated publication or source policy change is authorized.

Qualification adds the public ancestor diagnostic to the coherent existing55
cases, with explicit current membership and outcomes. Cover table/control-local
and conflicting globals, duplicate P, P-own declaration exclusion, local/default
shadows and late w14 caused by another P. Cover Ruby property and tail absent,
empty and foreign defaults, property-local alias rebinding, canonical w-vector
byte compatibility, already-local declarations, immutable repeated output,
changed-field Ruby source publication and unchanged original pairs/IDs. Pair
empty and paired P under foreign-w with canonical controls. Retain total growth
preflight and no-growth positives. Run affected all-target CHECK before the next
coherent build, then the complete named focused batch, frozen-probe successors,
fresh normal CLI and installed wheel cases and remaining scoped gates. Update
HLD03/04/10before freeze. Independent microscope and actual archives/CI/native/
full/hash/raster/typing/WASM gates remain due. The previous zero-public-delta
inventory is historical and must reflect this one additive method at final
freeze. Q04ownership, source push, comments and publication remain separately
held. Root grants implementation only after the owner acknowledges this revision.

## Pass11 D6 to D8 implementation contract

Root read the complete independent all-pass11, SHA256
3f174279b90e30df7049541fa06fa71e3111790c28f894145ac00302db46fb56:
three defects D6/D7/D8, zero smells and zero nitpicks. Source remains the frozen
3857-entry v3 map316b1009eff8697120faa83055616e2b229d4bb1feb9bcf34e92a95fd3e6614f.
Root authenticated all111 paused implementation members,105 ordered-source phase
members and70 valid public duplicate-content probe members. Both56-child batches
stopped at their first child, leaving55unrun. Current CLI/wheel successors were
not acquired. Passing compile/lint/build evidence is not runtime acceptance.

This revision adopts the independent D7 design57161918344dcb18b5a1928067d4e5f5b6fadd1a4c1801659da343f9b237df5c
with21authenticated source pins, and D6/D8 closureddb6b318d3017467731e5d2d642681dfd5277dc32837ecc5b4a0208495e0030a
with its four authenticated sealed members. These refine existing A13 source
preservation and actual parser ownership. No new user product choice remains.
The formal review is complete. Implementation is a separate remediation phase.

### D6, selected text retains its actual Ruby owner

The actual canonical and q/local-Word-w ordered controls start with Ruby1/base,
opaque producerOnly and deletion40/Prior. After scalar NEW, the prepared owner
has Ruby0 before publication. Original source/ID and deletion survive. This is
owner retirement, not a namespace assertion or permitted annotation deletion.

Keep the current scalar first-text-child and typed first-anchor policy owners.
In existing CT_P and CT_Sdt direct ordered-slot replacement, one replacement run
updates the actual run in place through existing replace_accepted_run_segments.
Do not insert/remove a one-to-one replacement. Existing recursive revision/control
routes retain the same owner. This preserves Ruby geometry, hyperlinks, revision
slots and source locator positions without introducing a forwarding wrapper.

For direct CT_P zero/many replacement, move existing Ruby objects out before
the current insertion/removal operation, then restore those same objects with
checked boundary mapping. Replacing slot i by physical count k maps boundary a
to a for a<=i and to a+k-1 for a>i. Check arithmetic, bounds and output length
before mutation/allocation. Preserve all other Ruby fields and opaque/property/
annotation data. Keep existing hyperlink and revision restoration beside it.

If zero logical payload would collapse the last physical base and discard its
excluded Ruby annotation, retain the selected original base as an empty residual
shell. Clear only selected modeled content. Use that actual physical count in
the boundary map. Existing clean non-Ruby slots may retire. Reverse multi-slot
processing must preserve the last original base at its own first anchor, including
a Ruby after an unrelated first typed anchor. Residual getter/run counts retain
A13's existing distinction from requested logical payload count. Do not move old
metadata onto new runs. Do not change global remove_runs behavior, create an empty
Ruby span, broaden placement or add a blanket refusal. Source IDs/pairs, COW clone
independence, resource admission and checked error atomicity remain mandatory.

### D7, explicit original consumer context and content-only source emission

The actual q=Word field with field-local foreign w and plain run plus deletion40
is initially admitted as an ordered owner. normalizeNEW/source pair/publication
then falsely refuse accepted source paths. It contains no Ruby. Canonical and
q/local-Word-w controls complete NEW. Retain the exact path guard, caller pending
scalar semantics and original raw source. No silent malformed output is claimed.

Keep the private XmlWriteRole enum and public serialization defaults unchanged.
Necessary existing private cores receive an explicit borrowed source_context:
Option<&[(&str, &str)]>. Source entries decode their actual immutable owner frame
once after admission. None is no proved consumer and never complete empty-default
knowledge or permission to borrow publication globals. Publication passes None.
Source intent remains an explicit private argument, not TLS mode or model state.
Fresh source boundaries suspend unrelated publication facts and legacy bits.
Descendants use the explicit source context and bounded own-local overlays.
Do not install original source scope as actual publication-root facts.

Factor CT_P's substantial existing content loop into one private content emitter,
including its boundary, hyperlink, Ruby, raw-tail and field bookkeeping. The full
paragraph core writes its root and invokes that same loop. ordered_paragraph_xml
accepts the original consumer and writes the content directly in Source role.
It creates no synthetic P and strips no declarations. A discarded wrapper cannot
supply a binding. Preserve current child order and source fast paths.

Obtain the Parsed Field frame before ordered normalization emits it, and use the
same frame for the existing reparse proof. Revision refresh overlays its original
parent parse prefixes with its retained root declarations before emitting content,
then patches the same shell and reparses under the same original parent. Propagate
the explicit consumer through actual CT_R/property, mixed segment, typed cache,
detached span, nested instruction/source edit, legacy comparison, hyperlink,
table/row/cell/control and boundary callers listed in the independent report.
Each nested Field/control/revision owns its own immutable scope. Source None and
fixed validation envelopes remain separate domains. Public and reentrant public
calls remain Publication and restore outer state on return/error/unwind.

Known Word roots and modeled attributes must resolve to Word in their effective
source consumer without rebinding opaque attributes or untouched descendants.
Keep canonical bytes where its existing canonical contract applies. Under foreign
w, prefer an existing nonempty Word alias, otherwise a deterministic bounded free
rdocxW prefix with a declaration on the actual surviving generated owner. Include
full source bindings and own root metadata in collision/selection decisions.
Default Word does not qualify attributes. All generated descendants must use the
selected qualifier or establish Word at their own safe generated leaf. Source run
metadata/local-binding records retain their original lexical and URI meanings.
Fresh relationship attributes under original foreign r retain the relationships
URI and opaque producer-r data. No generic QName rewrite or global w hoist.

Ruby raw vectors already complete their captured context relative to canonical w.
Its actual generated source root may establish that surviving canonical boundary,
then pass its effective source context into its modeled property/phonetic/base
writers. Existing retained run/root records still restore their own raw consumer
bindings. Do not borrow a stripped P or use this to weaken D6 retention.

with_run_wrapper_content also uses the shared content emitter. Its actual
placeholder::write_wrapper_back callers need original context. Recover operation-
local wrapper contexts through existing WrapperSource/parent links and original
raw root declarations, using the SAME current parser entry convention. Field
links expose actual prefixes. Raw links retain their existing canonical-w parser
convention with actual local overrides. Synthetic CT_P records are not universally
complete. Do not guess unknown globals, add Wrapper fields/markers, expand grammar
or create persistent state. This expressly authorizes the necessary edits to
existing rdocx-oxml/src/placeholder.rs beside the already authorized source files.
No new file/module/crate or second public coordination API is authorized.

Preflight checked TOTAL generated-root/name/declaration/attribute escaping growth
for the whole source composition before expanded construction, including children,
markers, relationship qualifiers and fresh prefixes. Reuse the SAME current private
work envelope from original source/scope/current model facts. Borrow URI data and
bound local overlays by admitted owners/depth. No per-child full URI copying,
expansion-derived allowance or raised limit. The post-write length guard remains
a backstop. Require zero expanded constructions on a genuine over-budget refusal.
Actual publication root/ancestor/default facts and old r/mc/wp/w14 policy stay.

### D8, source traversal follows the selected typed slot

The qualified normal public probe imports/saves/reopens single, empty-duplicate
and nonempty-duplicate controls with one PAGE7. Single/empty update1 and reopen1.
Nonempty raw duplicate causes inventory refusal, with ZIP and package parts exact.
This is a current imported-source counterexample, not old runtime or schema/native
acceptance. Preserve raw duplicate bytes and exact inventory/path/byte guards.

Mirror the ACTUAL typed parser: only its first QUALIFIED START sdtContent reserves
the typed slot. EventEmpty stays raw and does not reserve it. A paired empty Start
does reserve it. A foreign/local-name lookalike does not. Do not change this parser
behavior to simplify the source scan. Use document admit_story_sdt_content_slot's
existing correct event distinction as a reference.

Add one operation-local seen bit to the existing legacy source control stack.
Only a qualified Start directly under that modeled control enters content and
sets it. Skip later contents as whole raw subtrees with the existing checked reader,
without paragraph/path/field entries. Empty remains raw. Nested controls have
independent state. Preserve all Root/Table/Row/Cell grammar, namespace checks,
source ranges, body certificates and complete consumption. Do not broaden the
modeled collector or introduce raw-descendant/ordinal/substring fallback.

Close the same selection policy in lower source_has_ordered_field_owner, used by
revision deletion/moveFrom ownership admission, without constructing a model to
discover the bit. Existing dynamic source scanners already have seen state, but
must not mark or classify an EMPTY content as its typed slot. Reconcile the five
actual Start/Empty callers: scan_dynamic_table_spans, collect_simple_toc_diagnostics,
toc_bookmark_marker_offsets, whole_paragraph_bookmark_ids and toc_paragraph_insertions.
Their other owner/revision/namespace/graph rules remain. No ancillary runtime RED
is invented. Every raw byte still counts for resource admission and custody.

### Coherent successor and final debt

Preserve all actual failed binaries/streams/maps and sealed probes. New tests must
prove initial typed admission before requiring a typed projection. Reacquire the
whole56named union with explicit membership, all six ordered-source phase rows,
three public duplicate controls and the needed event-form/alias/nested controls.
Cover scalar NEW/empty, typed0/1/2 and protected residuals, multi-run/later Ruby,
nested control/revision/COW/late error, Source original alias/default/fresh-prefix
collisions, own revision local shadow, mixed/detached/wrapper/related-story callers,
reentrancy/foreign-r/source isolation and total-growth/no-growth controls. First
Empty then Start, paired-empty Start then later content, and foreign lookalikes
must distinguish actual D8 ownership. Retain original18/12probe successors, seven
ancestor rows, six saves/15cache8/SVG and fresh CLI/installed54binding controls.

Use existing carriers, scoped all-target CHECK before expensive builds, strict
affected lint/fmt/prose/adapters/diff and HLD03/04/10 before freeze. New contexts
without executed old-source probes remain source-derived expectations, not claimed
historical RED. The existing actual D6/D7/D8 counterexamples supply real stopped-
source witnesses. Independent microscope pass12 must review the resulting union.
Archive refresh, exact-source green CI and every final native/schema/full/hash/
raster/typing/WASM gate remain due. Q04/Issue430 and held reporter/publication
actions remain unchanged. Root grants the next sole implementation lease only
after the owner acknowledges this canonical revision and these exact design pins.

## D7 concrete leaf, drawing and mutation closure

Root adopts four independent source-only clarifications: leaf9242d9ec18681ed40c1e42291366ddb1fbee4d6d146673e270c3de02a470129c,
drawing/mutation7a346ceea09ed4e232f3884efbd8fec02b85ab992234a992f4104f92546fbf0b,
capture resource42ffda55ad51166958ae7e625edc45520abb5a671f4988a861a225ec8a921766,
and run-raw contract98be4167de82d57d7ab10daf5af687997d00c20fc84502f9123f97bea4162986.
Root read each complete report and checked its pins. The leaf report's table pin
is its historical live observation. Its concrete border/shading callers remain
present, and the later stable pause map independently pins the current table.
Other final clarification source pins match the paused tree. These are designs,
not additional executed RED/GREEN claims.

The worker returned all leases after scoped check6 PASS8.61s with one incomplete
section-property unused-context warning. Its partial3857-input map is
fa986aedae2b31478c4298cc1ae98faeae1812eee11375f7fe713709b2bb81c9.
Only ten authorized source files differ in that partial implementation. No
semantic test has run on it. D7 remains incomplete. The following supersedes
the earlier broad canonical-w Source Ruby ancestor wording.

### Public property behavior and explicit private leaves

Authorize necessary edits to THREE existing files: properties.rs, borders.rs
and drawing.rs in rdocx-oxml, beside the already authorized existing files.
No new file/module, public method, signature, field, trait, generic framework
or global source mode is authorized. Preserve public CT_Shd::write_xml's exact
caller tag and canonical w attributes, public CT_BorderEdge::write_xml_attrs,
edge/paragraph/table border and tab writers' existing behavior, and public
CT_Drawing::to_xml. Never infer public modeled attribute URI from arbitrary tag.

Factor substantive attribute work into private explicit-Word-qualifier cores.
Existing public entries retain canonical w. Private Source entries resolve a
safe generated Word root and modeled attributes against actual source context
PLUS the leaf's own retained declarations. Preserve opaque QName/value/order.
An own q=producer shadow must force a different Word qualifier, not a rebinding.
Cover all nine shading selectors, border val/sz/space/color, false/true toggles,
PBdr schema-order edges, tab val/pos/leader, existing table/cell collections,
run fitText and EastAsianLayout leaves. Existing canonical semantic comparison
snapshots in retained property paths remain canonical, never q snapshots.

For started border source_xml, retain the actual raw root and its own scope as
authority. Reuse surgical merge_defaults_scalar_attributes with canonical-w
MODELED-ONLY before/after snapshots. Opaque extra_attributes must not pollute
those semantic snapshots. Let the existing merge use its original Word alias
or bounded fresh qualifier. Do not pass caller-q snapshots absent from that raw
scope, rebuild the raw subtree, change the global merge helper or public branch.
Require unchanged source exact and one-owned-value edits preserving all opaque
attrs/children/comments/lexical bytes. All leaf growth belongs to the existing
TOTAL source preflight, not a new per-leaf allowance.

### Drawing and Ruby consumers retain independent raw authority

CT_R Source invokes a private Drawing Source entry with explicit effective
original context. Select a safe surviving Word drawing wrapper and matching end.
Keep public Drawing behavior unchanged. Raw inline/anchor bytes remain exact,
including own local r/wp/default shadows. Their inherited original bindings
must reach them through generated ancestry. New modeled Drawing payloads keep
existing concrete generated emitters and actual relationship declarations.

Ruby raw-vector capture completion does NOT prove arbitrary modeled base or
phonetic Drawing payloads are complete relative to canonical w. Source Ruby,
rubyPr, rt and rubyBase therefore use safe known Word qualifiers under their
actual original source consumer. Do not hoist canonical w over their common
ancestor. Pass that source context and only safe generated qualifiers into runs.
For the specifically approved Ruby raw vectors, satisfy their canonical-w
contract locally at each emitted raw root that needs it, using the existing
borrowed root plan against actual Source destination. Own foreign/local bindings
win. Do not mutate stored vectors, rewrite QNames or complete Drawing raw this way.

Extend existing CT_R retained namespace discovery to Drawing inline/anchor raw
under the FULL original run frame, including original Ruby/rt/base-local context.
Capture only externally consumed dependencies in the existing root-attribute
metadata carrier. Keep true run-local lexical records distinct. This Drawing
discovery MUST include canonical w when consumed, unlike the existing general
extra_xml canonical exclusion. A later q/fresh Source run cannot otherwise
prove Drawing's original w URI. No new flag, record kind, marker, registry or
model field. Raw Drawing bytes remain unchanged. Necessary new existing-carrier
metadata is an intentional context-preservation fix, not new source identity.

General CT_R extra_xml retains its EXISTING implicit canonical-w raw contract
and canonical capture exclusion, preserving canonical raw-slot fixed points.
Decode own original records before composing destination bindings. An explicit
own foreign binding and an actual raw-root local declaration override that
implicit contract. A containing Field's foreign w cannot redefine a pasted
canonical run's raw meaning. Prefer a surviving canonical run boundary when
safe for all its own payloads, declaring Word locally only when needed. If own
retained foreign dependencies make that unsafe, use a safe known Word run
qualifier and local completion ONLY of general raw roots with a proved canonical
contract. Drawing retains its separate own facts and exact bytes. Do not infer
two different original URI meanings from indistinguishable model records.
Mixed contexts must be represented by actual own/local dependency records,
not guessed payload origins. Preserve opaque root attributes and QName-valued
metadata. Public default serializers and unchanged Field bytes/IDs/pairs stay.

### Mutation precedes serialization and needs its own authority

Obtain original Field context BEFORE scalar/typed/annotated ordered mutation,
not only before ordered_paragraph_xml. Thread it through private slot/control/
revision mutation into a substantive private fallible remove-runs core wherever
zero/many replacement can generate collapsed hyperlink raw. Keep existing public
remove/replace signatures and ordinary legacy behavior. Explicit checked Source
paths propagate failures to the uninstalled outer candidate and never silently
suppress Source construction errors.

Revision edit closures need their original root-overlaid context before editing
as well as refresh. Existing legacy callbacks may still call an infallible
remover before refresh without that private argument. Close that route through
EXISTING HyperlinkSpan.extra_attributes capture of only original dependencies
needed by opaque attrs, boundary children/revision raw and QName-valued metadata.
Preserve local lexical declarations and canonical no-needed bytes. Foreign
w/r/default dependencies become retained declarations on that same owner, with
no new field/marker or guessed complete paragraph frame. Original source bytes
outside needed namespace completion stay exact.

Empty, preserved and collapsed generated hyperlink roots use their actual own
records, safe known Word/relationship qualifiers and matching ends. If a legacy
Source-None path lacks a local known guarantee, establish it on that surviving
generated root without rebinding opaque w/r. Do not hoist it to a paragraph or
mint fresh prefixes already used in retained names/values/raw. Approved placeholder
wrapper graph recovery supplies context to its private mutation path too.
Public callbacks remain public behavior, not an ambient Source mode.

TOTAL mutation-time collapsed-root growth and later source/vector/leaf growth
are admitted together before the first expanded collapsed raw construction.
Use the same original-source/current-fact envelope and an operation-local checked
reservation if needed. Never give each child a full renewed allowance or reuse
newly expanded raw as parent input. Existing checked revision/accepted-run paths
propagate errors and preserve outer atomicity. Public infallible signatures do
not become Result APIs. Final source-path/source-pair guards remain exact.

### Initial capture also requires preconstruction admission

Existing admit_word_xml scope-copy counters return copy totals. They do NOT
prove the inequality for new Run/Hyperlink namespace records. Add an aggregate
ORIGINAL-input capture preflight at existing independent admission boundaries
before model construction. Use the same actual modeled owner selection, including
D8 first-qualified-START policy, and existing borrowed external dependency indices.
Pool physical scope declarations once and keep stack indices for shadowing. Do
not use the URI-cloning discovery helper while planning or construct models just
to identify capture owners.

Compute checked TOTAL needed record/storage/declaration/escaping bytes across
the original input, with actual duplicate storage charged where required. Exclude
already represented/local dependencies. Preserve canonical general-raw exclusions
and the explicit Drawing exception. Original raw/current scalar/attr facts and
needed external scope once feed the SAME unchanged private affected-source work
arithmetic. A local URI already in original bytes is not multiplied into fresh
scope allowance. Returned copy_bytes, projected records and child expansion are
not new input budget. A huge unused URI selects zero new work and cannot trigger
a new metadata refusal. Matching already-local/canonical zero-growth cases pass.
Genuine N required copies above the unchanged allowance may resource-refuse,
with zero expanded record constructions, before allocation. This is necessary
metadata work, not unsupported format.

Cover actual independent main/header/footer/notes/comments/glossary parts, public
P/R/table/row/cell fragments/readers, controls, direct revision parses, Field and
exact-source helpers, and Ruby private roots invoked without an admitted parent.
Internal parser descendants consume the same parent decision/reservation, never
derive a fresh allowance from expanded XML. Keep admission's existing public-to-
crate return signature and public parse interfaces. Styles/numbering property-only
parses are excluded unless a concrete new capture owner is actually constructed.
Record that caller if found. Physical opaque bytes still count, while opaque
descendants cannot become modeled capture owners by local-name coincidence.

### Final necessary discriminators

Preserve the executed D6/D7/D8 witnesses. Add exact public custom-tag border/
shading/default-Drawing compatibility controls, Source root-local leaf shadow
and started raw-source surgical edits, inline AND anchor raw drawings in both
Ruby base/phonetic runs with distinct original local scopes, and simultaneous
canonical raw-vector dependencies. Exercise mutation-time empty/preserved/
collapsed hyperlink under Field/control/revision and a real legacy revision
edit callback, before refresh and after publication/reopen. Include pasted
canonical checked run/raw into foreign-w ordered Field and reverse foreign run,
MC value-only dependencies and represented mixed-owner context controls.

New capture tests include shared long needed URI across Run-Drawing/Hyperlink
owners, smaller success and actual total-over-envelope refusal before expanded
records, huge UNUSED URI and already-local zero growth, independent versus parent-
admitted readers, and no descendant allowance reset. Record actual original bytes,
needed external facts, predicted storage, allowance and expanded construction
counts separately. Existing raw parsing allocations are not claimed zero. All
new non-executed cases remain proposed discriminators, not historical RED.
HLD03/04/10 and independent pass12 must cover this closure. No public API/data
shape/version/baseline or final acceptance debt is waived. Renew implementation
only after explicit acknowledgement of this revised canonical contract.

## Retained field raw declaration closure on generated paragraph roots

The imported field discriminator
`fx211_selected_run_xml_replaces_checked_existing_owner_atomically` required a
retained field start tag as an exact substring. The approved D7 contract
already permits a declaration on the actual surviving generated owner when a
canonical paragraph root cannot carry the source alias. A retained field raw
root that uses a source Word alias under a generated canonical paragraph root
therefore receives the minimal missing alias declaration on that retained root.
This is declaration closure. It does not rewrite the field instruction, result
text, opaque producer attributes, child order or any byte outside the retained
root.

The discriminator checks the expanded field name, the instruction, the stored
result text, the producer namespace and the producer attribute. It accepts the
added alias declaration and exercises the default-prefix arm. No public
behavior, resource limit, baseline or source byte outside the retained root
changes. This closes the pre-existing red assertion against the approved
closure rule rather than weakening it to a text-only check.

## C1 user decisions, 2026-11-01

Q-C1-01 preserves exact raw-fallback acceptance. Original-input preflight must
match the existing supported typed versus raw outcome. It may not refuse an
otherwise supported over-budget typed candidate when semantic parsing would
retain that container as raw XML. The unchanged private work budget still
applies to actual retained capture work. This rejects conservative false
refusal as a substitute for exact qualification.

Q-C1-02 permits one narrow unchecked exception. Bound parser and checked
operation paths, including their internal descendants and composed sibling
work. Existing legacy infallible read helpers invoked directly on
caller-assembled models keep their signatures and outputs and remain explicitly
unchecked. This exception does not exempt parsing, checked operations or their
internal use of those helpers from the aggregate operation budget.

These are explicit user policy choices, recorded in
`/private/tmp/rdocx-cross-review/SOURCE_C1_USER_DECISIONS_20261101.json`.
They do not approve the pending C1 mechanism or close its review findings.
The next Source implementer must acknowledge both choices in the existing
F-X272 progress record before receiving the write lease. All other acceptance
criteria, resource arithmetic and final verification riders remain required.


## Root C1 implementation sequence, 2026-11-02

Root resumes the existing approved source-preservation invariants after the user's implementation takeover. The rejected V11 mechanism is not adopted. First repair the independently closed wrapper-dispatch undercharge: a nonempty Word smartTag or customXml retained by a paragraph is also parsed as a synthetic ordinary paragraph by run_wrapper_paragraph. Its direct runs, nested wrappers, hyperlinks and content revisions therefore participate in the same original-input capture reservation. Carry that existing paragraph dispatch fact through the original-input walk, without constructing a model, creating a child allowance or changing the direct legacy infallible read-helper exception. Do not generalize this to arbitrary opaque children or wrappers retained inside a run.

The new existing-module regression original_capture_expanding_wrappers_share_the_original_input_allowance covers a nested wrapper after a direct run and two sibling wrappers, plus a smaller admitted control. The separate original_capture_keeps_supported_raw_control_fallbacks_over_the_typed_budget proves the still-open exact fallback obligation. Both actually failed against the unchanged behavior on 2026-11-02, with32other capture controls passing. The wrapper repair is an incremental implementation of the existing contract, not approval of the unclosed constructor-verdict mechanism. Source C1 acceptance, Issue430 implementation, independent final review and every final rider remain pending until the complete contract holds.


### Shared instruction facts prerequisite

The next bounded C1 step separates shared field grammar from model publication in the existing text.rs. Tokens contain decoded text or a source-occurrence index, never an eagerly cloned nested Field. One concrete syntax record owns the trimmed raw spelling, actual Unicode-uppercase name, argument operands, ASCII-lowercase switches and physical nested order. Ordinary read construction materializes that same syntax with its actual nested fields. The complex-field close performs its unchanged initial scalar and nested-child admission before materialization. Scalar checked arithmetic is shared with initial_field_work_bytes, including64per argument, decoded text lengths, switch name and optional argument text lengths, cached result and exact source scope.

This introduces no public model field, signature, parser policy, trait, generic, file or budget. It is a concrete prerequisite for constructor qualification, not a claim that the qualifier is implemented. Existing populated nested fields still undergo the existing model and normalization guards, so their work is not declared empty. The remaining complete preconstruction facts need bottom-up weighted-owner/table-depth and affine normalization summaries, exact fallback propagation and a bounded single qualification schedule. Exact raw-control fallback remains RED until those routes are implemented. No per-field new source allowance or depth cutoff is introduced.

Local controls compare pre-materialization scalar work with the existing model calculation for surrounding whitespace, Unicode name normalization, decoded quoted text, flags and nested switch arguments. Existing parser and field tests verify unchanged instruction/result/source behavior. The complete original_capture batch must retain the known raw-fallback failure honestly. Final C1 and sprint gates remain pending.


### Negative constructor certificates for retained raw controls

A bounded implementation increment shares the existing six empty run measurement decoders (size, complex-script size, spacing, position, kerning and width) between ordinary RPr construction and a source-only constructor-failure scan. The scan produces only proven raw-control spans, never a claim of positive typed validity. It follows direct mandatory body/table/row/cell/paragraph/run paths and the actual SdtOwner.owns policy, first-started-content selection and nested-control failure boundary. Opaque, property-descendant, revision, field and other unqualified routes remain Unknown. A nested control failure is swallowed at its own control boundary, matching construction, rather than poisoning its containing control.

The explicit source event stack uses the same pooled namespace declarations and effective binding resolver as original capture accounting. A first scan records proven fallback spans, then original capture accounting suppresses only those discarded descendant model records. Actual control construction consumes prepared occurrence facts on the proved primary routes before allocating the control model. Unproved entry and projected-alias routes retain the negative source scan. A proved late mandatory scalar failure therefore avoids constructing discarded runs and expanded records. No throwaway model, special invalid-size test, ambient state, new allowance, public API or cutoff is introduced. A certificate never tightens the imported parser's scalar rules. Valid typed over-budget controls still refuse.

This increment does not close the complete constructor qualifier. Unknown property families, field/revision/ruby and other constructor outcomes still need exact facts, populated-cache model maxima, affine normalization summaries and full scheduling closure. Independent control-root admission still needs its owner context carried by the applicable original admission route. Preserve these gaps even when the present raw-control regression becomes green. HLD updates for this increment are deferred to the existing F272 impact list after Root reconciliation, because another worker owns HLD03/10/12 in this wave.


The negative certificate requires both a resolved Word binding and the actual constructor's activated Word alias. A source declaration activates that alias. An inherited encoded namespace marker alone does not. An unavailable certificate also stays Unknown and falls through to the unchanged ordinary constructor, rather than becoming an outer failure. Tests cover both distinctions. The implemented run-measurement regression is now green, but an actual paragraph-spacing constructor fallback remains a separate red control. This is not complete qualification.

The original admission performs one certificate scan sharing its declaration pool with capture accounting. Control entry currently repeats a source scan. Actual event counts on guarded document inputs at control depths1,8and32 are28,252and2508 against16,44and140 physical events. These are repeated scans, not linear work. The existing weighted-depth guard refuses depth257 before run construction on that document route. The public checked placeholder::try_replace_in_block route performs namespace/XML checks but its direct CT_Sdt call has no independently proved weighted-admission dominator here. Full operation/subspan facts reuse or a complete existing-envelope proof across every checked caller remains unclosed. No new cutoff or per-child reservation is added to hide that gap.

### Root shared paragraph scalar increment

Root extended the same negative-certificate mechanism with a concrete ParagraphMeasurement decoder used by ordinary paragraph/frame construction and source admission. It shares all18numeric spacing, indentation, division/outline and frame attribute slots, retaining signed i32 measurement rounding versus unsigned u32 integer parsing, missing-value behavior and raw attribute spelling. Actual mandatory pPr/rPr dispatch is represented without entering opaque or paired numeric children. No model or expanded capture record is created by the decoder. The constructor entry helper is now named control_has_property_measurement_failure to describe both property families.

This remains a partial constructor mechanism. It does not certify complete positive validity, fix compound run/border/tab/section/table/SDT or field constructor outcomes, supply populated-field maxima or affine work summaries, or close derived-input and checked placeholder scheduling. Q-C1-01/Q-C1-02 and every original budget/refusal/preservation rule remain unchanged. The two next compound run-property controls explicitly retain genuine missing-family failures rather than allowing the scalar matrix to be mistaken for full C1 acceptance. HLD reconciliation stays deferred until the shared HLD worker returns.


### Compound property scalar wave

The source-only negative facts cover mandatory compound run numeric properties, paragraph/run/table border measurements, direct tabs, selected section/table numeric properties and selected control-property scalars. Ordinary construction and admission use the same signed integer-measurement kernel and concrete unsigned/span/merge/control decoders. Compound run selectors are shared with their actual consumers. Table and section owner/action selection remains an explicit source event dispatch that follows the actual constructor, rather than claiming complete semantic validity. fitText width/id and eastAsianLayout id retain signed i32 parsing. Border size/space, table bands/counts and column counts retain u32 parsing. GridSpan preserves its positive-value requirement. Control alias/tag/id preserve normalized first qualified val lookup and required-value errors.

Only proved mandatory paths propagate failure. Empty table property numeric leaves are distinguished from paired retained leaves. Cell property parsing and grid columns accept their actual Start/Empty forms. First-started sdtPr and sdtContent selection is explicit, and nested control failure stays at its own boundary. Empty page size/margins and columns use their actual constructor actions. Paired page size/margins can be preserved locally by their constructor and remain Unknown here. Malformed pgNumType start stays retained rather than becoming a new failure. Opaque, ignored, revision, field, ruby and hyperlink paths remain Unknown. Missing mandatory Option outcomes such as absent fitText width are not yet a complete source certificate.

The actual former fitText/eastAsianLayout controls and the broad early/late scalar matrix are green. This does not claim complete positive constructor facts, all raw-fallback outcomes, populated field model maxima, affine normalization, original/derived input mapping, operation/subspan reuse or checked placeholder scheduling closure. Per-control rescans remain on explicitly unproved entries and projected contexts. The proved primary routes consume the operation-local occurrence facts described below. No budget, cutoff, public API or independent per-child allowance changes. Q-C1-01/Q-C1-02, raw XML preservation and final HLD/sprint gates remain binding.


### Operation-local constructor occurrence facts

One original admission prepares source control occurrences with the actual owner, original source span, negative raw outcome and closed descendant occurrence range. Construction consumes disjoint source occurrence ranges in parser order. Entering one control reserves its full subtree range from its parent's cursor before parsing the control. If that control fails or stays raw, its unconsumed descendants cannot select facts belonging to the next sibling. This is an explicit source occurrence mapping across capture/re-emission and root namespace injection on the proved constructor routes. Those transformations retain element order. Namespace activation is a separate prerequisite, not inferred from ordering. No derived byte offset, hash or equal-text lookup selects an occurrence.

Only the proved body/paragraph/table/row/cell/control routes carry this context. Unknown field/revision/ruby/drawing and synthetic-field routes remain explicitly outside the scheduling claim. Prepared negative facts also govern actual original weighted owner admission. A checked block replacement first retains its existing complete XML/namespace validation, then admits its original input with the actual Body root-control context before constructing a model. Within each prepared original input, one allowance remains aggregate across siblings. No ambient registry or independent child allowance is introduced by its internal constructor consumers. Full shared constructor action selection, positive/fallback facts, field maxima/affine work and derived-input scheduling remain mandatory debt.


The primary context-aware routes include the document root, public consumed body/paragraph/table/row/cell readers, paragraph/table/cell fragment roots and the independent checked block helper. The consumed paragraph reader carries an explicit virtual Paragraph constructor role while preserving its existing QName/close and sibling validation. Control-root admission requires its actual caller owner. story_raw_is_typed uses that owner rather than guessing Standalone. The public Result-returning standalone CT_Sdt reader remains an unproved entry because exact constructor error consumption must be represented before its discarded-owner admission can be changed safely.

An inherited encoded-only Word namespace binding can become an activated alias after source closure. Such contexts retain the old constructor scan and remain outside the one-pass qualifier claim. Source-declared Word aliases, foreign rebindings, opaque descendants, first-only properties/content and repeated equal-text controls have explicit source-occurrence controls. A context-backed constructor never selects facts by derived byte offsets or equal text. This mechanism is still based on the proved action subset, not a complete shared constructor-action/positive-validity kernel.

The independent try_replace_in_block helper is bounded for its original input after its existing full XML/namespace/root validation. Its public signature carries no outer operation reservation. Cross-crate replace_scoped_candidate and paired_scoped_edits still need one aggregate origin/reservation across selected and counterpart materializations. Reopening a package or validating physical correspondence alone is not that proof. edit_raw_block and its header/footer/textbox/note mutator uses are not exempt. Q-C1-02 permits only direct legacy infallible read helpers on caller-assembled models. Read helpers inside a checked operation also require its bounds. Resource failure propagation and atomic refusal through those mutator callers remain due.


### Exact borrowed-model normalization summaries

Root replaces repeated normalization descendant visits with one per-call iterative postorder over actual current Field allocations. Each subtree records coefficient A and offset B for exact work A times incoming ancestry plus B. Local byte accounting is unchanged, including raw source, current instruction and result, original scope multiplicity, legacy/default/author strings, checked original/current payload and ordered normalized/pending source. Current instruction edges add zero ancestry. The selected cached view adds one, and ordered accepted-run edges add one plus their actual revision ancestry. Every repeated child occurrence still contributes its full coefficient and offset. Shared original source identity is not current-model identity.

The unchanged former walk remains test-only as a numerical reference. Controls compare parsed ordered revision results, divergent current fields sharing immutable source and result owners, selected versus stale typed caches, legacy metadata, rebasing and duplicate occurrences. Shared model DAG depths2/4/8 retain exactly equal charged work. Actual old visits7/31/511 become5/9/17 summary computations, matching unique current allocations rather than charging less work. These are model-walk observations, not parser preconstruction facts or release timing. An initial Root expected-delta assertion omitted the existing dynamic factor124 on the64-byte argument term. Correcting that test premise produced the expected result. This was a test defect, not a production semantic RED.

The two new summary controls and the52 existing field-filter cases pass on the current native producer. Changed-crate all-target/all-feature warnings-denied Clippy and changed-source formatting pass. Parser-populated summary facts, complete weighted-owner/table maxima, aggregate initial child admission, derived source scheduling, namespace/action/Result outcomes and cross-facade mutator composition remain unfinished. No allowance, cutoff, public signature or model shape changed. Source snapshots and actual receipts are retained under /private/tmp/rdocx-s96-root-c1-implementation/root-affine-model-stage. No full C1, provider or sprint acceptance follows.


### Borrowed stored-resource independent maxima

Replace path-frontier revisits with one per-call iterative postorder over actual owner kind and allocation identity. Local weights retain the existing predicates: control1, table depth1, a Simple-source ordered Field1, and typed textbox16 only under its current shape predicate. Each subtree stores independent weighted-owner and table-depth relative maxima. Children combine by coordinate maximum, not sum, and the witnesses can differ. Every existing stored-resource edge remains, including original/current instruction operands, all cache/run/span alternatives, original control sources, revisions and both authoritative and legacy textbox models.

Apply each existing local entry refusal before enumerating its children. A completed memo hit checks its relative maxima against the actual incoming path. If either coordinate fails, follow the existing child/LIFO order through violating summaries to emit the same actual local boundary error, skipping only proved fitting subtrees. Do not traverse an entire invalid tail just to compute a root maximum. Empty charged owners retain their local weight. Inclusive256weighted and32table limits, original allowances and public/model shapes stay unchanged. The former iterative guard remains test-only reference. This borrowed-model guard does not supply parser-side facts, complete constructor outcomes, affine parser summaries or aggregate checked-operation closure.


The stored-resource implementation now enumerates the exact former model edges through WordOwner.resource_children and retains the old guard as a test-only reference. The per-call memo key includes owner kind and actual allocation address. Current Field clones sharing a source Arc remain separate nodes. Independent relative maxima have different witness branches where applicable. A memo hit that fails either incoming coordinate descends only through offending summaries in the original append/LIFO order, preserving the actual weighted-versus-table diagnostic order even when both coordinates fail. Local entry refusal precedes child enumeration, including an authored2048control invalid tail.

Controls compare reference acceptance and actual error text at256/257weighted and32/33table boundaries, separate witness branches, sibling alternatives, stricter/incomparable shared paragraph paths and Simple-versus-Complex source forms with ordered results. Every Field storage alternative and retired original control runs remain represented. Typed textbox tests require both authoritative body and stored legacy paragraphs, including an empty typed body charge. The shared composed graph observes363new popped owner entries and362computed summaries versus699old popped entries. These are actual guard loop observations, not release timing or parser preconstruction facts.

This completes only the bounded borrowed-model maxima increment. Parser-side constructor/positive/Option facts, populated source summary publication, affine parser facts, standalone Result and projected namespaces, original/derived source scheduling, aggregate selected/counterpart reservations and checked mutator error propagation remain due. Q-C1-01 exact raw fallback and Q-C1-02 direct legacy infallible reads only remain unchanged. Full C1, provider acceptance and all named final sprint gates remain open.


### Parser source-piece and publication facts

Parser field source is first an ordered borrowed plan of actual run, hyperlink, boundary marker, control and opaque source occurrences selected by the existing complex-field source producer. The plan retains their actual source slots and byte lengths, not equal-text or reconstructed-offset identity. Admission reads those same pieces through a concrete buffered source input before assembling a new source Vec. Contiguous source and segmented source share the actual namespace/capture/weighted admission implementation. Synthetic admission envelope bytes stay separate from original input allowance. No generic validator or throwaway model supplies facts.

Before parent field source capture and original-instruction publication, the parser must include actual populated instruction and result-cache children in its prospective affine/resource facts. Shared metadata and source selectors remain actual constructor consumers. Existing Field children may supply facts only at their real completed publication boundary. Cache alternatives and annotation retention remain explicit. This connected checkpoint may stage these source/publication prerequisites, but must not claim that already constructed run/drawing/revision/ruby children are preconstruction facts. Original whole-operation reservation, full source constructor actions, derived origin mapping and all unproved contexts remain explicit until implemented. No new allowance, cutoff or public surface is authorized.


Derived complete-complex result publication must prepare the exact replacement owner graph, selected ordered normalization children and current legacy/source scalar terms before source COW or normalized-source copying. It consumes actual already-owned result facts without constructing a temporary Field. Replacement resource selection shares the real Field storage policy, retaining original/current instruction, typed/simple/New/span alternatives while retiring only the parsed/explicit cache slots actually removed by the publisher. A local source Arc borrow must not force an avoidable COW clone after qualification. This later mutation does not inherit a stale initial-close summary. Whole constructor/clone/source scheduling remains due beyond this boundary.


### Owned zero-weight field cloning continuation

The parser publication facts do not bound recursive Clone on pure Complex chains. The actual owned zero-weight edges include current instruction arguments and switch arguments, cached_fields, and Field entries in simple, typed and parsed cache-run vectors. The proposed continuation replaces only Field derived Clone with concrete iterative Enter/Finish frames. Each owned occurrence receives its own completed clone, preserving instruction and cache slot order. FieldSource and FieldRunSpan stay Arc-shared, as in the existing derived Clone, and ordered-result paragraph and XML owners also retain their existing Arc sharing. No source representation, signature, budget or cutoff changes. Nonfield run contents keep their existing cloning behavior. Drawing/body/revision recursion is not thereby globally stackless and remains an explicit scheduling obligation.

Root's initial suggestion that cache re-entry was automatically limited by256/32was checked against the actual resource graph and withdrawn. Pure Complex cache/run chains refute that shortcut. This is a corrected Root design suggestion, not a failed implementation or runtime observation. Controls must cover standalone FieldInstruction cloning, all owned cache slots, exact source/span/ordered-owner identity, mixed argument/switch/cache chains and iterative dismantling of deep test models. This remains part of the current unfinished C1 author phase, not a new acceptance cycle.


### Contextual revision header action continuation

The actual paragraph revision constructor chooses raw on a missing or invalid id before constructing descendants. A new constructor-oracle control has a genuine RED: original weighted admission refuses257nested control names inside that rejected revision. The connected repair shares one real RevisionHeader decoder with actual CT_Revision construction and source qualification. It qualifies only proven paragraph, hyperlink and run-property revision slots with their actual author requirement. A rejected child revision becomes its own opaque source span and does not fail its enclosing control. Effective Word alias activation is shared by the qualifier and the small selected header prefix view, without expanding namespace URI records. Valid-header content remains Unknown until its real body constructor outcome is represented. Root revision caller author context remains unproved and is not guessed.


### Connected parser checkpoint after F283 integration

The current source retains prospective populated simple/complex facts and now checks completed ordered-result publication before normalized source copying and COW source publication. Its resource forest shares one current-allocation memo across every actual replacement alternative, preserving original/current instruction, retained simple/typed caches, New baseline runs and span owners while retiring exactly the parsed cache slots that publication removes. Ordinary initialization and completed ordered publication compare their prospective facts with the real populated model in test builds. Borrowed source pieces use the same admission reader, including every byte boundary across namespaces, entities and Unicode. These are actual publication boundaries, not a complete original-input proof.

Field Clone now uses iterative Enter/Finish scheduling for every directly owned zero-weight Field occurrence in current instruction arguments and switch arguments, cached_fields and all three owned cache-run vectors. Its exhaustive initializer retains every scalar and existing source/span/ordered Arc relationship. Actual all-slot and composed4096depth controls pass on a512KiB thread stack, observing4097owned occurrences for Field and standalone FieldInstruction clones. Their resource maxima are exactly0/0. Deep fixtures are dismantled iteratively and use no recursive equality oracle. Existing nonfield drawing, run-property revision and body cloning, plus general recursive Drop, remain unproved by this mechanism. No historical crash is claimed.

The actual revision header decoder is shared with contextual source actions. A genuine RED refused257control names under an invalid or missing revision id although the real Option constructor retained the revision raw. It is GREEN, including200namespace-expanding descendant runs with zero discarded run constructions. Qualified first-attribute selection, decoded numeric entities, local alias shadowing, paragraph optional-author versus other required-author dispatch, Start/Empty events and valid enclosing-control charges are covered. Valid-header content outcomes and unproved root author contexts remain Unknown. Failed control and failed revision actions are distinct. A failed control selects its existing raw-capture policy. A failed revision stops typed descendant construction while preserving the enclosing capture policy.

An initial combined action record incorrectly borrowed control raw-capture policy for property revisions. The actual112case batch exposed111PASS/1FAIL in the existing property-leaf preservation control. The failing executable/source generation is pinned as revision-raw-capture-regression. The explicit action discriminator and inherited unconstructed flag repaired it without changing that fixture, an allowance or a refusal rule. Final connected batch145PASS, separate completed ordered-boundary1PASS, changed-crate all-target/all-feature warnings-denied Clippy0 and changed-source rustfmt0. Counts overlap and are not summed. First Clippy found two needless borrows and one needless return introduced in this stage's empty-simple helper. They were repaired without lint suppression. Source, commands, child/wrapper outcomes and executed binary hashes are retained in parser-field-facts-stage.

This is a continuation checkpoint under the same author lease, not author STOP, review acceptance or provider completion. Original constructor positive/Option outcomes, earlier run/drawing/revision/ruby construction, simple caller capture, whole-original and derived aggregate reservation, standalone Result/projected-alias actions and cross-facade mutator propagation remain open. Recursive complete complex-result qualification/transfer and grouped source/span publication are the next connected scheduling target. No facade, binding, HLD, Issue430, full native/schema/hash/installed/CI/packaging/release work was performed. Root retains those exact final obligations.


### Complete complex-result scheduling continuation

The next private refactor extracts the existing local geometry selector unchanged. Qualification walks borrowed actual cache occurrences with an explicit pending stack. Transfer moves the selected owned Field from its already owned grouped vector and uses concrete parent frames containing the Field, its local geometry, remaining cache slots and completed child occurrences. A parent takes the next actual cache slot only after the preceding child has published, preserving occurrence reuse errors and child-first mutation order. Local geometry is re-evaluated at transfer entry against the then-current projection, exactly as the existing recursive transfer does. Each actual owner receives the same qualification and transfer-local checks. This does not invent reuse of coordinates across derived or namespace-transformed input. No dummy Field, throwaway model, new cutoff or public surface is introduced. General nonfield clone/drop and original aggregate reservation still require separate proof.


### Grouped field source and span publication continuation

Grouped fields can enlarge each retained raw source after close facts were computed, then add original outside-run span owners. The connected publication guard recomputes exact affine work by replacing only the actual root Parsed raw-source byte term. It computes prospective resource maxima over the actual retained children and the exact replacement span-run alternative from field_span_parts, before COW source cloning or span model copies. The same child selector receives explicit optional ordered-result and span replacements. Group raw length is checked from borrowed source parts before assembly. Actual final fields are compared with their refreshed prospective facts in test builds. The source selector and field_span_parts fallback remain unchanged. Outside-run model reparsing still precedes these span facts, and no whole-original aggregate or renewed per-part source allowance is asserted.


### Complete-result and grouped-publication source observations

Complete complex-result qualification and transfer now use explicit borrowed and owned frames with the existing local geometry selector. The actual24level parsed result chain observes24qualifications and24maximum pending transfer frames, with every ordered result populated as cache. Related field, Issue425and revision controls79PASS. Grouped source/span publication then adds fresh exact source-root affine and prospective resource facts before actual COW source copies and span clones. Its actual shared-run case records two source enlargements, preserves all outside text and both fields, and agrees with the old numerical reference after publication. A synthetic overflow control observes zero source copies before refusal. These are new controls, not claimed historical failures.

The grouped generation passes82related field, Issue425, revision and affine controls and changed-crate warnings-denied Clippy0. Exact binaries and full conservative input inventories are pinned as complete-result-schedule-green and grouped-publication-green. The first scheduling check missed the qualified Range type and was corrected before successful check/runtime. That is an author type error, not a semantic RED. The final source has no qualified geometry recursion or owned complete-result transfer recursion, while unrelated model Clone/Drop recursion remains outside that observation.

This still does not close C1. Original constructor/action positive and fallible outcomes, earlier run/drawing/ruby construction, simple caller capture, outside-run reparsing inside field_span_parts, whole-original and derived aggregate scheduling, standalone Result/projected alias context and facade mutator propagation remain due. Refreshed facts are local actual-publication facts, not a claimed inherited allowance or original occurrence proof. The same canonical author lease continues.


### Shared run child actions and Option outcomes continuation

The next genuine RED is original_capture_special_child_option_outcomes_charge_only_actual_raw_records. Actual CT_R parsing retains a special child with foreign attributes raw and stores the repeated external namespace record. Original capture classification currently labels every Empty sym, ptab and empty-special candidate modeled, so its source guard omits those actual records. The source producer returned Ok instead of the existing work refusal. The coherent repair shares a concrete Start/Empty RunChildAction selector with actual CT_R dispatch and extracts only primitive special values from the existing symbol, positional-tab and empty-character readers. Actual construction turns those values into RunContent. Source classification consumes the same Option outcome without creating a throwaway RunContent model. Typed unused-namespace and raw zero-growth controls must remain accepted. Shared strict comment-reference scalar decisions can then follow the same actual Run action, with permissive note-id fallback preserved. This remains a constructor/action foundation and not complete positive run/drawing/control qualification.


### Original declaration facts continuation

Actual supported20owner source shows20escaped declaration-length scans for only2unique physical source pairs. That work-counter regression is a genuine RED and source acceptance itself remains unchanged. The operation-local declaration pool will retain each exact escaped-length result once, including a deferred overflow outcome so an unused declaration cannot create a new refusal. Every physical retained record still adds its full length and duplicate charge. A randomized hash bucket accelerates pool lookup, with exact prefix and namespace equality required inside each bucket. Hashes do not identify source occurrences, constructor facts or captured subspans. The existing deterministic source-order indices remain the actual record identities. URI strings stay owned once per pooled pair, with no URI copy in the hash index. This supplies reusable numeric source facts without changing a budget, factor, cutoff or public API.


### Effective namespace state continuation

Root identified reverse ancestor scans in capture binding and Word-alias activation. The actual supported raw-input control confirms a genuine RED. At depths16,128,512the source qualifier sees42,266,1034events, but namespace helpers visit792,34840,532504scope frames, with76,524,2060binding lookups. Full parsing of those same opaque payloads succeeds and preserves their bytes before the work assertion. Declaration-length caching alone does not close that work.

The coherent replacement retains an operation-local per-prefix history of exact pooled index, lexical depth and physical-declaration origin. Real Start/Empty declarations enter before dispatch and leave at the matching physical scope. Current binding and alias activation consult the effective entry directly. The existing ancestor-at-run-frame query searches the prefix's declaration history by depth, preserving inherited-depth conventions and local-shadow distinctions without walking empty frames. Bare inherited Word aliases stay distinct from encoded-only inherited bindings. Source-order inherited reduction retains the last exact prefix occurrence in linear passes. Actual scope visits, map lookups and history comparisons remain separate counters. No whole-original scheduling or global linear-time claim follows until the remaining routes are proved.


Simple-field constructor closure shares one concrete header/syntax producer between both actual factories and source actions. Missing instruction is immediately raw. Empty-element empty-name syntax is immediately raw. A started empty-name syntax remains present until the existing late check after cache parsing, so preceding child errors retain their boundary and order. Source qualification certifies only these actual early raw outcomes, with inactive first Word-instruction aliases kept Unknown for unproved namespace projection. The actual missing-header RED covers both descendant weighted owners and namespace records, with direct paragraph and inline-control consumers.


### Checked replacement error propagation reconciliation

Root authorizes four doc-hidden fallible twins with unchanged arguments and Result<usize>: try_replace_in_paragraph, try_replace_regex_in_paragraph, try_replace_in_header_footer and try_replace_regex_in_header_footer. Existing usize entrypoints retain compatibility error policy. Shared actual traversal feeds native replace_in_body, replace_in_headers_footers and both Result-returning compiled-regex routes. Checked callers forward errors and install only successful candidates. These APIs expose no budget, context or model-state internals. This decision does not close original-range aggregate reservation. HLD03 and HLD10 describe the final coherent propagation behavior.


### Scoped original-operation technical reconciliation

One inert doc-hidden ScopedWordOperation<'a> may borrow the unexpanded original part. Its constructor new(original: &'a [u8]) returns Self without scanning or funding the whole part. The approved methods take the actual original Range<usize> and borrowed BTreeMap<String,String> inherited scope: paragraph returns Result<CT_P> using the existing nonchecked fragment reader, document_block returns Result<CT_Document> using the exact canonical Word document/body wrapper, and replace_block additionally takes paragraph: Option<usize>, placeholder: &str and replacement: &str and returns Result<(Vec<u8>,usize)>. Internal admission facts and numbers remain private. Named consumers are scoped selected replacement and paired counterparts and their paragraph/document-block structural projections. The separate checked template table reader is not covered by these methods.

Metric-specific reservations use the deduplicated union of original physical ranges actually reached, necessary original declarations and current facts. Nested or overlapping bytes and original current-fact occurrences fund once. Namespace scope credit preserves the existing exact prefix/URI pair pooling. Each actual materialization consumes the corresponding existing metric again. Capture-record arithmetic is not a universal budget for ordinary source copies. Exact field/current-value and normalization terms retain their existing formula. Expanded bytes, unvisited ranges and unrelated part bytes never fund a reservation. Incremental registration preserves selected-first errors, expected mismatch and zero-count returns before counterpart work. Scope declarations must correspond to actual original offsets or genuinely external original context. The standalone template external-context route remains pending its own exact integration.


The scoped operation keeps a separate cumulative added-xmlns insertion reservation using the existing retained_binding_work arithmetic with current_bytes equal to0. Credit is the reached original raw-range union plus the existing multiplier on exact original prefix/URI lengths, pair pooled once. Only actual added escaped declaration bytes are debited, following the current closure policy for all missing in-scope declarations and skipping xml and local declarations. Ordinary raw copy, parse and model bytes are not debited to this metric. Unused bindings fund only this metric when actually injected and cannot inflate retained capture or field budgets. This technical application does not claim the previous native closure helper already enforced a cumulative cap.


### Root implementation continuation

The prior agent stopped after an automated classification error. Root owns the
canonical source and build lease at ed588356. Its eleven product/HLD files still
match the 158-case checkpoint. CI changes remain in the separate F283 worktree.

The concrete remaining alternate-drawing route uses CT_R.alt_drawings, which is
stored and cloned alongside direct RunContent::Drawing. Apply the same existing
Anchor weight and children to both locations in the production model traversal
and its retained reference. No new limit, model field or public signature is
introduced. A small direct-versus-alternate functional control will establish
the prior outcome before this correction. Aggregate-operation work remains open.


Root confirmed the alternate-drawing omission with a small functional control: the existing model traversal reported0instead of16for one stored alternate anchor. Both production and reference traversals now visit CT_R.alt_drawings. Direct and alternate drawings agree for authored/legacy shape storage at depths1,16and17. The new control and existing independent-maxima control pass, scoped check and changed-crate all-feature Clippy pass. Evidence is retained in root-alternate-drawings. This closes that concrete source lead only. The earlier missing Box in the new fixture was a setup error, corrected before the actual failing outcome.


### Shared namespace-closure implementation

Root adds one doc-hidden cross-crate helper in the existing document module,
close_content_fragment_namespaces(&[u8], &BTreeMap<String,String>) -> Result<Vec<u8>>.
The existing facade entrypoint preserves its native error mapping and test-only
diagnostics. The operation planner and facade use one lower closure plan/emitter,
so neither duplicates namespace insertion policy. This is an additive internal
Rust bridge, with no public facade signature or output-policy change. Planning
retains pooled binding indices and exact insertion lengths before allocation.


### Scoped operation capture and insertion slice

Root connects the approved operation type in a bounded first slice. Its reached
original range union, physical root-attribute facts and exact pooled namespace
pairs fund separate cumulative capture and added-xmlns metrics before closure
allocation. Existing paragraph, canonical document-wrapper and replacement
readers retain their current admission and error behavior. This slice does not
claim complete original/derived field, current-value or normalization facts.
Those metrics and replacement of independent derived parser admissions remain
required before C1 completion or provider acceptance. Source preparation failure
does not publish reservation credits. Root remains the sole canonical writer.


### Physical field-span range retention

Root replaces copied payloads in private FieldSpanPart with exact original
source ranges, including every repeated run shell and property range. The
constructor consumes outside models without assembling unused field sources.
Source-fragment and run-field writer consumers assemble only selected parts
from those ranges, preserving the existing bytes and fallback decisions. The
Outside model reparse is retained for now, so this slice alone does not close
its original-operation allowance or field/current/normalization aggregation.
No public API, new budget, cutoff, file or module is introduced.

The retained ranges are relative to the supplied field-span input. That input can itself be assembled from paragraph/run sources. Composing these local ranges back to the original part remains required before any operation-level credit. The range-retention slice does not relabel assembled input as original-part provenance.


### Scoped original constructor reuse

The operation's original capture planning already produces constructor outcomes.
Reuse those outcomes through internal paragraph, canonical document-wrapper and
block-replacement admission seams, retaining weighted traversal while avoiding
a second capture allowance on namespace-expanded input. Map negative-action
positions through the known root namespace insertion and canonical wrapper
prefix. Control occurrence identity stays at its original source location.
Public standalone readers retain their existing admission. Original field/current
and normalization provenance and the remaining derived readers are still open.
No new public API or numeric budget surface is introduced.


Prepared outcomes are reused only where the original inherited or root-local
w binding keeps the historical admission context unchanged. Without that
binding, paragraph admission's bare-w mode and the document wrapper's synthetic
w can differ from original actions. Those modes retain their existing ordinary
admission until explicitly represented. A small genuine Root regression showed
a newly activated valid w control consuming a later q raw-control certificate
and losing live text. Its frozen failing source and binary are retained. This
is a Root integration defect caught during engineering inspection, not a new
peer pass, a library input refusal or model error. The fallback is context-aware
and preserves the prior reader route. It does not prove complete context
equivalence or close the still-pending derived admission work.


### Simple constructor summary reuse

The next bounded slice returns the already-computed simple Field and its
ParsedFieldSummary together. The paragraph retains started-simple summaries by
the actual run index at construction and supplies them to the existing complex
cached-child summary consumer. Unobserved producers retain the current model
summary fallback. Empty-simple dispatch is not broadened. This is transient fact
reuse, not original byte provenance, a new ledger or a source-copy quota.

The normalization envelope is an output/work allowance, not an ordinary copy
balance. Debiting every unchanged raw capture against it would add an unapproved
refusal boundary, for example two unchanged1MiBcaptures of one original source.
That earlier engineering suggestion is rejected before implementation. Current
source/scope/current-value provenance and genuine generated-work aggregation
remain pending, with unused or synthetic context excluded from new credits.


### Raw-block source refusal classification

The existing raw-block mutator must separate ordinary source refusal from a
resource refusal before typed construction. One private WordSourceFailure in
document.rs carries Input or Resource with the unchanged underlying OxmlError.
Public reader outputs retain that original error. Mixed producers tag actual
arithmetic/limit sites, while XML/attribute/UTF8 and reserved-namespace failures
stay Input. Resource-only leaves are tagged at their actual admission boundary.
No message matching, public error variant, quota, source credit or model state
is added. Declaration-length failures remain deferred until their selected use.

The actual raw-block consumer admits recognized paragraph/table/body-control
original bytes. Successful facts feed its existing concrete constructors.
Ordinary Input retains the old parser/fallback route. Resource escapes checked
header/footer, textbox and note callers before publication. This connects a
real consumer rather than adding an unused admission facade. Later resource
failures inside unproved constructors and complete original/derived aggregate
scheduling remain separate debt.

Local controls pair the actual400character namespace/200opaque-run source-work
refusal with an earlier header edit, literal and regex modes, and raw paragraph,
table and body-control roots. Refusal must leave original paragraphs/raw bytes
unchanged and construct zero refused runs. A reserved xml binding in a foreign
opaque subtree and duplicate namespace/attribute syntax retain historical raw
preservation. Invalid-property raw controls stay preserved under the existing
qualified action policy. These are exact original input cases, not fresh
allowances funded by synthetic wrappers.


### Wrapped run constructor outcomes

Hyperlink direct runs and started simple-field result runs use the existing
fallible run constructor. Source qualification now dispatches those same direct
run occurrences into its shared primitive actions. A simple header has three
source states: Unknown namespace projection, an early raw None, or permission
to parse its result. The last is not complete positive field validity. Only
that proved result mode descends into run actions. Early None keeps its child
raw boundary. A later child Err propagates through the simple factory into its
enclosing control, as the actual factory does. Revision Option boundaries stay
unchanged and do not acquire this propagation policy.

Actual local RED covers hyperlink, valid simple instruction and started empty
simple instruction before a later257control chain. All three actual constructors
return raw but source admission refused. The unchanged fixture becomes GREEN.
The missing-instruction control stays typed at its parent and still refuses
later genuinely typed over-depth controls. No new quota, public API, model
construction or original source credit is introduced. Drawing, revision content,
positive Ruby prefixes and broader scheduling remain separate obligations.


### Ruby prefix proof for the shared Empty bold action

Source qualification and the actual run-property reader now share the selected
Empty Word bold decoder. It uses the existing permissive toggle semantics,
including omitted and unknown legacy spellings. Started bold stays on its
separate raw path. A Run defers positive property success until its concrete
property frame closes. That frame starts potentially known but keeps positive
proof only for the shared Empty bold primitive and its existing ignored-event
policy. Unsupported property children make it Unknown. Existing negative scalar
facts and preceding errors still propagate.

The actual new RED records one successfully constructed prefix run followed by
Ruby None for three bold spellings, while source admission wrongly charges the
unconstructed namespace-expanding tail. The unchanged fixture becomes GREEN.
Its counterpart without Ruby rejection still refuses the genuine200run work,
and an earlier invalid size still fails its enclosing control before Ruby None.
No complete positive property validator or broad drawing/revision proof follows.


### First drawing extent negative action

Both actual drawing container readers now share their exact Empty extent
attribute decoder. It preserves decoded values, literal unprefixed dimension
selection, partial dimension updates, signed parsing and errors from decoding
unselected attributes. Source qualification selects only an actual Run drawing
and direct Start inline/anchor using the existing lexical container selector.
It records a primitive error only for the first direct Empty local-name extent.
Any prior element permanently leaves later events unproved. Started extents,
subtree-consuming children and nested early-end cursor modes are not promoted.
A successful extent does not prove the whole container successful.

Actual local RED covers inline/anchor and canonical/foreign extent names. Each
actual bad numeric constructor makes the enclosing control raw before later
257controls, while source admission refused. The unchanged fixture becomes
GREEN. Corresponding valid numeric drawings with required docPr metadata remain
typed and still refuse genuinely live over-depth controls. New discriminators
retain bad extents hidden in consumed children, Started extents and events past
an early nested container End. A separate actual-reader control preserves a
prior dimension across a later partial extent and decoded signed numeric input.
This is a bounded negative action, not full drawing cursor qualification or
closure of all drawing scalar and metadata failure paths.


### Physical simple-header fact consumption

The next increment retains the shared simple header at its actual qualified
opening occurrence and passes it to the actual simple factory. Lookup uses an
exact byte offset only while the cursor borrows the same live input slice as
the consuming reader. It does not match instruction text, run ordinal or the
later FieldSource construction ID. Unknown/projected contexts and captured
subparses without a proved input binding keep the existing factory path. A raw
control cannot slide its unused header into a later direct field, because no
sequential header cursor or content search is used.

Original admission stores separate original and current lookup opening spans.
The existing insertion transform shifts only lookup positions for prepared
closed/wrapped readers. This is identity relative to that actual admission
input, not a claim of complete original OPC-part mapping. The scoped range's
whole-operation origin integration and all complex-field receipts remain due.
The cursor's borrowed input keeps the physical backing live and prevents an
unrelated allocation at a recycled address from qualifying. Child control
cursors drop this binding until capture correspondence is explicitly proved.

The lexer/classifier also retains its exact deferred instruction-only scalar
sum, excluding result cache, inherited scope and legacy metadata. The actual
initial-field admission adds those actual current terms with the unchanged
formula. Deferred overflow does not reject unused headers. This is not an
aggregate original credit, unused-namespace credit or a source-copy quota.

The discriminator has two equal-instruction direct fields with different dirty
flags and cache lengths, separated by a raw control containing a third distinct
header. Before reuse the three source headers and two factory headers are read
five times. Afterward only the three physical source headers are read. Repeat
with a long unused namespace and preserve Unicode uppercase expansion, quoted
arguments and switches, actual caches, flags and parser early/late error order.
A standalone unbound factory still uses its ordinary shared header reader.


The initial simple-header reuse slice passed303related controls and bounded
final1, but Root then reproduced a real context regression in that new source.
Admission can supply an implicit bare w alias absent from the actual paragraph
reader's root scope. A q:fldSimple with unbound w:instr wrongly became typed,
and an unbound dirty attribute could also acquire a value. ActualRED reports
two fields rather than one. This is a Root defect missed by Root's delegated
bounded audit, not an external model error. The source and binary are retained
in simple-header-origin/CONTEXT_RED.json.

Reuse now additionally requires the same first qualified instruction attribute
and, when parsed, dirty attribute in both contexts. The shared header records
those exact attribute indices under its actual early/late selection policy.
The consuming cursor compares the actual start/prefix selection before cloning
a prepared header. A mismatch uses the unchanged ordinary factory. This proves
only that selected simple-header context, not all constructor namespace modes.
Four physical/context controls pass, including the unchanged alias regression.
The broader historical admission/default-alias reconciliation remains due.


### Declared-root and paragraph-part context reconciliation

Complete document, paragraph-fragment, header/footer, notes, comments and
glossary readers now pass their actual empty inherited scope into admission.
Their real readers learn Word aliases from declarations. Empty inherited scope
encodes that initial context across qualification, capture and weighted admission.
Header/footer, notes and comments additionally select an explicit private reader
role because their actual global wrapper loops retain namespace updates after
End. This role is necessary caller context, not a schema-validation mode.
Consumed legacy readers, genuine inherited fragments and resolver-driven drawing
entries retain their existing defaults and supplied bindings.

Source constructor routing includes actual notes/comment paragraph owners and
glossary docPartBody paths. Note/comment item content is a paragraph collection.
The three complete part readers select their real global wrapper and item
actions outside those item readers. Other complete subtrees are Unconstructed
only where the actual reader captures or skips them. A recognized nested global
wrapper remains traversable, with its declaration updates retained after End.
The existing discarded-span channel excludes only their descendant typed work.
It does not invent a positive constructor outcome or remove required ancestor
raw retention. Unknown constructor families elsewhere remain Unknown.

Header/footer, notes and comments keep the admission they already performed and
pass its concrete cursor through their actual paragraph readers. This reuses
control and direct-header facts over the same live full input. Captured subparse
bindings retain their explicit fallback. Glossary's captured body route remains
separate and is not silently supplied mismatched original offsets.

Actual local RED covers six wrong resource/error boundaries, a live q-control
lost to a phantom implicit-w occurrence, and three raw-block collections wrongly
refused at257unconstructed controls. A separate actual32level test observed
2447qualification events against143physical events for header, and2449against145
for each notes/comments case. Focused GREEN now restores all outcomes and exactly
one physical qualification scan per part. Ordinary strict glossary XML refusal
is retained. No new quota, public API or unused-namespace credit is introduced.


### Origin-aware aggregate field accounting

ScopedWordOperation remains the one operation owner. Capture, namespace
insertion, field normalization output and constructor-entry work retain their
existing separate metrics. Original input, retained source/context and actual
current facts are not interchangeable. Original physical ranges and current-fact
occurrences fund once, while every genuine generated materialization consumes
its applicable reservation again. Selecting an antichain of affected model roots
only schedules work. Summing their independently renewed normalization envelopes
does not establish aggregate funding. All existing local guards still apply.

For field normalization, lift the existing local terms into origin-labelled
registration before applying their unchanged arithmetic. Original raw ranges
contribute their union. Retained source mapped to those ranges cannot credit
them twice. Existing retained sources at a checked model ingress remain explicit
baseline facts where original OPC correspondence is unavailable. They are not
relabeled proved original OPC ranges. Necessary original scope follows the
existing exact accounting, with unused and synthetic context excluded from new
field credit. Actual current instruction, cache, legacy and incoming payload
facts retain their existing scalar terms and W_NS.len() + 64 multiplier.
Existing checked payload, stored normalized-source and pending payload terms
retain their actual baseline identities. Newly generated output cannot supply
those terms again during the active operation.

Two distinct current model occurrences can share immutable retained source and
have different current values. Their current facts stay distinct. A candidate
clone, temporary projection, repeated inspection or derived parse of one current
occurrence forwards its existing identity. It does not establish another input.
Replacing current values selects the prospective replacement facts before the
affected expansion, without accumulating old and replacement values as separate
credit. Source IDs, equal text and a new allocation alone prove no original
physical correspondence. Lookups must remain anchored to the borrowed admitted
input or actual frozen baseline owner relationship.

The existing affine summary preserves all local child-edge and revision-ancestry
work. It remains a local planning/reference calculation, rather than a reason
to mint repeated aggregate input credit. Prospective output sizes must use the
actual writer's emission decisions before expanded allocation. A size check
after allocation is only a backstop. Formatter callbacks execute once and their
actual result supplies the prospective current facts.

Ordinary unchanged capture, inspection and retained-byte transport do not become
normalization output merely because another buffer is allocated. The existing
two unchanged1MiBcapture acceptance control remains binding. A genuinely repeated
template publication is a new output occurrence with the same baseline source
identity. An unchanged first occurrence retained in place remains ordinary
retention. Inspection, formatting and replay must forward the same operation,
without funding it from copied item.raw or expanded wrapper bytes. The precise
prepared-formatting bridge on the existing operation must carry those actual
mapped inputs, rather than accepting an unrelated raw buffer as original. Its
concrete cross-crate signature will be reconciled at that integration seam.

Required connected controls cover original input A, retained source/context B
and actual current values C all different, equal-text distinct physical fields,
shared source with divergent current models, repeated projections of one current
model, overlapping original ranges, changed caches and payloads, actual ancestry,
collective duplicate-materialization refusal and atomic publication. The
origin-labelled decomposition must also agree term-by-term with the existing
single-owner reference before aggregate integration is accepted. This defines
remaining implementation work, not completed C1 or a new quota.

### Root-adjudicated mapped-output carrier API

Root adjudicated the following concrete additive doc-hidden carrier signatures
as a technical reconciliation within the prepared minor release. This is not
an independent design approval, peer reset, new quota or completion claim.

The doc-hidden ScopedWordSource has private bytes, source-piece correspondence
and an Arc operation identity anchor. It has no public constructor, Clone,
numeric identity or exposed map. Its only direct accessor is
as_bytes(&self) -> &[u8]. The existing ScopedWordOperation retains its current
methods and receives these concrete methods:

```rust,ignore
pub fn source(&mut self, range: Range<usize>, scope: &BTreeMap<String, String>)
    -> Result<ScopedWordSource>;
pub fn source_paragraph(&mut self, source: &ScopedWordSource) -> Result<CT_P>;
pub fn format_paragraph(
    &mut self,
    source: ScopedWordSource,
    child_edits: Vec<(Range<usize>, ScopedWordSource)>,
    occurrences: &[(usize, Range<usize>, String, Option<CT_RPr>)],
) -> Result<ScopedWordSource>;
pub fn repeat_source(&mut self, source: &ScopedWordSource) -> Result<ScopedWordSource>;
pub fn finish_source(&mut self, source: ScopedWordSource) -> Result<Vec<u8>>;
```

The source method takes an actual original range and scope. source_paragraph
serves template parsing, numbering and formatting inspection without granting
another baseline allowance. format_paragraph consumes a parent and actual child
outputs. Child ranges address the parent's current mapped bytes. All operation
anchors, ranges and nonoverlap are checked before splicing, and correspondence
is spliced with the bytes. Equal output content is no identity proof.
repeat_source creates a fresh generated publication occurrence and consumes
applicable demand without minting baseline credit. finish_source consumes its
carrier and returns bytes only after applicable reservation. An unchanged first
retention remains ordinary retention. A foreign-operation carrier refuses.
Actual native/template integration must use these operations rather than
reparsing copied bytes as independently admitted original input.

Reservation updates are staged and committed only after successful
materialization. source_paragraph consumes the existing mapped facts and work
context. format_paragraph retains actual current inputs separately from child
XML. repeat_source charges only the existing applicable metrics, without a
universal normalization cap on fieldless template copies. finish_source is final
transport and cannot debit again. Named real consumers are required for every
method before completion. Preserve supported Send and Sync behavior and include
this additive surface in the final strict SemVer inventory.

### Stored-result P provenance at the active operation boundary

Root clarified the concrete parser case before aggregate ledger implementation.
Simple ordered parsing copies raw[result_bounds] into normalized_xml, and its
initial local summary includes result_bounds.len(). Complex completion has the
corresponding original result slice. NormalizationInputWork must remain exactly
equal to the old local role formula before origin-union accounting. This does
not require the post-union operation allowance to equal a sum that counted the
same original result bytes twice. All local guards and references stay unchanged.

At an active parser boundary, a P copy of the mapped original result range joins
the same original raw-range union as A. Storing that copy creates no additional
aggregate P credit. A P that already exists at checked-model ingress is instead
an actual selected baseline stored/current fact, funded once through its frozen
owner relationship. A new normalized output or P copy made inside the active
operation only forwards original/current/generated provenance. It cannot mint
a new baseline fact or enlarge that fact by being stored. The same rule applies
across reparse and candidate reopen. No source-id, text equality or new model
allocation establishes a new input. This is an origin-aware application of the
existing metric, not a new quota or a change to the local numerical reference.

### Current-fact lifetime across required phases

Root adjudicated the actual template sequence before ledger coding. The first
template_programs call flushes the ingress current model C0 before formatter
results C1 exist. Preserve C0 facts that an earlier required planned phase
actually uses or publishes. A later actual formatter-supplied C1 is genuine new
incoming data and may fund its applicable scalar terms once. Record phase and
lifetime identity and actual use. The no-old-plus-replacement rule excludes
unused or superseded facts credited for a single prospective expansion and
clone/derived/current re-registration. It does not erase required earlier-phase
input. C1 made by internal normalization, splicing or reparse is generated
provenance, not new incoming data. Do not count a callback result twice through
clones or relabel emitted bytes as another C1 input. Local pre-union arithmetic
stays equal to its reference and origin/retained unions remain separate.

### Connected replacement-result piece boundary

The replacement-result branch shares one borrowed text/break selector with its
actual XML writer. The selector stores only the remaining input and one pending
separator. It scans each UTF-8 scalar once and retains no event buffer. Its
typed materialization allocates at most one content entry per input byte plus
one empty-text entry, with total owned text bytes no greater than the incoming
value length. The checked result character validation still precedes this
materialization and locked-child dispatch. This removes the generated run XML
and throwaway parse at that boundary. It does not claim that later child
normalization, revision refresh or aggregate reservation is complete. Existing
local admission arithmetic and parser fallback remain unchanged.

Root adjudicated the eliminated intermediate metric boundary. The direct
RunContent replacement consumes only its actual applicable constructor/work
metric before allocation. It does not emit run XML. Later actual source XML
consumes normalization-output demand with publication multiplicity. Hypothetical
run XML removed by this factoring cannot be charged, and final byte transport
cannot debit that later producer twice. Local scalar/reference guards remain
unchanged. No universal typed-byte normalization cap follows.

### Root-adjudicated incremental producer reservation

Root approved incremental checks at each actual applicable producer rather than
a monolithic future topology plan. One inherited operation context must check
every applicable allocation or growth before it occurs. Existing local guards
and aggregate metric arithmetic remain unchanged. Selected frozen ingress facts
are registered once through proved original or retained owner relationships.
Actual clone input is reached for that clone's work, but allocation or reach
does not prove another physical source or current input. Clone, edit and reparse
producers forward identities and transform surviving child slots explicitly.
An ordinal reused after topology changes is no same-owner proof.

The output sink checks actual chunks before Vec growth. Inner escaping,
BytesText, BytesStart and clone allocations retain their own prior applicable
checks. A sink alone is insufficient. Actual retained transport stays separate
from normalization output. Resource errors from the sink must retain their
producer provenance through I/O adapters. No callback replay, hypothetical XML
charge, new heap-byte quota or constructor fallback relaxation follows. Public
state stays staged and atomic. Text, Tab and Break item counts remain constructor
units. Actual text lengths receive their applicable unchanged local current
checks separately. Canonical parser-entry observations must not be relabeled
as complete leaf-action or heap-byte observations.

The model-ingress branch snapshots selected scalar roles before candidate
cloning, using normalization_children rather than the wider stored-resource
graph. Distinct current Field occurrences stay distinct. Reached retained
FieldSource and stored-result records deduplicate through their actual frozen
owner relationship. Deduplication keys are transient and are not a registry
where later allocations can seek credit. No extra model/source Arc is held
solely as an identity anchor, because that can alter COW behavior. Necessary
external scope selection shares the actual retained binding policy and preserves
encoded versus bare record roles. This branch does not claim original OPC
correspondence or initial-parser P overlap accounting.

Direct replacement leaf allocation prepares its cumulative action reservation
before allocation and commits after the infallible materialization completes.
Actions and text lengths are checked separately against the existing local
guard. This model constructor branch remains incomplete until every inner
writer, revision refresh and derived parser carries the inherited context and
actual source-piece output reservation. Existing checked setter payload
validation is itself an emission producer, not a pure baseline inspector. Its
generated annotated-source length cannot become fresh incoming payload credit.

### Root-adjudicated fixed multi-origin and model ingress

Root approved the following exact additive doc-hidden methods on the existing
ScopedWordOperation. No new public type, trait, generic parameter or exposed
fact identity is introduced. Existing new(original) and source behavior stays
unchanged. This is technical reconciliation, not publication approval.

```rust,ignore
pub fn with_inputs(
    originals: Vec<(&'a str, &'a [u8])>,
    current_paragraphs: Vec<&'a CT_P>,
) -> Result<Self>;
pub fn source_part(
    &mut self,
    part: &str,
    range: Range<usize>,
    scope: &BTreeMap<String, String>,
) -> Result<ScopedWordSource>;
pub fn source_model_paragraph(
    &mut self,
    paragraph: &'a CT_P,
) -> Result<(CT_P, ScopedWordSource)>;
```

with_inputs freezes exact borrowed OPC/name and existing current owner
relationships before clone, flush or staging. It remains inert. Cataloging
performs no XML scan, materialization or funding of unrelated entries. Duplicate
part names refuse deterministically. source_part first checks the fixed name,
then validates the requested range before slicing, planning or registration.
Only reached physical ranges in that actual original domain enter its union.
Two different original parts retain different domains even at equal offsets or
with equal contents. Names alone never prove physical input.

source_model_paragraph first checks membership in the actual frozen model
catalog. A foreign owner refuses before cloning. Its actual producer binds
selected existing B/C facts, normalizes and serializes once under the inherited
context, and returns both the prepared model and mapped output. Repeated calls
forward the same frozen input fact identities and ledger. Generated bytes never
become another original, retained baseline or current ingress. Model-only B
remains distinct from proved A unless the actual producer supplies correspondence.
Every surviving child identity follows its actual clone/edit/reparse producer.

Named consumers are render_grammar before clone/template_programs, the current
model preparation/C0 source path, and related-part/note programs plus raw textbox
selection. One checked publication retains the operation across reached parts,
serializers, parsers and revision refresh. It cannot create a ledger per part,
program or item. Default datum/filter provenance remains a separate open
implementation seam, not a capability granted by these three methods.

The final strict SemVer hidden-public inventory must include ScopedWordSource
and its as_bytes accessor, the six previously listed source/format/repeat/finish
methods, and these three methods. All are additive within the prepared minor
release. Existing public signatures remain unchanged. No other public addition
is authorized by this reconciliation. HLD mechanism prose is updated only when
the corresponding implementation exists.

### Root-adjudicated operation-produced text copy metric

Root approved a narrow internal metric amendment for actual operation-produced
text carriers. Their materialization and copy byte demand is tracked separately
under the same unchanged source/current work allowance and inherited context.
This is not a relabeling of A16 physical run/revision-cache fact-copy visits.
Actual text.len is not added to constructor entry counts or claimed as emitted
XML. No generic heap profiler, exact heap-byte claim or new numerical quota is
introduced. Existing local scalar/reference terms remain unchanged.

Reserve before actual String allocation, growth or copy at the concrete native
default-format, template and filter producers. Generated results forward their
input identities and never create new C credit. A frozen datum funds once by
its actual occurrence, while a genuine external callback result follows the
existing once-per-event input rule without replay. Borrowed and in-place
no-copy forwarding do not pretend that a materialization happened. Unchanged
retained transport keeps its existing separate treatment.

The discriminating example is an operation that derives100MiBof text from a
1MiBfrozen datum, then repeatedly copies that generated value through trim
filters before one final XML publication. Individual lengths, constructor
entry counts and final XML alone do not bound those repeated real copies.
No-op trims should forward the borrowed or in-place carrier where possible.
Genuine repeated copying consumes this separate byte-valued demand before
allocation. The existing allowance is not renewed from the generated100MiB.
Controls must preserve old local/simple cases and show actual repeated-copy
refusal before allocation with atomic publication. Any discovered existing
accepted-limit or public-signature difference outside approved additions is
returned to Root at that exact boundary.


### Incremental normalized XML emission boundary

The existing `ordered_paragraph_xml_in` producer uses a concrete private byte sink. Its actual emitted chunks reserve normalization-output bytes against the same inherited operation ledger before `Vec` growth. The selected field keeps its unchanged local `normalization_work_budget` byte limit in that producer. Revision refresh has no additional invented local limit. A sink shares the operation ledger through a concrete private synchronized anchor, including nested serializers. It cannot replace the ledger, register input or renew credit. Ordinary retained source capture and final byte transport do not enter this producer metric. The emitted buffer has the applicable selected local and operation byte bounds. Its bookkeeping is constant space and each chunk is processed once. Inner escaping, source cloning and surgical replacement producers still require their own prior applicable reservations. The outer sink does not qualify those allocations.


### Root-adjudicated retained revision transport

Root rejected the superseded deferred-refresh candidate in favor of preserving the existing actual `try_from_raw` error boundary. An immutable `CT_Revision.raw_xml` record forwarded directly by its Source writer is retained transport, including a record already generated by its actual refresh producer. That direct write retains local output-length and overflow checks. It does not become another normalization-generation producer merely because the containing buffer grows. The actual refresh still consumes generation demand. A private scoped transport receipt binds this one synchronous write to the borrowed record range and advances by actual chunks. Equal copied bytes, unrelated raw writes and nested generated serializers receive no exemption. The receipt is dropped on success or failure. Applicable existing capture, source-copy and text-copy metrics remain separate. A genuine repeated publication still consumes demand at its real producer. No source credit, quota, error-boundary relaxation or generic raw bypass follows.


### Selected model input registration on the scoped operation

The scoped operation owns the same private field work ledger used by actual normalizers and writers. `with_inputs` remains inert. `source_model_paragraph` first verifies the borrowed paragraph against that frozen catalog, before any clone. It then admits the existing model, selects its actual writer and accepted-normalization field inputs, and registers their frozen retained/current roles once. Repeated selections retain prior membership and cumulative demand. Source/P record identity and current occurrence identity remain separate. These model records provide B/C provenance, not original A proof. A derived output inventory can supply constructor policy facts but cannot register its bytes or parsed models as another baseline. The actual paragraph clone, normalization and Source serializer share the operation ledger. The returned carrier has generated correspondence until actual source pieces are forwarded. No generated buffer is called original input.


### Selected frozen result formatting facts

Root adjudicated that actual containing-run properties selected by scalar Field result emission are incoming model facts outside the unchanged local Field equation. The scoped model producer registers that frozen property relationship once before cloning. Retained raw records contribute their actual B byte lengths. Current strings contribute UTF-8 byte lengths. Numeric and enum values contribute canonical lexical byte lengths, using the existing writer spellings and allocation-free numeric formatting. Boolean values contribute the canonical single byte 0 or 1. The existing affine current-value coefficient applies to C, with no XML serialization to manufacture credit. Unselected outer properties, candidate clones and repeated publications do not renew the allowance. This is selected input coverage under the existing metric, not a new quota. Raw-only coverage does not qualify the full current formatting role or all nested routes.


### Initial original current-fact occurrence identity

Root adjudicated a private original current-fact token minted only by the initial successful canonical constructor delivery. The key combines actual proved physical part ranges with the actual canonical marker partition. For a simple field it is the delivered field opening/source occurrence. For a complex field it is the proved physical run plus the canonical scanner's actual raw-child marker slot. The index alone is never permission to register an original fact. Two actual markers in one shared raw run remain distinct current occurrences. Raw source and P still contribute only their physical range union. Failed, raw-fallback, unselected and unproved markers supply no activated input credit.

The original versus produced reader mode travels in the existing private source carrier and cursor. A generated or reparsed reader cannot mint another original token. It may forward an already registered token through an actual surviving source-piece correspondence. COW and grouping forward the token at the real clone/source-edit producer, before the original owner relation is changed. Source IDs, equal bytes, candidate allocation and a shifted field index cannot create a token. Current facts retain the unchanged initial local scalar terms. Activation is separate from inert constructor facts, and repeated reaches reuse the same operation ledger. Complete source-policy coverage and preconstruction producer reservations remain required. This reconciliation adds no public API, numerical limit or independent approval pass.


Retained formatting extra-attribute tuples are genuine stored B records, measured by their actual stored name and value lengths. They are not labeled raw lexical XML. They do not receive the affine C multiplier merely because a namespace declaration survives serialization. Modeled font, color and language strings and canonical primitive values remain C. Namespace C requires the separately proved actual necessary binding role. The Source writer's retained attribute-value syntax is preserved, with actual large unused declaration and escaped URI/unknown-value controls required. Generated escaped XML is never counted as incoming data.


### Inner XML event production reservation

Root adjudicated irreversible reservation of actual inner event materialization under the existing normalization-output metric before payload growth. A later synchronous unchanged event transfer uses a private one-shot prepaid-production receipt bound to that exact produced payload range and operation ledger. It is demand already spent, never input credit. Unprepaid XML delimiters and scaffold are charged normally. Equal copied buffers receive no receipt. Failure or raw fallback after payload production does not refund that spent work.

The concrete first consumers are authored simple-field start events and generated Field instruction/result text. One shared five-character escape policy supplies actual payload length and actual emission, matching pinned quick-xml 0.41.0 including unchanged carriage returns. Simple header recipes have seven bounded optional borrowed pieces. Text emission scans borrowed bytes, reserves the exact escaped payload, then allocates once at the measured length. The transfer receipt has constant space and advances through actual synchronous payload chunks. Each byte is visited a bounded number of times. Existing fieldless transport, local bounds, source-fallback/error policy and scalar equations stay unchanged. Surgical source helpers and other inner producers still require their own connected reservation. No generic heap metric, new quota, callback replay or public surface is introduced. The existing Vec writer and normalized-output sink are the two real destinations of the shared concrete Field writer helpers.


### Bounded ordered-simple source composition

The ordered-simple path whose existing recursive publication would transport the projected source unchanged uses one borrowed final record recipe. The substantive source-metadata predicate is shared with the existing unchanged selector. Normalization and source-range checks keep their preceding order. Nested effective instructions still refuse before simple source-root lookup. The actual lexical instruction selector is shared with the retained fallback, including the first qualified attribute, lexical QName/quote preservation and original iterator errors. A root/result overlap that cannot use the proved original coordinates keeps the old staged path.

The recipe contains at most five borrowed or escaped pieces: retained prefix, optional current instruction value, retained middle, actual normalized result and retained suffix. It composes the corresponding source map from the existing raw and stored-result receipts. Replacement instruction bytes are Generated. The effective current instruction and projected raw length receive the old prospective local publication guard before allocation. One final changed record is reserved irreversibly before growth and emitted once, with exact synchronous mapped/prepaid transfer. No mutable whole-source COW is needed in this selected route. The dirty, lock, typed-cache, legacy-metadata and rebuild selectors remain on their existing paths. Their later composition and the remaining apply path are not qualified by this increment.


### Root lock-publication namespace reconciliation

The actual retained lock writer must preserve the qualified identity of an unprefixed Word owner and any foreign canonical-prefix binding. The existing parser admits a default-Word fldSimple with a nonempty Word attribute alias and foreign w. A lock update cannot append a second xmlns:w or rebind the opaque descendants. Reuse a genuinely bound nonempty Word alias for the new attribute, retaining prefixed-owner, None-removal and unchanged-source behavior. This fixes an inherited writer defect within the existing namespace preservation contract. It changes no public signature, resource equation, new input credit or fallback quota. Focused native-library RED/GREEN and lexical/reopen controls qualify this bounded producer. Other C1 composition, constructor and native/template work remain open.


### Q-C1-03, approved template-copy operand accounting

The user explicitly selected Count selected operands on2026-11-03 in reply to Q-C1-03 and its exact proposal. The template text-copy metric counts64plus canonical scalar-text UTF-8 bytes per actually used operand, deduplicated by real input occurrence. Null and empty scalar values therefore still supply one actual operand. This is an explicit new template operand rule. It is not preexisting FieldArgument credit, a heap-size measurement or a new factor in the local Field equations.

Use the existing current-work multiplier W_NS.len()+64. Actual consumed Word-source/model bytes require original-range or frozen-model proof and count once. Word defaults/separators require their real source/model operand provenance. Unused JSON siblings grant nothing. Computed/filter/imported values never renew input. A real successful callback result contributes once at its validated event. Borrowed, moved and in-place forwarding consume no nonexistent copies. Actual String allocation, growth and copying consume the separate template copy demand before allocation. Template copy credit cannot increase Field constructor/normalization allowances, and template XML output receives no universal fieldless byte cap from this rule.

The user decision resolves the fieldless empty/null join choice. Required evidence includes1000empty and null joins, selected operand deduplication, unused huge siblings, genuine Word literal source proof, callback once, numeric and Unicode parity, no-copy trim and cumulative actual-copy refusal with atomic publication. The prior exact proposal is retained at /private/tmp/rdocx-s96-root-c1-implementation/native-datum-copy-support/Q-C1-03-PROPOSAL.md. Its alternatives are not operative. No implementation or provider acceptance follows from this answer.

Root authorizes the five concrete doc-hidden methods in the source-pinned native datum proposal: with_template_inputs, template_datum_key, template_datum_index, template_scalar_input_bytes and reserve_template_text_copy. The existing workspace serde_json dependency may be added to rdocx-oxml so those methods borrow the actual Value rather than creating serialized identities. Grammar/filter policy stays native. The unimplemented format_paragraph occurrence tuple uses borrowed text/properties, a shared optional callback-event Cell and actual selected JSON dependencies. None covers computed/literal/filtered values without a callback. Existing new/with_inputs methods and public binding surfaces stay unchanged. Add these methods and the earlier rewrite_imported_source bridge to the exact final hidden-public SemVer inventory before completion. No public datum type, numeric identity API, new module, feature or independent quota is authorized. Word-literal receipt forwarding and parser allocation guards still need their actual connected implementation.

The exact additive doc-hidden signatures are:

```rust
pub fn with_template_inputs(
    original_parts: Vec<(&'a str, &'a [u8])>,
    current_paragraphs: Vec<&'a CT_P>,
    data: &'a serde_json::Value,
) -> Result<Self>;
pub fn template_datum_key(
    &mut self, parent: &'a serde_json::Value, key: &str,
) -> Result<Option<&'a serde_json::Value>>;
pub fn template_datum_index(
    &mut self, parent: &'a serde_json::Value, index: usize,
) -> Result<Option<&'a serde_json::Value>>;
pub fn template_scalar_input_bytes(
    &mut self, value: &'a serde_json::Value,
) -> Result<usize>;
pub fn reserve_template_text_copy(&mut self, bytes: usize) -> Result<()>;
pub fn format_paragraph(
    &mut self,
    source: ScopedWordSource,
    children: Vec<(Range<usize>, ScopedWordSource)>,
    occurrences: &[(usize, Range<usize>, &str, Option<&CT_RPr>, Option<&std::cell::Cell<bool>>, &[&'a serde_json::Value])],
) -> Result<ScopedWordSource>;
pub fn rewrite_imported_source(
    &mut self,
    source: ScopedWordSource,
    edits: Vec<(Range<usize>, ScopedWordSource)>,
) -> Result<ScopedWordSource>;
```

Construction freezes the actual JSON root without serialization or eager traversal. Key/index selection proves actual parent membership and selected child identity. Containers and unselected siblings provide no operand funding. Scalar accounting proves actual selected membership before allocation and deduplicates that occurrence. Numbers use allocation-free canonical lexical counting before reserved emission. The format occurrence callback Cell marks only a genuine successful external callback event. None means no callback. The actual native event owns its shared Cell through alias replay, and lower Field input registration consumes it once. Some requires empty JSON dependencies. Imported edits are actual stage-specific producer transcripts. Validate anchor, bounds, nonoverlap and overflow before growth, compose once, carry unchanged source pieces and mark replacements generated. Neither bridge mints original Field credit. Keep no-op transport and existing importer error order. Word literal receipt forwarding remains a separate connected producer obligation, not an assumption from these signatures.


Native template ingress freezes borrowed original package entries and actual current
body/table/control/note/glossary paragraph owners before candidate cloning or
flushing. Reuse the existing paragraph visitor with explicit borrowed-owner
lifetimes instead of creating another walker or snapshot. Its existing read
consumers and the actual input freezer are concrete consumers. Selection is
inert and grants no unvisited model/source bytes. The same operation flows through
main evaluation, raw text boxes and charts. JSON lookup registration alone is a
connected intermediate, not completion of scalar/literal copy reservation or
Field/source publication. Native commit must release every frozen input borrow
before replacing the original document. Preserve native path-kind error order
and deferred lookup behavior while operation failures propagate explicitly.


The existing FieldSourceEdit::apply producer uses the inherited Field operation
context. After verifying the source length, reserve the full prospective changed
source record before allocating its replacement buffer or growing the source
Vec. This is the existing inner-XML production metric, not a new equation or
independent quota. A later metadata/lock rewrite is its own actual production.
No unchanged-copy receipt is fabricated for that subsequent producer. The old
local projected-length/overflow checks and byte-level escaped splice remain.
Preallocation refusal leaves the caller buffer unchanged. Compound source
coalescing and original map forwarding remain separate required work.

Nested-field source transport may borrow the original raw buffer only after its
existing complete scan, span/identity checks and recursive child validation have
finished and produced no edits. Changed output remains owned. Both real callers
must preserve the borrowed transport until an actual later producer requires
ownership. Do not skip validation, mint original credit from an equal copy or
charge an allocation for an unchanged borrowed buffer. Changed nested edits and
whole source recipe/map closure remain mandatory beyond this transport step.


The Word template text-input role uses actual canonical decoded Text bytes.
Union selected decoded UTF-8 intervals once under the actual producer identity.
For source input, the real qualified Text decoder supplies the whole physical
Text-event identity. Its selected decoded offsets are not XML byte offsets.
Entity spelling therefore requires no decoded-to-encoded mapper for this role.
Frozen-model input uses its real immutable owner and content slot with decoded
intervals. Do not add the JSON64term or Field credit to Word literals. Preserve
origin offsets through actual splits, wrappers and projections. Generated
replacements never renew input. A private carrier-owned fact sidecar must feed
the existing placeholder walk directly, with no public CT_Text/CT_P field,
String pointer/equality identity or second semantic validator. This is a
technical reading of the approved Q-C1-03actual consumed Word operand-byte role,
not a new user policy decision. The prior preparation note's lexical mapper
requirement was overfactored and is withdrawn. The full source/model/Generated
producer connection remains to be implemented.

Native first JSON scalar text production admits the actual used scalar and
reserves its measured canonical copy before String construction. Number emission
uses one pre-sized buffer after counting. A borrowed input string's first
trim/case output reserves its actual output length before materialization.
Joining selects each actual array member before its scalar copy. This does not
qualify final join allocation, owned generated filter growth, Word literals,
parser strings, callback or later execution/formatter copies. Those remain
mandatory. Until Word text input credit is connected, no complete native copy
policy acceptance or provider gate may be claimed from these first-copy checks.

Ordered simple-field flag composition must preserve the actual existing root
rewrite policy. With unchanged typed/legacy/child metadata, allow dirty/lock
changes in the borrowed recipe. Any such flag change reaches the old metadata
root rewrite, even if only lock changed. A nonforeign w binding yields a
canonical w root/end, canonical xmlns:w/instruction and normalized dirty. A
foreign w binding preserves the owner QName and instruction attributes, with
only the actual first instruction value changed by the earlier lexical edit.
Its existing dirty-prefix selection remains. The later changed-lock policy
removes qualified old locks and appends the actual selected owner/attribute
prefix. Keep imported unrelated attributes and the original error order.
Perform the old post-result/instruction local guard before metadata-parser
failures, then reserve the one final record before its materialization. Empty,
ambiguous or unsupported source geometries retain the old route rather than
inventing a new fallback refusal. The actual normalizer and namespace publisher
retain their separate responsibilities.


The bounded ordered-simple flag producer shares the content-metadata predicate
with the existing unchanged selector, then handles dirty and lock differences
without mutable whole-source COW. Header attributes remain borrowed pieces.
The actual namespace walk reads original header, normalized result and original
closing through WordXmlInput, using word_prefixes_at for every reached Start and
Empty event. Instruction value escaping cannot change that namespace context.
Root metadata errors follow the old post-result/instruction local guard, and
lock-prefix errors follow the full descendant walk. Canonical root and actual
End are rewritten together. Foreign bindings retain the original owner and
first-instruction lexical identity. Literal quote ambiguity in a later lock
header and unsupported root geometry retain the old staged path. One final
source record carries original/result slice maps and generated header pieces.
This closes only the selected no-typed-cache, unchanged legacy/children route.
Changed nested/typed/legacy composition and whole C1 remain open.


Ruby property prefix qualification shares the actual child's decoded scalar
outcome with CT_RubyPr::read_child. Successfully modeled Empty children preserve
positive prefix proof, including existing permissive alignment/toggle spellings.
Malformed or missing half-point values remain raw rather than errors. Started children and unproved raw captures remain unknown. Never decode a
started property's scalar before its complete capture, or a later child after
the actual property reader has returned None. Earlier real errors still prevent
a later None from discarding constructors. No fresh allowance is introduced.


### Selected JSON dependency bridge for actual Field formatting

Root reconciles the format occurrence tuple with the already approved genuine
incoming current-value role. Its sixth field is a borrowed slice of actual
selected scalar Value references from the same frozen operation. This supplies
the missing provenance connection without a new public type or numeric identity
API. It does not change either allowance equation. Direct scalar evaluation
retains the actual selected leaf. Trim/case filters forward its dependencies,
and join records the real selected scalar array members. Containers, generated
filter results, unused siblings and default Word literals add no JSON dependency.
Execution/loop/MCE/copied-note carriers forward references rather than re-resolve
paths or discover inputs by text equality.

Before expanding an actual affected Field cache, validate anchors, scalar slot,
nonoverlap, all datum memberships and scalar kinds. Activate only canonical
UTF8input bytes under the existing Field current-value multiplier, once globally
per real frozen Value occurrence. Template64never enters the Field ledger.
Equal but distinct actual JSON nodes remain distinct. Replaying one actual
datum into multiple Fields does not renew input credit, while every actual
producer still pays its construction/output demand. Fieldless/chart/raw-textbox
uses do not activate Field current-value credit. Stage all validation before
mutating membership or growing output.

When a genuine callback supersedes the prepared JSON text before Field
publication, pass an empty dependency slice with the true callback flag. Do not
count both unused original JSON and its replacement as new current input. Only
an earlier actually consumed Field phase can retain its already registered C0.
Computed/literal/filter values still pass false, with genuine JSON dependencies
only where actually present. Callback event deduplication remains separate.
The doc-hidden tuple revision enters the existing final SemVer inventory. No
implementation or acceptance of this bridge is claimed by this reconciliation.


The concrete template join producer retains actual member/scalar selection and
error order, then checks complete joined output length and reserves it before
String allocation. Its separator contributes actual output demand but no fresh
JSON operand credit. Upper/lower reserve measured standard conversion output
for both borrowed and owned intermediates. Owned trim transports its existing
buffer using the current in-place edit and does not invent a new allocation.
No filter result is re-admitted as input. Exact-exhaustion controls retain the
1000empty/null positive and standard contextual Unicode conversion controls.


### Empty raw Ruby property capture versus deferred completion

Root extends the same actual property outcome proof to successful Empty Raw
children. The real capture_empty_element writes the exact Empty event into Vec,
after the actual namespace/scalar selection has succeeded. It does not decode
ordinary raw attributes or traverse descendants. Keep known prefix state at
this point and mark deferred fragment completion using the RubyProperties
frame's existing content_seen flag. No new quota or source record is granted.
A later direct-character None bypasses completion in the actual factory, so it
may stop future constructors. With no None, the Properties End makes prefix
proof unknown because raw-fragment completion remains unproved. Started child
capture and its later scalar read stay unknown. A real reached scalar error
before None still fails, and no scalar is decoded after the recorded None.
Missing/malformed half-points and foreign Empty properties exercise this path.


### Actual grammar input selection for template copy work

The actual visible-text grammar producer may select each complete decoded Text
interval it visits while aggregating parser stretches. The native parser reads
all those returned strings to discover markers/tokens, including non-token
bytes. This is consumed Word input under the unchanged approved metric, not
only an eventual default literal substring. Register the actual original
Text-event or frozen immutable owner/content slot before aggregation grows its
String, and union each decoded interval once. Skipped descendants, opaque
payloads and unvisited siblings provide none. Row/direct/all visits retain
deduplication. The first canonical lower Text input constructor is input
formation, while later aggregation/lexer/unescape/default/execution copies are
demand. Generated/default output creates no new credit. With actual grammar
input selected, a separate literal JSON-unescape provenance mapper is not
required. Original/source-model producer lineage remains mandatory. No new
public callable is approved by this selection clarification alone.


Root adjudicated the callback event field as Option<&std::cell::Cell<bool>>
in the existing six-field format occurrence tuple. None means no external
callback. Some refers to the actual native Rc-owned event shared through
clones and MCE aliases. False means not activated for Field input, true means
consumed. Fieldless formatting leaves it false. Successful actual affected
Field input registration flips it once. Later failure does not refund that
activation. Some requires empty JSON dependencies. Native must not reset or
reconstruct the event on replay. This replaces the earlier bool classification
in the additive doc-hidden SemVer inventory, without a new public type or quota.

The lower producer shares a private borrowed prepared-occurrence kernel with
the existing owned public replacement method. The scoped path threads its
actual WordOperationWork through wrapper content serialization, simple Field
reparse and final normalization/Source emission. Raw wrapper origin requires
actual parser-produced capture maps. Existing WordProjectionSourceMaps.extras
is the candidate seam to expose privately from the same parser, preserving
its exact retain/reorder decisions. No byte search or equal-text registration
may substitute for that producer map. Wrapper content maps use actual content
slices and generated paragraph shells. New replacement text stays generated.


The approved private parser seam extends
CT_P::from_xml_fragment_admitted(xml, controls, extra_maps) and
CT_P::from_xml_with_revision_facts(reader, prefixes, root, revision_facts,
control_facts, extra_maps), where the final optional argument is a mutable
WordExtraSourceMaps companion. This crate-private alias contains optional Arc
WordSourceMap entries aligned to the final extra_xml vector. Actual projection
and paired retain decisions transform both. Generated root records append None.
The existing public entrypoints request no companion. Scoped formatting and
actual recursive wrapper parsing consume it. No public model field is added.


### Actual parent-Field text emission context

The operation-wide field_output_metric flag does not classify every ordinary
paragraph Text event as Field output. A private FieldXmlEmission holds only the
shared normalized-output ledger and its existing enabled state. Actual ordered
Field result paragraphs and typed/new cached-run producers supply this context
to the existing run emitter. Ordinary model-source paragraphs supply None even
when their operation has another Field. This avoids a generic fieldless cap.

The substantive private run and paragraph-content emitters accept the optional
context. Mixed cached-run segments forward it from their actual parent. Reuse
the existing five-character escape sizing/emission and exact one-shot prepaid
payload transport. Reserve escaped Text/DeletedText before its owned buffer
allocation. Borrowed unescaped text still reaches the existing checked sink
before growth. Names/attribute/property builders, raw revision stages, control
context propagation and final formatting integration remain separate mandatory
producer work. No whole-wrapper sizing replay, new quota/public type or fresh
source/current baseline follows from this private context.


The next concrete property primitive is a qualified single-val Empty event.
FieldXmlEmission::write_word_value measures its element QName, qualified val
name, escaped value and exact attribute punctuation, reserves that payload
before QName formatting/attribute escaping/event-buffer growth, then emits the
unchanged quick-xml event with exact payload prepayment. The checked sink still
charges delimiters once. RPr style, simple numeric and simple enum val events
share it. Existing fieldless callers retain the ordinary producer and no cap.
RPr root namespace construction, fonts/color/language retained multi-attribute
records, toggles, compound properties and positioned-raw stages remain mandatory
work. This does not qualify a whole RPr serializer or modify Field input terms.


Modeled-only RPr root/font/language events use a shared exact generated-Word
name plan before variable QName or attribute allocation. Existing generated_word_start
uses the same selector after its actual local-declaration overlay. Preserve
Source context versus empty-original selection, any foreign-w test, first
nonempty Word alias order, exact fresh-name collisions and checked index growth.
Publication RPr keeps its existing fixed-w/no-new-declaration rule. Child font
and language plans use the original caller context, retaining existing redundant
declarations. No sorting, inferred root-scope update or fresh input credit.

FieldXmlEmission measures the planned QName, optional namespace declaration and
actual borrowed modeled attributes before rendering the prefix or growing an
event. Fixed-size optional attribute slots carry no XML buffer. Empty retained
attribute collections select this producer. Nonempty retained collections and
color/toggle/compound/positioned-raw/control stages remain explicit mandatory
work. Existing non-context callers keep their ordinary serializer.


### Genuine callback input for the separate text-copy ledger

Root technically reconciles Q-C1-03 with its two actual input roles. Callback
text contributes its real UTF8 bytes times the existing template-copy multiplier
once per actual successful validated external event. The JSON operand64term
is not assigned to callback or Word inputs. This is Root's explicit narrow
interpretation of the existing selected-input rule, not a claim that the user
verbatim specified callback arithmetic and not a new quota. Field equations
and their separate input roles remain unchanged.

The additive doc-hidden ScopedWordOperation method is:

```rust,ignore
pub fn register_template_callback_input(
    &mut self,
    value: &str,
    delivered: &std::cell::Cell<bool>,
) -> Result<()>;
```

Native creates one private Rc-owned event with separate copy-delivery and
Field-delivery Cell<bool> values, both initialized false only after actual
callback success, existing character validation, and the existing complete
property probe. Preserve its exact namespace closure, strict checked parse,
error shape and callback order. The ordinary temporary XML/model validation
path retains its existing guards. It is neither Field output acceptance nor
a Q-C1-03 named native text-carrier copy. No generic heap exemption or profiler
is introduced by this existing metric boundary.

The earlier prepared-text clone into the callback carrier still reserves its
actual copy against the real preceding Word/JSON inputs. The validated result
then moves into the prepared carrier. Register actual callback copy input once
before its first later native carrier clone, and debit every actual such copy
before allocation. Do not charge moves, replay the callback or fund retroactive
carrier demand. Non-emitting probe work grants no Field/source or JSON input.

Copy registration validates checked addition before modifying allowance or
consuming its delivery Cell. A delivered immutable event is inert. Existing
copy refusal stays sticky. Shared aliases/MCE replay retain the same Rc and
flags without resetting them. A separately invoked real callback for a copied
physical note creates a distinct event. The existing format six-tuple receives
only the Field-delivery Cell, which activates only for an actual affected Field.
Copy delivery on a fieldless path does not consume Field eligibility.

Required controls cover once-event alias reuse, distinct events, independent
flags, fieldless copy before later Field use, no64for empty callback, refusal
stickiness/overflow, no event on failed full validation, unchanged error/callback
order, and every actual native carrier-copy site. Include the actual returned
signature in the final hidden-public SemVer inventory after implementation.
This technical reconciliation alone implements or accepts none of the bridge.


### Native Word grammar carrier seam, Root approved

Root approves this technical carrier contract on2026-11-04 after reviewing the
exact V2 supplement, actual source call order, substantive writer boundaries and
incoming Field scalar authority. This is design/API authorization, not source
implementation, another design peer, provider completion or native acceptance.
Existing methods and signatures stay unchanged. No new public type, model field,
module, trait, generic framework, origin-ID API or quota is authorized. The eight
additive doc-hidden methods below belong to existing ScopedWordOperation and
must enter the final Rust SemVer inventory and actual external compile canary.
CT_Text and Field name the existing crate::text types. Range and BTreeMap are
the standard-library types, not new aliases or public carriers. Current cache
membership is Word-only input, distinct from Original OPC and Field facts.

Source preparation identity: V2 patch SHA256
eb89844bdf21ffe63dc794cfb0c76f19dcf7d11472e0863d1c529288c1075b54,
applied to plan5b81417f71c0c894857622fd6542cbf0e6835909fa4f56ac4adfb36350369b52.
The exact normative contract is inline here. Temporary proposals and source
line observations are evidence custody, not a second plan authority. Current
emitter work is unaffected. Native implementation must consume this contract
before starting. Issue430 V13 remains separately unadopted and unimplemented.


The Native private `TemplateWordLineage<'a>` is operation-local correspondence,
not funding. At the actual staging clone producer, it records structured
owner/control/table/run/content paths to actual frozen borrowed `CT_Text` and
needed Field input facts. It includes direct block-control runs as well as
paragraphs, tables, notes and glossary owners. At an actual Field-only run
replacement it forwards ordinary sibling Text slots and the actual surviving
Field facts. Changed or newly inserted output without that producer witness is
Generated. No later String pointer, ordinal, hash or value equality discovers
an origin. Native resolves paths under an immutable candidate borrow immediately
before the substantive writer, producing short-lived current/frozen pairs.

```rust,ignore
pub fn with_template_word_inputs(
    original_parts: Vec<(&'a str, &'a [u8])>,
    current_paragraphs: Vec<&'a CT_P>,
    current_bodies: Vec<&'a CT_Body>,
    current_texts: Vec<&'a CT_Text>,
    current_fields: Vec<&'a Field>,
    current_word_sources: Vec<&'a [u8]>,
    data: &'a serde_json::Value,
) -> Result<Self>;
pub fn source_model_body_item(
    &mut self,
    frame: &CT_Document,
    item: &BodyContent,
    ancestors: &[&[(String, String)]],
    text_origins: &[(&CT_Text, &'a CT_Text)],
    field_origins: &[(&Field, &'a Field)],
    raw_origins: &[(&[u8], &'a [u8])],
) -> Result<ScopedWordSource>;
pub fn source_model_document(
    &mut self,
    document: &CT_Document,
    text_origins: &[(&CT_Text, &'a CT_Text)],
    field_origins: &[(&Field, &'a Field)],
    raw_origins: &[(&[u8], &'a [u8])],
) -> Result<ScopedWordSource>;
pub fn source_model_notes(
    &mut self,
    notes: &crate::footnotes::CT_Footnotes,
    endnotes: bool,
    text_origins: &[(&CT_Text, &'a CT_Text)],
    field_origins: &[(&Field, &'a Field)],
    raw_origins: &[(&[u8], &'a [u8])],
    prepare_drawing_namespaces: bool,
) -> Result<ScopedWordSource>;
pub fn source_model_glossary(
    &mut self,
    glossary: &crate::glossary::CT_GlossaryDocument,
    text_origins: &[(&CT_Text, &'a CT_Text)],
    field_origins: &[(&Field, &'a Field)],
    raw_origins: &[(&[u8], &'a [u8])],
) -> Result<ScopedWordSource>;
pub fn source_current_word_record(
    &mut self,
    record: &'a [u8],
    range: Range<usize>,
) -> Result<ScopedWordSource>;
pub fn compose_sources(
    &mut self,
    pieces: &[(Option<&ScopedWordSource>, Range<usize>, &[u8])],
) -> Result<ScopedWordSource>;
pub fn close_source_namespaces(
    &mut self,
    source: ScopedWordSource,
    scope: &BTreeMap<String, String>,
) -> Result<ScopedWordSource>;
```

The new factory freezes actual membership without inspecting or crediting all
text. Existing `with_template_inputs` is unchanged. Frozen current-source slices
must be complete actual held cache, serialization-proof or incoming model-
retained XML records selected by Native. Fabricated part names and arbitrary
generated buffers are excluded. Exact duplicate
registrations share membership. Overlapping views of the same backing record
must not become independent origin domains. Native retains the owning Arc for
the operation lifetime and presents that ingress-frozen record only from the
actual selected transport branch. This CURRENT SOURCE role grants Word text
provenance only, never Original OPC/A or Field C0 eligibility.

Before any writer/output growth, validate frozen membership and duplicate
current Text/Field keys in the pair lists. Current borrows and list borrows are
shorter than `'a` and remain immutable through the synchronous actual writer.
An invocation-local address index may identify those borrowed current objects
only during that invocation. It is not a persistent current-pointer catalog.
The lower membership check does not itself prove pairing: the Native actual
clone and mutation producers supply that correspondence. Field pairs are
consumed by the existing real Field source/normalization input bridge, forwarding
already eligible frozen facts rather than registering candidate values. They
must not be unused API placeholders or supply Word cache bytes to Field work.

`source_model_body_item` shares the exact existing
`CT_Document::body_item_xml_with_ancestor_declarations` implementation, including
ancestor-envelope admission, actual part root, foreign-w paragraph Source branch,
and ordinary Publication branch. It neither adds normalization nor substitutes
a standalone paragraph serializer. The whole-document method shares the exact
existing `write_part(PartWrite::Whole)` path, including root bindings, declaration
and actual w14 insertion positions. Notes share the exact existing
`CT_Footnotes::to_xml_footnotes` or `to_xml_endnotes` substantive writer selected
by the existing family. The Native single-note fragment producer uses the same
existing one-note container writer and actual parsed fragment interval.
Glossary shares its existing `CT_GlossaryDocument::to_xml` substantive writer.
No forced reserialization of retained records is permitted to manufacture tags.

`ScopedWordSource` privately carries Word text event/decoded-interval facts
beside its existing Field correspondence and constructor facts. Typed Text tags
are created at the actual CT_R Text emission, from the supplied current/frozen
witness. Internal Field cloning, cache selection or normalization forwards its
own actual source and Field facts before the selected writer changes owner.
Raw Field/wrapper/revision transport uses actual source pieces, not a match to
rendered or cached display text. No new original or current role is inferred
from a parsed Field cache or an equal duplicated representation.

`compose_sources` is the one slice/assembly method. For Some(source), bytes must
be empty and range selects that same-operation carrier. For None, range must be
0..0 and bytes are actual newly generated output with no origin. Validate every
operation, bound, overflow and layout before growth. A single Some piece serves
slice callers. Namespace closure shares the existing exact closure plan and
forwards surviving facts through its actual insertions. Constructor facts are
transformed only under their existing prepared-context proof or requalified
through the actual reader. Neither composition nor closure renews allowance.
Actual Field materialization/transport reservations remain applicable, distinct
from lazy Word text selection and subsequent native text-copy demand.

#### Actual caller and closure matrix

| Actual producer or consumer | Required connection and discriminator |
|---|---|
| render_grammar ingress and clone_for_staging | Freeze original parts, current model owners and preexisting current records before cloning. Build the private lineage at that actual clone, without credit. Unsaved Text absent from OPC must retain its frozen owner. |
| prepare_current_field_models | Forward ordinary sibling slots at the exact run replacement. Changed Field outputs retain only actual eligible source/current facts. A generated equal sibling receives no witness. |
| MainPartBlocks::serialize changed-child branch | The new body-item method is the actual emitted byte producer, with real ancestor declarations. Preserve foreign-w and opaque root attributes byte for byte. |
| MainPartBlocks full writer fallback | The whole-document method is the actual fallback producer. Carry Text tags through exact root/w14 insertion and scanner body/item spans. No whole-writer-only shortcut omits changed child emission. |
| Preexisting cached Arc blocks and serialization-proof copy | Transport the actual frozen CURRENT SOURCE record. Select an exclusive source/model origin role only for each final initial program source before parse_program. Cache buffers produced during this operation retain prior origins and never become new input. |
| Program item slicing and row/direct/all visits | Use same-operation carrier slicing/assembly, namespace closure and actual physical item spans. Repeated visits union one origin. No output ordinal or equal-byte lookup establishes identity. |
| Actual CT_R parser and placeholder TextWalk | Attach facts only when actual decode constructs RunContent::Text, including its actual consumed text span and decoded offsets. Custom note marks and unselected opaque owners are not stamped Text. The same walk registers visited intervals before String aggregation grows. |
| Simple Field and smartTag/customXml wrappers | Forward actual captured maps, content slices and generated shells into recursive parser companions. Follow wrapper parents during selection. Changed/complex cache displays not visited by this walk gain nothing. |
| Retained revisions and actual refresh | Unchanged raw transport carries its real source facts. Actual refresh serialization, wrapper assembly and reparse forward emitted facts. Do not recover cache owners by text equality. |
| Footnote/endnote additions and whole related-model flush | Use the notes writer on actual witnessed owners, then carrier fragment slicing/append. Preserve old family root, separators and error order. Existing imported XML takes its actual source route. |
| Related cache refresh and glossary flush | Reparse with the carried event/slot facts, and forward them at actual model installation. The glossary method covers the real typed flush. Unrelated comments/opaque parts remain unselected. |
| Raw textbox prepass | Preserve its actual program-range exclusions and original-only token discovery. Its consumed input is disjoint from checked program ranges. New replacement pieces are Generated. |
| Structural replay, copied notes, loops and MCE aliases | Forward ParsedProgram/ProgramExecution facts and shared callback events. Copied notes clone existing execution rather than parse_program again. Every actual copy pays demand, while original Word intervals never reactivate. |
| Default and separator operands | Their grammar strings use the selected actual original/current Text facts. Decoding or filtering those literals creates generated demand, never JSON64 or new Field input. |

Required deciding cases include equal original/generated text, unsaved model
edits, ordinary Text beside a normalized Field, actual block-control runs,
partial-body and whole fallback parity, cached current-source choice, repeated
row/table/alias visits, unchanged raw wrappers, revision refresh, note append and
cache reopen, glossary input, and copied-note execution without new parse credit.
A reopened Document forms a new genuine input domain. Each case observes actual
Native grammar aggregation and cumulative refusal before later copy allocation,
not merely a helper map. Original Field equations, no renewed generated funding,
full schema/native/installed/hosted riders and whole C1 obligations remain intact.

These signatures and actual writer hooks are authorized together. If an exact
notes/glossary or raw-refresh caller cannot consume the
proposed carrier without changing its existing byte/error route, that exact
producer blocks the complete batch and returns for a specific amendment. It is
not silently classified as ordinary typed Text or declared complete by the
main-body tests. No implementation or acceptance follows from this supplement.


#### Incoming Field scalar Word origins, distinct from Field work

The existing `field_origins` pair parameter has a concrete Word-text consumer.
A newly authored `Field::new("CUSTOM", template_text)` may hold its incoming
Word grammar solely in `cached_result`, with no frozen CT_Text. Do not call the
later generated Text uncredited merely because it was not a CT_Text at ingress,
or credited merely because the surrounding Field has C0. A private Word-only
origin identifies the actual frozen Field, its `CachedResult` scalar slot and
a decoded UTF8 interval. It supplies only the actual bytes selected by grammar,
with neither template64 nor any Field A/B/C/P input term. The factory's actual
frozen Field membership is inert until this scalar is selected by its actual
writer route. Duplicate cache representations and equal text are not new slots.

At the actual Native clone and Field-only replacement producers, Field pair
facts carry the selected incoming slot relationship, not every future value of
that Field. Lower internal Field clone producers forward the same scalar origin
explicitly while the cache authority is unchanged. An incoming model cache edit
made before operation ingress belongs to that frozen scalar. A cache computed
or replaced after ingress is Generated unless its actual producer forwards an
existing source/decoded interval. Never attach the old scalar origin to all
future `cached_result` values merely because the Field pair survives.

The actual authority sequence is `write_field_with_result_properties`, followed
by its retained/ordered/New form selection and `write_field_result_run`.
`typed_cache()` wins when its existing display predicate selects those runs.
`new_result_runs()` is the next physical-run authority. Those paths use actual
Text or nested Field facts in the selected runs. They do not also activate the
parent scalar origin. Only the scalar fallback selects the current incoming
`cached_result` slot. Empty PAGE/NUMPAGES fallback text `"1"` is generated and
receives no origin from the empty cached_result. Existing equality predicates
remain the semantic authority selector, never the provenance-discovery method.

The scalar event producer is `write_field_result_content_inner` and its actual
`FieldResultParts` traversal into `write_field_text_with_source`, not only the
ordinary CT_R Text branch. Forward a private borrowed selected-scalar witness
through this exact call. Preserve each actual part's decoded offsets from the
selected incoming cached_result. Tab and break parts are not Text intervals
and supply no Text bytes to a Text-only grammar walk. Do not construct a new
public FieldResult variant or recover offsets by searching emitted strings.

The same witness must reach the actual scalar replacement producer in
`update_simple_field_source` and `update_complex_field_source`, including their
first-result and missing-result branches. Newly emitted Text XML is generated
for Field-source ownership but carries this separate Word-only incoming scalar
origin when that actual branch uses the frozen cache. Existing raw source
pieces retain their own origin instead. Both roles remain separate in the
carrier so this never manufactures Original A or renewed Field current input.

For ordered normalization, `ordered_normalization_plan.cache_override` selects
`requested` from the incoming Field cache. An actual
`replace_ordered_result_text` copies/splits that requested scalar into the
replacement runs and must forward its intervals there before subsequent
serialization/reparse. Normalization derived from child changes instead keeps
the child Text/Field source facts actually used. The later parent
cached_result/derived_parent_cache assignment from accepted_text is a computed
summary, not a new scalar input origin. Existing output/source recipes carry
surviving facts at their actual edit and parse producers, without comparisons
to equal display strings. Unchanged Parsed Simple raw transport is the separate
retained-source route, not this scalar fallback and not a second cache credit.

Required discriminators add a newly authored CUSTOM Field whose only cache is
a default/separator template string, a dirty incoming model cache assignment
before ingress, an explicit incoming cache override normalized into ordered
runs, selected typed-cache precedence, an unchanged Parsed Simple raw template,
and a fresh computed display equal to a frozen scalar. The computed duplicate
gets no renewed input. Repeated writer/row/alias reads and duplicate cache
displays union the same selected scalar interval once. All tests reach actual
Native grammar aggregation, including note/glossary consumers where applicable.
Field and Word counters must demonstrate their independent input roles.

`compose_sources` describes only an ordered physical producer transcript of
surviving carrier slices and newly produced bytes. It is not a new generic XML
validator. Existing original/current Field guard coverage, actual constructor
fallbacks and all other C1 resource debt remain separately required. These
specific additions are part of this approved bridge contract and supply no
implementation or acceptance proof.


### Deferred incoming result roles, Root approved

Root approves this complete private-role reconciliation on2026-11-04 after
reviewing the actual omitted OUTER cache authority, the original mapped2/mapped3
refusals, initial source constructor and frozen-model entry paths, and existing
Field/property input equations. This is technical design authorization only,
not another peer, implementation acceptance, provider completion or permission
for a conservative refusal. It adds no public method, type, model member, quota
or per-node fixed term. The local helper/reference equation stays identical.

The exact approved behavior is inline below. Source preparation SHA256 is
763d9d2b3b0ef30b181d0c3944758b5dad5686c641602c5bd5b039fbd060b599.
Original source and frozen model each require their own actual ingress proof.
Produced carriers and Original A correspondence alone create neither role.
Existing successful-constructor facts remain unchanged, while actual previously
omitted ordinary result operands may supply their existing C term once when
selected. The valid raw-only reproductions and their model equivalents must
pass actual end-to-end operation tests before this class is accepted.

#### Existing equations retained

The actual Field current helper is normalization_input_work (text.rs:4582), calling initial_field_work_bytes (2026) and legacy_field_work_bytes (2059). Its cache operand is cached_result.len, with no64perText. initial_field_work_base (1999) is raw instruction + instruction name + cached display + supplied scope. Argument/switch64 terms stay restricted to their existing instruction roles. affected_source_dynamic_work uses the unchanged124coefficient. The local Field helper/reference MUST remain numerically identical.

The actual selected property helper is CT_RPr::selected_input_work (run_properties.rs:660). It computes retained-owner-work(B, empty-scope, C), where B includes existing retained source/attribute records and C is selected_current_bytes (710): actual UTF-8 strings, canonical lexical numeric/enumeration bytes, existing1byte bool representation, and2hex bytes for tint/shade. No serialized RPr or sizeof estimate is an input. Unknown retained namespace/attribute records do not become affine semantic C.

At operation level, selected original role addition has the existing form deltaA_union +124*deltaC. Legitimate frozen-model B/C remains the existing separate ingress route. A raw record reconstructed during a source parser is NOT a fresh model-B baseline. New deferred source property bytes already covered by A cannot also add the same retained B again.

#### Exact role inventory and owner

Keep two distinct private identities:

- Original Field identity/facts for that actual successful physical Field occurrence, with the unchanged original Field helper current.
- Original result-operand identity bound to the NEAREST successful containing Field result owner, actual original run/atom occurrence and concrete role. This is supplied by original capture/constructor proof, never an index after edits or Original pieces alone.

The result role can be a small private enum with Scalar(content_slot) and Properties for an actual original run occurrence. Its key includes the containing original Field key, original physical run range/partition and the actual first constructor's stable content slot. An ordinal is valid only as a component of that canonical factory delivery proof. Do not add a generic graph, public ID or source_id identity.

A provisional capture receipt carries the result-owner frame until that containing Field actually succeeds. At successful delivery bind it to that Field's original fact and record the cache authority actually selected by the constructor: ordered accepted text versus direct run cache versus skipped raw result. Late None/error does not bind or fund it. These facts stay inert until a selected producer consumes the corresponding original result input.

#### Scalar measurement and nested exclusion

After actual original run/Text construction under that receipt, retain borrowed canonical display length facts, not a second source text decoder. For ordinary Text/DeletedText selected by the actual accepted/current mode, use CT_Text.text.len. Existing cached_run_display (text.rs:6537) and CT_R::content_text (7580) define the other original cache spellings: tab/newline/carriage return one UTF-8 byte, page/column break one byte with their cache control code, selected custom note marks their actual string bytes, and font-encoded symbols/drawings/comment markers no display bytes. Reuse those concrete policies. A first implementation must not silently assume every RunContent variant is Text.

STOP at RunContent::Field. Do not count its cached_result/projected_text as a new deferred scalar of OUTER. It has its own original Field current authority and may have its own deferred result owner. Also do not use paragraph.accepted_text, CT_R::text or cached_run_display's aggregate String as the new registration input: those aggregate helpers include nested Field display. The new visitor uses their same NONFIELD scalar selection decisions only.

This exclusion does not rewrite historical base terms. An original Field constructor that already aggregated a nested display into its actual cached_result retains that old local current term exactly. It receives no extra deferred role for that same included operand. The supplement prevents new duplicate funding, not a retroactive change to the local equation.

For raw ordinary text omitted from initial OUTER.cached_result, activate its genuine selected original scalar bytes once. For an ordered OUTER whose accepted-text cache already included that same physical scalar, mark its deferred role covered by the base authority and add zero. Do not decide coverage by equal output text. The original constructor branch and actual physical operand membership prove it. Equal strings at distinct physical occurrences remain distinct original inputs. Multiple edits/spans or split runs within one original atom do not renew it.

#### Properties measurement and selection

Use the SAME actual source writer's property selection. Original decoded run properties become eligible only when the selected producer actually preserves/clones/emits them as input, before the relevant generated allocation. Parsing an unrelated wrapper does not finance its properties. Existing explicit outer-result properties versus cached fallback selection must not be replaced by a second guessed selector.

Factor selected_input_work's EXISTING measurement into a private substantive terms method, e.g. selected_input_terms(&self) -> Result<(usize,usize)> for retained B and semantic C. selected_input_work then applies the exact old retained_owner_work formula. The two actual consumers are existing frozen model ingress and new deferred original-source selection. This is a numerical decomposition of the current helper, not an additional formula.

For source receipts, use actual original property-record physical ranges to intersect the operation A union and take current C only from genuinely selected original typed semantic values. Do not add the returned aggregate B+124C blindly when its B is already original A. Retained name/value tuples created by parsing are not extra incoming stored-source B. If a retained item has no exact physical correspondence and did not exist at checked-model ingress, record that missing proof instead of treating the new tuple allocation as baseline.

A later supplied replacement RPr is genuine new caller input under its approved formatting role, not original properties. Old original properties may fund a prior required phase only if that phase actually consumed them. Unused/superseded old properties cannot be added simply to finance their replacements. The earlier C0/C1 phase-lifetime clarification applies unchanged.

#### Narrow private API/data responsibilities

1. The actual simple/complex constructor creates a provisional ResultInputFrame for its actual result region/context and selected cache authority. Original Raw capture records a view of it beside WordExtraSourceMaps. At successful containing Field delivery, registry binds that frame to the physical Field fact. No public Field/CT_P sidecar is added.
2. WordFieldInput and captured cursor transport the deferred view through actual slice/shell/capture operations independently of initial_delivery=false. The registry stores inert source-role facts keyed by actual original atom/run occurrence and owner. Actual run construction may report canonical scalar lengths and property physical spans through this private cursor/return-out seam.
3. Existing pre-edit mapped text walk builds the actual original wrappers under their real context. It stages successful deferred scalar/property facts but does not activate unrelated parents/operands just because text_runs visited them.
4. The selected-match/Field-ancestor pass in placeholder.rs:171-205 already knows which Field producers will run. Extend its substantive registration to consume the selected containing Field/result inputs needed by those actual producers, before candidate cloning/splitting and generation. Use the proven original result-owner receipt, not candidate addresses. If a whole selected Field source phase actually consumes other original cache runs, that actual producer may activate those original inputs once. An unread unrelated raw branch remains inert.
5. A private WordOperationWork method such as register_original_result_inputs(owner_receipt, actual_selected_role_views) performs checked A-union/C-dedup updates under the existing ledger. Private registry methods resolve actual successful delivery proof and previously covered base roles. They do not compute authority from arbitrary supplied lengths.
6. Derived candidate clones, split Text, wrapper splices and refreshed Field reparse forward consumed role keys only. They cannot call the original-result registration with a new role or recredit their generated cache. set_simple_field_source_in forwards old original facts only. Actual source/current maps remain independent from authority.

These are authorized internal signature shapes in the existing modules, not public additions. The same actual parser/writer policies must supply role/view information. If a path cannot produce the proof, keep that exact path open rather than apply a permissive map-only fallback.

#### Context, changes and errors

The capability distinguishes original capture context from actual consuming context. Raw wrappers use detached implicit w plus local declarations. Stored simple Field result uses its actual parse scope. A borrowed unchanged capture preserves the same receipt. Generated/changed source does not mint one. An independently captured unchanged child can retain its receipt across an outer edit only when the actual producer forwards the same child and unchanged reader context.

For an unread outer wrapper edited before any original child/result role was proved, Original map fragments alone do not establish original dispatch or namespace meaning. That remains a compatibility obligation needing original producer facts, NOT an approved new refusal boundary. This supplement grants no permission to reject formerly accepted valid operations.

Ordinary malformed/unsupported None never creates original result facts. Resource errors retain actual producer classification and propagate in existing order. Already reached successful deliveries and spent work remain sticky after later failure. Public candidate publication remains atomic.

#### Complete discriminator set

- Exact mapped2/mapped3 original fixture, no inserted direct typed cache run.
- Same ordinary Raw result with NO INNER, and with an empty/tiny INNER cache. Report actual results, no source-derived runtime claims.
- Large ordinary original Text plus two small property edits, and large original semantic font string. Original raw/property bytes already in A must not fund a second B copy.
- An OUTER ordered accepted-text cache already including the same original raw Text: delta deferred scalar credit is zero.
- Nested Field cache only versus adjacent ordinary Text: parent deferred visitor adds ordinary atom bytes only, not child display. Direct-child aggregate terms retain the old reference exactly.
- Multiple edits/splits/repeated parses and MCE/template replay of one original operand: no renewal. Two distinct original equal atoms remain distinct.
- Unused huge raw wrapper and two unchanged1MiB captures stay inert/accepted under existing applicable checks.
- First read after independently proved unchanged-child transport versus unread namespace-envelope change. No fabricated original role from A pieces.
- Actual malformed None/resource failure and public atomicity controls, exact helper/reference numerical equality, selected input/property term equality with the old helper.

Root authorizes the complete original and frozen-model role below. No new role is already implemented by that authorization, and restoring INNER credit alone does not prove the full class.


#### Amendment: actual frozen-model ingress

The capability must cover BOTH existing ingress roles. A private DeferredOwner enum can distinguish Original(actual original successful owner/capture witness) from FrozenModel(actual selected frozen Field/current-owner membership and retained Raw-slot witness). FrozenModel is not an Original locator, supplies no OPC part/range and cannot fund A. The same scalar/property measurement and base-covered exclusion rules above apply.

For source_model_paragraph, bind the FrozenModel witness while source_model_paragraph still borrows the admitted current paragraph and register_model_fields still walks its actual frozen Fields, BEFORE clone, normalization or writing. The operation's existing paragraph/Field membership is authority. Freeze the actual Field-to-result/retained-source/Raw-slot relationship and consuming context there. Record current branch authority and which operands are already covered by the initial cached_result/current terms. No eager raw decoding or credit follows. An unsaved caller-assembled model is a real supported ingress, not generated callback output.

A Raw slot key is its actual frozen owner relationship plus captured initial slot/record identity, with the existing retained source/paragraph relationship anchoring it. Slot number is only a component of that frozen proof, never a candidate ordinal after edits. A shared immutable Raw slot reached through two projection edges retains the SAME baseline role identity and credit once. Distinct genuinely existing current Field values retain their separate existing current-input identities, even if their retained source is shared. Preserve that distinction rather than using one pointer, source_id or equal text for both rules. The owner enum authorizes the relationship, while the role dedup key identifies the actual baseline operand occurrence.

At the actual clone/serializer producer, explicitly forward that FrozenModel receipt through the real paired Field/result/raw-slot correspondence and source-piece maps into the produced carrier. This receipt does not relabel the produced carrier as frozen model input. Ordinary produced(map,registry) cannot recreate it. The first real eligible raw reader may stage decoded scalar/property facts under the preserved receipt, and actual selected Field-result consumption activates them once. Generated normalized sources, candidate clones, callback strings and later parse allocations do not mint new FrozenModel membership.

B remains the existing retained-input union. The frozen Raw/source record may already have been selected by register_model_fields and funded once through retained_inputs/stored_inputs. Do not add that B again when its deferred operands are decoded. C additions contain only the actual selected, previously uncovered ordinary scalar/semantic property bytes. No Field64, no scope guessing, no generated-output C, and no change to normalization_input_work/local reference. Property raw records already included in that B likewise cannot add another retained contribution through selected_input_work.

Use the same small internal registration consumer for both variants, with origin-specific proof resolution and common checked role dedup/measurement. Actual native body-item/frozen-model entry must select these witnesses before staging/flush and forward them through its real producer. A post-flush paragraph cannot be registered as another frozen model. Add the unsaved incoming-model analogue of the no-INNER fixture, repeated source_model calls, shared-slot DAG and distinct divergent-current controls, retaining source/None/resource ordering and no runtime claims here. The initial OPC-only proposal by itself does NOT close this model path.


### Root approved finite imported XML payload and publication coverage

This technical contract uses the existing normalization-output metric and two
existing opaque public types. It adds one doc-hidden operation method:

```rust
pub fn source_imported_field_xml(
    &mut self,
    source: &ScopedWordSource,
    owner: Range<usize>,
    value: &str,
    escape: bool,
) -> Result<ScopedWordSource>;
```

The still-unimplemented rewrite_imported_source signature above now consumes
moved ScopedWordSource payloads in its edit list, replacing the earlier owned
Vec payload argument. The new method performs actual raw or five-character
escaped XML fragment production after exact reservation. It is not a paragraph
parser or a new XML validator. Both modes preserve the existing native byte
policy, including raw values and unchanged carriage returns. No template-copy,
Original, model-current or callback input credit follows from producing it.

Private payload-only metadata binds the actual parent source occurrence and
nonempty source owner interval. The existing opaque payload owns its produced
buffer and exact one-shot prepaid receipt together. No ambient exemption stays
active while it waits in an edit list. A pointer lookup, equal bytes, matching
length or Generated source map cannot manufacture that receipt. Unrelated
prepared sources are not prepaid scalar payloads. Moving the payload into the
rewrite transfers its actual receipt, consumed only during the synchronous
copy of that exact produced buffer. Fresh changed-record Field prefix/suffix
and other unprepaid output still reserve before growth. Dropping or failing
never refunds spent work. Later repeat or changed-record production pays again.

Publication coverage is private metadata independent of Original/input roles.
Actual successful simple and complex constructors report selected current
source child/shell intervals. Model Field writers mark their actual output.
Source map slice, assembly, closure and splice transcripts forward this same
coverage. Nested overlap is unioned in current coordinates, while separate
published copies each consume demand. Source equality and inverse Original
lookup cannot establish coverage. A pending or failed constructor is not a
successful Field and unknown coverage is not a proof of fieldless input.

For successful shared-run splitting, preserve its actual selected Field ranges,
including its existing invisible-outside folding. When the existing splitter
falls back, retain independently committed actual constructor-owned child and
necessary shell/property intervals. Leave unowned outside bytes ordinary.
Do not infer the hypothetical successful splitter result, mark the whole shared
run Field, add another parser or reject the formerly accepted fallback. Capture
annotations must come from actual run and wrapper producers and follow their
real assembly transcript. A run-only solution does not qualify owned hyperlink,
control or companion-source paths. Failed/incomplete candidates grant nothing.

The payload owner witness must be wholly selected Field or wholly ordinary.
Ordinary payloads receive no Field debit or receipt. Empty insertion coordinates
have no side-based ownership. An empty scalar uses its actual nonempty enclosing
qualified event only when its full coverage proves the owner. Mixed ownership
cannot be rounded, prorated or made a new valid-input refusal. Supply actual
producer subrange proof or retain that case as unfinished compatibility work.
The rewrite validates all anchors, parent bindings, edit containment, bounds,
nonoverlap and length overflow before its output grows. Empty edit lists move
without a fabricated copy. Equal-byte allocated replacements remain production.

Named actual native consumers are patch_word_values, patch_relationship_ids,
add_identity_attribute_edit, add_attribute_value_edit, the complex instrText
replacement, replace_fragment_store_id and uniquify_drawing_ids_in_xml. Their
selected replacements are existing UTF-8 strings. Preserve raw versus escaped
policy and all decoder, REF parser, namespace and source-span rules. Keep final
payload recipes private in existing native files. Complete each stage's old
scan and semantic checks before payload reservation, preserving no-match and
numeric range-exhaustion errors. Selected docPr integers may use a ten-byte
stack decimal producer after those checks. Do not move the REF parser lower or
replace unrelated arbitrary FieldSourceEdit Vec users through this finite seam.
Decoded instruction accumulation and unescaped scratch strings are not emitted
XML or automatically template text-copy demand. Their existing applicable
constructor/scalar bounds still need actual qualification, without a new heap
quota. A later XML escape/output producer reserves its own applicable demand.

Required controls cover raw/escaped identity, every escape and CR, deletion and
empty values, ordinary versus Field versus mixed owner, foreign operation and
parent, failed composition with sticky work, single-use receipt and fresh
repeat, duplicate numeric ids at exhaustion, malformed later scanner input,
no-match store errors, and complete bytes/error/map parity. Confirm the actual
orphan-instrText shared-run fallback discriminator before claiming that route
covered. Large outside text, nested fields, paired tabs/breaks/custom marks and
wrapper assembly must preserve exact selected coverage and unchanged behavior.
This approval supplies no implementation, provider completion or final rider.

### Genuine scoped caller-literal current input

Root technically reconciles the existing incoming C1 role with the actual
scoped literal API. A current public control replaces a102-byte PAGE Field's
TOKEN cache with an actual caller-owned65536-byte String and succeeds. The
shared-work connection must preserve that acceptance. The replacement remains
Generated in source correspondence. Its distinct actual incoming caller-value
role may fund only the existing applicable Field current-scalar term once.
Neither original A credit, template64, a new multiplier nor unused padding
follows from this role. Internally produced template replay is not this input.

Authorize one additive doc-hidden substantive method on ScopedWordOperation:
replace_block_with_literal_input, with the existing replace_block arguments
and an additional borrowed Cell<bool> delivery flag. It owns the real existing
prepare, complete block parse, edit and final publication path. The existing
replace_block compatibility entry creates one local flag for that real external
call and enters the same substantive implementation. No new opaque type,
module, generic, public facade signature or independently tunable budget is
authorized. Add this method to the final hidden-public SemVer inventory and
compile its exact signature in the existing external-consumer canary.

The two actual Native public invocation routes create one local delivery flag
per real caller replacement. Private candidate and paired-alias calls forward
the same borrowed replacement and flag synchronously. A paired alias cannot
renew it. A different literal invocation gets its own event even when text is
equal. This is an explicit producer protocol, like the genuine callback event,
not text equality, pointer discovery or an unforgeable public capability. A
computed replay cannot manufacture such an event. Do not reuse one event for
different values or repurpose the callback or JSON input roles.

At the unchanged literal match-selection boundary, validate the actual accepted
match and containing Field ancestors before output growth. Activate exact
replacement UTF8bytes under the existing Field current-value multiplier only
when that real incoming value is consumed by an affected Field. Empty needle,
no match, a rejected crossing match and fieldless replacement add no Field
credit. Register all selected inputs and checked arithmetic before mutating
credit or the flag. Ordinary validation/error order and selected-first count
mismatch or zero-count return stay unchanged. Actual earlier consumed C0
remains governed by the existing phase-lifetime rule. Candidate models, emitted
XML, reparsed output and wrapper maps cannot renew input. Sticky refusal cannot
be cleared by another event, and failed operations do not refund used credit.

Required discriminators include the observed65536-byte accepted caller value,
empty and zero-match unused values, fieldless output, exact accepted-match
selection across run/wrapper boundaries, repeated paired aliases with one
event, equal distinct real invocations, generated replay without a new event,
and a late resource/namespace failure leaving original bytes and maps intact.
Preserve one final Field writeback for all actual literal matches. Raw/control
map companions and the actual final-writer transport proof remain separate
mandatory producer obligations. This clarification authorizes implementation,
not acceptance or provider completion.

### Retained literal text mutation metric units

Root reconciles the scoped literal carrier with the existing4410to4424metric
units. Updating an existing CT_Text slot does not create a new Text constructor
entry. Unescaped model String allocation or copying is not emitted XML and is
not automatically Q-C1-03template-copy demand. The narrow operation-produced
text-copy amendment names the concrete Native default-format, template and
filter carriers. It does not authorize a generic aggregate cached-String byte
quota over every lower model edit. Do not sum such byte lengths into
constructor_actions, XML output or the template ledger. No new counter or
accepted-input restriction follows from the literal connection.

Actual applicable local current/scalar checks and complete physical-fact copy
reservations still need their genuine before-growth producer coverage. A
passing XML guard proves neither earlier String allocation nor complete current
role qualification. Report those boundaries precisely and preserve any
unqualified existing-metric obligation. This clarification does not waive
preallocation checks, new typed result action reservations or final C1/provider
acceptance. It prevents an unrelated heap quota from being invented while the
coherent shared-work carrier is reviewed.

### Explicit scoped template TextWalk consumer

Root technically approves one additional doc-hidden method on the existing
ScopedWordOperation after tracing the actual cross-crate boundary. Native
parse_program currently calls public placeholder::replaceable_texts after the
lower paragraph parser returns. A crate-private lower walk alone cannot connect
that caller to the operation's Word interval facts. Source creation or automatic
walking inside source_paragraph would incorrectly select unvisited input.

```rust,ignore
pub fn template_replaceable_texts(
    &mut self,
    source: &ScopedWordSource,
) -> Result<Vec<String>>;
```

The explicit method accepts only a source from this same operation. It shares
the existing paragraph admission, exact parse/map companions and Field fact
registration, then runs the same substantive lower placeholder TextWalk. That
actual walk activates only the proved Word intervals it visits before String
aggregation grows. It uses the existing selected-input and text-copy equations,
without a new quota or eager all-Text grant. Repeated calls retain consumed role
keys and cumulative demand. Generated equal text has no new input authority.
Malformed source and resource failure keep their existing ordering and sticky
operation policy. No input credit follows merely from constructing a source.

Existing source_paragraph retains its existing behavior and grants no automatic
Word text selection. Existing public placeholder::replaceable_texts and its
legacy direct-read exception stay unchanged. Native paragraph grammar replaces
its detached parser/public-reader pair with this explicit scoped consumer.
Row direct/all visits preserve the actual whole-row validation and refusal
order, existing paragraph visitation and marker predicates. They use actual
physical paragraph slices with carried facts and namespace closure, rather
than value matching, invented traversal or extra input per repeated visit. An
incompatible concrete row consumer requires explicit technical reconciliation
before any additional API is added.

This raises the final hidden-method inventory from30to31with no new public
type, model field, module or generic. Update the external compile canary and
actual SemVer inventory. The change is technical design authorization, not
implementation, another peer, provider completion or runtime acceptance.

### Explicit scoped Native Field preparation consumer

Root technically approves one further doc-hidden method on the existing
ScopedWordOperation after verifying the actual early Field normalization
cross-crate boundary. Native prepared_current_field clones before source-model
serialization, and its public lower normalizer creates fresh local work. Later
Field pairs or source writers cannot recover newly created Text origins or
replace that earlier producer's call order.

```rust,ignore
pub fn prepare_template_field(
    &mut self,
    current: &Field,
    frozen: &'a Field,
    normalize: bool,
) -> Result<Field>;
```

Native retains its existing preflight_current_field call and raw-preservation
bypass, then supplies the actual inverse bypass result as normalize. Validate
frozen membership and establish the genuine clone producer's correspondence
before cloning. With normalize=false, preserve the existing imported-raw
clone path and inert provenance. Do not add eager semantic validation, admission
or Field-current activation that defeats that raw acceptance. Existing
non-template prepared_current_field callers keep their current behavior.

With normalize=true, retain the existing recursive order through instruction
arguments and switch arguments, actual cached children, then the parent owner.
Use the existing substantive normalization producers with inherited operation
work. Preserve each existing local Field equation, error and raw-fallback
boundary. Forward the selected incoming CachedResult interval at the actual
requested copy, FieldResultPart construction, nested assignment, selected
writer and reparse. Computed derived_parent_cache and accepted-text summaries
remain generated. No value match discovers an origin.

The method returns the normalized Field and retains its actual produced source
and constructor correspondence in operation-private facts. Those facts must
survive Native's real model installation and later synchronous writer pairs,
without public model fields, a self-referential return type or a persistent
candidate-address catalog. Frozen Field identity and actual producer/slot
relationships distinguish divergent current values sharing retained source.
The explicit Native producer supplies correspondence, as with the approved
writer pair lists. Membership alone is not proof that arbitrary values pair.

Word CachedResult input remains inert until the actual scoped TextWalk selects
its proved Text intervals. No JSON64, eager all-Field Word credit, generated
cache input or renewed role follows from preparation. Repeated preparation
retains the same input identities and cumulative demand. Existing actual
normalization and fact-copy reservations remain applicable before growth.

The final hidden-method inventory is32across the same two opaque types. Update
the external compile canary with this exact return and both normalization
branches. This technical authorization implements or accepts none of the
bridge, adds no public type/module/trait/generic or quota, and does not reset
review caps or complete the C1/provider gate.

### Frozen retained XML Word source correspondence

Root technically extends the existing Word-only CURRENT SOURCE catalog to
complete actual retained XML byte records in the incoming model, frozen before
Native staging. The earlier cache/proof-only wording did not already cover
those records. This explicit extension preserves genuine dirty-model raw
textbox input without promoting a serialized candidate buffer to new input.
It changes no numeric allowance or Field/local-input formula.

The four existing approved source_model_body_item, source_model_document,
source_model_notes and source_model_glossary signatures above now each have
the final argument raw_origins: &[(&[u8], &'a [u8])]. There is no new type or
method in this signature amendment. Native's actual clone and mutation
producers supply structured retained-record current/frozen correspondence,
including Drawing and run extra XML. Membership is frozen in the existing
current_word_sources constructor argument without eager decoding or credit.

Before writer growth, validate frozen membership, exact view bounds and
duplicate current keys in the invocation-borrowed pair list. The concrete
retained XML emitter forwards copied intervals from that actual current record
to its paired frozen record. Its lexical context lasts only through the
synchronous writer. No equality search, inferred model ordinal or persistent
candidate-pointer catalog discovers a source. Namespace insertions and new
replacement bytes remain generated. Encoded root-attribute metadata is not a
Word Text witness.

Copied records and overlapping views retain one frozen input domain. Separate
real input records remain distinct. Actual existing clone/slot relationships
forward facts through further publications. Only an actual accepted Word Text
parser and the selected scoped walk activate decoded intervals, once. The
CURRENT SOURCE role grants no Original A, Field C0, JSON64 or eager all-XML
credit. Keep the existing exclusive origin choice for each initial program
source. Missing correspondence cannot become permission for a new refusal
boundary or a fresh baseline.

This amendment is technical design authorization for the already required raw
textbox/current-source compatibility. The hidden-method inventory remains32,
and the four unpublished signatures and external compile canary must agree.
No implementation, new peer or provider completion is claimed. The separate
raw textbox scoped consumer still needs its exact final signature recorded
before implementation.

### Scoped original-literal textbox consumer

Root technically approves this additional doc-hidden method on the existing
ScopedWordOperation after tracing both raw textbox visits and simple-Field
reachability. The earlier part-name-only proposal is not adopted. Actual
post-staging raw sources may differ from frozen package bytes, so the method
consumes and returns same-operation carriers with their real provenance.

```rust,ignore
pub fn replace_original_template_literals(
    &mut self,
    source: &ScopedWordSource,
    excluded: &[Range<usize>],
    replacements: &[(&str, &str, &[&'a serde_json::Value])],
) -> Result<(ScopedWordSource, usize, Vec<String>)>;
```

Share the exact existing replace_original_literals_in_xml_part substantive
walker, duplicate/nested-box handling, paragraph capture and exclusions.
Preserve its actual namespace, validation, matching and error order. The actual
paragraph decoder and scoped TextWalk receive the consumed source spans and
constructor companions. Do not discover an origin from a callback paragraph
ordinal, equal bytes or generated serialization. Select proved decoded Word
intervals once and reserve actual discovery/aggregation copies before growth.
The returned source carries the actual copied/generated rewrite recipe for
Native publication and reacquired programs. Replacement bytes remain generated.

Native main input uses its real source_model_document output. Related parts
reuse actual flush/publish carriers. Neither substitutes frozen part-name
bytes for the actual current source or creates an input baseline from output.
Actual incoming retained raw XML uses the separately approved frozen-record
pairs. Preserve checked program exclusions and original-only token discovery.

Each replacement triple supplies its original token spelling, prepared text
and the exact borrowed JSON operands actually used to produce that text.
Validate membership before relevant growth. Affected Field selection uses the
same existing actual operand roles as format_paragraph, without treating a
derived String as a new caller-literal event. Join forwards its actually used
children. Word literals and default-only results have empty JSON dependencies.
Unused Null/default inputs receive no inferred credit. Word text-copy and
Field-current roles remain distinct. No JSON64is added to a Field equation.
Raw formatter expressions retain their existing early refusal, so this method
adds no callback flag or new callback behavior.

Empty replacements perform discovery and return the actual original sources.
For Native's nonempty rewrite call, scoped mode may avoid collecting its
genuinely unused second discovery result and return an empty sources vector,
while still performing every actual matching walk and required validation.
If the implementation retains that allocation, it must reserve its real copy.
The existing public legacy tuple always retains its current semantics. No
new conservative refusal, renewed source credit or omitted scan is authorized.

The final hidden-method inventory is33across the same two opaque types. Update
the external compile canary and four amended raw-origin writer signatures.
This is technical API/consumer authorization only, not source implementation,
another peer, provider completion or native acceptance.

### Move already-owned first Text into scoped stretches

Root approves a producer optimization within the existing scoped TextWalk API
to preserve zero-input no-op consumers. An empty PAGE or NUMPAGES Simple Field
emits generated1, which remains generated and contributes no Word input.
Copying that parsed result into a new one-byte stretch would require credit
that the input does not supply. The method already owns its required parsed
paragraph and wrapper models, so it may move the first uniquely owned Text
String directly into each returned stretch instead.

Share the same TextWalk selection and boundaries with the existing reader.
The built walk may first register exactly the proved Text intervals that it
will actually visit, without copying Strings or parsing again, then aggregate
owned pieces. This prevents an earlier generated append from refusing solely
because legitimate later visited input has not yet been registered. It does
not select unvisited owners, attributes or opaque content. Actual genuine Word
selection still occurs when the first String moves and is not skipped merely
because that move costs no new copy.

Move only an already-required uniquely owned parsed Text buffer. A clone,
Cow-to-owned conversion, Arc::make_mut copy or extra parse cannot manufacture a
free buffer. Shared/nonmovable pieces retain their actual charged copy path.
Reserve every actual later copy or append before growth. An all-generated
multi-piece stretch still refuses when its actual demand is unfunded. No
quota, allowance, input role, public signature or model field changes.

The borrowed public replaceable_texts behavior stays unchanged. Full original
construction and validation/error order precede owned consumption, and shared
caller models are never mutated. Controls must establish actual generated1
no-op success, zero new input credit, refusal of a later real unfunded copy,
borrowed-reader output/boundary parity and a reachable mixed generated-prefix/
later-genuine-input case. Unsupported wrappers alone do not prove the generated1
bug. The hidden-method inventory stays33and the current canary signatures are
unchanged. This is implementation authorization, not a completed acceptance.

### Immutable access to actual retained model records

Root authorizes four additive doc-hidden borrowed-data accessors on existing
lower model types for the actual Native pre-clone input census and synchronous
current/frozen writer pairing. The glossary root, each doc part and each
revision retain separate private raw buffers that their existing publishers
copy. Public typed projections and equal package bytes are not substitutes
for those actual incoming records.

```rust,ignore
impl CT_GlossaryDocument {
    pub fn retained_xml(&self) -> &[u8];
}
impl CT_DocPart {
    pub fn retained_xml(&self) -> &[u8];
}
impl CT_Revision {
    pub fn retained_xml(&self) -> &[u8];
}
impl CT_Sdt {
    pub fn retained_run_xml(&self, content_index: usize) -> Option<&[u8]>;
}
```

Each accessor borrows only its stored buffer, without cloning, serializing,
normalizing, selecting Word input or granting any work. The three Vec-backed
owners return an empty slice if their stored record is empty. The control
lookup uses the actual stored absolute SdtContent index, not run ordinal.
Missing capture returns None. This exposes no private source/provenance type,
mutable access, public model field, new module, trait or generic.

Stored control capture does not assert current serialization eligibility. Its
existing private writer guard still checks the original run and pending Field
replacement. Pending-edit buffers remain distinct. Native freezes records
before clone and carries only actual clone/mutation owner and slot relations.
An index after arbitrary public content splices cannot establish provenance.
No equality discovery, package-byte search or persistent candidate-address
catalog is authorized. Actual writer consumption remains the authority for
copied intervals and lazy TextWalk selection remains the authority for input.

Add exact borrowed lifetime calls and optional-control lookup to the existing
external canary and SemVer inventory. There remain33authorized methods across
the two opaque carrier types, plus these four accessors on existing model
types. No new type or binding is added. Final union compilation and actual
Native consumer tests remain required. This finite access amendment is no
implementation acceptance, peer reset, provider completion or new refusal.

### Movable prepared Field receipt identity

The actual prepare_template_field return moves into Native staging and later
writer pairs. A shared FieldSource Arc cannot identify divergent incoming
CachedResult occurrences. Root therefore explicitly permits one private inert
producer receipt identity on existing Field, with its substantive payload
remaining in operation-private facts. This satisfies the existing ban on a
public model field and persistent candidate-address catalog.

Use the existing operation anchor and a checked operation-local occurrence
identity, for example Option<(Weak<()>, usize)>, or an equally small private
concrete identity. Do not store source bytes, a model, a frozen reference,
provenance map, Work or mutable allowance on Field. The ordinary path is None
without an extra allocation. Actual prepared clone/move forwarding preserves
the identity. Ordinary constructors/parsers do not manufacture it. Dropping
the operation releases its facts, with no source retained by the token.
Foreign or expired receipts are not authority and create no new broad legacy
refusal. Existing Field equality, serialization and public behavior stay the
same.

Receipt presence is not proof that public cached_result or nested children
still contain original values after mutation. Actual requested-copy,
normalization assignment, selected writer/reparse and Native clone/mutation
producer relations must forward the proved intervals. Changed/generated bytes
receive no inherited input merely because their Field retains a token. The
operation resolves the receipt only under those actual producer facts, with
no String address/equality discovery or candidate pointer history.

Repeated preparation and actual clones may have separate producer receipts,
but their Word input roles still resolve to the same frozen operands and union
once. Divergent ingress Fields sharing FieldSource remain distinct. Controls
must distinguish those cases, actual move/clone delivery, generated mutation,
foreign/dead receipts and ordinary None behavior. Field and Word allowance
rules are unchanged. There is no additional public method/type or canary
signature change, and no implementation or provider acceptance follows.

### Issue430 child partition and concurrent source ownership

Root approves the complete existing Issue430 contract as F-X272a, whose sole
canonical design is .claude/plans/F-X272a-design.md. All45A430rows,14G430gates,
five original issue rows, prior review history and Q-A430-28 remain intact.
This section supersedes older unadopted/paused/adoption-after-C1 wording only
for that child. It changes none of this parent's carrier API signatures,
equations, C1 obligations or acceptance debt.

F-X272 retains the comparison/source C1 author and existing worktree, branch,
Base and dirty source. F-X272a may use a distinct clean reviewed worktree and
existing baseline APIs after ordinary claim/start. Its native comparison,
revision custody/cleanup and lower note-type parser work has no demonstrated
dependency on a new unfinished C1 carrier callable. Shared source files are
serial integration resources. An actual required unfinished operation would
remain a genuine blocker and must not be bypassed.

The parent cannot complete or unblock F-X277/F-X279/F-X280 until its own work
and the integrated child meet actual local gates, with all final riders still
pending where the workflow permits the provider checkpoint. Child approval,
claim or scoped results alone never complete the parent. No peer cap resets,
hash baseline changes, extra aggregate effort, reporter comments or publication
follow from this partition. Root final and actual combined reconciliation remain
mandatory. The parent author acknowledges this scope change before affected
note-comparison work, without pausing its unchanged C1 implementation.

### Issue430 created-part special inventory correction

The approved child F-X272a now resolves Root contract defect D1 from
F-X272a-created-specials-contract-codex-final-1.md. Comparison creation keeps
exactly the edited input special inventory, including no specials, rather
than manufacturing conventional records. Existing examples with synthesized
specials refer to records already in the edited input. All45childacceptance
rows,14gates, destination identities, original-unused preservation and Q28
remain unchanged. F-X277b retains authored defaults, no future provider is
assumed. The parent completion rollup consumes this corrected child contract.
This changes no C1 API, equation, metric or source/build lease.

### Concrete mapped final-buffer transport

The existing final Word block serializer uses a live Writer<Vec<u8>> proof.
Its final byte copying is transport outside the normalization producer metric,
as stated in Incremental normalized XML emission boundary. The connected
format_paragraph carrier may use the existing private NormalizedWordOutput
at this exact final caller to return its actual source map and Word events.
This preserves the old Vec transport behavior and changes no API or equation.

The private sink binds only the same live mutable final Writer and existing
ledger identity. Its lexical guard restores on ordinary error and unwind and
holds no metadata lock during Writer reentry. A matching active final-writer
identity is required at that sink. Publication role, equal bytes, a caller
boolean or merely having an ambient final lease cannot authorize transport.
Nested normalization writers remain ordinary full-demand sinks even in that
same operation. No source input, Field Original/current credit or prepaid
receipt follows from this final-buffer role.

Actual inner escaped XML, changed records, generated Field output and namespace
closure keep their existing pre-growth reservations, byte demand, errors and
sticky refusal. The final sink changes only outer copying that the old Vec
never charged, and obtains maps/events during that same serialization.
Retain the existing concrete Vec validation path. Any new generic factoring
must have both actual Vec and Normalized caller instantiations in the working
implementation, not a proposed second caller or a forwarding-only wrapper.

Required controls compare actual old Vec and connected final bytes, errors,
normalized demand and map/event coordinates. Prove zero outer-copy allowance
parity with actual unchanged Vec behavior, full nested-record demand on the
same ledger, generated Field and closure charges, wrong-writer isolation,
error/unwind restoration and no live receipt after completion. Ordinary None
paths retain their behavior and allocation policy. This technical clarification
approves no implementation or completed carrier/provider gate. The active
owner acknowledges this exact revision before changing the final sink.

### Q-A430-29 child recognition amendment approved

The user explicitly chose Preserve valid histories (Recommended) in
call_3b89d767d87c471bb71e618747fba848. The child operative amendment treats a valid
unmatched B as retained tombstone or completed imported metadata under the
existing lifetime rules, never binding the following unrelated reference by
adjacency, wrapper or numeric coincidence. Positively identifiable inconsistency
refusal remains, with the accepted inability to detect every missing-N corruption.
No marker redesign, quota, C1 API or Q28 exception change follows. The child
owner must acknowledge and implement the revised contract, preserve all45A/14G
obligations subject to that explicit amendment, update HLD and obtain Root review.
Its separate D1 opaque-owner repair is not waived. Parent C1 work continues.
No provider readiness follows merely from the answered decision.

### Explicit scoped Native row TextWalk consumer

Root technically approves one additional doc-hidden method on the existing
ScopedWordOperation after tracing Native template.rs parse_program and its row
visitors. The paragraph-only method remains unchanged.

```rust,ignore
pub fn template_row_replaceable_texts(
    &mut self,
    source: &ScopedWordSource,
) -> Result<(Vec<String>, Vec<usize>, Vec<usize>, bool)>;
```

The tuple contains owned actual stretch buffers, direct-view indices, all-view
indices and the row-local preserved-sibling or content-control hazard. Each
index addresses the returned buffer vector. Direct is not a subset of all:
the existing direct control visitor includes SdtContent::Run Text concatenation,
while the all visitor omits that variant. Preserve both substantive visitors,
including their order and multiplicity. Direct still runs before all. Share a
buffer between views only for the same actual parsed occurrence and stretch,
never because strings compare equal. Returning two borrowed views must not
materialize the same buffer again. Distinct physical equal text stays distinct.

The method checks the same-operation source, carries its actual namespace and
source maps through the existing wrapped-table checked parse once, and retains
the existing first-row and error behavior. Actual successful parser companions
provide physical Field ownership and selected Word facts. Commit only returned
owners. Paragraph slices alone do not prove complete row coverage. Unknown
capture or fallback remains unknown, never inferred fieldless. No second row
parse, guessed traversal or extra input funding is authorized.

Native retains program_marker and its dedicated-marker predicate on borrowed
str views. The returned hazard covers only the existing row extra_xml preserved
siblings and nonempty content_controls. Native still applies item.marker_unsafe,
item.field_owned and inherited_field in their existing order. The lower helper
does not adopt Native grammar or change marker acceptance. Shared applicable
input and text-copy budgets remain unchanged, with actual registration and
reservation before growth, sticky failures and no repeated baseline allowance.

Implement and test direct-only control runs, shared paragraph views, nested
cells/tables/controls, zero and multiple stretches, equal distinct occurrences,
namespace closure, malformed whole rows, preserved siblings and existing marker
refusals. Compare direct/all strings and marker outcomes with the old real
visitors. Repeated calls consume existing role keys and cumulative demand.
Add exactly this one method to the external compile canary and final strict
SemVer inventory. No public type, model field, module, trait, generic or new
quota is added. This is concrete technical reconciliation of the existing row
consumer, not provider completion or acceptance of an unimplemented method.
The sole owner acknowledges this contract revision before affected source edits.

### Same-producer core capture transcript

Root technically approves one additive doc-hidden core callable after reading
oxml-core raw_xml.rs capture_element and the actual scoped textbox seam.

```rust,ignore
pub fn capture_element_with_source_ranges(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    input: &[u8],
    opening: Range<usize>,
) -> Result<(Vec<u8>, Option<Vec<(Range<usize>, Range<usize>)>>)>;
```

Each pair is a known absolute input range and its captured-output range for
actual copied bytes. Output gaps are Generated. Removed source bytes have no
copy pair. The caller supplies the already consumed opening-event range.
The core producer verifies the same live reader/input backing, actual consumed
position and physical header payload. Numeric occurrence bounds establish the
candidate mapping, byte checks validate it. No search for equal content or
expired-owner identity is authorized.

One private substantive existing capture loop serves the unchanged public
capture_element and the actual mapped lower textbox consumer. Transcript
collection is disabled without allocation in the legacy path. Preserve the
old writer events, nesting, error and EOF behavior exactly. Default and trimmed
Text modes may map only after the actual pinned QuickXML consumed-window and
payload rules are established. Track trim_start and trim_end independently,
including all-whitespace skipped before markup or EOF. Removed whitespace has
no pair. CDATA, entity and markup event boundaries retain their actual producer
semantics. Expansion and other unproved modes, BOM coordinate uncertainty or
failed transport proof yield the same captured bytes with None, not a new
refusal, altered parse or ordinary/fieldless proof. Preserve one actual Word
operand identity and its decoded-range funding, never a new fixed input term
per copied fragment. Both indentation and whitespace-bearing text payloads need
discriminating controls, not just unchanged TOKEN text.

For End normalization, copied name/punctuation fragments can remain mapped
while removed whitespace has no pair. DocType maps only actual copied payload
and closing delimiter, with canonical keyword/space Generated. Every other
event needs its actual producer transcript. Generated framing creates no
Original, Field, Word or current-input credit. This byte transport does not
certify complete Field ownership or resolve a lost actual Field map by itself.
Existing capture/output and sticky work rules remain unchanged.

Require real old/plain versus mapped byte/error parity, source/output bounds,
all existing event variants, normalized End/DocType, entities, nesting, EOF,
reader position, unproved modes and the disabled-transcript legacy path.
Connect the actual lower scoped walker rather than adding a speculative
consumer. Include this exact core callable in the external compile canary and
final strict SemVer inventory alongside the operation methods. No new public
type, module, trait, generic, quota or lower duplicate parser is approved.
The sole owner acknowledges the exact revision before affected edits. This is
technical API reconciliation, not implementation or provider acceptance.

### Existing-API F-X277b scheduling clarification

The integrated Issue430 child F-X272a may supply its approved genuine LOCAL
provider checkpoint to F-X277b's A277-2/3/6 implementation after actual prefix
completion. Source audit found those rows use existing baseline mechanisms,
not unfinished C1 carrier APIs. Separate-worktree implementation and serial
semantic reconciliation preserve this parent's sole source writer and Base.
F-X277's umbrella, F-X279/F-X279b and F-X280's umbrella retain their parent
barriers. F-X280b may use the separately reviewed completed LOCAL F-X272a plus
F-X277b chain for its exact raw-validation and note-provider prerequisites.
Its final actual C1 open/error/resource/snapshot and both-family OPC composition
remain mandatory, with no unchecked validator or count-based custody proof.
This parent remains in-progress. Its own completion still requires C1, the
integrated child and actual applicable gates, with all final riders retained.
No C1 implementation, budget, API or review-counter change follows here.

### Same-operation whole-document factory bridge

Root authorizes one concrete additive doc-hidden method on the existing lower
operation type in crates/rdocx-oxml/src/document.rs:

```rust
pub fn source_document(
    &mut self,
    source: &ScopedWordSource,
) -> Result<CT_Document>;
```

The named consumer is Native from_package_reusing's actual fresh main-document
factory, reached by import_fragment_content_with_state. Its current
CT_Document::from_xml call receives bytes only. The existing private
from_xml_admitted already owns the single substantive Reader and actual
body/paragraph/table/control constructor chain. Reuse that chain with the
source's actual map, operation registry and unchanged initial-delivery flag.
Validate the existing operation anchor before work. Preserve the existing
admission, grammar, resource and error order. Return the real model and retain
only operation-private facts observed by those actual factories. Do not add a
public disposition type, expose private parser context or change document_block
into a whole-document parser with different wrapper semantics.

A successful model and joined returned Fields are still insufficient for
ordinary or complete coverage. Record actual reached, constructed, fallback,
opaque and unvisited outcomes as needed by the existing private consumer join.
Unvisited or unsupported regions stay unproved. Only genuinely observed
selected ownership and ordinary disposition may classify a payload owner.
Later smartTag/customXml wrapper factories and other relevant actual consumers
must join their own occurrences. A transient REF/PAGEREF inspection Field is
not automatically a published XML owner. Do not infer raw fldSimple promotion
from a helper definition without an actual caller. No extra parser, validator,
normalization, output quota or input allowance is approved by this bridge.

A private Native optional carrier context replaces the fresh factory call.
It must not parse again after the old factory or reserialize a model to invent
a source witness. Preserve root/background/declaration semantics and the
hinted/kept alternatives. Those alternatives may carry actual prior producer
facts or stay unproved. Do not disable reuse solely to manufacture observation.
Bind facts to the exact actual main source occurrence. Repeat-produced packages
carry their generated main source through from_part_content and
import_content_staged. Opening those packages does not renew Original or
frozen-model credit. Original raw fragment bytes captured before preparation
remain distinct from later generated dependency or glossary output.

Use existing regression carriers for ordinary versus source-aware document
factory parity, empty and multiple Fields, imported aliases/shadows, retained
root/background/raw bytes, generated-repeat origin, foreign-operation refusal
and failed-factory nonpublication. Reuse the existing nested-wrapper/control
late-construction discriminator rather than equating an empty outer Field list
with ordinary XML. Preserve actual source/error/bytes parity and the disabled
ordinary path. Batch affected local controls and gate them on the frozen source.
Final Native, installed, resource/history, SemVer and all sprint riders remain
mandatory. This callable alone supplies no completeness or provider checkpoint.

This makes35authorized carrier names across the existing two opaque types,
including as_bytes, with the existing accessors, two earlier helpers, core
capture callable and residual helper inventoried separately. Add source_document
to the exact external compile client and strict SemVer inventory after actual
implementation. The addition is a Rust0.17.0 additive surface. No Python/WASM
binding, new file/module/crate, feature, trait or generic is introduced. Existing
HLD03/HLD04 own the private parser/fact handback mechanism and HLD10 retains the
public compatibility boundary. Existing public-API/parser/shared-consumer risk
riders apply. Plan status stays approved, all unresolved work stays pending and
no peer cap resets. The sole author acknowledges this revision before edits.

### Borrow stored ordered Field input without materialization

Root authorizes promoting the existing allocation-free
Field::stored_ordered_result method in crates/rdocx-oxml/src/text.rs from
pub(crate) to doc-hidden pub, preserving its exact body and signature:

```rust
pub fn stored_ordered_result(&self) -> Option<&CT_P>;
```

The concrete consumer is Native template-input membership freezing before clone,
preparation or the actual model writer. The current public ordered-run getters
compute a display String, and the owner-projection getter creates an owned
normalized paragraph. Neither is a substitute for an inert borrowed view.
This method only borrows stored state. It does not certify current cache
eligibility, select Word input, decode raw XML, clone, normalize or grant work.
Writer and reader branch guards remain authoritative when an operand is used.

Freeze actual model identities and carry correspondence from the real clone or
preparation producer. Do not recover it later through equality, changed ordinals
or a persistent candidate-address catalog. Preserve Sdt Run inputs missed by the
paragraph visitor and use absolute SdtContent indices for retained captures.
Use the existing Field/Word regression carrier and final external client to
check the borrowed lifetime, optional result and unchanged selection behavior.
No standalone test binary, model field or public provenance type is added.

The operation carrier count stays35. This is one additional borrowed accessor
beside the four retained-XML accessors, included separately in the Rust0.17.0
additive SemVer inventory and exact compile-client calls. No Python or WASM
surface changes. Existing HLD03 and HLD10 describe the input/read boundary.
Public-API, shared-consumer and resource riders remain due. This visibility
change supplies no ordinary certificate or whole Native/provider completion.
The sole writer acknowledges this revision before the visibility edit.

### Actual body seeds for inert Word membership

Root reconciles the existing with_template_word_inputs signature above by adding
current_bodies: Vec<&CT_Body> immediately after current_paragraphs. No other
parameter changes. This unpublished hidden Rust constructor now has seven
parameters. It introduces no public type or additional carrier name. Update
all real callers and the exact final external client together. The previously
prepared six-parameter client remains historical preparation, not final proof.

Native supplies actual borrowed main and glossary bodies before staging, plus
the existing borrowed paragraph roots for notes. Its existing paragraph visitor
keeps its read behavior and is reused for paragraph seed collection. A body
seed reaches paragraph-free block controls and their lower-private retained
original runs, which cannot be enrolled through direct current Run callbacks.
The lower constructor uses the existing WordOwner::resource_children physical
graph, including stored Field caches, revisions, Ruby, controls and drawing text.
Inspect Text leaves and actual stored raw records at each visited owner. Dedup
actual owner identity before expanding shared immutable subgraphs. Extend only
membership catalogs. No parsing, lazy model projection, normalized String,
resource admission, selected-input funding or model clone occurs here.

Body, paragraph and Field seeds may overlap. Their same physical descendants
are frozen once. Absolute control content indices select retained captures.
Stored Field owner membership does not bypass current writer/cache eligibility.
Keep Field private source/current relationships with the actual frozen Field.
Raw records not reached by this physical graph remain explicitly outside this
increment, with actual note/glossary root records supplied separately when used.
No equal-bytes lookup, post-preparation ordinal matching or persistent candidate
address catalog establishes correspondence. Actual clone and preparation
producers still owe the approved TemplateWordLineage handback.

Use existing lower and Native test carriers for a paragraph-free block control,
retained original control runs, nested stored Field Text, revision and Ruby
leaves, overlapping seeds and unchanged zero allowances before actual selection.
Foreign candidate identities must still refuse. Existing Native read visitors
must keep their prior outputs. HLD03/HLD10 and the Rust0.17.0 API inventory own
this callable revision. Existing public-API/resource/shared-consumer and final
riders remain due. Root acknowledges the reviewed revision before edits. This
is no provider completion or peer-cap reset.

### Positive ordinary root events from the actual document reader

The existing whole-document factory may report private positive ordinary
intervals only for its actual qualified document opening, body opening or Empty
body event, and document closing event. Record the same Reader's before/after
positions at those actual branches. Capture the body opening end before its
child parser consumes the body. Commit these facts only after the whole factory
succeeds. Ordinary from_xml calls collect no companion. No second reader,
recovered QName search, whole-container ordinary inference or parsed-model
serialization supplies a witness.

These facts concern only the emitted/consumed root events, including their
attributes. Unknown body children, background payloads, retained wrappers and
the complement of known Field intervals stay unproved. Empty Field lists and
returned_fields_joined do not certify them. Actual Field ownership takes
precedence when later relevant consumers prove overlapping intervals. Root
facts add no Original, current, Word, namespace or output allowance.

The existing private publication coverage and source maps carry these positive
ordinary intervals separately from Field demand. Slice, composition, namespace
insertion, repetition and predecessor release preserve only the actual copied
images. Inserted/generated bytes remain unknown unless their own actual producer
proves a role. Existing Field-copy reservations count Field intervals only,
never a generic complement or ordinary XML quota. This changes no public
signature, carrier count, model field, parser policy or binding surface.

Existing tests must distinguish actual qualified opening events from body
contents, aliases and shadowed foreign names, partial/failing factories, copied
versus inserted intervals, predecessor drop and overlapping Field priority.
HLD03/HLD04/HLD10 describe the partial private facts. This is one bounded
consumer handback, not whole relevant-consumer closure, implementation of the
two imported-payload methods, actual Native lineage or provider acceptance.

### Positive ordinary property events from actual readers

Extend the private positive ordinary sidecar to the existing successful qualified
Empty pStyle and rStyle property-reader branches only. The actual property
reader supplies event bounds after its current value-attribute decoder succeeds.
Forward those facts through the already-required scoped property parse and
actual capture copy transcript. Do not add a parser, QName pre-scan or generic
ordinary classification. Ordinary None callers allocate no new transcript.

Facts remain staged through the actual retained constructor-return chain.
Discard them when a temporary successful property belongs to a run or owner
that is later discarded or preserved only as raw fallback. Whole document
success cannot commit observations from discarded candidates. Apply actual
Field coverage first, including selected Field cache properties. A style element
inside a proved Field remains Field-owned. Start variants, unknown property
children and whole pPr/rPr subtrees remain unproved. A retained smartTag or
customXml wrapper can forward only its own later actual property-reader result
through its captured map. Initial document parsing cannot invent that result.

Local evidence distinguishes qualified Empty observations from retained Start
forms, foreign or shadowed names, value decoder and parent failures, mapped
capture normalization, aliases/BOM, Field priority and copied-versus-generated
transport. This private extension changes no public signature, valid-input
acceptance, allowance or quota. Relationship, drawing and other raw identity
owners still require their own actual consumer handoff. HLD03/HLD04/HLD10
remain the impact files. Final union and all prior residuals remain mandatory.

Root technically approves this bounded extension after the clean independent
design pass. Implementation and its constructor-return evidence remain pending.

## Checked source transport copies, existing budget reconciliation

Native assembly exposed an unclosed existing copy boundary. compose_sources
materializes a new buffer from actual selected carrier images. This is checked
operation work, not the narrow Q-C1-02 legacy-read exception. Apply the same
existing Field publication-copy reservation used by repeat_source. Select the
actual assembled source-map publication and positive delivered Field coverage,
union their overlapping output images and count each distinct output occurrence.
Equal provenance across two output copies cannot merge those output occurrences.
Unknown regions remain unknown, and ordinary ranges grant no copy allowance.

Validate anchors, ranges, checked lengths and constructor policy before the
applicable reservation and byte growth. A failed preparation cannot consume a
copy reservation. Shared refusal stays sticky. rewrite_imported_source already
reserves its actual parent/payload copies before materialize_source and must
not be debited again through a common helper. Namespace closure also owes its
actual copy reservation before closure.write. A genuine no-copy move may avoid
copy work while retaining the existing admission, constructor and delivery
policy. Generated namespace insertions do not acquire source credit.

This is reconciliation of the approved existing metric, with no new quota,
public signature, parser, ordinary-complement certification or transport
exemption. Keep separate Word-scalar and actual writer-publication handback
gaps explicit. A producer lacking a positive fact cannot recover it by adding
a parser solely to guess origin. Source inspection and local copy controls do
not complete Native consumers or the final provider gate.

## Checked w14 declaration source handback

The next actual producer bridge uses this additive doc-hidden method on
ScopedWordOperation. It adds one method to the previous35-method carrier
inventory, with no existing signature removal or change. Include it in the
external caller and final strict SemVer inventories.

```rust,ignore
pub fn declare_source_w14_on_root(
    &mut self,
    source: ScopedWordSource,
) -> Result<ScopedWordSource>;
```

The named existing consumers are source_model_document and the actual scoped
MainPartBlocks assembly. Share the existing declare_w14_on_part_root decision
and insertion location through a concrete private planner in the existing
text module. Do not duplicate its uses-w14 or already-declared predicate.
Determine physical insertion from the actual live reader/input origin, including
UTF8 BOM. Preserve ordinary helper output and error policy, apart from fixing
an actual BOM position bug if the shared planner exposes one. Record any such
output correction explicitly with its actual discriminating evidence.

The method validates operation identity and sticky refusal. A no-insertion
result retains the original bytes without copying or renewing an input role.
For insertion, prepare the exact selected before/after source images and the
actual declaration bytes, then use checked composition as the actual buffer
producer. All applicable existing reservations precede that byte growth.
Generated declarations gain no source credit. Constructor facts, mapped
publication, Word events and delivered coverage follow the real producer.

Native assembly consumes this actual Source before obtaining any separately
required legacy record through paid repeat_source/finish_source. It must not
mutate a legacy Vec first and recover a Source by equality afterward. Keep
target-event replay and later writer/parser handbacks separately pending.
No new public type, generic, quota or namespace policy is introduced.

Local controls distinguish no-op moves, root binding already present including
a foreign URI, missing declaration with actual use, no-use input, aliases/BOM,
foreign operation, exhausted known-Field copy work and unchanged caller state
on failed preparation. Compare ordinary/scoped output through the actual
shared helper, and require a real generated w14 main program with its Source
handback. No extra parser solely to qualify origin is permitted. HLD03/04/10
remain the existing impact list. Whole C1 and every final rider remain pending.

## Returned model Text companion for actual reparses

The actual canonicalization/optimized reparse consumers cannot pass ordinary
returned CT_Text through the original frozen-model reference arguments. Close
that callable gap with these two additive doc-hidden operation methods. Existing
36 signatures remain unchanged except the explicit note-producer mode amendment
below. The carrier method inventory becomes38, and
there is no new public type or constructor.

```rust,ignore
pub fn source_model_document_from_read(
    &mut self,
    document: &CT_Document,
    source: &ScopedWordSource,
) -> Result<ScopedWordSource>;
pub fn source_model_body_item_from_read(
    &mut self,
    frame: &CT_Document,
    item: &BodyContent,
    source: &ScopedWordSource,
    body_index: usize,
    ancestors: &[&[(String, String)]],
) -> Result<ScopedWordSource>;
```

Successful source_document retains an operation-bound private reader companion
on the existing opaque Source. Capture the actual retained Text paths and
Word origins through the same substantive factory. Discarded temporary owners
and raw fallbacks contribute no returned-model facts. No extra parser, copied
value lookup, candidate-address catalog or fresh frozen baseline is allowed.

The actual Native producer explicitly pairs the returned document or selected
body item with that Source. Whole-document consumers are bookmark/TOC and
drawing/relationship reparses. Body-item consumers are parse_one_child and
read_source_child followed by the actual MainPartBlocks writer. The wrapper
body_index is0 for a one-child parse. Real assembly remaps selected positions
at its handback, not by guessing later ordinals. Old/current and updated Sources
remain distinct throughout optimized qualification.

The new writers validate operation identity and successful read facts, then
bind actual borrowed model selections to the retained origins for the lifetime
of the substantive writer. Reuse existing Field registry/minted receipts and
retained-reader raw maps. They do not enroll returned Text as another input.
Real edits invalidate or remap facts at their producer. A payload Source copy
does not establish that its model was read. No-op transport may preserve the
same actual receipt without granting input credit.

Retain old signatures for initial frozen-model writers. The new methods are
only for their actual paired reader result and surviving producer paths. Source
and model mutation outside that pairing cannot recover authority by equality.
No new quota, source policy or deep security investigation is added. Final
SemVer/external caller inventory includes both additive calls.

## Related-reader and linear control accessor closure

Adopt the same returned-model companion for the actual note-family reader and
writer. The full additive operation inventory becomes40, preserving the prior
36 signatures and the two document/body methods above, subject to the explicit
note-producer mode amendment below.

```rust,ignore
pub fn source_notes(&mut self, source: &ScopedWordSource) -> Result<CT_Footnotes>;
pub fn source_model_notes_from_read(
    &mut self,
    notes: &CT_Footnotes,
    endnotes: bool,
    source: &ScopedWordSource,
) -> Result<ScopedWordSource>;
```

Use the existing substantive CT_Footnotes factory for both note families. Its
successful actual read captures the same private Source companion. Native note
cache refresh and allocated-note discovery supply the actual paired result to
whole-note and selected-note writers. The existing source_paragraph factory also
captures its actual returned paragraph facts. Its selected body writer uses
index0. No wrapper reparse, new public type, new input credit or equality-based
origin recovery is introduced. Partial, discarded and raw-fallback owners do
not become returned typed facts.

Publish these existing borrowed CT_Sdt methods as doc-hidden public methods,
with unchanged signatures and bodies:

```rust,ignore
pub fn retained_resource_run_xml(&self) -> impl Iterator<Item = (usize, &[u8])>;
pub fn original_resource_runs(&self) -> impl Iterator<Item = &CT_R>;
```

The borrowed accessor inventory becomes7. Both the existing lower freezer and
the actual Native clone/prepared-writer traversal consume the linear iterators.
Repeated per-index retained_run_xml searches are not a replacement. These
accessors expose existing retained runs without copying, new mutable authority
or a new ownership policy. Include all six additions in the external callable
inventory and affected local checks. This is functional handoff closure. The
deferred security assessment and final union gates remain pending.

## Actual selected-note writer handback

The existing Native serialized_note_fragment emits a selected note under a
one-note part root. A returned note at physical index greater than0 cannot use
the whole returned-vector receipt as index0. Adopt this additive doc-hidden
method, preserving the existing40 carrier names and seven accessors:

```rust,ignore
pub fn source_model_note_from_read(
    &mut self,
    note: &CT_Footnote,
    endnotes: bool,
    source: &ScopedWordSource,
    note_index: usize,
) -> Result<ScopedWordSource>;
```

The carrier inventory becomes41. The actual selecting producer supplies the
physical index from the successful note-family reader, together with its
selected note or actual clone. Use only that note's returned paragraph/raw facts
and existing Field receipts. Emit the same one-note part root through the same
concrete note-family writer, then use existing checked selection and namespace
closure for the actual fragment. Do not serialize unrelated notes, rediscover
the owner by id/value/address or run an extra parser to qualify its origin.

Reuse the existing note writer with a borrowed one-note slice where practical,
avoiding another typed clone solely for serialization. Add a selected later-note
control with distinguishable preceding content in both families. Keep whole-note
flush and initial frozen-clone writers on their existing methods. Include the
additive call in external caller inventory. No new public type, quota, input
allowance or broader note-editing contract is introduced.

## Exact Native note drawing preparation mode

The same initial source_model_notes method serves two actual Native producers.
Whole-new-part flush already prepares fresh inline/anchor drawing namespaces.
The selected one-note fragment producer does not. Preserve that distinction by
adding the final prepare_drawing_namespaces bool to the canonical signature
above. This supersedes only the earlier unchanged-signature statement for this
one unreleased doc-hidden method. Carrier names stay41 and accessors stay7.
All other40 carrier signatures remain unchanged.

Pass true only from the actual whole-new-part producer that already performs
authored drawing namespace preparation. Pass false from the initial selected
one-note container. The selected returned-reader writer retains its ordinary
selected-fragment behavior. Use a private scoped FieldXmlEmission context and
the same substantive note/drawing writers. Fresh raw_xml=None inline/anchor
roots receive the existing local r/wp declarations only in the prepared mode.
Parsed/raw drawings and opaque bytes stay unchanged. No posthoc equality-based
origin, additional parser, public type, quota or new namespace policy is added.

Check actual producer parity separately for selected append and whole-new-part
flush in both families, including inline/anchor cases and the charged downstream
Text consumer. Do not label a whole-part preparation fixture as selected-append
parity. Public special-note replacement also uses a selected-record producer
and must retain its existing behavior.

The pinned local v0.16.4 document.rs contains neither ScopedWordOperation nor
source_model_notes. This is an amendment to the unreleased C1 bridge, not a
change to a published signature. Preserve earlier external compile observations
as historical and update the final external caller to the amended signature.
No new output delta or user-visible policy change is approved here.

## Scoped cache decline and authoritative reader

A pre-operation MainPartBlocks parsed-child cache has no successful-reader
companion for this operation. In the scoped path, decline that optional model
reconstruction and use the existing authoritative whole-document factory once
with its actual Source. This replaces the chosen reader path. It is not an
additional parser that qualifies a cached model after the fact. Keep the
separate qualified kept.main_document reuse and ordinary None-context cache
paths unchanged. No value/address origin inference, renewed input, conservative
valid-input refusal or extra baseline is permitted. Performance riders remain
mandatory and may expose an optimization need, without blocking the initial
correct functional path on an unproved cache-origin mechanism.

Actual whole-reader output supplies its ordinary Text/raw companion to the
next substantive writer. Field-bearing and replacement-child factories may
retain their existing optimized route only when an actual same-operation
companion is available. Ordinary optimized Some controls remain unchanged.

## Derived Field preparation handback

The repeated-fragment loader returns Fields from an actual same-operation
reader. Those Fields are not frozen operation inputs, so Native must not call
ordinary prepared_current_field or register their model as new ingress. Add
this unreleased doc-hidden operation to preserve the existing preparation order:

```rust,ignore
pub fn prepare_template_field_from_read(
    &mut self,
    current: &Field,
    source: &ScopedWordSource,
    normalize: bool,
) -> Result<Field>;
```

The final carrier inventory becomes42 names with7 borrowed accessors. All
existing41 signatures remain unchanged. Require this operation's successful
model-read Source and an actual mapped parsed Field from the paired returned
model. The caller retains that pairing as for the returned-model writers.
Forward existing original/source and prepared-root facts through the same
Field clone and normalization implementation. Do not add a local model to the
frozen catalog, create another operation or grant a new input allowance.

The Native preflight still chooses unchanged imported bypass versus actual
normalization. Failed semantic preparation remains atomic and sticky resource
refusal stays on the shared operation. Actual downstream read/prepare/write
controls must include a parsed Field, a prepared model Field, nested cached
content, bypass behavior and the next charged Text consumer. Foreign-operation
Sources, Sources without a successful read and unregistered Fields must refuse.
Same-operation caller pairing remains an explicit responsibility, as with the
existing returned-model writers and the deferred pairing assessment. Extend the external compile inventory after
source closure. This is functional producer closure, not completion or a
waiver of deferred assessment and final riders.

## Scoped revision installation and comments carrier closure

Root technically approves the following additive closure of the existing C1
consumer map. The signatures are unreleased doc-hidden Rust bridges. No budget,
quota, input baseline, Python/CLI method or public context type is added. This
supersedes earlier carrier totals only. The final inventory becomes46carrier
names and7borrowed accessors. All existing42signatures remain unchanged.
Implementation, external caller compilation, HLD reconciliation and every
existing provider and final gate remain required. Approval of these contracts
is not implementation completion and does not reset an exhausted peer cycle.

### Actual prepared-run installer

Native prepare_current_field_models currently installs a prepared run through
the ordinary CT_P setter. Its Revision path creates local Work and discards
returned XML correspondence. Existing prepare_template_field methods only
return Fields and cannot own the containing run installation. Add:

```rust,ignore
pub fn install_template_prepared_fields(
    &mut self,
    paragraph: &mut CT_P,
    path: &AcceptedRunPath,
    fields: Vec<(usize, Field)>,
    frozen: Option<&'a CT_P>,
) -> Result<bool>;
```

The operation owns the actual current-run clone and Field-slot substitutions.
Validate the selected path, unique existing Field slots, actual same-operation
prepared facts and initial frozen paragraph membership before installing.
Reuse the existing checked private accepted-run installer with the shared Work
and emission context. Preserve the Native semantic-equality no-op decision and
per-Field preflight/preparation order. Do not accept an arbitrary replacement
run as evidence that its ordinary Text was cloned from the current run.

The actual successful Revision factory may retain a private, numeric
raw-coordinate and paragraph-read companion anchored weakly to the existing
FieldOriginRegistry. This is an actual owner handback, not a global address
catalog or a new input source. Stage it until the existing constructor succeeds.
It contains no borrowed model, Source or Work reference. A dead or foreign
registry supplies no authority. Clone forwarding follows the actual raw and
paragraph clone, including Arc::make_mut before a scoped edit. Ordinary changes
to raw or payload state invalidate the companion. Metadata-only source-ID
rebinding must retain the actual successful constructor correspondence.

Scoped refresh retains the full mapped paragraph writer result, combines the
actual old root prefix and suffix with that produced body and feeds the same
existing Revision factory with produced WordFieldInput and initial_delivery
false. Materialize needed numeric publication coverage before old collectors
disappear. Forward the actual fresh Field maps and facts through the existing
registry-aware retain_ordered_field_source_identities_in handoff before source
identity COW. Commit the updated revision and companion only after the same
factory succeeds. Do not add another parser or normalize under a local Work.

Initial clone lookup may remove its blanket Revision exclusion only once this
installer and actual registry-aware handoff exist. Keep the prepared-state and
exact replaced owner/path/slot exclusions. Untouched sibling slots after refresh
rely on the existing validated accepted-path/content-slot relation, not equal
text or a guessed later ordinal. Replaced slots never reacquire initial roles.
Changed revision branches use their own actual returned companions. Unchanged
branches keep their actual initial clone pairs and must not lose Word input
because a different revision was changed.

The decisive local control prepares two distinct Field runs inside one initial
insertion with ordinary A&B Text. The first install forces refresh, the second
still uses the same operation and its actual untouched slot relation, both
changes survive and the next charged Text read succeeds. Include same-run
multiple Field slots, a nested revision, semantic no-op retention, a repeated
query of an already installed slot, ordinary mutation invalidation and foreign
or dead companion refusal. Preserve exact error order and held input on failure.
These are functional controls, not a replacement for deferred deeper checks.

### Comments part reader and substantive writers

The dirty comments flush and its later refresh currently use ordinary
serialization and parsing. Paragraph serialization alone cannot preserve the
comments headers, paragraph IDs, opaque owner order or whole-part factory
handoff. Add these three operations on the existing ScopedWordOperation:

```rust,ignore
pub fn source_model_comments(
    &mut self,
    comments: &CT_Comments,
    text_origins: &[(&CT_Text, &'a CT_Text)],
    field_origins: &[(&Field, &'a Field)],
    raw_origins: &[(&[u8], &'a [u8])],
) -> Result<ScopedWordSource>;

pub fn source_comments(
    &mut self,
    source: &ScopedWordSource,
) -> Result<CT_Comments>;

pub fn source_model_comments_from_read(
    &mut self,
    comments: &CT_Comments,
    source: &ScopedWordSource,
) -> Result<ScopedWordSource>;
```

Share the existing substantive comments writer. Preserve its validation order,
conditional canonical w/w14 declarations, retained attributes, emitted binding
scope, root/comment extra_xml order, author/date/initials/id and the real
paragraph_ids argument to each paragraph writer. The ordinary Vec writer and
the scoped normalized writer are two actual consumers. Use that same writer
with the shared operation and emission context, then the existing checked w14
root declaration policy. Do not serialize ordinary paragraphs and wrap their
strings into a newly qualified comments part.

At the existing comments factory, collect actual comment-vector and paragraph-
vector companions only on successful pushes. Preserve its part admission,
namespace/EOF/error behavior and paragraph-ID alignment. Empty owners supply no
invented Text. Capture root and per-comment opaque XML maps at their actual
capture and push events too. Opaque maps only forward selected ranges and do
not certify ordinary or Field semantics. Commit the companion only after the
same whole factory succeeds. No additional parser qualifies an ordinary result.

Native dirty flush supplies the actual current/frozen clone pairs, including
root and comment opaque slots. Capture a private frozen comments-root handle
at the existing staging clone producer and include its actual root/owner raw
records in existing operation membership without activating them. Preserve the
actual drawing-namespace preparation producer and its returned correspondence.
After refresh, retain the Source next to its actual installed package Arc and
use the returned-model writer for that paired model. Do not introduce a second
Document cache or recover provenance by comparing serialized bytes.

The initial scoped fragment loader and the comments branch of the existing
story setter also use source_comments at their existing factory call. Thread
the actual related comments Source through the private loader context. Do not
parse ordinarily and call the scoped reader afterward. Initial frozen authoring
uses the model writer, and derived imported comments use their actual successful
reader handback. These are current distinct consumers, not future routes.

Local controls compare complete ordinary/scoped comments bytes and live model
content with zero/one/many comments, paragraph IDs, retained root/owner XML,
Fields, an inline control and a revision. Exercise initial frozen authoring,
actual imported-fragment read, dirty flush, refresh and subsequent emission,
including the next charged visible Text consumer. Preserve opaque content and
caller-held input on refusal. Schema, bindings and native Word checks retain
their final integrated timing. No selected-comment emitter or mode flag is added.

### Inventory and HLD adoption

Extend the existing single external caller inventory from42to46methods before
its final compilation. Keep its earlier versions as historical evidence.
HLD impact remains exactly03-architecture.md for shared operation and reader/
writer lifecycle,04-opc-and-packaging.md for comments parts and opaque/namespace
preservation, and10-bindings-spec.md for hidden Rust carrier inventory. There is
no Python or CLI method addition. All prior C1 and provider residuals remain.

## Internal template numbering projection closure

Root approves this private scheduling clarification after the actual canonical
join regressions. It changes no public API, quota, input allowance or review
cap. The existing optional TemplateMainWriter distinguishes private internal
template staging from ordinary/final public package projection. With that
context present, the shared substantive flush emits the actual live main model
and its matching live numbering definitions once. It performs no numeric
permutation and returns empty remap tables. The operation-bound Source and
its temporary package bytes therefore describe the same numbering ID space.

Without that context, retain the existing canonical numbering projection,
unconditional live-model/body restoration and emitted maps. This includes
rendering's final prepare_and_reopen_staged and ordinary save, sign, output
and same-document reopen paths. Public canonical output and original live
numbering handles retain F277b's contract. Internal fragment imports still
use their existing actual allocation/remap policy. No blind inverse, new
source catalog, output-equality origin lookup or second Field materialization
is added. Remove the superseded duplicate pre-publication Source/descriptor
capture and temporary fragment workaround when the common producer already
emits a consistent live package.

Local composition controls preserve the exact original callback location,
actual same-document selected instance and start override through scalar and
repeated rendering, and a prepared Field with numbering present. Existing
numbering publication/restoration/refusal and scoped template/loader/import
controls remain applicable. The earlier base-level-reader diagnostic is an
invalid premise, with its correction preserved. The corrected actual instance
override counterexample and all real Field-copy refusals remain retained.
These are ordinary functional regressions, not a deeper security assessment.
HLD03/04/10 describe the private live-ID staging and final canonical projection.
All existing provider, final native/oracle/full/installed/hash/CI/release
requirements remain. No provider or sprint completion follows this approval.

## Current local provider implementation acceptance

Root accepts the completed local implementation at the source identities in
/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/returned-reader-functional-batch/ROOT_TAKEOVER/root-canonical-join-1/LOCAL_PROVIDER_EVIDENCE.json.
The named genuine LOCAL regression gate passes30Rust Issue425 cases,35focused
comparison cases,2CLI cases and4freshauthenticated opt2Python nodes. Supporting
scoped cohorts, external46/7client compilation and affected warnings-denied
all-target/all-feature Clippy pass. Whole-provider Root final1reports zero
defects and zero smells. Current source includes the reviewed canonical joins
and private internal-live/final-canonical numbering clarification. No known
ordinary local implementation finding remains in the approved contract.

This is the approved LOCAL provider checkpoint for run-sprint after actual
structured handoff validation and integration. It permits recording provider
implementation completed in existing run state and done delivery dependency
cells only when the integrated source is an ancestor of the consumer Base.
It does not itself claim integration or unblock a consumer. The original
recorded recovery Base706e857f and later source reconciliation HEAD1d2c5a80
remain distinguished, with all failed attempts retained.

Keep this plan approved and the final checklist item pending. AS_BUILT and
SPRINT_TRACKER wait for actual final acceptance. All named final Word/SDK/LO,
full installed/typing, corpus/hash/pixel, hosted CI, performance, WASM/archive,
workspace and sprint review obligations remain. The narrow Q-A430-28 positive
native observation stays unproved under its user-approved exception. Generic
typed note attribute recovery is outside the operative contract, not a claimed
repair. Optional deeper security work stays deferred separately. Publication
still requires final sprint close and a separate approval at the reviewed SHA.
