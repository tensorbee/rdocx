# F-X272, Preserve existing revisions during document comparison

**Status**: approved
**Sprint**: S96
**Size**: L
**Depends on**: F-X229, F-X230, F-X245

## Problem

`crates/rdocx/src/comparison.rs` rejects compared stories containing modeled revisions. Issue 425 requires a second reviewer to retain a prior redline while tracking changes elsewhere, or a documented limitation and usable workaround. PR427 head `c1e261ca63ebda5b7da1adec886ee40364a4e3de` supplies bounded owner preservation, with material exclusions. Its stale donor measurements and 116 missing main commits cannot be adopted.

## Spec reference

- `docs/hld/03-architecture.md`, checked comparison and physical story ownership.
- `docs/hld/04-opc-and-packaging.md`, "Producer read and writer authority" and "The package".
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

Q-S96-04tracks normal-note identity, standard tracked representation and selected
resolution/package cleanup. Use actual reference/owner association rather than
positional pairing or bare numeric ID equality. Preserve special separator owners,
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
