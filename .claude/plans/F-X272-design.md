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

F-X229, F-X230 and F-X245 are done in BACKLOG. F-X271 is an execution-order tooling prerequisite, not a new formal product dependency. Final native/installed evidence stays pending. Q-S96-01 is answered. No claim until all S96 designs close.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Exact Issue 425 two-reviewer reproduction | First author's two revisions retained, new author deletion/insertion added elsewhere | Rust/Python/CLI | Adapt `issue425_comparison_keeps_existing_revisions_and_tracks_changes_elsewhere` and exact Python repro | Installed 3.9/3.12 and native retained revision inspection |
| A2 | Prior revisions in body, table, control, header/footer, footnote/endnote, text box, section properties | Unchanged owners retain XML/authors/dates/ids/location, noncolliding new ids in compared scope | All compared stories | Extend existing regression carrier with per-owner cases | Combined source-qualified native save/reopen |
| A3 | New author accepted or rejected | Reproduce edited or original input respectively while preserving first review | Native all-story resolution, Python, CLI | Author-filtered resolution and independent byte/order assertions | Installed suite and native author-filtered controls |
| A4 | Changed/deleted/moved/inserted revised owner, final paragraph mark, changed revised section | Under the approved bounded policy, location-qualified refusal and no source/live-owner/output change | Rust/Python/CLI | Adapt `issue425_comparison_refuses_changes_that_touch_an_existing_revision`, add moved/table/control cases | Limitation and executable checked workaround, Q-S96-01 |
| A5 | Existing revised comment plus comment graph change, ignored comments/stories | Whole comment replacement refuses if it would overwrite revisions, ignored sources follow current exclusions | Comparison/comment graph | Existing comment-package refusal plus ignored-story controls | Installed and combined graph acceptance |
| A6 | Identical redline, no-op and late invalid input | Zero new counts, exact no-op, late refusal leaves original and output intact | CLI human/JSON, native, held Python handles | New-only count and late-failure controls | Current installed runtime and strict typing/stubs |
| A7 | Unsupported edited revised owner | Documented limitation plus tested checked workaround, no preservation claim after resolving baseline | Existing `gaps.md`, public APIs and CLI help/stub docs | Execute public workaround on same input | Issue 425 alternative acceptance, Q-S96-01 approved |

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

Contributed acceptance adds pinned full issue/PR rows and omission mapping. Parser/serializer and XML preservation add alias/schema-order and opaque-byte controls. Compatibility, multi-consumer and cached projections add same-input native/binding/CLI and held/fresh/reopen comparisons. Public behavior/binding docs add final WASM compile, installed typing/stubs and affected archive dry-runs/size. External native oracle adds actual producer/tool/input/output identities. Metadata adoption adds affected final policy checks. No new generic, trait, module, crate or source file is planned.

## Hash harness

Expected unchanged 49 hashes/7 pixels. New comparison behavior is a separately labelled commit. No baseline change is predicted from this story.

## Implementation checklist

- [x] Resolve Q-S96-01 and close concrete workaround signatures before approval.
- [ ] Reconcile own PR427 commits against current Base without donor measurements.
- [ ] Guard all revised-owner mutations, preserve exact old bytes.
- [ ] Prove newly allocated id postconditions and new-only CLI counts.
- [ ] Cover all A-IDs and held/fresh/late-refusal controls locally.
- [ ] Align public docs/stub/help and existing gaps record.
- [ ] Update listed HLD intent, scoped gate, independent microscope, prepared handoff.
- [ ] Retain final native/installed/archives debt until actual integrated acceptance.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision and recommendation | Status and answer/source |
|---|---|---|---|
| Q-S96-01 | F-X272 A4/A7, combined S96 acceptance | Recommend bounded unchanged-owner preservation plus documented/tested checked resolve-and-compare workaround. Changes within a revised paragraph, whole table/control or revised section remain unsupported. Alternative: expand implementation to carry revisions through such edits, increasing scope and requiring a revised design/estimate | Answered 2026-10-28 by the user: "Receommeneded", selecting recommended bounded preservation plus tested checked workaround. The issue permits this alternative |

## Contract closure

Q-S96-01 is answered. No new public method or option proposed. Exact Python defaults and checked all-story author/id resolution signatures are transcribed above from current binding/stub and `revision.rs:251to278`. The executable workaround still needs actual local evidence during implementation. Documentation alone cannot assert a usable route. Compatibility source distinction is verified, native saved-artifact proof remains due.

## Intake identity

Read-only GitHub refresh on 2026-10-28 found no body or contribution-head changes from the S96 intake. Original payloads and comments are retained under `/private/tmp/rdocx-s96-run-refresh`.

- Record 425, updated `2026-10-27T20:08:05Z`, body SHA-256 `3f4f9b5577506d36ec13cdb09f145d551cce7701e93f4fc53419f0a74f78d26b`.
- Record 427, updated `2026-10-27T20:59:51Z`, body SHA-256 `5cef6bc8e0440076c8da5c46778ee22595c208f6698f367d2fbe9635208a46f4`.
  Contribution head `cc3473d0c05659167847a941e0c7eb2bbd7217f5`.

The draft initially assumed a `gaps.md` file existed. Repository inventory shows none. Put the limitation and checked workaround in existing README and CLI comparison documentation instead. This is a documentation-location correction, without changing the accepted bounded policy or adding an unapproved source file.

Design closure on 2026-10-28: consumer map, observable A-IDs, discriminating cases, public signatures, imported states, HLD impact and local/final evidence were checked against current source. Q-S96-01 is answered in F-X272. Approvals follow provider contract order. No material user decision remains. Pending native/hosted/installed observations are retained final acceptance debt, not recorded passes.

## Implementation-time contribution refresh

Read-only refresh on 2026-10-28 after the user scope answer found PR 427 rebased onto released main `24af95d3d94c25e95f61a75e7e21f18fa95b6363`, with current head `c1e261ca63ebda5b7da1adec886ee40364a4e3de`. Its current body SHA-256 is `dc490c52d857c06751586bf2a945ee1ecb6cc81922b03c21cf40b0b3544c643d`, updated `2026-10-28T06:43:32Z`. Issue 425 body is unchanged. Historical initial pins above remain attributed to that intake.

The current four own commits are `56523be00a3c58b4b730717af99ce66e57911f88`, `e2796f0567d26c47688fdc8057d8cde21417ca06`, `9d0744301b604fdac003686d47690a67028a5100` and `c1e261ca63ebda5b7da1adec886ee40364a4e3de`. Comparing per-file added/removed payloads confirms all product code, tests, stub and HLD contribution changes are identical to the original pin. Changed README/archive measurements remain excluded and will be remeasured on the final union. The issue acceptance and Q-S96-01 decision are unchanged.

Raw refreshed bodies/comments/PR metadata, commits and exact current diff are retained under `/private/tmp/rdocx-s96-run-refresh`. The running worker keeps Base `c8acc502b74f29ed7e9546939f17007709a9f22b` and its dirty implementation/evidence. Its owner must acknowledge this canonical pin revision before further implementation. No reset/rebase or source replacement is authorized by this factual refresh.

## Source relocation contract clarification

Independent namespace investigations identified lost source scope across comparison descendants and revision paragraph merging. This is an A2/A3 preservation repair within the approved bounded policy. It does not authorize edits inside prior revision-bearing owners. The existing `revision.rs` renderer is the private provider for main and related-story accept/reject operations, including all, author, id, id-set and date selection. Its paragraph merging, kept-wrapper promotion and property replacement must preserve the expanded names of transported elements and attributes.

Use actual source ancestry and destination context at relocation. Preserve the first merged paragraph shell and attributes, the last paragraph's property precedence and each child's own source scope. Carry incoming wrapper promotion into a merged shell, preserve child-local shadowing and an absent default namespace, and never globally rebind producer w. Ordinary unchanged-context traversal retains its bytes. Add only namespace closure needed for actual relocation. Keep selection, resolution counts, schema order, atomicity and both comparison reconstruction postconditions intact. Existing parent links and local declarations provide source ancestry. No new public signature, parser, trait, module, crate or source carrier is authorized or needed.

Local evidence adds the actual final-mark failure with inherited q properties and foreign producer w, plus direct accept/reject resolution with conflicting source paragraph scopes, a child-local shadow, opaque element/attribute identity, save/reopen and a related story. Preserve genuine pre-repair outcomes for newly added controls. Batch the affected cases with `unwrapped_revisions_keep_wrapper_namespace_bindings`, `rejected_property_changes_keep_owner_namespace_bindings`, `contextual_paragraph_markers_merge_the_adjacent_paragraphs`, `contextual_cleanup_preserves_foreign_property_shells`, `paragraph_mark_removed_before_a_table_removes_an_emptied_paragraph` and `rejecting_paragraph_property_change_keeps_mark_and_section_properties`. Retain prior comparison, ignored-source, workaround and revised-owner refusal controls. These are focused local regressions, not complete suites or native oracle acceptance.

The existing HLD03/HLD10 work list remains unchanged and must describe this shared source authority. Namespace investigation pass 2 records the observed 12-test result, 11 passed and 1 failed, separately from source-traced challenges and prior setup/semantic failures. Full feature microscope and every final rider remain pending. Before implementation resumes, the running owner must acknowledge the canonical revision containing this clarification, preserve Base and dirty work, and reconcile this section into its annotated worker plan without resetting it.

## Revision identity clarification

Independent investigation pass 3 establishes N3, separate from N1 transport: an unrelated public comment edit removes redundant paragraph namespace declarations while retaining identical old revision bytes and effective bindings. `CT_Revision` equality currently mistakes private parser-context enumeration order for an owner change. The fresh authenticated Python observation is three passing cases and one semantic failure, retained with its original wheel/module identities. No actual revised-owner edit or new user decision is involved.

Repair only identity of `CT_Revision`'s complete private parser context in the existing `crates/rdocx-oxml/src/revision.rs`. Retain the original parser vector and its order for parsing, editing and writer prefix selection. Existing PartialEq may compare those complete entries without depending on enumeration order, preserving multiplicity and every binding value. Every other field remains compared, including raw XML, kind/id/author/date, typed and projected content, nested revisions and parse policy. Do not sort the shared namespace scanner, discard context, add a semantic XML comparator, normalize published original bytes or weaken revised-owner refusal.

Add actual pre-repair low-level controls in existing carriers for empty, paired and nested/projected revision identity under reordered equivalent context, with unchanged raw bytes and parser input. Meaningful foreign/Word/default/undeclaration differences remain unequal. Preserve observable raw and projection differences. Add the matching native public comment reproduction, including original paragraph/revision bytes and comment graph after save/reopen and accept/reject of the new author. Rebuild and authenticate a successor wheel before rerunning the same four selected Python nodes. Preserve the failed predecessor and all historical observations.

Scoped checks now include the affected `rdocx-oxml` crate, in addition to the existing changed native, CLI and binding crates. Refresh the affected bounded local tests, including all 22 prior namespace, relocation, ignore, refusal and workaround controls. HLD03/HLD10 remain the complete HLD work list. The running owner must acknowledge this canonical clarification and reconcile it into the annotated worker plan before implementation, without changing Base or resetting dirty work. Full microscope and every named final gate remain due.

## Property source and publication closure

Independent investigation pass 4 clarifies D2 and reopened N1 without changing the bounded user policy. Parsed opening boundaries fix the quoted-attribute parse failures, but the actual successor loses an opaque rPr attribute after property reconstruction. Typed CT_RPr deliberately does not retain opaque root attributes, so comparison must carry private ORIGINAL and EDITED raw property-source snapshots beside typed formatting precedence. No new public model field/type, parser for unknown attributes, public API, trait, module, crate or source carrier is authorized or needed.

Retain each side's direct pPr/rPr source shell, root QName, effective namespace context, attributes and raw children before adding comparison markers. Carry the selected snapshot through emitted paragraphs, trailing and moved predecessor marks, final insertion/deletion and block-control/related-story callers. Preserve existing own/merged accepted/rejected precedence and the final paragraph's own section break. Patch only intended direct Word property/change slots, respecting schema order. An absent property owner differs from an empty owner carrying opaque metadata. Unknown-formatting policy remains left-biased where it already was. Audit ordinary paragraph and run-property regeneration under the same raw-source authority instead of leaving typed-only equivalents beside the final-paragraph repair.

Every fresh canonical Word container must close at its actual publication boundary, including generated rows, row properties, cells, paragraphs, tables, controls, fields and property early returns for equal or ignored formatting. Reuse the existing qualifier policy. Retained raw fragments keep their actual source context. Do not globally rebind producer w or put a source producer-w attribute onto a fresh Word-w root with changed meaning. A marker's local Word binding does not qualify its ancestors. Preserve old revision bytes, atomic refusal and both reconstruction postconditions.

The existing consumer map is extended by pass 4's cited publisher inventory. Local checks retain all original quoted-opening cases and add distinct own/merged property values, absent versus metadata-bearing empty owners, a source Word q:rPr with opaque producer-w attributes and a supported block-control or related-story final path. Assert expanded identities, opaque bytes, intended formatting and both author-resolution sides after save/reopen. Obtain current-source RED for the canonical/alias/foreign-w row diagnostic, then a bounded inserted/deleted row, nonempty property and existing move/control matrix that separates fresh containers from retained fragments. Do not add a new move capability or use the predecessor wheel as current-candidate qualification.

D1's stable kind/part/owner count repair has actual local CLI evidence. N3's equality-only repair remains unchanged. Refresh only affected scoped evidence and build/authenticate a fresh successor wheel for changed native source before the same selected Python checks. HLD03/HLD10 remain the work list. Full microscope pass 2 and every final sprint rider remain pending. The implementing owner must acknowledge this canonical revision and reconcile this section into its annotated plan before resuming, preserving Base, dirty source and every historical observation.
