# F-X268, Bound Word owner inventory and unchanged package staging

**Status**: completed
**Amendment note**: Codex root approved the refresh4 mechanism 8 contract on 2026-10-25 after the user answered Q24. The baseline is `d1f42516eb3747442fe4304a6331eab946dc9a0e`. The scope approved through refresh3 stays approved. F-X268 stays unclaimed. F-290 keeps its existing approved implementation and review. No gate pass is claimed.

**Current implementation readiness**: Integrated on sprint/s95.4, canonical owner codex. All eight mechanisms and subsequent bounded repairs are implemented and reviewed. Pass30 is clean. Actual affected debug corrections and A7 memory, lifetime and fallback review pass. The dated execution boundary below identifies the remaining final debt. Approval snapshots retain their historical unclaimed and not-started states. This paragraph is not a full acceptance claim.
**Sprint**: S95.4
**Size**: L
**Depends on**: F-X266, F-X267
**Estimate**: 12 to 21 developer-days for the combined scope, as a provisional Root planning judgment after the historical 10 to 17 plus a provisional 2 to 4 for this increment. Reassess it after decomposition. It is not a measurement, not a capacity promise, and not a new promise. Actual work is unknown. Do not add it to the historical seven-story figure of 27 to 41 developer-days. The 10 to 17 figure stays historical for the scope before PR422. The earlier 7 to 12 figure stays historical for the scope approved before the 419 and 420 amendment.
**Amendment**: Codex root approved the 2026-10-24 refresh2 extension. That approval stays. Codex root approved the 2026-10-24 refresh3 extension after reading the full corrected contract and the four-file delta. The canonical HEAD at that approval is `d6cefc239ab7de0d4dae63e13d730da558d3a3a5`, plus this working draft. This record is that Root review. It is not a self-review. The prior approved scope stays approved. Codex root approved the refresh4 extension on 2026-10-25 at baseline `d1f42516eb3747442fe4304a6331eab946dc9a0e`. No implementation is underway. No gate pass is claimed.

## Execution boundary, 2026-10-27

The API, preservation, refusal and acceptance contracts below remain binding. Problem descriptions, approval-time unrun statements and local provider-prefix checkpoints retain their historical source and scope. They do not describe the final integrated execution state. At this checkpoint, the native union and affected corrective cases, installed Python suites and typing, native Word and PowerPoint qualification with explicit source applicability, A7 measured memory and retained-history review, A9, the original Issue407 gate, hash49, lint, WASM, rustdoc, supply chain and policy observations are available. Actual producer identities remain attached to each observation. No old binary is relabelled as a new build.

The remaining integrated acceptance debt at this checkpoint is Cell.add_picture's required flat-cost gate, the active README example batch, the clean22-package publication dry run and the final full-union record. The original failed attempts remain evidence. Status stays approved until the required final acceptance passes. When this plan becomes completed, its final source-bound delivery entry in docs/sprints/AS_BUILT.md supersedes these historical execution-status statements. Completion changes neither the contract nor any threshold. Build-only hosted wheel qualification belongs to release preparation before final release approval. The subsequent F-X270 intake prepares fresh0.16.4/0.14.4 patches within this sprint because the old0.16.3/0.14.3 candidates were already published. It is not an additional acceptance criterion in these eight sprint designs. Publication remains separately approved at the exact reviewed main SHA.

## Problem

Direct body insertion and later package staging still rebuild more of a Word document than the changed owner, and a retained namespace owner can leave a body insertion unsavable only after the model has already changed. The current native methods are infallible and mutate before they return:

- `Document::add_paragraph` at `crates/rdocx/src/document.rs:26608`
- `Document::add_table` at `crates/rdocx/src/document.rs:27239`
- `Document::insert_paragraph` at `crates/rdocx/src/document.rs:27272`
- `Document::insert_table` at `crates/rdocx/src/document.rs:27292`

Python publication already goes through `prepare_word_revision`, `publish_api_candidate`, `publish_owner_candidate`, and `publish_form_candidate` in `crates/rdocx-py/src/document.rs`. Those paths are the binding consumers of the inventory. They are not a new public API.

The cache and piece facts below are copied from the pinned PR417 and PR418 trees. They are explained here and have not been executed on this branch.

## Spec reference

- `docs/hld/14-development-backlog.md`, F-X268. The test gate is regression.
- `docs/hld/03-architecture.md`, Owned publication and causal paths.
- `docs/hld/04-opc-and-packaging.md`, The package, and Checked Word operation source ownership.
- `docs/hld/10-bindings-spec.md`, Python Word document operations.
- `docs/hld/12-testing-strategy.md`, Binding tests, and Word document-operation parity controls.
- `docs/hld/02-scope-and-non-goals.md` is unaffected. This story does not edit it and it is not an HLD impact.

## Approach

One story, eight serial mechanisms. The first four stay as approved. The fifth and sixth are the approved 2026-10-24 refresh2 extension. The seventh is the approved 2026-10-24 refresh3 extension. The eighth is the approved 2026-10-25 refresh4 extension. It is the same F-X268 story. It is not a ninth story and it adds no stable public API, dependency, or module. The Q19 and Q21 behavior is its own labelled commit. Performance and memory steps are separate commits. No stack is reimported. The reviewed F-X266 and F-X267 source stays. F-288, F-289, and F-290 keep their approved contracts. Their fact-changing publishers are named in the consumer map. The F-289 and F-290 publication proofs and related-owner proofs stay. They are not regressed by the seventh or eighth mechanism.

Captured increment boundaries, not the stacked diffs against main:

- Old PR415 `0fbf3e96709198dbd01b9e810181109d106d9f68` to new PR415 `2839e22c14f96aeaeab5245581b5681509774559` is a diverged endpoint tree of 9 paths: `README.md`, `crates/rdocx-oxml/README.md`, `crates/rdocx-oxml/src/document.rs`, `crates/rdocx-py/src/document.rs`, `crates/rdocx-py/tests/test_core.py`, `crates/rdocx/src/document.rs`, `docs/hld/03-architecture.md`, `docs/hld/10-bindings-spec.md`, `scripts/readme_doctests.py`.
- PR417 `bdf49f0005d1d15078b76b5da83c5f3e2b8cce12` has six own commits after the new PR415 head. Its own tree is 14 paths: `README.md`, `crates/rdocx-oxml/README.md`, `crates/rdocx-oxml/src/borders.rs`, `crates/rdocx-oxml/src/text.rs`, `crates/rdocx-py/src/document.rs`, `crates/rdocx-py/src/lib.rs`, `crates/rdocx-py/src/paragraph.rs`, `crates/rdocx-py/src/run.rs`, `crates/rdocx-py/src/table.rs`, `crates/rdocx-py/tests/test_core.py`, `crates/rdocx/src/document.rs`, `crates/rdocx/src/lib.rs`, `docs/hld/03-architecture.md`, `scripts/readme_doctests.py`.
- PR418 `dba496fcd8b4d482995a23ab81d27ad2bfcbf5e7` has six own commits after merge `0912a40e19950b223745ef83e8164266b4c36b66`. The merge parents are the PR417 head and PR416 head `768f8c109bd43fb9c1f6ecdfe2d913382bcc0368`. Its own tree is 8 paths: `README.md`, `crates/rdocx-py/src/document.rs`, `crates/rdocx-py/src/lib.rs`, `crates/rdocx-py/tests/test_core.py`, `crates/rdocx/src/document.rs`, `crates/rdocx/src/lib.rs`, `docs/hld/03-architecture.md`, `scripts/readme_doctests.py`.

Own increments only, still candidates and not accepted code:

- PR419 `9916bcdc01bc4acc065ae45afa8bd8ae07a2d30c` follows PR418 `dba496fcd8b4d482995a23ab81d27ad2bfcbf5e7`. It has 3 own commits and 7 paths. The own diff SHA-256 is `0de97b8eb0df312584bb86976676eb8fc66d7be694dd96a99830b0647836cc53`.
- PR420 `37851676af3e82a0500c4743ef86da417950d2c3` follows that PR419 head. It has 5 own commits and 11 paths. The own diff SHA-256 is `5b94cfeb7854c8d86a5f76c1d6a32424be8faa6d6af1b565af6eeb6fc28da728`.
- PR421 `ab038a21bb88a38d46c8f640587c3a20b57f6770` follows that PR420 head. It has 7 own commits and 10 paths. The own diff SHA-256 is `50c390ff8038d69b95d19510122ab51d2fd3c37b30f5b5d79a89dc610f805e35`. The file is `/private/tmp/s95.4-refresh3/pr-421-own-increment.diff`. The stacked diff against main is not this increment. The 10 paths are `README.md`, `crates/rdocx-oxml/README.md`, `crates/rdocx-oxml/src/borders.rs`, `crates/rdocx-oxml/src/document.rs`, `crates/rdocx-oxml/src/text.rs`, `crates/rdocx-py/src/document.rs`, `crates/rdocx-py/tests/test_core.py`, `crates/rdocx/src/document.rs`, `docs/hld/03-architecture.md`, and `scripts/readme_doctests.py`. `content_control.rs`, `field.rs`, and `settings.rs` are not in those 10 paths.
- PR422 `cf7bf55edfab1e0617f6fe35810355cb02126018` follows that PR421 head on a linear first-parent chain of 5 own commits. The compare is ahead by 5 and behind by 0. The merge base is the PR421 head. The own diff is 60325 bytes and its SHA-256 is `ee0a8ce239fafea714c7bc1a49601fab4f2c864c3931b8ff4c98c3fb9dce189c`. The file is `/private/tmp/s95.4-refresh4/pr-422-own-increment.diff`. The stacked diff against main is not this increment. The 5 paths are `README.md`, `crates/rdocx-py/tests/test_core.py`, `crates/rdocx/src/document.rs`, `docs/hld/03-architecture.md`, and `scripts/readme_doctests.py`. Comments and reviews are 0. PR422 stays an unmerged candidate and is not accepted code.

Actions run `40092396584` on the PR419 head is donor evidence only. Word fidelity job `120867623753` failed while downloading LibreOffice 26.2.5.2, with TLS `CERTIFICATE_VERIFY_FAILED`, before any fidelity comparison. The evidence file was absent. CI gate job `120878395382` failed because that Word job failed. The hash harness on that same donor run succeeded. No Word comparison was performed. A later green PR420 check is not a PR419 pass. The PR419 head did not move. A later Actions run `40103891571` concluded success, including Word fidelity job `120901503513` and CI gate job `120914750403`. That later run stays donor evidence. It is not local acceptance. PR421 checks were incomplete at cutoff 2026-10-24T17:02:28Z. No PR421 acceptance is inferred. At cutoff 2026-10-24T23:05:39Z the PR422 head has 19 completed donor check runs, with success or skip conclusions, and the donor CI gate succeeded. The legacy combined status endpoint still records `pending`. That field is recorded literally. It is not a claim that CI is incomplete and it is not acceptance. Donor CI is not our pass. Final fidelity on the own union stays mandatory. No product hash or render defect is inferred from those logs.

Donor timings and the donor 68 MiB claim are unverified. Reported 57 and 18 call timings are not copied. They are not added to the historical Q23 set of 73 flat calls, the current Q24 set of 79 flat calls, or the 5 linear calls. README package text, an M5 Max figure, and an allow-dirty note are not copied. README and crate README numbers in those trees are not copied. The donor figure is not accepted.

### 1. Qualified body splice and namespace replay

Amend the integrated F-X266 qualified body splice so an unrouted paragraph facade position still uses the checked route. Nested namespace replay is scoped to the exact current XML, model, and namespace-owner facts. A proof is not reused after an edit. Compatible positive owners stay. Do not adopt the donor's incidental redundant saved-namespace byte delta. Existing package bytes stay unchanged except for the requested edit.

### 2. Owner pieces

Keep one piece inventory per body block and per related physical part. Reuse on equality needs serialization and fact provenance. Semantic `PartialEq` alone is not proof. Cover the existing field source for form, span, derived parent cache, default, current, and typed nested cache, the tab source occurrence, and bookmark prefix facts, through the public and native publishers that change them. Reconstruct the same owners, parents, routes, projections, read facts, order, scope births, generations, and retired set as the checked full old inventory. Grammar is the exact qualified physical grammar, including independent text box ownership. Before any partial publish, an unsupported shape, a duplicate physical part, an invalid relationship, or a context change uses the old full rebuild. An error does not shift semantics.

### 3. Opened package staging

Share the opened package, the parsed main part, the retained owners, and the check record with staged clones. Reopen resets that baseline. Reuse only checked unchanged main bytes plus model serialization, and only the relationship, section, drawing, namespace, media, related-part, signature, and embedded facts that were checked. `OpcPackage::reread` keeps its current `trusted_baseline` precondition. Public part maps are not assumed immutable. A save-equivalent compare requires exact bytes and the same errors in the same order. A new or in-memory document falls back to the old ZIP compare. `render_template`, `compare`, and `rebuild_toc` keep the old `to_bytes` error order. A related part that changed is still checked when the main part is unchanged. A failed cache check is not success. A scope-preserving property edit, callback, or direct candidate assignment invalidates the facts it changes even when the generation is unchanged. An opaque change is not ignored. Legacy form defaults, current values, placeholders, custom XML, protection, permissions, and mail-merge stores and settings compose. Dead owners are not revived.

### 4. Binding publication

Python publication uses the prepared inventory. It allocates every result object and revision token before commit. A no-op and a refusal keep the old bytes, revisions, and live handles. Conservative publication still retires prior scopes. Append keeps only eligible owners. An owned old snapshot stays old. A returned handle uses the prepared proved route or the old fallback. It does not start a new whole-document walk.

### 5. Eligible in-memory reopen

`reopen_prepared_staged_with_limits` may reuse OPC reread only for a document opened from bytes, with `PackageReadLimits::UNBOUNDED`, and with trusted provenance. A new or in-memory document means there is no trusted opened-byte baseline, so that commit keeps the checked ZIP path. The same document may qualify after it is legitimately saved and then reopened from those bytes. A Flat OPC first commit, a limited reopen, a failure hook, and any unproved context keep the checked ZIP path. A previously authored document is not excluded forever. The carried main parse, nested namespace owners, and canonicalization record, including a cached error, move only when the main bytes and the original contextual dependencies are exact. Relationship, identifier, section, and related-part facts stay checked. A per-part identifier definition cache may replay only when duplicate ordering and errors match, including a retyped part. Exact bytes alone are not enough when relationship or type context changed. Both reopen routes carry the mechanism 6 address-read memo only under an exact current writer frame. Every subsequent lookup rechecks each aligned child byte and item coverage before reuse. A changed child drops its own facts and invalidates the body fingerprint, while unchanged prefix and suffix facts remain eligible. The carry step uses the existing qualified writer partition instead of another whole-model write. Parsed source, writer eligibility and relationship facts remain independently qualified. A debug whole-path comparison is internal agreement. It is not a substitute for release regressions.

### 6. Incremental direct-body addressing

`to_xml_with_body_spans` must reproduce the exact existing `to_xml` bytes, the direct-body order, and the head, tail, and final section context. A block source, owner, item, cell, or text box fact is reused only when the bytes match and the namespace, relationship, owner grammar, and context match. A root or frame change, and every changed fact, is invalidated. Duplicates match position and occurrence. There is no ordinal alias after prefix, suffix, or middle alignment, or after a reorder. Preserve the F-289 corrected accepted physical providers, per-part physical ordinals, and independent story scope for fields and controls. Preserve default, current, glyph, font, placeholder, and protection facts, and F-290 metadata. Do not overwrite `content_control.rs` from the donor. The existing direct id lookup and `collect_control_ids` may use an index while they preserve uniqueness, errors, and physical order. `direct_body_item_location` must equal the full story snapshot lookup for zero, one, and many items, nested controlled content, tables, cells, text boxes, and moved, removed, or repeated blocks, before and after both reopen routes. An unsupported independent block scan takes the checked full scan. An error is never an empty success. Whole-main hash and serialization stay separately counted and disclosed. A debug whole-path comparison is internal agreement. It is not a substitute for release regressions.

Direct modeled paragraphs use the same checked `direct_body_item_location` route. Controlled paragraphs retain their existing two-segment physical lookup. The changed-before-reopen control extends `main_body_layout_reopen_reuses_unchanged_children`: at 50 and 200 paragraphs, through both reopen routes, one changed unread child requires exactly one fresh child scan and returns the same full-reader location. Frame changes and unsupported bodies retain the full reader. No additional cache field, retained history or public contract is introduced.

Namespace span rebasing shares immutable declaration and binding vectors and uses the closed admitted owner-name set. It changes offsets without copying scope strings. This is representation sharing inside the existing namespace layer, covered by Q22. Exact frame, scope, partition and field guards remain unchanged. Qualified child reads borrow their cached actual parse. Consumers that need an owned document still copy the children. A source-model equality query checks every root/body frame field and every child, including a later unproved child after an earlier mismatch. It grants no writer eligibility. Opening serialization proofs use the existing exact writer-qualified piece path and its resource checks and authoritative full-writer fallback. The carried-record equality shortcut additionally requires value-reusable children and an independently successful current opening serialization proof. Fields, private facts and unsupported shapes keep the old normalization comparison.

Controlled-cell routing may prepare a complete map from one ordered physical owner inventory and nonoverlapping direct items. It preserves the old within-container cell ordinal and checks the same modeled pointer before selecting an owner. Unsupported ordering, containment or kinds decline the map and retain the original lookup and error order. Both existing batched cell projections consume this same optional map. No public method or retained cache field is added.

### 7. Main-part body-child staging

This is the approved 2026-10-24 refresh3 extension. It extends mechanisms 5 and 6. It does not replace them. The private design candidate is `MainPartBlocks` inside the existing `crates/rdocx/src/document.rs`. It is not a new file, module, or dependency. Its fields are the head bytes, the root range inside that head, the tail bytes, the kept block scans, an optional flushed part, a frame `OnceLock`, and nested namespace owners. A comment in the candidate says a child equal to a kept child that holds no field serializes to its kept bytes, and a child that holds a field is read again. That rule is the contract. `PartialEq` on a field-containing block is not reuse proof.

`CT_Document::body_frame_xml` is `#[doc(hidden)]` and returns `Result<(Vec<u8>, Vec<u8>)>`. The pair is the bytes written before the first body child, through the body start, and the bytes written after the last body child. That helper runs before the full writer declares `w14` on the root. The stored head used by `MainPartBlocks` then applies the same conditional `w14` declaration the full writer applies. `CT_Document::body_item_xml(&self, item: &BodyContent) -> Result<Vec<u8>>` is `#[doc(hidden)]`. `BodyContent::RawXml` is a passthrough without model admission, and the helper does not reject raw XML. The helper matches the whole writer on admission, bytes, error, and binding context. The item need not belong to the body. It writes the item under `root_binding_scope`. Parent binding context is exact. Inherited body bindings, unknown source bytes, root bindings, and frame bindings stay. A changed root invalidates the kept frame and the kept blocks. An alternate Word prefix on the document root is such a change. A fragment spliced into that part uses the checked whole-part read.

`BodyItemOwners` gains `holds_field: bool` as its 10th field on the final callable contract. The detector is a conservative byte-word scan for `fldSimple` and `fldChar`. It covers prefixed aliases and the default namespace because it matches those words and does not require one prefix. False positives are allowed. False negatives are not allowed. A positive result refuses reuse of that child by value. The implementation serializes or reads the changed child, or the field-bearing child, anew. Other proved blocks may still be reused. The checked whole-part fallback runs only when the child proof, the frame proof, or the context proof is unavailable. A positive field result does not by itself require that whole-part fallback for every field-bearing block. Field source, run span, nested cache, tab occurrence, and bookmark prefix publishers already named in the consumer map still invalidate the facts they change. Physical duplicates keep occurrence, order, and context. There is no ordinal alias.

Every fast path has a checked fallback. The fallback returns the same bytes and the same errors in the same order as the full writer and the full reopen. Release proof of that equality is mandatory. A debug whole-path comparison is internal agreement. It is not that release proof. An unproved child, an unproved frame, an unproved context, a failed parse, unknown source that the block path cannot prove, or a binding mismatch uses the whole-part fallback before publication. A field-bearing block whose child, frame, and context are proved is serialized or read anew, and the other proved blocks stay reusable. That case does not require the whole-part fallback. An error is never an empty success.

`take_staging_work` changes from the retired `[usize; 4]` counter to `[usize; 6]` under `debug_assertions` and `#[doc(hidden)]`. The four-slot array is not the callable counter. Python `_staging_work` returns the same six-slot array. It is not a stable export. The order is parsed bytes, scanned bytes, canonicalized bytes, ZIP bytes written, serialized children, then children read. The first four slots are byte counts. The last two slots are child counts. A child count is not a byte count. Setup calls sit outside the counter window. Cold work on the first changed edits is measured separately and is not amortized into the warm window. The warm window starts after two changed edits and then records two distinct edits.

Final-gate query clarification: masked relationship queries may compose a fresh result from two independently writer-produced exact partitions with identical head bytes, independently checked empty frames and identical root/body namespace vectors. Unchanged closed children supply actual relative XML namespace and relationship facts at the same physical occurrence. Changed children and the new tail are checked anew. Fields, private unsupported shapes, changed heads, incomplete partitions and crossing spans keep the checked fallback. This is a temporary read-only query certificate. Generic frame invalidation, stored proof contexts, actual parsed ASTs, writer eligibility and decoder facts stay unchanged. Replay and opaque-occurrence masking remain authoritative, including an opaque occurrence sharing a modeled ID. No new retained cache layer or public API is added.

### 8. Scoped replacement proof and child canonicalization

This is the approved 2026-10-25 refresh4 extension. It is the eighth serial mechanism on the same F-X268 story. Implementation is in progress as recorded above. It does not add a ninth story, a stable public API, a dependency, or a module. Mechanisms 1 through 7 stay as approved.

Scoped replacement proof uses physical body children. Picture canonicalization and relationship renames use those children and the tail. The old whole paths stay the authoritative fallback.

Count, order, context, root, and span proofs are exact before a fast path is accepted. In-bounds disjoint edits do not panic on a `Vec` splice. Sentinel collision and physical selection use a double parse and a write normalization that matches the current whole algorithm. The fast path is not a weaker acceptance than that whole path.

A body block that contains a field, and a block that does not round-trip, is normalized and read anew. `PartialEq` alone is not reuse. Other proved blocks may still be reused. Preserve F-288 default, current, and provenance. Preserve F-289 physical form owners and handles. Preserve F-290 metadata and the bound store.

Release equality and release error order are required. A debug whole oracle is internal agreement only. There is no blanket hypothetical OOXML validator. Unknown foreign attributes, entities, head, tail, or an unaffected child must raise where the whole path raises, and they must raise before publication. A repeated call and a refusal leave bytes and live handles unchanged. Any unproved child, frame, or context chooses the whole fallback with the same errors, the same order, and the same bytes.

The private helpers below are captured from the PR422 own patch. They are not a public cross-crate contract change. No forwarding wrapper is copied only to destructure a tuple. `direct_story_content_items` calls `direct_story_content_items_in` because the caller already holds the scanned items.

```rust
fn canonical_retained_main_part(
    retained: &[u8],
    remap: &HashMap<String, String>,
    modeled_drawing_ids: &[u32],
    nested_image_relationship_ids: &HashSet<String>,
    canonical_drawing_ids: &[(u32, u32)],
) -> Result<Option<(Vec<u8>, CT_Document)>>;

fn relationship_remap_edits(
    xml: &[u8],
    remap: &HashMap<String, String>,
    inherited_bindings: &[(String, String)],
) -> Result<Vec<(usize, usize, Vec<u8>)>>;

fn xml_decodes(xml: &[u8]) -> bool;
fn scoped_proof_marker(original: &[u8], restored: &[u8]) -> Result<String>;
fn scoped_proof_sentinel(selected_closed: &[u8], marker: &str) -> Result<String>;
fn direct_story_content_items_in(
    xml: &[u8],
    owner: &StoryOwnerSpan,
    story_items: Vec<StoryItemSpan>,
) -> Result<Vec<StoryItemSpan>>;
```

On the existing `MainPartBlocks` type, the captured methods are `edited(&self, edits: &[(usize, usize, Vec<u8>)]) -> Option<Self>`, `bytes(&self) -> Vec<u8>`, `relationship_remap_edits(&self, remap: &HashMap<String, String>, bindings: &[(String, String)]) -> Option<Vec<(usize, usize, Vec<u8>)>>`, and `canonical_retained(...) -> Option<Option<(Vec<u8>, Self, CT_Document)>>` with the same four arguments as `canonical_retained_main_part` after `self`. `BodyBlockScan::decodes(&self, root: &[u8]) -> bool` and `MainStoryLayout::owner_kinds(&self) -> Vec<StoryKind>` stay private. `Document::scoped_main_source_by_blocks` returns `bool` and `Document::prove_whole_scoped_main_source` stays the whole `Result<()>` path. A `false` fast result means the children do not tell, and the whole proof decides.

`BlockStagingReads` gains `decodes: OnceLock<bool>`. The stored fact is whether `xml_decodes` reads the bytes of `alone(root)`. A false result means those bytes are not a fast proof. Replacing the block builds a new scan, so a changed fact does not keep the old `OnceLock`. A false error is invalidated when the bytes or the root context that produced it change.

Q22 applies to this new fact and to the current state. The measured points stay clone, reopen, and drop, with no accumulating history. A whole-model clone, hash, or copy stays a disclosed linear cost. There is no universal O(1) claim.

### Public Rust contract

Q21 adds four stable Rust methods. There are no optional parameters and no new units. There are no Python `try_` names, no WASM methods, and no CLI methods.

```rust
impl Document {
    pub fn try_add_paragraph(&mut self, text: &str) -> Result<Paragraph<'_>>;
    pub fn try_insert_paragraph(&mut self, index: usize, text: &str) -> Result<Paragraph<'_>>;
    pub fn try_add_table(&mut self, rows: usize, cols: usize) -> Result<Table<'_>>;
    pub fn try_insert_table(
        &mut self,
        index: usize,
        rows: usize,
        cols: usize,
    ) -> Result<Table<'_>>;
}
```

Preflight order on the four `try_` methods, all before mutation: the insert index range first, then paragraph XML characters, then the retained namespace guard. An index greater than `content_count` returns `Error::Other`. It does not panic. Text is checked with the existing `oxml_core::xml::reject_non_xml_characters`, and the native OXML error is retained. Existing Python primitive, type, and index checks still run before that guard. The native error hierarchy is unchanged. The new early namespace error is intentional only on the affected checked paths. Zero rows and zero columns stay accepted, as the existing table builders accept them. Other row and column semantics stay as they are. There is no new resource budget. When namespace replay or a shadowed serializer binding would make saving fail, the method returns the existing native `Error::Other` with the same underlying diagnostic, before any model, layout, or revision change. The bound check runs before the edit. There is no panic, no silent normalization, and no side effect on that failure.

The four existing infallible methods keep their signatures and their legacy behavior. Q21 exempts them from Q19. They are not described as checked. The four new methods are the checked Rust API. Python `add_paragraph`, `insert_paragraph`, `add_table`, `insert_table`, `Paragraph.insert_paragraph_before`, `add_heading`, and `add_page_break` use the checked path. Their public signatures, error family, primitive checks, and index validation stay. The early atomic refusal is the labelled behavior commit. It is not the performance commit.

### Hidden native helpers

The internal shared validation method is `#[doc(hidden)] pub fn preflight_checked_paragraph(&self, index: Option<usize>, text: &str) -> Result<()>`. Its actual consumers are the four checked Rust methods and Python paragraph character-error ordering. Python checks the native policy before its existing splice or fallback publisher. This adds no stable Python export and does not replace the source-preserving route.

Adopt a donor helper only where the Python binding or `Document` is an actual consumer. The declarations in the PR417 and PR418 block below are methods, not free functions. The later refresh2 block adds the free function `take_reopen_work`. None of these is a stable Python, WASM, or CLI export. `BodyItemContext` is the whole fact source for piece reuse, and a caller must hold that proof. No forwarding wrapper, trait, generic, module, crate, feature flag, or dependency is added. No new trait or generic is planned. The repository rule that needs two implementers or two instantiations applies to traits and generics, not to this method. `CT_Body::paragraph_count_before` is the concrete count. `Document::paragraph_count_before_content` is one call chain into that method, not a second independent consumer. A public cross-crate count helper needs the actual `Document` consumer. It does not need an arbitrary second consumer.

Optional eligibility `None` chooses the checked full path. A helper `Err` may take that full path only where the full path returns the same checked result, error, and error order the old path would have returned before publication. A full-path `Err` propagates through the existing native and Python mapping. Invalid ownership or invalid XML never becomes empty facts or success. `unsavable_body_edit` returning `None` is not a general proof that the document is savable.

Captured from PR417 `crates/rdocx/src/document.rs` as nine fields. That nine-field capture stays historical. The struct below is the final refined callable contract. It keeps those nine fields and adds `pub holds_field: bool` as the tenth field. The structs are public doc(hidden) cross-crate Rust types. The functions are methods:

```rust
#[doc(hidden)]
#[allow(clippy::type_complexity)]
pub struct BodyItemOwners {
    pub snapshots: Vec<StoryItemSnapshot>,
    pub cells: Vec<(ContentLocation, usize, usize, StoryId)>,
    pub paragraphs: Vec<(ContentLocation, Vec<u8>)>,
    pub links: Vec<(ContentLocation, LinkInfo)>,
    pub cell_stories: usize,
    pub text_boxes: usize,
    pub paragraph_routes: Vec<Option<Vec<usize>>>,
    pub tables: usize,
    pub cell_paragraphs: Option<(HashMap<(usize, usize, usize), Vec<ContentLocation>>, bool)>,
    pub holds_field: bool,
}

#[doc(hidden)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BodyItemContext {
    part_name: String,
    root: [Vec<(String, String)>; 5],
    plain_local_edit_blocked: bool,
}

impl BodyItemContext {
    pub fn part_name(&self) -> &str;
}

impl Document {
#[doc(hidden)]
pub fn body_item_owners(&self, item: &BodyContent) -> Result<BodyItemOwners>;
#[doc(hidden)]
pub fn body_item_context(&self) -> BodyItemContext;
#[doc(hidden)]
pub fn body_children(&self) -> &[BodyContent];
#[doc(hidden)]
pub fn story_part_sources(&self) -> Result<Vec<(StoryKind, String, Cow<'_, [u8]>)>>;
#[doc(hidden)]
pub fn story_part_owner_snapshots(
    &self,
    kind: StoryKind,
    part_name: &str,
    xml: &[u8],
) -> Result<(
    Vec<StoryItemSnapshot>,
    Vec<(ContentLocation, usize, usize, StoryId)>,
    Vec<(ContentLocation, Vec<u8>)>,
    Vec<(ContentLocation, LinkInfo)>,
)>;
#[doc(hidden)]
pub fn story_hyperlink_relationships(&self, part_name: &str) -> HashSet<&str>;
#[doc(hidden)]
pub fn unsavable_body_edit(&self) -> Option<&str>;
#[doc(hidden)]
pub fn paragraph_count_before_content(&self, slot: usize) -> usize;
#[doc(hidden)]
pub fn saves_as(&self, other: &Document) -> Result<bool>;
}
```

Captured from the new PR415 tree at `sources-old415-vs-new415/crates/rdocx-oxml/src/document.rs`. The impl starts at line 3397 and the method at line 3423:

```rust
impl CT_Body {
    #[doc(hidden)]
    pub fn paragraph_count_before(&self, index: usize) -> usize;
}
```

`cell_paragraphs: None` is failed cell routing. It selects the checked full path. It is not an empty success.

`saves_as` is the PR418 method now declared on `impl Document`. It may use `OpcPackage::reread` only while the current trusted-baseline precondition still guarantees the same bytes, the same error, and the same error order as `to_bytes`. Otherwise it uses the ZIP compare. The donor comment that a reread is automatically byte-identical is not adopted as proof.

The refresh2 amendment adds these hidden cross-crate contracts. They are not stable Python, WASM, or CLI exports. No new trait, generic, module, file, or dependency is added.

```rust
impl CT_Document {
    #[doc(hidden)]
    pub fn to_xml_with_body_spans(
        &self,
    ) -> Result<(Vec<u8>, std::ops::Range<usize>, Vec<std::ops::Range<usize>>)>;
}

impl Document {
    #[doc(hidden)]
    pub fn direct_body_item_location(&self, slot: usize) -> Result<Option<ContentLocation>>;

    #[cfg(debug_assertions)]
    #[doc(hidden)]
    pub fn main_body_block_scans() -> usize;
}

#[cfg(debug_assertions)]
#[doc(hidden)]
pub fn take_reopen_work() -> [usize; 3];
```

The approved refresh3 extension adds these hidden contracts. They are not stable Python, WASM, or CLI exports. `MainPartBlocks` stays private to `document.rs`. No new trait, generic, module, file, or dependency is added.

```rust
impl CT_Document {
    #[doc(hidden)]
    pub fn body_frame_xml(&self) -> Result<(Vec<u8>, Vec<u8>)>;
    #[doc(hidden)]
    pub fn body_item_xml(&self, item: &BodyContent) -> Result<Vec<u8>>;
}

#[cfg(debug_assertions)]
#[doc(hidden)]
pub fn take_staging_work() -> [usize; 6];
```

`holds_field` is the tenth field of the final `BodyItemOwners` declaration above. It is not a free-standing field.

`body_frame_xml` and `body_item_xml` are consumed by the actual native staging and reopen callers in `Document`. `holds_field` is read by the piece inventory and by block reuse. `take_staging_work` feeds Python `_staging_work`. The array order is fixed above.

`to_xml_with_body_spans` is consumed by the actual native `Document` story-source and direct-slot callers. `direct_body_item_location` is consumed by the actual Python document and table callers. `None` means absence only. Failed fast eligibility uses the checked full lookup. A genuine `Err` is preserved. `take_reopen_work` is a free function, exported from the native `lib.rs` for the Python `_reopen_work` counter. The array order, taken from the assessed counter body before any local test, is ZIP package writes, then main document part parses, then identifier-part bytes scanned. The implementation counts those actual sites. `main_body_block_scans` is the debug count of main-part body children scanned rather than reused, and it feeds Python `_main_body_block_scans`. Debug hooks are not stable exports and are not release exports. A release installed full test may guard a helper only for a counter case. Mandatory debug proof stays. There is no broad skip.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Checked insert | `try_add_paragraph`, `try_insert_paragraph`, `try_add_table`, `try_insert_table`. Python `add_paragraph`, `insert_paragraph`, `add_table`, `insert_table`, `insert_paragraph_before`, `add_heading`, `add_page_break` | Q21. Old four infallible methods stay legacy. Binding keeps its signatures | Evaluate the guard, then update, then serialize. Reopen the saved package |
| Namespace replay | `unsavable_body_edit` against current XML, model, and owners | Proof dies when the owner, local namespace, text, or typed field changes | Before mutation and again after a fact change |
| Piece inventory | `body_item_owners`, `body_item_context`, `body_children`, story part helpers | Binding publication is the consumer. Context equality is caller proof | Read, project, serialize, reopen |
| Staging | `saves_as` and the opened baseline shared with a staged clone | OPC `trusted_baseline` stays. Related parts stay checked | Save, error order, reopen, drop |
| In-memory reopen | `reopen_prepared_staged_with_limits` and `take_reopen_work` | UNBOUNDED opened-from-bytes trusted provenance only. Every other context keeps the ZIP path. Identifier replay keeps duplicate order and errors | ZIP path, in-memory path, limited reopen, both routes after mechanism 6 |
| Direct-body address | `to_xml_with_body_spans`, `direct_body_item_location`, `main_body_block_scans` | Exact `to_xml` bytes. `None` is absence. Fast-path failure uses the full lookup. `Err` stays `Err` | Full snapshot, both reopen routes, cache off, invalidated context |
| Main-part body blocks | `body_frame_xml`, `body_item_xml`, private `MainPartBlocks`, `BodyItemOwners.holds_field`, `take_staging_work` | Checked fallback matches full bytes and error order. No field-block reuse by `PartialEq`. Changed root invalidates. F-289 and F-290 publishers stay | Warm N and 4N counts, cold work separate, cache off, alternate root, save and reopen |
| Form and settings facts | Direct `cached_result`, `set_cached_runs`, annotation cache, `cached_field_mut`, refresh, attachment, and `set_legacy_form_field_value`. F-289 `validate_form`, `fill_form`, `reset_form`, `protect_form`, `set_document_protection`, `remove_document_protection`, and `add_story_permission_range`. F-290 `set_mail_merge_package`, `remove_mail_merge_package`, and `prepare_mail_merge_data` | Approved callables only. No second form model. Property callbacks and direct candidate assignment invalidate changed facts even when the generation is unchanged. Private field source form and equality that ignores serialization facts use proved invalidation or the full fallback. They must not publish a different truth. Dead owners stay dead | Held value, fresh value, native value, save, and reopen |

Imported metadata and an absent cache fact stay absent. A missing fact is not filled with a default by the cache.

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| Index above `content_count` | `Error::Other` before mutation. No panic | Current `insert_paragraph` panics. Q21 replaces that only on the `try_` methods | `checked_body_insertions_refuse_unsavable_namespace_owners_atomically` |
| Non-XML text | Existing `oxml_core::xml::reject_non_xml_characters`. Native OXML error. No mutation | Existing character check | Same native test |
| Zero rows or columns | Accepted, matching current table builders | `Document::add_table` | Same native test |
| Namespace or shadowed prefix | Existing `Error::Other` and the same diagnostic, before model, layout, or revision change | Q19 refined by Q21 | Native test and `test_checked_body_insertions_preserve_owners_on_namespace_refusal` |
| Old four methods | Signatures and legacy behavior unchanged, including a later save failure | Q21 exemption | Explicit legacy case in the native test |
| Package bytes | Unchanged except the requested edit. No redundant saved xmlns delta | Integrated F-X266 fragment rule | `word_owner_piece_sources_match_the_full_physical_inventory` |
| `saves_as` | Exact bytes and exact error order, or ZIP fallback | `OpcPackage::reread` trusted baseline | `unchanged_main_staging_preserves_save_bytes_and_error_order` |
| Work counters | Real walk, parse, and serialization counts. Not a constant entry counter | F-X266 linear owner bound retained | `unchanged_main_staging_reuse_counts_actual_work` |
| Body-child staging | Same bytes and same error order as the full path. `holds_field` blocks value reuse. Alternate root reads the part whole | PR421 candidate only. Not accepted code | `staging_reads_only_the_main_part_blocks_an_edit_changes`, `test_staging_writes_and_reads_only_the_main_part_blocks_an_edit_changes`, `test_a_fragment_imports_into_a_main_part_whose_root_has_another_word_prefix` |

## Dependency readiness

F-X266 source `9c6ef377971182a73b2e3cef07ccedde579d2f37` is integrated at `9ee48bf0902d976de016553b1dd0c79fed856612`. F-X267 source `70573c91030be7073d29e3451383060eb207207a` is integrated at `d0f17a6ac7ef675d0ed2a86cea641ce21d65bfdd`. Both are reviewed and in progress. Neither is done. Approval of those plans is not final acceptance.

The local prefix is approved contract. Canonical `0d8dfb68396e33ddfcaa0d9400c30c006e12d2b7` added `## F-X268 local provider checkpoint` to `.claude/plans/F-X266-design.md` and `.claude/plans/F-X267-design.md`. Codex root approved both checkpoints on 2026-10-24 for use by F-X268. Actual prefix completion has not been performed. Neither story is done. The checkpoints do not weaken the final performance, hash, or Word riders.

The prefix names every local scoped row already in those approved sections. Outcomes and commands come from the existing progress records. The approved contract does not invent a combined pass. No gate pass is claimed here.

F-X266 test plan, all 9 local scoped rows: `direct_routes_match_full_body_bytes`, `append_keeps_eligible_owners`, `conservative_retires_all_scopes`, `failed_splice_publishes_nothing`, `fragment_omits_unused_w14`, `owner_scope_precedence_unchanged`, `held_fresh_reopen_match_old_semantics`, `body_work_counters_follow_owners`, `missing_run_path_refuses`.

F-X267 test plan, all 13 local scoped rows: `remove_part_returns_shared_bytes`, `reread_matches_write_and_read`, `reread_matches_shared_zip_name_limits`, `hand_built_baseline_uses_fallback`, `untouched_slide_bytes_stay`, `fit_text_ignores_unrelated_broken_slide`, `fit_text_at_selects_size_valid_in_distinct_shared_contexts`, `footer_proof_is_part_local`, `parsed_baseline_tracks_real_edits`, `fx211b_native_opaque_root_edit_retains_invalidated_signature`, `embedded_replace_counts_commit_zip_write`, `unnormalized_opaque_part_counts_reread_zip_fallback`, `commit_counters_follow_real_work`.

The prefix also consumes the actual scoped checks, clean Root reviews, and the current applicable artifact proof. An old artifact is not reused as a fresh qualifier. Rerun only the inputs a later change invalidates, including a checked build and import where binding evidence is affected. No full oracle rider runs at the prefix. Word, the full suite, the hash harness, the final release 79-call and 252-call runs, the F-X267 248-call run, and the memory rider stay pending. Before any F-X268 claim, verify that the F-X266 and F-X267 source heads above are actual ancestors, that this prefix is actually done, and that binding source proof exists where the binding is affected. `AS_BUILT.md` and `SPRINT_TRACKER.md` are written once, after the full union. A formal consumer claim waits for the actual done checkpoint.

The next paragraph is the state before the 2026-10-25 local prefix. It is dated historical for that earlier readiness.

The current F-289 reviewed local integration is `cbcb0c9f07f392d541362f08bd6e6af6311f3417`. Root whole-feature pass 2 is clean, with 0 defects, 0 smells, and 0 nitpicks. The fresh local checks are native 91 and Python 13 on artifact `130f13ce533f8a78858b09a7113febc33198ebfdf778682e1196d85683906105`. Those checks are local only. Final Word 16.114 A7 and the full gate stay pending. F-289 is not done. That checkpoint is not delivery credit and it is not a formal dependency change. F-X268 stays unclaimed. F-X268 does not start while the shared leases are held by F-289 and then F-290.

## Local prefix readiness 2026-10-25

Root approved genuine local provider completion for F-X266 and F-X267 on canonical `79047bf193bcd9dfb229a0b37a93d979dc0997a7`. This record is that Root review. It is not a self-review. Both providers are done for that local prefix. Neither is final-accepted. Plan status for F-X266, F-X267, and F-X268 stays approved. The sentences above that say the prefix has not been performed stay the earlier contract text.

Source head `9c6ef377971182a73b2e3cef07ccedde579d2f37` and integration `9ee48bf0902d976de016553b1dd0c79fed856612` are ancestors of that HEAD. Source head `70573c91030be7073d29e3451383060eb207207a` and integration `d0f17a6ac7ef675d0ed2a86cea641ce21d65bfdd` are ancestors of that HEAD. Current artifacts are `/private/tmp/s95.4-fx266-verified-_rdocx.abi3.so` sha256 `ae3879224176459797da677f26058757a5aa0592817921004c985ded1d8d17d6` and `/private/tmp/s95.4-fx267-local-debug-_rpptx.abi3.so` sha256 `9fab4e917aef5ee74457ab95e85ffddaa262698fc7d9cbeaa1e88cc9cd0e56ca`. The `9fab4e91` file is a fresh development build and import proof only. It is not the native 13 and it is not a release build. The old `1fe91d2d` archive is not the current qualifier. The seven fresh F-X267 runs fill missing dedicated GREEN logs. They are not a claim that a later change invalidated those rows. The F-X266 linear-proof paragraph that still says the counter proof stays pending is the earlier contract state. The current local counter GREEN is `body_work_counters_follow_owners` in the 9 pytest rows. The signature row passed 1 with `--features digital-signatures`. The earlier 0-test filter is not a pass. The preferred retained PowerPoint proof for `parsed_baseline_tracks_real_edits` and the other names in the post-counter block is `.claude/scratch/F-X267-progress.md`.

F-290 is integrated at `79047bf193bcd9dfb229a0b37a93d979dc0997a7`. Root whole-feature pass 3 is clean at `.claude/reviews/F-290-whole-feature-pass-3.md`. Public local checks are consumed. F-290 stays in progress and is not done. F-289 stays in progress and is not done. Shared leases are released. This readiness record does not claim F-X268. Final debts stay pending: the 79 flat gates, the 252-call catalogue, the 5 linear calls, the original 407 gate, the PowerPoint 248-call run, the original 10, 40, and 80 timings, the M21 digest, memory, Word, the hash harness, the full suite, strict mypy, stubtest, WASM, and packaging. `AS_BUILT.md` and `SPRINT_TRACKER.md` stay unchanged.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Q19 and Q21. Four new Rust methods and every named Python delegate. An old document whose save already fails. A valid document. Bad index, non-XML text, zero columns | Refusal is atomic. Model, bytes, and handles stay at the old unsavable state. The old four methods stay legacy. Valid bytes match the checked source. Errors keep the existing family | Native `Document` and Python document, paragraph, heading, and page break | `checked_body_insertions_refuse_unsavable_namespace_owners_atomically`, `test_checked_body_insertions_preserve_owners_on_namespace_refusal` | Local behavior commit. Not the performance commit |
| A2 | New PR415 unrouted facade position and nested namespace proof. Then an owner, local namespace, text, or typed field edit | First proof matches the checked source. The stale proof is not reused. Accepted aliases and unknown bytes stay. No incidental namespace byte change | Facade and fragment writer | `word_owner_piece_sources_match_the_full_physical_inventory` | Local regression |
| A3 | Piece inventory against the full inventory for body, nested cell, text box, comment, note, shared header, duplicate part, invalid relationship, identical blocks, and a direct returned owner. The publishers named in the consumer map, including direct candidate assignment and property callbacks. Direct-body span bytes and block reuse | Equal owners, parents, routes, projections, read facts, order, births, generations, and retired set. Span bytes match existing `to_xml`. Unsupported shape falls back before publish. Held, fresh, native, save, and reopen agree. Ignored serialization facts do not count as proof | Binding and native readers | `word_owner_piece_sources_match_the_full_physical_inventory`, `word_owner_pieces_follow_splices_and_fact_changing_publishers`, `body_spans_match_existing_to_xml_on_current_fixtures`, `main_body_block_reuse_preserves_physical_owners_and_publication`, `test_body_block_reuse_preserves_forms_settings_and_live_owners`, `fx268_changed_existing_context_uses_full_inventory_before_publication`, `fx268_case_equivalent_story_parts_use_the_full_inventory`, `test_word_piece_shared_header_identity_uses_opc_part_names` | Local regression on the integrated ancestor that contains those publishers |
| A4 | Splices at head, middle, and tail, plus a missing item. A later edit. No-op and refusal | Route and later edit stay inside the work bound. Conservative births and retirement match the old path. Dead owners stay dead. No-op and refusal preserve bytes, revisions, and handles | Python publication | `word_owner_pieces_follow_splices_and_fact_changing_publishers`, `test_word_owner_pieces_follow_edits_splices_and_related_owners` | Local regression |
| A5 | Opened cache and `saves_as` for new, opened, clone, source-edited, relationship, media, signature, and embedded packages. Main unchanged with a changed related part. Eligible in-memory reopen and a document with no trusted opened-byte baseline | Exact bytes and native error order, or the ZIP fallback. Related changes are checked. Identifier order matches. A warm edit does not reparse an unchanged main part. A document qualifies for in-memory reread only after a trusted opened-byte baseline exists | OPC and binding | `unchanged_main_staging_preserves_save_bytes_and_error_order`, `unchanged_main_staging_reuse_counts_actual_work`, `test_unchanged_main_staging_keeps_forms_and_settings_authoritative`, `in_memory_reopen_matches_zip_for_new_flat_limited_and_hook_paths`, `in_memory_reopen_keeps_related_part_errors_and_drops_stale_main_checks`, `identifier_scan_replays_exact_bytes_categories_and_old_error_order`, `parsed_none_reopen_grows_zip_parse_and_identifier_bytes`, `test_header_story_text_reopen_counts_the_changed_header_part`. Native read-limit coverage is in the first of those four native tests. Python evidence uses the existing `Document.set_story_text` API only. Forms and settings evidence stays the planned `test_unchanged_main_staging_keeps_forms_and_settings_authoritative` | Local regression |
| A6 | The 79 targeted calls and the full 252-call catalogue on a fresh release install | The 79 meet the Q20 flat bound. The 252 results and residuals are reported without xfail hiding. The original 407 gate still holds for its three calls. The four extra PR419 test-map keys are not silently added | Installed Python 3.12 | Named in the performance section. Not a worker run | Final union |
| A7 | Q18 and Q22 memory workload | Actual per-layer and combined RSS samples at the named points. No summed RSS fabrication. No 68 MiB pass line and no hard cap | Fresh release process | Named in the memory section | Final union |
| A8 | Installed Python, strict mypy, stubtest, rustdoc, publish archive, WASM, hash, and the prior seven stories | Prior Word, SSIM, risk, and hash gates stay mandatory and stay unrun here. Q19 compares the previous save failure, model, and handles. Two failing `to_bytes` calls are not atomicity proof. Final own-union fidelity is mandatory. The PR419 donor log is not that pass | Final union | None in the worker | Final union |
| A9 | Q23 linear calls at 250 and 1000 paragraphs | Each call meets `large <= 5 * small + 5 ms`. Each relevant pass count meets `large <= 5 * small`, with additive allowance 0. Setup counts stay outside that pass. Whole-main hash and serialization are reported separately | Installed Python 3.12 | `body_lookup_work_is_linear_at_250_and_1000` for counts. Timings only in the final union | Final union |
| A10 | Fourteen public calls at 20 and 80 paragraphs, and native table row, section, list, paragraph scoped replace, cell scoped replace, and story picture edits at 50 and 200 paragraphs. Setup outside the counter window. Warm samples are two distinct edits after two changed edits. Cold work is a separate measurement | All six staging slots are compared as one `[usize; 6]` counter. The retired four-slot counter is not this contract. The first four slots are byte counts and are zero, except `Document.import_fragment`, whose bounded source-document byte work stays and is equal at both sizes. The last two slots are child counts, serialized children then children read. Those child counts are equal across the two sizes. They are not byte counts and they are not required to be zero. Reopen ZIP writes are 0. Main-part parses are 0, except the fragment call, where that slot is 1. Identifier-byte absolute delta is at most 16. Native warm staging has the first four byte slots at zero and the first two reopen slots at zero, with the same warm counts at both sizes. A cache-off mutant for the scope proof, the canonicalization, or the rename does actual work that scales. A missing helper is a compile failure, not that RED. These are counted-work obligations. They are not new timing gates | Python document and native `Document` | `test_staging_writes_and_reads_only_the_main_part_blocks_an_edit_changes`, `staging_reads_only_the_main_part_blocks_an_edit_changes`, `test_a_fragment_imports_into_a_main_part_whose_root_has_another_word_prefix` | Local counted-work regression. Approved with mechanism 8. Independent of the 79 timing gates |
| A11 | Release fallback and error-order proof. Aliases, the default namespace, a shadowed namespace, `w14`, opaque attributes, entities, tail, a field, a tab, bookmarks, and a duplicate physical position. Valid positive controls stay | The fast path matches the current whole algorithm on bytes, errors, and error order, or it takes that whole path. Unknown foreign attributes, entities, head, tail, and an unaffected child raise where the whole path raises, before publication. Repeated calls and refusals leave bytes and live handles unchanged. An unproved child, frame, or context uses the whole fallback | Native `Document` and Python document | Named with the mechanism 8 group. Not a timing promise | Local regression. Approved with mechanism 8. Not a timing promise |

The original 65 targeted calls are 62 catalogue keys plus `Run.xml`, `Table.xml`, and `Cell.xml`. The pinned 252 does not cover those three reads. The refresh2 amendment adds 8 calls. That historical Q23 flat set is 73 unique calls, which is 70 catalogue keys plus the same 3 XML reads. Q24, answered 2026-10-25, adds 6 calls. The current flat set is 79 unique calls, which is 76 catalogue keys plus the same 3 XML reads.

Targeted calls from PR415, 13:

`Document.insert_paragraph`, `Document.remove_content`, `Paragraph.insert_paragraph_before`, `Document.remove_content_range`, `Document.add_paragraph`, `Document.insert_table`, `Document.add_table`, `Document.add_heading`, `Document.add_page_break`, `Paragraph.xml`, `Run.xml`, `Table.xml`, `Cell.xml`.

Targeted table, row, and cell setters from PR417, 26:

`Table.indent`, `Table.alignment`, `Table.width`, `Table.set_borders`, `Table.set_border`, `Table.set_cell_margins`, `Table.grid_widths`, `Table.set_column_width`, `Table.set_cell_grid_span`, `Table.set_cell_vertical_merge`, `Table.autofit`, `Table.first_row`, `Table.last_row`, `Table.first_column`, `Table.last_column`, `Table.banded_rows`, `Table.banded_columns`, `Row.height`, `Row.height_rule`, `Row.cant_split`, `Row.is_header`, `Cell.width`, `Cell.vertical_alignment`, `Cell.shading`, `Cell.set_border`, `Cell.set_margins`.

Other PR417 calls, 16:

`Cell.add_paragraph`, `Paragraph.text`, `Paragraph.add_run`, `Paragraph.add_hyperlink`, `Paragraph.add_equation`, `Run.remove`, `Run.add_field`, `Document.split_run`, `Document.add_comment`, `Document.reply_to`, `Document.resolve_comment`, `Document.add_chart`, `Document.set_text_watermark`, `Document.set_image_watermark`, `Document.set_note_policy`, `Settings.odd_and_even_pages_header_footer`.

PR418 calls, 10:

`Document.set_font`, `Document.embed_font`, `Document.remove_embedded_font`, `Document.remove_font`, `Document.embed_true_type_fonts`, `Document.save_subset_fonts`, `Document.embed_system_fonts`, `Document.add_numbering_definition`, `Document.add_numbering_instance`, `Document.link_style_to_numbering`.

That original list is 65 unique calls. `Run.xml`, `Table.xml`, and `Cell.xml` are the three additional required read timings. They are not catalogue keys.

The 8 added flat calls, none of them in the original 65, are:

`Document.replace_image`, `Document.replace_image_for_story`, `HeaderFooter.add_paragraph`, `HeaderFooter.add_table`, `SectionHandle.set_text_watermark`, `SectionHandle.set_image_watermark`, `SectionHandle.remove_watermark`, `Cell.add_table`.

The historical 73 and the current 79 use the Q20 bound. The 100-paragraph time is at most 1.5 times the 25-paragraph time plus 5 ms, under the existing sample rules. The four extra keys in the PR419 test map are `FootnoteEndnote.add_paragraph`, `FootnoteEndnote.add_table`, `Document.add_custom_xml_store`, and `Document.set_chart_legend`. They are not added to the 79 and they are not silently gated. Under the historical Q23 set, `Run.add_picture`, `Cell.add_picture`, and `Cell.merge` were 252 residuals. Q24 moves `Cell.add_picture` into the 79. The current residuals of that trio are `Run.add_picture` and `Cell.merge`.

The six Q24 additions, none of them in the historical 73, are `Paragraph.replace_text`, `Cell.replace_text`, `Document.replace_text_at`, `Document.add_picture`, `Cell.add_picture`, and `Document.add_hyperlink_to_story`. All 5 A9 linear calls stay at 250 and 1000 paragraphs, including `Cell.replace_text` and `Document.add_picture`, which are also in this flat set. The 252-call catalogue, the original 407 gate, and the F-X267 248-call run stay unchanged.

### A10 counted work

The nine calls named in the PR421 own diff stay. The approved refresh4 extension adds five calls on producer-retained XML: `Paragraph.replace_text`, `Cell.replace_text`, `Document.replace_text_at`, `Document.add_picture`, and `Cell.add_picture`. The counted-work set is those 14 calls. A10 is independent of the Q24 timing gate. `Document.add_hyperlink_to_story` is one of the 79 flat calls and is not one of the 14. Python sizes stay 20 and 80 paragraphs. The producer keeps the XML. Setup stays outside the window. Warm samples are two distinct edits after two changed edits. Native sizes stay 50 and 200 paragraphs, and the existing native group adds paragraph scoped replace, cell scoped replace, and story picture. The six counter units stay the same. Bounds stay the same, including the fragment byte exception and the reopen identifier bound. A real cache-off mutant for the scope proof, the canonicalization, or the rename must expose growing actual work. A missing helper is a compile failure, not that RED. A10 does not add a timing promise of its own. Donor timings are not accepted.

### A11 release fallback

A11 is approved with mechanism 8. It proves release fallback and error order for aliases, the default namespace, a shadowed namespace, `w14`, opaque attributes, entities, tail, a field, a tab, bookmarks, and a duplicate physical position. Valid positive controls stay. The fast path matches the current whole algorithm, or the whole path runs with the same errors and the same order. Debug agreement is not this release proof.

### A9 linear calls

The native count gate opens separate serialized fixture bytes for every family and size, without sharing parsed or scan caches across families. Its first edit is cold after that family's required target acquisition. After freezing each cold, warmup and warm count vector, independent semantic checks verify the exact edited owner, value, control identity and destination, picture dimensions or insertion, or comment endpoints. Cell checks cover every other physical cell. Native test-profile count execution with debug assertions disabled retains test instrumentation and production guards while omitting debug-only agreement oracles. Debug source-proof controls remain separately required. This count execution supplies no installed Python timing evidence.

The five Q23 calls are `Cell.replace_text`, `Document.set_picture_size`, `Document.add_picture`, `Document.move_content_control`, and `Document.move_comment`. Use the same pinned builder, adapted to 250 and 1000 paragraphs, with the same six formatted runs and the same target geometry between sizes. Fresh target, call setup, and open stay outside the clock. The full public lookup, validation, and publication are inside the clock. Alternate 3 fresh changed targets at the small size and the large size. Retain every raw sample. The minimum follows `costs_ms`. The cold first sample stays visible. Warm statistics stay separate. There is no compilation concurrency. For each call, the large time is at most 5 times the small time plus 5 ms. Count the actual visited nodes, blocks, identifier operations, and lookup operations at 250 and at 1000. Pin that counted work before measurement. For each relevant pass, the large count is at most 5 times the small count, and the additive allowance is 0. Count setup separately, outside the measured pass. Retain the raw counts for each phase. A constant entry counter is not proof. Report whole-main hash and serialization separately. If the first cold scan is itself linear, disclose that. A semantic mutant or a cache-off build must expose the count regression. Absence of a new helper is a compile failure, not that RED.

### Proof-only writer comparison

Repeated equality checks may compare an existing qualified writer record's head, child and tail bytes directly. They validate current resource limits, exact frame bytes, every current private-fact admission, exact staged-child equality and field-token exclusion. This optional path requires an explicit `w14` root declaration, which makes the existing writer's conditional declaration step inert. Frames without it, fields, bookmarks, tab stops, changed metadata and failed correspondence retain actual serialization. Read provenance remains separate from writer eligibility.

An opening retains its independently produced current writer record in the existing `MainBodyScans.serialized` slot. That positive current record takes precedence over an older carried record. There is no new cache field or public API. This initializes the existing writer cache earlier, so Q22 includes its footprint at open as well as its bounded current state after edits. It may retain normalized bytes and staged child values earlier than before, which must be measured rather than called free.

`unchanged_writer_comparisons_keep_the_record_without_assembling_output` counts actual assembled output bytes and checks that equality does not replace the writer record. Its real cache-off control must violate the same zero-assembly condition. `writer_piece_comparisons_match_whole_bytes_and_decline_changed_frames` covers empty, one and many children, unequal expected bytes and changed root, tail and final child. `writer_comparisons_keep_private_and_conditional_declaration_fallbacks` requires actual writers for fields, bookmarks, tab stops and conditional `w14` frames. Publication continues to produce owned bytes. These local mechanism checks do not replace the release timing, memory or native gates.

### Owned output from an existing normalization

`serialized_main_part` may copy the opening's existing normalized serialization proof only after resource admission, the ordinary block-proof switch and a complete current `writer_bytes_match` result of `Some(true)`. Every other result keeps the established serializer. The returned record remains writer-only with `part = None`. No raw-source, parsed-model or namespace authority is inferred from the normalized output. This reuses an existing retained proof and adds no cache field or retained output copy.

The owned output still copies all bytes. `writer.copied_normalized_output` counts that actual copy separately from assembled output and the six staging slots. Those bytes remain a disclosed linear cost. `unchanged_main_output_copies_normalization_without_assembly` distinguishes producer source from normalized output, compares complete writer bytes and requires the real force-off route to assemble. `normalized_main_output_keeps_changed_and_private_fallbacks` covers changed children with a stale opening proof, fields, bookmarks, missing explicit root declarations and new documents. These local observations do not imply a flat timing pass or replace memory qualification.

Root approved an eager private `BodyBlockScan` field-token fact at the d4736c7619d3812548ddb3f66f3a7ba899e7c581 checkpoint. The fact is exactly the case-sensitive `fldChar` or `fldSimple` substring test over the block's owned immutable bytes, including lexical false positives. Writer, source and child-scan constructors compute it. Empty frame bytes use false. Clones and frame-only copies preserve it only with identical bytes. Every byte replacement reconstructs the block. The nine block-specific consumers reuse this fact, while the independent whole-item owner scan keeps its existing scanner. Current mutable-model and private-owner checks, staged equality, namespace and source proofs, resource admission, error order and forced fallback remain mandatory. No public API, additional retained byte buffer or history is added. Actual struct size before and after is recorded locally, with this inline fact included in Q22's parent body-scan layer rather than treated as free memory.

`field_word_facts_avoid_warm_byte_scans` counts actual visited comparison windows for each token, including early exits, and requires zero warm rescanning after construction. An actual uncached traversal of the same block bytes must exceed that bound. Constructor, frame-copy, token-change, repeated-edit, unchanged-neighbor and private-owner controls preserve the byte fact's lifetime and fallback behavior. These local mechanism checks do not replace the unchanged 79-call timing gate, full catalogue, memory qualification or native Word checks.

### Immutable inventory and dense control index

Picture inventory validates each discovered embedded relationship against the owner already read from the same immutable source. A shared private relationship validator retains relationship lookup, type, internal mode, target resolution and missing-part error order. The public relationship API still checks the current story and fingerprint first. Linked-image validation stays unchanged. No retained cache is added.

The content-control XML index assigns IDs at each append position, including empty elements. Closing an element uses a checked direct ID lookup instead of scanning the growing index. Nested ranges, namespace identity, parent IDs and malformed-input refusal remain the same. Actual close candidates are counted per caller, including failed lookups. Whole writer assembly calls and output bytes are separately visible in the native work report under the existing five-times rule. These measurements do not cover every allocation or replace release timing.

Scoped replacement reuses its already checked canonical source when selecting the retained physical owner, within the same immutable candidate phase. The retained owner inventory, canonical versus retained owner comparison, scoped source proof and all publication checks remain required. Other fragment callers still perform their own current canonical check.

`picture_inventory_assembles_one_source_for_all_relationships` requires one actual source assembly for four pictures and checks the public relationship results separately. `control_index_closing_work_is_linear` exercises actual changed control movement at 25 and 100 paragraphs, with the same five-times and zero-allowance rule. `control_index_preserves_nested_and_empty_ranges` checks dense IDs, physical ranges, parent and namespace identity, and malformed endings. `retained_fragment_view_reuses_checked_canonical_source` compares every discovered story against the independently checked fragment reader while observing that the retained-view helper does not construct another canonical source.

### Post-pin public surface audit

The immutable 252-call catalogue stays. Its timed coverage checker predates F-288, F-289, and F-290. The final union public stub inventory enumerates every added or changed method and maps each new read-only, factory, global, and local method to its F-ID, its A-ID, and its timing disposition. Do not relabel a local method `WHOLE_DOCUMENT` only so the old checker passes. A transparent qualification driver may adapt only the post-pin coverage classification, with the new methods and the reasons named. It preserves all 252 old keys and all 79 required timings. It does not skip or drop broadly. A new local method that is not in the 79 stays outside the Q20 timing promise unless the original authoring plan already owns a performance criterion. Correctness, runtime, and stub tests for those methods stay mandatory. Report the expanded inventory count separately. Do not claim that 252 covers it.

The full catalogue is 252, which is 250 timed calls plus 2 reads: 216 literal timed calls, 15 CoreProperties calls, 13 Section calls, 6 App calls, and 2 reads. Of those, 145 were the old growing set and 107 were the old non-growing set. Run all 252 without hiding status behind xfail. Report each slow residual. Do not copy a donor pass count. Existing non-growing calls that are not in the 79 keep the original bound of 2.5 times N plus 5 ms. Each of the 79 uses the Q20 bound: the 100-paragraph time is at most 1.5 times the 25-paragraph time plus 5 ms.

The final consumer is a fresh release install of Python 3.12 from the actual union source head, with the artifact SHA, interpreter, and config recorded. The external builders use 25 and 100 paragraphs, six formatted runs, a picture or table every 10th paragraph, two sections, a header, comments, and a control, for the qualified targets in the pin `4646db50ed783ddf4a2220f4b95a15ad7a225944`. Sampling alternates N and 4N, up to 3 fresh changed targets, and keeps the minimum the way `costs_ms` does. Two samples may stop early only when both fastest samples are above 50 ms. Setup, open, and build stay outside the clock. Every raw sample is retained. There is no concurrent compilation. The first sample is cold and stays visible. Warm statistics are separate. A cold walk and a fallback cost are disclosed. A warm minimum is not a constant-time claim. The original Issue 407 gate stays: mean of 3, 100 versus 1000 paragraphs, six verbose run-property runs, `large <= 3 * max(small, 1 ms)`, plus a separate fresh median of 7, for both cold and warm, on the original F-X266 methods. A native benchmark does not replace the Python public lookup, validation, and publication.

Local counters sit at the real walk, parse, and serialization sites. Separate the owner-projection phase from the total public call. On a qualified case, 20 versus 80 or 50 versus 200 unrelated body blocks of equal piece count, the piece owner projection reads at most 2 pieces, and at most 1 piece after a qualified top splice or tail removal. That phase does not serialize an unrelated body block. Underlying staging may still serialize the current typed main part as the exact byte proof. Count those whole-main serialization bytes and that time separately. Do not hide them inside a zero reparse or zero canonicalization count. After the first unchanged main check, an eligible warm edit does 0 main reparses, 0 main canonicalizations, and 0 ZIP compare writes. That is not zero serialization for the whole call and it is not an O(1) whole call. The total public Python time covers lookup, validation, publication, and serialization. The existing F-X266 qualified direct-read and splice bound of no unrelated serialization stays unchanged. Misses and visited metadata are counted, so an O(N) metadata walk is disclosed and is not called O(1). Registry owner work keeps the existing F-X266 linear bound and is not quadratic. The cold proof body walk is counted separately. Native counters cover the current main parse, owner scan, canonicalization, and ZIP writes. Related-part checks are still counted. A constant entry counter is not proof. A build with the cache forced off, or a mutant that skips invalidation, must fail these asserts. Absence of a new helper is a compile failure, not that semantic RED. The debug full-path oracle is internal agreement, not external parity. Debug counters and release guards do not skip the production check.

Memory, final Q18 and Q22 rider only. The workload is 20000 paragraphs, six runs each, with fixed text and run formatting. Measure and report the actual main XML length. The donor 6 MiB claim is unverified. Use 3 paired alternating fresh processes. Record the median current RSS and report each peak, with the raw samples and the aggregation named, after open, after the first font edit, after 100 distinct warm edits, after clone and drop, after reopen and drop, and after the document drop. Pair a pre-cache commit baseline with the final release, sharing the inventory and otherwise equal. Record the baseline source, build, artifact, allocator, and process boundary. The shared cache may retain the opened baseline, one parsed model, and one serialized copy, and the current changed parts, plus bounded projection facts for each live piece. It does not retain a history per edit. N edits and 4N edits must not retain one baseline per edit. Clone ownership and reopen release stay safe. There is no 68 MiB pass claim and no hard budget the user did not choose. The donor 68 MiB figure is not accepted. Measure authored and new-document fallback, and the unsavable namespace, duplicate part, signed, and embedded fallback times, separately. Correctness is mandatory on those paths.

Q22 measures each cache layer on otherwise equal release builds, and it measures the combined footprint the same way. The layers are the PR418 parsed model and serialized copy, the PR419 reopen and identifier caches, the PR420 body bytes and scans, and the refresh3 `MainPartBlocks`, `BodyBlockScan`, frame, parsed, namespace, relationship, drawing, numbering, and identifier caches. Each of those added layers is measured independently and in the combined footprint. The points stay current state, N versus 4N, clone, reopen, and drop. There is no per-edit history. Use true equivalent pre-layer and final variants. Retain the original fixture, statistics, raw RSS, peaks, clone, reopen, drop, and the N versus 4N history test. Do not add RSS deltas together and call that sum the combined footprint. Q22 already covers this caching. The approved refresh4 extension adds the `decodes` fact to the same policy. The points stay current state, clone, reopen, and drop. There is no accumulating history. Whole-model clone, hash, and copy stay disclosed linear costs. There is no universal O(1) claim. Q22 is not reopened. Q24 is a timing-gate choice, not a memory-policy choice, and it is answered.

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | Zero, one, and many body blocks. Direct and nested. Alias and shadowed namespace | Guard and inventory match the full path | A stale proof after a text edit fails |
| C2 | Current versus default form value. Implicit and explicit protection | Facts change even when the generation does not | Semantic equality used as the only proof fails |
| C3 | New document and reopened package. Partial write and whole write | Bytes and error order agree, or the ZIP fallback runs | A main-unchanged skip that ignores a related part fails |
| C4 | No-op, refusal, and conservative splice | Handles and retired scopes match the old publication | A refusal that mutates, or a dead owner that returns, fails |
| C5 | Opened UNBOUNDED package, new document, limited reopen, and both reopen routes | In-memory reuse matches ZIP bytes, errors, and identifier order only on the eligible path | A limited reopen or a changed relationship context that skips the ZIP path fails |
| C6 | Direct slot zero, one, and many, nested controls, tables, cells, text boxes, and moved or repeated blocks | Span bytes match `to_xml`. Slot lookup matches the full snapshot before and after both reopen routes | An ordinal alias, an empty success, or a donor overwrite of form providers fails |
| C7 | Field-bearing and field-free blocks. Default namespace and an aliased root. Duplicate occurrence order. Cache on and cache off | Kept bytes match the full writer. A field block is not reused by `PartialEq`. Counts match A10. F-289 and F-290 owners stay | A false-negative field reuse, a changed-root hit, or a fallback that drops an error fails |
| C8 | Aliases, default and shadowed namespaces, `w14`, opaque attributes, entities, tail, field, tab, bookmarks, and a duplicate physical position. Valid positive controls. Cache on and cache off | A11 bytes, errors, and error order match the current whole path. Refusal and a repeated call leave bytes and live handles unchanged | A fast acceptance weaker than the whole path, a dropped error, or a changed refusal byte fails |

Omitted permutations are the full cross product of those axes. Each axis above, including C7 and C8, has one discriminating case. Same-consumer round trip does not replace the native, Python, stored, layout, and reopen comparison named in A3 and A5.

Native groups go in the existing `crates/rdocx/tests/integration_test.rs`:

- `checked_body_insertions_refuse_unsavable_namespace_owners_atomically`
- `word_owner_piece_sources_match_the_full_physical_inventory`
- `word_owner_pieces_follow_splices_and_fact_changing_publishers`
- `unchanged_main_staging_preserves_save_bytes_and_error_order`
- `unchanged_main_staging_reuse_counts_actual_work`
- `in_memory_reopen_matches_zip_for_new_flat_limited_and_hook_paths`
- `in_memory_reopen_keeps_related_part_errors_and_drops_stale_main_checks`
- `identifier_scan_replays_exact_bytes_categories_and_old_error_order`
- `parsed_none_reopen_grows_zip_parse_and_identifier_bytes`
- `body_spans_match_existing_to_xml_on_current_fixtures`
- `main_body_layout_raw_blocks_match_the_old_scanner`
- `main_body_layout_whole_walk_and_fallback_counts`
- `main_body_layout_mixed_owners_follow_old_scans`
- `main_body_layout_serialized_facts_and_frame_use_scanner_bytes`
- `main_body_layout_reopen_reuses_unchanged_children`
- `main_body_layout_multiple_body_owners_use_full_reader`
- `main_body_layout_coverage_steps_scale_with_children`
- `main_body_block_reuse_preserves_physical_owners_and_publication`
- `body_lookup_work_is_linear_at_250_and_1000`
- `staging_reads_only_the_main_part_blocks_an_edit_changes`

Python groups go in the existing `crates/rdocx-py/tests/test_core.py`:

- `test_checked_body_insertions_preserve_owners_on_namespace_refusal`
- `test_word_owner_pieces_follow_edits_splices_and_related_owners`
- `test_unchanged_main_staging_keeps_forms_and_settings_authoritative`
- `test_unchanged_main_staging_counts_actual_work`
- `test_header_story_text_reopen_counts_the_changed_header_part`
- `test_fx268_body_locations_keep_physical_slots`
- `test_body_block_reuse_preserves_forms_settings_and_live_owners`
- `test_staging_writes_and_reads_only_the_main_part_blocks_an_edit_changes`
- `test_a_fragment_imports_into_a_main_part_whose_root_has_another_word_prefix`

`body_lookup_work_is_linear_at_250_and_1000` instruments each lookup pass separately: actual XML read attempts, consumed bytes including failed reads, and examined candidates including rejects. It covers direct and controlled cell replacement, picture resize/insertion, block and nested control movement, and comment movement. Cold and warm phases keep separate vectors. Each warm sample follows two distinct changed warmups after the cold edit. Setup, target acquisition and post-call save/reopen observations remain outside the window. Mandatory mutation validation remains inside. Only audited debug agreement oracles pause this counter, not the production empty-frame qualification that also uses `uncounted_scan`. Every pass and unit uses the same five-times bound and zero additive allowance. The old per-paragraph picture-prefix reader must fail that same gate. The smaller controlled-cell regression uses balanced 25/100 fixtures and an actual legacy-route control, not synthetic work increments. Lookup read and candidate units do not measure all writer or model-copy work or Python publication. Separate output units count assembled and copied normalized bytes. Their separate timing and memory gates remain required. Implementation alone is not acceptance.

The implemented mechanism 6 span case is inline in `crates/rdocx-oxml/src/document.rs`. The implemented address cases are inline in `crates/rdocx/src/document.rs`. `test_fx268_body_locations_keep_physical_slots` is that binding node. It does not replace `test_body_block_reuse_preserves_forms_settings_and_live_owners`, `body_lookup_work_is_linear_at_250_and_1000`, or the A3 and A5 forms, settings, provenance and composition criteria. Those broader composition checks are separate from address coverage. `main_body_block_reuse_preserves_physical_owners_and_publication` is the implemented broader A3 composition proof. The seven address cases do not replace it. The last two staging-counter slots are populated by mechanism 7. The broader native and Python A3/A5 compositions and A10 controls are implemented. Their current execution and remaining acceptance are recorded in the progress notes. The original five native groups and four Python groups stay. A9 timings run only in the final union. The new named tests cover a forced fallback, a cache-off build, an invalidated context, and an actual RED on the old provenance. Inline unit helpers already in `document.rs` stay allowed.

`crates/rdocx-py/tests/typing_smoke.py` changes only if a public Python signature changes. None is expected. The native test calls all four new `try_` methods. Each local batch names one RED case on the old unsavable mutation, one GREEN case on a valid document, and one negative control. Worker verification is scoped `cargo fmt` and clippy on the changed crates. It is not the full suite, Word, the hash harness, or the performance rider.

The Q19 RED captures the model, the bytes, and the live handles at the point where the old save already fails. It does not treat two failing `to_bytes` results as proof that nothing changed.

## Rejected alternatives

- Importing PR415, PR417, PR418, PR419, PR420, or PR421 as a wholesale merge.
- Trusting semantic `PartialEq` as the only cache proof.
- Dropping the full-rebuild fallback in order to keep a fast path.
- Making the old infallible Rust methods panic or return `Result`. Q21 rejects both.
- Claiming the whole 252-call catalogue is constant, or copying a donor benchmark as credit.
- Treating the 14 A10 counted-work calls as a substitute for the 79 flat gates.
- Leaving the six Q24 calls measured only after the user required all 79 flat-cost gates.
- Reusing a field-bearing block because `PartialEq` returned true.
- Persisting a second shadow form model beside the integrated F-288, F-289, and F-290 facts.

## Test plan

| Timing | Category | Test | Asserts |
|---|---|---|---|
| local scoped | regression | The original five native groups and four Python groups | Old-source semantic RED on the unsavable mutation. Fixed-source GREEN. Negative control on cache-off or a skipped invalidation |
| local scoped | regression | The four added native groups and two added Python groups | Forced fallback, cache off, invalidated context, and actual RED provenance. A9 timings stay out of this row |
| local scoped | regression | A10 native group and the two A10 Python groups | Equal six-slot counts. The first four slots are byte counts and are zero, with the fragment byte exception. The last two slots are child counts and stay equal. They are not byte counts. Reopen bounds, cache-off scaling, and alternate-root whole read stay. The approved set adds the five Python calls and the three native cases. A10 is not a timing promise |
| local scoped | regression | A11 and C8 | Release fallback, error order, and the named context cases. Approved with mechanism 8 |
| local scoped | regression | Four `try_` methods | Index, text, zero-column, and namespace results. Old four methods unchanged |
| final sprint | performance | Release Python 3.12, the 79 calls, the five A9 calls, the 252 catalogue, and the original 407 gate | A6 and A9 bounds and an individual residual report |
| final sprint | memory | The 20000-paragraph RSS rider | A7 samples. No invented cap |
| final sprint | risk rider | `/verify --full`, hash, rustdoc, WASM, publish dry-run, installed stubs, and the prior seven Word and SSIM gates | A8. Unrun until the union |

The test gate from HLD 14 is regression. Local regressions are the readiness evidence. The performance, memory, Word, hash, and full-suite riders stay pending until final acceptance. No worker runs those riders.

## Approved owned transaction and complete splice packet clarification

Root approved the three bounded changes at 6542884725608c2baabf149dcc2496485d85b1d1. The existing borrowing Rust APIs and stable Python APIs keep their contracts. The additive doc-hidden internal bridges are `try_replace_text_at_owned`, `try_replace_text_in_cell_owned`, `try_replace_text_in_story_cell_owned` and `insert_picture_to_story_with_options_owned`. Each consumes an already private candidate. The three replacement bridges return `Result<Result<Option<(Document, usize)>, ReplacementCountMismatch>>`. None means no change and supplies no candidate. A mismatch or other error supplies no candidate. The picture bridge returns `Result<(Document, ContentLocation)>`. Concrete target and picture arguments remain the same as their borrowing counterparts. Both routes share each actual checked operation body. Old Rust methods still clone, then commit only a checked changed result. The binding keeps its outer private clone and conservative publication.

A consuming operation retains its actual source package map and invalidation flag and moves the two complete engine Mutex values before mutation. The same concrete finalizer used by ordinary native commit compares that source for signature invalidation and transfers the engines. This is a real package-map clone, with shared payloads. It does not retain another typed document, substitute an old Parsed package or alter mutex poison state. XML, target, media, dimension and option validation retains its original order. Reopen faults, no-op and expected-count mismatch leave the original live document and handles unchanged.

`PreparedMainRead` borrows one immutable candidate and owns actual canonical XML, its optional writer record and checked body/child spans. It is never retained in Document or a binding. Only the scoped replacement pre-edit phase and the reopened picture story/owner selection share it. Related sources are obtained at the original point after main serialization and before owner selection. Producer/source proof, private facts, namespace/decoder checks, resource admission, error order and force-off remain. The view ends before source edits, preparation, identifier allocation/remapping, reopen or publication. The ordinary span calculation and full fallback remain shared and unchanged.

Splice planning validates canonicality, observes only the two requested slots and counts scope cardinalities in one current inventory walk. The immediately consumed plan replaces the all-slot temporary map and the later duplicate scope-sizing walk. Preparation materializes inserted owners during its existing state assignment pass. Commit moves that complete packet instead of rescanning every kept owner to extract inserted state slices. All removal, projection, parent/cell ordering, scope retirement/birth, current owner rebasing and mandatory indexing remain eager. The six consumers are paragraph/table add and insert, remove_content and remove_content_range, with their existing AppendBody or Conservative publication. No persistent mutable-model success cache or lazy index is added.

Actual old-algorithm RED uses authenticated debug artifact 72d9c391fc205407fdee1eb783d73de3273f4318a11548a080c6ae42200a35d9. The five named public publishers each show a nested clone, while separate read-phase checks show two actual current outputs. Simple one-run splice controls show the old actual redundant walks and map inserts. The required Original407-shaped control separately uses python-docx's exact six verbose Arial 10 black runs with explicit false formatting flags at 100/1000, cold and two changed insert warmups, and counts actual setup owners. Do not equate the simple fixture with that shape. Debug counters record actual named clone calls, output calls/bytes, slot insertion attempts and visited records at the real sites. The minimal source-package origin clone is a separate debug unit. These are not changes to the six staging slots, a constant-time claim or timing acceptance. Real legacy borrowed and duplicate-query controls remain in local tests. All five publishers and six splice consumers require scoped composition checks, then one integrated final qualification. Source changes can alter peaks, so 65428847 memory receipts do not automatically qualify the new artifact.

## HLD impact

- `docs/hld/03-architecture.md`, Owned publication and causal paths
- `docs/hld/04-opc-and-packaging.md`, The package, and Checked Word operation source ownership
- `docs/hld/10-bindings-spec.md`, Python Word document operations
- `docs/hld/12-testing-strategy.md`, Binding tests, and Word document-operation parity controls
- `docs/hld/14-development-backlog.md`, the F-X268 definition

`docs/hld/02-scope-and-non-goals.md` is unaffected. No edit. It is not in the impact list.

Implementation leases, existing files only:

- `crates/rdocx/src/document.rs` and `crates/rdocx/src/lib.rs`
- `crates/rdocx-oxml/src/document.rs`, `crates/rdocx-oxml/src/text.rs`, and `crates/rdocx-oxml/src/borders.rs` only where a fact helper needs them
- `crates/rdocx-py/src/document.rs`, `crates/rdocx-py/src/lib.rs`, `crates/rdocx-py/src/paragraph.rs`, `crates/rdocx-py/src/run.rs`, and `crates/rdocx-py/src/table.rs`
- `crates/rdocx/tests/integration_test.rs` and inline units already in `document.rs`
- `crates/rdocx-py/tests/test_core.py`
- `crates/rdocx/src/content_control.rs`, only the existing indexed element lookup and `collect_control_ids`, composed with the final F-289 provider
- `crates/rdocx-oxml/README.md`, only when an actual final inventory needs it

No new source file, module, or test binary. README, the rdocx crate README, and `scripts/readme_doctests.py` change only for an actual final archive inventory, on the same pattern as a reviewed measurement, and not by copying donor figures. Stable Python is unchanged, so no new stub export is added.

## Risk routing

Matched rows of `.claude/skills/risk-routing.md`, each added to the existing final union and not run per worker:

- Public API of a published crate. The four `try_` methods are stable Rust. State the semver note. `cargo publish --dry-run` and the `.crate` size assertion stay under 10 MiB.
- WASM or PyO3 bindings. `cargo check --target wasm32-unknown-unknown`. Workspace tests keep `--exclude rdocx-py --exclude rpptx-py`.
- Any parser or serialiser. Namespace replay and saved bytes stay exact apart from the requested edit.
- Cached read facts or projections. Every fact-changing publisher is in A3, including an unchanged generation.
- Performance budget or bounded-edit optimization. A6, A9, and the local counters.
- Shared semantic behavior with multiple consumers. Native, Python, stored package, layout invalidation, and reopen.
- Public API accepting raw XML bytes. Opaque bytes stay. The new refusal is only the namespace-owner save failure, not a general XML validator.
- Grammar, scope, compatibility mode or empty collection. The C1 through C8 matrix.
- Contributed PR or new issue acceptance. The PR415, PR417, PR418, PR419, PR420, PR421, and PR422 heads are pinned and are not accepted code. The PR419 Actions logs are donor evidence only. PR421 CI was incomplete at the refresh3 cutoff. PR422 donor checks are not our pass. The legacy combined status `pending` is a literal record and is not a claim that CI is incomplete.
- Compatibility, numeric conversion or collection ordering. Owner order, error order, and the retained 407 gate.

Rustdoc warnings are denied. Installed stub tests run in the union. F-288 A6 stays the native Word gate in that plan: Word-created originals and library-created files for all three kinds, with stored-property equality separated from Word-open formatting and calculation effects. It is unchanged and it is not an SSIM gate. F-289 A7 stays equal page count, exact text and checkbox semantics, per-page SSIM at least 0.95 at 144 DPI, matched fonts, deterministic drawing, and documented margin alignment. The complete eight-row native comparison passed at c62e0a71, with exact current source and input applicability still required. Other original story riders stay unchanged. No new Word oracle and no performance comparison against Word is invented. F-X267 M21 behavior stays on that story and is not attributed to F-X268.

## Hash harness

Expected unchanged for existing fixtures touched by F-X268. Do not re-record a baseline to hide a byte change. The Q19 commit may make a previously unsavable insertion fail earlier, and only on the new checked Rust methods and the Python checked path. That commit is labelled and reviewed on its own. No new render byte change is claimed. The F-X267 M21 digest stays that story's declared behavior.

Production read-phase reporting separates current writer calls, normalized opening copies, canonical record copies, namespace replay outputs, normalized retained/current/restored model outputs, masked rewrites and occurrence writers. Only the successful optional proof's audited debug whole-reference comparison is paused. Force-off and private-tab production fallbacks must expose their actual work and are not held to the qualified one-canonical-producer bound. Splice reporting requires actual successful preparation and consumption, counts every mandatory final owner index visit, and separately counts general-publication index visits. Pinned cold/warm 100/1000 six-run controls require one consumed packet, exact post-edit owner cardinality in the mandatory index, and no general-publication index work. An empty reopened document has no body-story owner, so its actual public insertion must decline the packet after classifying the section and expose all three general-publication index visits. Append paragraph and table survivors retain their existing contract while held ordinary runs become stale.

### Shared picture normalization reduction

Scoped child correspondence may reuse a completed normalization fixed point only when its actual first and second writer bytes both equal the actual canonical child bytes under the same immutable normal frame and current writer document. The fresh producer builder may supply its actual previously checked source child after the complete source partition proof and matching namespace metadata. Other callers perform the original first read. Both shortcuts require reusable owner facts and no lexical field tokens in either child or either frame. Fields allocate private source identities during parsing, so fields, bookmarks, tab stops and unsupported frames retain the full old rounds. A failed left normalization still evaluates the full right side. First writer bytes are dropped before any fallback right normalization. No source read grants writer eligibility, and the complete picture inventory remains authoritative.

Count actual parser attempts and writer calls/bytes by side and round. On the original seeded normalization route, controls show positive actual producer seeds and fixed points and retain the historical6P+13 to3P+13 reduction. The stronger complete producer proof below has its own positive admission controls and may omit that normalization entirely, without a whole-parse escape. A real uncached reference executes both right rounds. Unequal bytes, private owners, field-bearing frames, missing w14, changed frames, marker collisions and force-off controls retain their original outcomes and actual fallback work. This mechanism evidence does not qualify the unchanged79 flat gates, five linear calls, Original407 or memory limits. Marker collisions at a field-free exact fixed point may reuse the already executed pure rounds, but still reject at the unchanged absence guard and retain the authoritative whole fallback. They do not require redundant right-side parsing.

### Complete producer source and current writer authority

Root approved this bounded repair after repeated final normal picture timing misses and independent escalation to /root/review_shared_proof_batch at c62e0a71. The source/current comparison is a different proof from generic scoped normalization. Keep every producer root/body/scope, closed-partition, actual one-child read, field/private and frame guard. Qualify the complete actual source against the current immutable CT_Document with proved_document_matches. None declines, Some(false) retains the existing normalization route, and every later child must still qualify after an earlier mismatch. Retain the existing fallible body-frame/resource preflight before either success path.

Only Some(true) may additionally offer the normal record's exact actual part bytes to writer_bytes_match against that same current document. The complete current writer, explicit w14, exact frame, every staged child, resource and private-fact guards remain. Both children and all four frames must be field-free and actual owners reusable by value. Missing or false writer proof runs the original normalization path. This composes actual read authority with separately qualified writer authority. Source blocks stay staged=None, no ScopedChildCorrespondence is manufactured, and generic callers retain their actual normalization relation, marker checks and eager right-after-left-failure behavior. There is no new cache, retained graph edge, public API or process switch.

Temporary actual pinned whole-public probes establish the mechanism before a normal rebuild: both Document.add_picture and Cell.add_picture at25/100 positively satisfied both guards, all12 prefix windows and exact old/new/OFF saved/reopened parts and returned picture facts agreed. Actual children_read181/430 became125/291, second normalization reads and both normalization writes56/139 became zero, serialized_children22 stayed and all four staging byte slots stayed zero. Source child/shell parsers and frame writers remain separately counted work. All residual writer, namespace, decoder, inventory and binding revision work remains visible. The old6P+13 and3P+13 observations are historical intermediate mechanisms, not a requirement to perform redundant normalization after this stronger proof succeeds. Seed/fixed-point tests retain a genuine source/current mismatch to exercise their original route.

Native guards use real byte-opened canonical producer packages, whitespace and zero/one/many/duplicate children, complete root/background/section/current-model comparisons, late unproved source after an earlier mismatch, changed normal bytes, staged proofs, frames and missing w14. Private fields/bookmarks/tabs, aliases/default/body scopes, source-versus-writer separation, generic memo independence, five owned publishers, resource/error/reopen/signature authority and whole-reference controls remain required. Twenty-four isolated prototype cases passed before this port. Canonical scoped evidence and independent review still follow. The temporary OFF switch and owned-model clone diagnostic are not production interfaces. The unchanged79 flat gates, five linear calls, Original407, memory and final union remain pending on the actual successor normal artifact. Mechanism success does not predict a timing pass.

### Warm picture inventory read reuse

The complete picture inventory remains authoritative for both picture insertion bindings. Preserve story ordering, every owner and paragraph scan, namespace-scope and closure checks, relationship validation, duplicate physical correspondence, extent checks and atomic publication. Do not replace the inventory with an early target-only lookup. Actual normal warm profiling at ef048ffb identifies repeated standalone paragraph parsing inside that inventory as a growing cost after the producer normalization route has already disappeared.

Root selects a bounded read-only reuse of existing successful actual child parses for direct main-body paragraphs. Require the exact current output partition, complete head and tail context, recorded root and body declarations and bindings, and the exact whole physical paragraph span. Read only an already populated successful ParsedChild. A writer-staged model, value equality alone, a missing read proof or a newly populated memo cannot grant eligibility. Start with ordinary text, tab and break runs in a no-drawing paragraph. Decline opaque run or paragraph children, alternate drawings, fields, private owners, controls, revisions, equations, rubies, hyperlinks, comments and bookmarks. Drawing-bearing, nested, related-part, changed-frame, nonexact and unavailable cases retain the original fragment parser at its original error-order position. No retained cache, public API or sampling change is added.

Literal frame provenance composes the immutable record constructors with the current exact partition. Writer assembly refreshes child records when head or tail changes. Tail edits and the leading-whitespace bootstrap refresh them too. Source splices carry a child only under the same head and tail. Cloning retains the identical frame. The paragraph query adds no getter that populates these facts. Changed-frame controls must exercise these actual publishers rather than fabricating an internally inconsistent record through private field assignment.

The local discriminator must positively count actual warm reuse at the original parser site for both pinned public picture calls and compare all twelve cold and warm windows against the real force-off route. Preserve saved part and media bytes, returned occurrence identity, relationship facts and reopened results. Count the actual remaining fragment-parser attempts and the direct proof hits separately. Add stale source, missing proof, changed namespace/frame, private/field and opaque controls, plus an earlier unrelated inventory failure that still prevents atomic insertion. Mechanism success alone does not qualify the unchanged timing gates. Run a normal release successor only after actual warm hits and correctness controls pass, then measure both unchanged picture limits before proceeding to the remaining final union.

### Canonical readback and drawing no-op publication

The normal eec45 picture gates remain failed after actual warm paragraph reuse. Root selects a coherent reduction across relationship readback, drawing no-op publication and retained-source readback while retaining the exact actual model. Keep relationship occurrence and live-registry discovery, relationship edit production, drawing slot validation, canonical ID reservation and complete drawing edit production. An empty actual relationship edit vector may offer bounded readback reconciliation. The final remapped retained source may offer the same reconciliation after its original slot and mandatory-read checks. A drawing vector whose every checked source slice already equals its replacement may offer a model no-op. Do not infer either from integer IDs or the remap map, and do not remove identity edits from the authoritative fallback.

An offered current record must prove the exact canonical source partition and complete frame and namespace context. Drawing no-op qualification reads only already successful actual child facts. Relationship and retained-source paths execute their original mandatory source-child qualification on their original final edited records before deciding whether all-child copying is redundant. Extract that actual read stage from parsed_document without changing its order, parsing, failure or private allocation. A failed qualification immediately takes the original authoritative fallback without retrying reads. After success, require the same child count, successful actual reads in the current scopes, no lexical field tokens in children or frame, reusable private facts and no reads_anew requirement. The current model must also pass the private-value policy. Qualification borrows existing reads without filling another memo or allocating private identities.

Relationship and retained-source reconciliation must produce the complete original actual readback, even when it differs from the current model. Compare every source child with its corresponding owned current child. Prepare actual source copies only for unequal safe children, and preserve the equal owned children. Compare complete actual frame fields, including section properties, namespace declarations, background XML and root attributes. Save displaced current values before applying prepared replacements and the actual frame. Validate the resulting CT_Document with the exact original resource check at the original boundary. If it fails, restore every displaced child and frame field before taking the unchanged old fallback. There is no observable partial model or successful publication on failure. Success installs the same qualified edited writer record with part=None as the original reconstruction. No metadata, parser or namespace difference may be ignored merely because output bytes agree.

Drawing success still requires complete proved_document_matches Some(true) and the original resource validation. Missing, failed, unequal, private and field-bearing drawing cases retain the entire original path. Relationship guard declines execute the original owned assembly after the same mandatory reads without repeating source qualification or changing private allocation. A writer-staged value is insufficient. All sites continue every registry update, related story, image and signature path. Retained-source authority comes from its exact final remapped source and actual reads, never from current writer bytes. No new retained cache, model Arc, public API, process switch or resource limit is introduced.

The private canonical_retained publication contract receives the existing candidate model by mutable reference and returns Option<Option<(xml, record)>> without an owned model tuple. None and Some(None) must leave the candidate unchanged. Preserve relationship remap, slot discovery, cardinality and ordering checks, and Some(None) rejection before mandatory parsing. Prepare final output bytes and the final source record before any model swap. Successful reconciliation or the original owned phase-two parse updates the candidate only after validation, then the caller executes its unchanged package and main-part record publication. Outer None still selects canonical_retained_main_part and its actual owned result. There is one production caller, and no from_package or opening-AST ownership redesign. Actual v2 counters identify one retained whole-model copy per warm call, including141 borrowed child clones at100 paragraphs. This is the evidence for batching the retained caller with the other two sites.

The first temporary probe admitted all eight warm drawing windows but declined relationship admission in all twelve windows because the new child had not reached the original mandatory reader. The second preserved those reads and established complete source/current inequality in every relationship window. Its actual clone counts and exact ON/OFF output agreement remain preserved. The specific differing fields are not yet established. These observations justify exact changed-value reconciliation, not weaker equality or missing-read admission.

After the complete cached-read guard succeeds, reconciliation consumes each exact immutable ParsedChild directly through fallible get and parsed.as_ref access. Repeating read_child would repeat the already successful staged/private revalidation for every child. The same shared self, initialized immutable locks, exact scopes and false reads_anew ensure the original accessor would return that precise borrowed value. Keep every qualification, source comparison, namespace certificate, resource check and rollback. A getter-only diagnostic control must restore the old accessor while retaining the other three optimizations, and count the actual accessor and revalidation entries at their real sites. No source authority or private policy is relaxed.

The next coherent local discriminator must record actual differing frame fields and child indices and kinds, then positively exercise relationship and retained-source reconciliation and warm drawing no-ops. Count actual Cow borrowed-child clones, owned moves and frame copies at their real sites by caller. Require the disabled shortcut to restore complete original owned-model work without moving it into other or reopened consumers. All twelve saved/reopened part, media, relationship and returned occurrence comparisons remain exact. Test unequal source values, nonidentity bytes, count mismatch, absent or failed reads, late unavailable children, private fields, namespaces, resource failure with full rollback, repeated IDs, original signature authority and unrelated inventory failure. Mechanism evidence precedes another normal release build, and the unchanged two picture gates still precede the remaining final qualification.

### Ephemeral story scan and namespace context sharing

Current warm profiling identifies repeated story projection and per-element namespace vector copies as growing work. The earlier owned source-splice proposal remains unimplemented and is not selected. Root selects a bounded scanner repair against the existing immutable source within one picture transaction. No public contract, performance bound, retained cache or parser admission changes.

For an exact source without `holds_field_words`, the first pre-insert scan supplies both the direct-content filter and the raw body section-property boundary search. The post-reopen first scan supplies both direct-content selection and the final projected index. Preserve the complete direct-filter pass before later placement errors, the original section classifier and closing-owner fallback, and the missing inserted-direct-slot error before projected lookup. Keep the extra legacy validation scan for `after=Some`. The direct filter borrows a slice and returns owned selected spans, preserving existing caller return types. For field-bearing source, independently before and after publication, retain the original scan schedule. Typed field admission can allocate private source identities, so equality of projected spans alone cannot authorize fewer field scans. False-positive lexical field tokens retain that same fallback.

Scanner-local `XmlElementFrame` and ancestor/paragraph contexts share immutable `Arc<Vec<String>>` prefix storage. Each namespace declaration still validates its prefix and decodes its value in the original order, then uses copy-on-write and performs the exact original ordered retain and append operations, including repeated identical declarations. Every attribute syntax check still runs. Ordinary attribute values retain their original decoding policy. A missing Word element prefix uses copy-on-write only when it must be appended. Unchanged elements share context without copying strings. Preserve all eleven helper call sites and empty/scope-derived roots. Typed consumers keep their borrowed slice interface, and owned `ComplexFieldEdit` contexts receive an explicit value copy at the original recording boundary. There is no parent pointer or persistent history edge.

Before another normal release build, one temporary normal-cfg diagnostic batch must compare all twelve pinned public picture windows with actual old-scan and old-prefix-copy controls, exact prior saved/reopened parts and returned occurrence facts. Count actual scan entries, events and bytes, eligibility scans, vector/string copies and namespace updates at their real sites. Native controls compare exact ordered prefix values and first errors for aliases, defaults, shadowing, undeclaration, repeated declarations, malformed attributes and entities, parent/sibling isolation, opaque/control/revision contexts and owned replacement bridges. Field controls preserve actual scan and private allocation schedules. Body sections, cell owners, placement errors and physical versus projected locations retain their original outcomes and atomicity. The complete picture inventory remains mandatory. Scoped checks and independent microscope precede one normal Word wheel-first successor. Both unchanged picture gates must pass before the remaining Word, memory and full-union qualification.

### Plain table inventory and empty complex admission

After the normal scanner successor passes Document picture timing but misses Cell timing, Root selects the remaining growing table-paragraph parsing identified by the current Cell profile. Extend the existing already-successful source-read proof to completely plain picture-free main-body table blocks. Keep the exact partition, immutable frame provenance, root/body declarations and bindings, lexical field exclusion and reusable private facts. A missing, failed or staged-only read never qualifies, and qualification never fills a memo. A source parse remains distinct from writer-byte authority.

The table predicate rejects every raw/control/revision carrier at table, row, cell, grid and property levels, including nested border/shading opaque storage, property exceptions and root-attribute sidecars. Cells contain only paragraphs satisfying the unchanged complete plain-paragraph predicate. Nested tables, controls, drawing-bearing content and unsupported containers decline the whole table. The source-to-model argument follows the actual table/row/cell paragraph parser and the unchanged discovery grammar. It does not claim every ignored property byte was modeled. Unsupported property descendants cannot grant an accepted owner or paragraph, and actual discovery remains authoritative.

Use a sorted, disjoint, ephemeral vector of exact certified body-block ranges. Direct-paragraph queries retain exact start/end equality. Table coverage additionally requires an actual TableCell owner and both its complete owner range and the complete nonempty paragraph range inside the same certified block. A direct paragraph cannot contain a whole cell owner. Reject inverted, crossing and unrelated ranges. Source part and root checks, every owner and paragraph scan, namespace resolution, fragment closure and first-error order stay before the query. A hit skips only redundant standalone paragraph parsing and its known-empty drawing projection. No new retained cache, public API, early target-only inventory or weaker malformed-input policy is introduced.

Separately, the story scanner may bypass its original paragraph-by-item complex-admission loop only when a pure, zero-allocation scan finds no complex item. Keep the complete positive loop, admission order, errors, private source allocations, retain, sort and dedup unchanged. Count the actual precheck and predicate work. This removes vacuous quadratic predicates without adding a candidate-reference collection.

Before a normal rebuild, one temporary normal-cfg batch compares all twelve pinned picture windows with independent table-proof and complex-precheck OFF controls, exact prior saved/reopened parts and physical occurrences. Require positive table-cell skips, actual restored parser/predicate work, and visible unchanged discovery/scope/closure work. Native controls cover complete shape and nested sidecars, real source versus staged/absent/private reads, frame/scope changes, exact and crossing ranges, aliases/defaults, opaque containers, malformed unrelated paragraphs, genuine drawing-table fallback and the original complex-field positive path. Scoped checks and an actual independent microscope precede one normal wheel-first successor. Existing timing bounds, all prior failures, memory and final-union obligations remain unchanged.

## Implementation checklist

- [x] Add the four `try_` methods without changing the old four signatures.
- [x] Teach the named Python delegates to refuse before mutation.
- [x] Keep the qualified splice and namespace proof on current facts.
- [x] Add the hidden piece and staging helpers with the captured signatures.
- [x] Count real walk, parse, and serialization work.
- [x] Land the behavior commit separately from the performance commit.
- [x] Leave the 79-call, A9, 252-call, memory, Word, and hash riders for the final union.
- [x] Add the eligible in-memory reopen and the direct-body span helpers with the captured signatures.
- [x] Keep `content_control.rs` to the indexed lookup and `collect_control_ids`, composed with the final F-289 provider.
- [x] Add the hidden frame and item helpers, `holds_field`, and the six-slot staging counter.
- [x] Add the mechanism 8 helpers with the captured signatures, the A10 additions, and A11. The contract is approved. Local execution and independent review apply as recorded in the current readiness and progress notes. Final sprint acceptance remains pending.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision and recommendation | Status and answer/source |
|---|---|---|---|
| Q18 | F-X268 A5 and A7 | Allow caching with measured memory use. Actual memory and fallback measurements and unchanged correctness are mandatory. No hard cap was chosen. The donor 68 MiB figure is unverified. | Answered 2026-10-24 by the user. "Allow caching with measured memory use (Recommended)". Recorded in the S95.4 issue and PR refresh, 2026-10-24. |
| Q19 | F-X268 A1 | Refuse atomically in Rust and Python before mutation when a retained namespace owner would make saving fail. Valid documents keep working. Separate labelled behavior commit. | Answered 2026-10-24 by the user. "Refuse atomically in Rust and Python (Recommended)". Refined by Q21. |
| Q20 | F-X268 A6 | Targeted 100-paragraph time is at most 1.5 times the 25-paragraph time plus 5 ms. Report every 252-call residual. Keep the original 407 gate. | Answered 2026-10-24 by the user. "Use these bounds and report residuals (Recommended)". |
| Q21 | F-X268 A1 | Add the four checked Rust methods. The original four infallible methods keep legacy behavior and are exempt from Q19. Python uses the checked path without new names. | Answered 2026-10-24 by the user. "Add checked Rust methods (Recommended)". |
| Q22 | F-X268 A7 | Apply the measured-memory policy to the PR419 and PR420 caches. Measure each layer and the combined footprint. Do not invent a cap and do not accept the donor 68 MiB figure. | Answered 2026-10-24 by the user. "Apply the measured-memory policy (Recommended)". Root approved the extension the same day. |
| Q23 | F-X268 A6 and A9 | Add the 8 flat calls to the Q20 bound and add the five linear calls at 250 and 1000 paragraphs. Residuals and the 407 gate stay. | Answered 2026-10-24 by the user. "Use the added flat and linear gates (Recommended)". Root approved the extension the same day. |
| Q24 | F-X268 A6 | Require all 79 flat-cost gates. The current set is 76 catalogue keys plus the same 3 XML reads. The six additions are `Paragraph.replace_text`, `Cell.replace_text`, `Document.replace_text_at`, `Document.add_picture`, `Cell.add_picture`, and `Document.add_hyperlink_to_story`. The bound stays 100 at most 1.5 times 25 plus 5 ms. All 5 A9 linear calls stay at 250 and 1000, including the 2 that overlap, and all 252 stay reported. The historical 73 stays the Q23 record. Current residuals from that old trio are `Run.add_picture` and `Cell.merge`. | Answered 2026-10-25 by the user. `Require all 79 flat-cost gates (Recommended)`. Root approved the refresh4 contract the same day at baseline `d1f42516eb3747442fe4304a6331eab946dc9a0e`. |

Q1 through Q24 stay answered and Q1 through Q23 are not rewritten. Q18, Q19, Q20, and Q21 belong to the earlier approved contract. Q22 and Q23 belong to the approved refresh2 extension. The approved refresh3 extension opened no question. Q22 already measures the added cache layers, the combined footprint, and the refresh4 `decodes` fact. No open question remains. Private implementation and the Q22 memory policy are Root judgment, not a new user question.

## Contract closure

Codex root approved the contract that preceded this amendment on 2026-10-24 after reading the full draft and the revision delta. The canonical HEAD at that approval is `c0b58c7fa8e4e919b9bfe5c02932f74827cea36b`, plus the working draft of that day. Q18, Q19, Q20, and Q21 stay answered. That approval left the original 65 targeted calls, the split of 62 catalogue keys plus 3 reads, the 252-call catalogue, the cache error rule, the counter split, the memory rule, and the public APIs as then written. The 7 to 12 developer-day figure stays the historical estimate for that approved scope.

Codex root approved the refresh2 extension on 2026-10-24 after reading the full revised contract and the four-file delta. The canonical HEAD at that approval is `0d8dfb68396e33ddfcaa0d9400c30c006e12d2b7`, plus this working draft. The extension revises the combined-scope estimate to 10 to 17 developer-days and adds mechanisms 5 and 6, the 8 flat calls, and A9. No implementation has started. No gate pass is claimed.

Codex root approved the refresh3 extension on 2026-10-24 after reading the full corrected contract and the four-file delta. The canonical HEAD at that approval is `d6cefc239ab7de0d4dae63e13d730da558d3a3a5`, plus this working draft. Fact cutoff is 2026-10-24T17:02:28Z. The extension is the same F-X268 story and its seventh mechanism. It adds A10 and C7. The new helpers stay the captured signatures. The staging counter stays six units. The 73 flat calls, the 5 linear calls, and the 252-call catalogue stay the same. Q22 covers the added cache layers. No new question is opened and no timing promise is added. The 10 to 17 developer-day figure stays the historical Root judgment for the scope before PR422. It is not a measurement and not a capacity promise. PR421 stays an unmerged candidate and is not accepted code. No implementation has started. No gate pass is claimed. This paragraph records that Root review. It is not a self-review.

Codex root approved the refresh4 extension on 2026-10-25 after reading the reconciled four-file delta. The user answered Q24 the same day with `Require all 79 flat-cost gates (Recommended)`. The canonical baseline is `d1f42516eb3747442fe4304a6331eab946dc9a0e`. Fact cutoff remains 2026-10-24T23:05:39Z. The extension is the eighth mechanism. It adds A11 and C8 and extends A10 counted work from 9 calls to 14. The current flat set is 79. The combined planning judgment is a provisional 12 to 21 developer-days. The 10 to 17 figure stays historical and is not added to the seven-story 27 to 41 figure. PR422 stays an unmerged candidate and is not accepted code. F-X268 stays unclaimed. F-290 is unaffected. No implementation has started. No gate pass is claimed. This paragraph records that Root approval. It is not a self-review.


### Q26 final native reference debt

The user approved requalifying the four missing PowerPoint SmartArt/timeline reference checks on installed PowerPoint 16.114.1 on 2026-10-27. The bounded contract, historical 16.104 pins, unchanged thresholds and independent capture requirements are in F-X267's Q26 section. This is existing final-union reference debt, not an M21 application oracle or performance comparison. F-X268 remains in progress. Native reference capture, trusted pin review, four actual outcomes and final applicability remain pending. No hash baseline or product behavior is changed by this plan amendment.


### Debug reference isolation and actual cost accounting

Whole-reference save and ZIP reopen comparisons use an independent cold clone for their main-part state. They discard Parsed, MainPartBlocks and MainBodyScans derived state while retaining all ordinary package and typed source data. The four debug reference entrypoints share this policy. Two paused main-parse agreement assertions inspect an initialized OnceLock without filling a cold slot, and still require unconditional fresh-model agreement. A stored failed parse is checked rather than treated as an uninitialized slot.

The six auxiliary frame, namespace, canonical-fragment, relationship-membership, closed-body and namespace-span units follow the existing staging pause at these debug-only comparisons. No new pause boundary is introduced. Production successful work and refused proofs still count outside the pause. The two existing native namespace cost bounds stay unchanged. Discriminators use actual fresh proof machinery, real refusal, readonly parse success/failure, and rich byte-opened alias, field current/default, bookmark, unknown property and W15 state. Original memo identities, serialized state and proof-slot occupancy must remain unchanged after the reference save/reopen. New source-only counterfactuals are labelled separately from the four actual original native failures. The final README/script lease also covers correction of the existing broken official docx4j comparison URL to the independently verified immutable first-party module commit6ae8731f989def24f62b96c2e6bf60eeac48a464. The comparison claim is unchanged. The original404 and corrected200 captures remain separate evidence. Release cfg excludes these debug helper/caller/counter changes and all new tests. Existing artifact and gate applicability requires an explicit independent successor assessment, never a producer relabel.

The final archive inventory lease includes the existing README rows for oxml-layout, oxml-pdf, rdocx, rdocx-layout, rdocx-oxml, rpptx, rpptx-layout and rpptx-render. Reconcile only actual final archive bytes, normalized member bytes, member count and measurement date. The 64-byte compressed tolerance and exact normalized inventory remain unchanged. These rows do not alter public or runtime contracts.


## H9 cold piece-construction repair

Actual secondary407 cold insertion and removal fail the unchanged three-times limit. Root approves a bounded reuse of the existing checked direct_body_paragraph_facts producer at word_block_owners for actual paragraph blocks. The existing native helper alone owns plain/no-section eligibility, borrowed model admission, namespace-complete serialization and actual fragment parsing. Only successful Some facts use existing plain_paragraph_owners. Err and None retain the exact full fallback. Do not introduce a typed-model shortcut, new eligibility predicate, dormant retained facts, allowance or API. First activation still validates current native state and builds current authority.

Independent source review reports zero concrete defects or smells. The predecessor test compares actual old owner facts and then fails because the checked direct-fragment producer was not selected. Source review and this mechanism RED do not prove the cold timing bound. Fixed unit/scoped checks, microscope and actual ordinary release timing remain required. All five flat-cost failures and their thresholds remain recorded separately.

## Cold activation, finite opened-paragraph reader invariance

This is an explicit new finite reader-invariance proof within the existing F268 cold407 remediation. It is not a claim that a successful writer alone proves arbitrary typed read equivalence. The current H10 cold result remains failed against its unchanged bound. No timing law, count, supported input, API outside the hidden support surface, Source allowance or review cap changes. Root approved this exact finite mechanism on2026-11-07 before implementation.

The proposed additive doc-hidden Native query is `opened_plain_body_paragraph_facts(&self) -> Result<Option<Vec<(String, Vec<rdocx_oxml::text::AcceptedRunPath>)>>>`. It returns a temporary batch in exact physical body-slot order only when the complete main body qualifies. No new public type, module, file, crate, generic, trait, retained cache or lazy authority is introduced. The existing single-paragraph direct_body_paragraph_facts remains the full checked writer/parser route for arbitrary detached paragraphs.

### One actual whole-opening proof

Check actual kept_main_part membership, including current main-part name and current package bytes matching the retained opening. Require a successful opening serialization_proof, exact full current CT_Document equality with the actual retained main_document parse, equal current versus retained root/body namespace declarations and body bindings, and the existing plain-local-edit namespace eligibility. Do not infer these from equal text, a value search, a hash collision, a source parser record or a nonempty proof alone. The proof covers the actual current complete document member, not caller-supplied paragraphs. It does not require raw source bytes equal normalized writer bytes because the finite payload grammar below separately proves these read facts invariant under that normalization.

Validate current model resource limits through the existing authoritative model admission. Apply one whole proof and one physical body traversal per cold activation, never whole-document equality once per paragraph. Only after all entries qualify may the query return facts. The trailing body section is not a body slot, remains subject to complete model/resource/frame proof, and supplies no paragraph entry. Any body child that is not an actual plain paragraph makes the optional query decline.

### Exact finite payload eligibility

Every paragraph has PPr=None. Its hyperlinks, comments, bookmarks, extra_xml including root metadata, content controls, revisions, equations and rubies are empty. Every run has empty extra_xml, extra_xml_positions and alt_drawings. Every content item is Text and every character is ASCII alphanumeric or U+0020 space. Empty text/run/paragraph cases are admitted only if the implemented writer/parser discriminator proves their accepted facts agree. Tabs, breaks, symbols, special characters, non-ASCII, CR/LF, Fields, drawings and all other content decline. Stored preserve_space is retained and is not an admission shortcut for broader text.

RPr may be absent. For present RPr, construct a default CT_RPr and copy only font_ascii, font_hansi, font_east_asia, font_cs, bold, italic, small_caps, strike, color, sz, sz_cs, underline, shading, vert_align and rtl. Require complete derived equality with the original RPr. Thus every nonlisted modeled field and every raw, namespace, root/child extra-attribute, revision marker/change/raw-position channel stays at its exact default. In particular font_extra_attributes, color_extra_attributes, language_extra_attributes, revision_markers, change, revision_xml and revision_xml_positions cannot carry data. Do not handwave modeled-only shading: reconstruct CT_Shd from its typed val/color/fill/theme_color/theme_tint/theme_shade/theme_fill/theme_fill_tint/theme_fill_shade fields and require full equality, so extra_attributes is empty. No border, fitText or eastAsianLayout is admitted through a copied nested object. The finite property list admits the intended rich407 formatting only if its actual parse satisfies this exact predicate. Execute that positive eligibility assertion before claiming the optimization covers the fixture.

### Reader invariance and ownership

The actual run writer emits each Text as one qualified w:t with the same text and optional xml:space. ASCII alphanumeric/space text has no entity or line-ending normalization. The actual run parser appends a Text for that paired w:t. The paragraph writer/parser preserve the actual run sequence in this grammar. Known RPr attribute order, canonical spelling and empty-property normalization cannot create, delete, wrap or reorder paragraph runs or modify their Text. The accepted text reader concatenates those strings and the accepted run-path reader emits each actual Run(index), independent of these formatting properties. This finite proof is for these two read facts only, not full typed model identity after normalization.

The query takes accepted_text and accepted_run_paths from the actual current borrowed paragraphs in physical body order after this proof. It returns no Source, receipt, input byte count or Work credit. It never changes package bytes, model, namespace policy, writer errors, stored IDs or revision serial. No per-paragraph facts survive as dormant caches. The immutable opening record retains its existing ownership and scope.

### Two actual consumers and fallback

The only new binding consumer is WordPieces::assemble when its previous block list is empty. Fetch the batch once for that cold build and map element i only to the actual body child i, after exact length/order admission. Existing plain_paragraph_owners constructs the same local facts and actual body context, and existing placed_route rebases physical slots. Never search values to find matching paragraphs. A populated old piece list, including a context-change rebuild, retains its existing warm qualification, prefix/suffix reuse and per-piece producer behavior. Do not add a whole-document proof to warm edits.

A declined or failed optional batch proof follows the exact existing per-block/full-authority path. It must not manufacture a successful result, replace the authoritative error, eagerly refuse a valid fallback, or grant a fresh resource allowance. Related-story source extraction, duplicate/case-equivalent part policy, link qualification, full-owner fallback, first activation and positive dormant publication remain unchanged. A body batch is not whole-document acceptance and cannot skip those other consumers.

### Required discriminators before actual release timing

Use existing Native and binding test modules. Capture genuine predecessor RED using a new batch-producer counter with actual full facts/addresses compared before the failing mechanism assertion, not an absent API compile failure or a timer. Include actual rich407 parsed RPr and six-run paragraphs with executable positive eligibility, exact accepted text/run paths and body-slot addresses against the existing full producer. Include no RPr, empty paragraph/run/Text where admitted, multiple equal-valued paragraphs and the trailing section. A changed current model or package, missing/failed opening proof, foreign or local namespace metadata, PPr, raw RPr/shading attributes, a nonlisted RPr field, tracked/control/Field content and invalid or non-ASCII text must decline into the unchanged full route with matching successful facts or authoritative error. Preserve missing-link, related-story and duplicate-part tests, dormant empty state and first-activation publication. Warm edits must show the batch counter did not run, while actual original prefix/suffix reuse remains operative. No actual Native/Rust/Python or timing result is claimed by this design.

After implementation, run affected scoped checks/tests, independent bounded review under the existing cap and Root final review, then create an authenticated ordinary shipping artifact for the original cold407 measurements. A debug mechanism counter is not performance acceptance. Preserve H9/H10 failures and existing thresholds. Full final union/native/release gates remain unchanged.

Root RED clarification: the predecessor binding regression first compares full actual owner facts and addresses, then asserts existing take_direct_fragment_bytes is zero at first ensure_word_authority. Its failure proves actual per-paragraph writer work. An absent additive API or compile failure is not RED. New batch-call diagnostics are successor-only warm-path controls.

Root namespace clarification: exclusions mean actual retained or effective namespace/raw/capture metadata. A source-only unused declaration discarded by the existing qualified parser remains a positive control when old/full accepted text and run facts agree. This read-only query adds no Source credit or publication authority. Do not add a scanner or ban unused source spellings. Retained foreign attributes, changed actual root/body bindings and effective raw owner metadata still decline.

## Shared prepared comment-parent read, 2026-11-07

Root approved the bounded reply remediation at reviewed checkpoint39287e8c, against private patch SHA256 a7da14ba01e6064298654a1b14e41f13bde3da36f4ab4471d74d2075fda23209. The current reply timing failure remains an acceptance failure. This amendment changes no status, threshold, review cap, public API, source credit, resource allowance or binding transaction.

Factor the complete comment ownership and all-paragraph factory read into one private CommentAnchorRead value. Its projection retains the existing accepted-run, flow, source marker, duplicate and malformed-owner rules. Reply reference qualification reuses the actual ownership and parsed paragraphs from that same successful read. Keep every part, cardinality, marker order, reference payload and revision check. Drop the owned read before identifier reservation or mutation. Public snapshot reads retain their clone and flush. The reply already flushes its private candidate before the prepared read. If that producer is not equivalent to the old clone and second flush, retain the original boundary and share its successful read facts. Do not infer source identity from equality or skip unrelated paragraph admission.

The binding and outer Native staged transactions remain. Binding double cloning, Cell paths and public owned callables are outside this increment. The existing canonical cold-owner batch and its errors remain unchanged.

Use the existing Native comments test module. A test-only counter at each actual all-paragraph factory records a pass only upon its first real paragraph parse, and counts that pass's parsed paragraphs. Bracket only the parent-read phase before identifier allocation. The genuine predecessor has two passes and the repaired path one. Accepted-comment projection probes and later anchor publication are not counted or claimed absent. First assert the real successful reply, markers and reopened state before the mechanism assertion.

Compare the actual first-flush prepared read with the old clone and second flush on unsaved/current rich Fields, aliases and related stories. Cover unrelated malformed-header paragraph admission and unknown-parent precedence. Preserve duplicate or split-story markers, invalid reference count or payload, and late rollback. Run scoped checks and affected tests, strict lint and format. Existing final release, timing, native, memory and integrated acceptance debt remains.


## Shared checked picture preparation

Root approves additive doc-hidden insert_picture_to_story_with_options_and_inventory_owned with the existing owned insertion arguments and Result<(Document, ContentLocation, Vec<PictureOccurrence>)>. Existing tuple2 insertion methods retain their behavior and perform no added collection. Document.add_picture and Cell.add_picture consume the complete checked inventory instead of preparing it again. The substantive collector is shared with pictures. After original authoritative location reconstruction, a field-free main and related source may share the same prepared read. Field-bearing sources retain the original post-finish fresh collector. Collector errors propagate without retry. Existing staged finish changes signature and engine state, not the immutable inputs read by the collector.

End the existing debug transaction read-phase2 guard after authoritative location reconstruction and before shared collection, preserving original phase0 collector accounting. Do not pause counters. Capture actual successful predecessor duplicate-preparation RED, full PictureOccurrence equality including physical source, both binding consumers, Field fallback, errors, unchanged insertion-only behavior and held-input atomicity. No source credit, quota, parser, retained cache or unchecked admission changes. Cell physical preselection remains its existing checked path because the separate indexed-fingerprint hypothesis failed. Fresh ordinary picture timing and all final acceptance riders remain due.


## Checked current replacement selector preparation

Root approves additive doc-hidden try_replace_text_at_selector_owned(self, story_kind, part_name, owner_index, item_kind, index_path, placeholder, replacement, expected), returning the same owned optional changed candidate/count or ReplacementCountMismatch result as try_replace_text_at_owned. Its one binding consumer is Document.replace_text_at for a current bound recognized paragraph, table or content control with exactly one index. The private CheckedItem target is constructed only by this same owned operation after complete fresh story discovery, actual source-owner and item-span checks, without intervening mutation. It skips only the duplicate initial range qualification. Shared story_item_span retains exact path, bounds and kind errors for both consumers. No stored fingerprint or external proof is trusted.

Complete initial story discovery and its unrelated errors precede selection. Missing-story selection precedes replacement-character validation, which precedes item validation. Bound stale or foreign validation remains first. Constructed or numeric selectors, unknown or unsupported kinds and nested paths retain the original resolver. The full postflush canonical/Source/C1 pipeline, copy reservations and conservative publication remain unchanged. No new cache, parser, quota, Source credit or old API signature change is introduced.

Capture successful predecessor RED with the existing actual spanned source producer counter at the preflush boundary, then require one current preparation instead of two. Compare full packages, current source and live-format behavior, related, Fields, retained raw aliases, exact malformed-source and invalid-character precedence, no-match, mismatch, stale/foreign, nested and constructed fallback. Existing affected controls may supply cases without duplicating the complete matrix. Fresh ordinary artifact timing and all final integrated acceptance riders remain due.
