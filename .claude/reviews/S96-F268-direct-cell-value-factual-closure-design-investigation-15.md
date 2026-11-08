# S96 direct cell value factual closure, design investigation 15

## Disposition

The factual conditions left open by investigation14are now met for the proposed model domain. The finite contract below is supported for explicit Root plan/HLD adoption and implementation. It is not an elapsed sufficiency claim, an implementation review or final acceptance. No product cache has yet been added by these observations.

Zero existing correctness defects, zero smells and zero nitpicks were established. Only this report was written. No product edits, compiler, import, runtime, test, diagnostic execution or benchmark occurred during this investigation. Root remains sole product/Cargo writer. Pass numbering and final riders are retained. The actual latest ordinary reply and Document.add_picture failures remain failures, and this proposal makes no Document picture repair claim.

## Authenticated factual progression

The packet is `.claude/scratch/S96-final-evidence/f268-direct-cell-value-investigation-20261108-1`. Three distinct investigations remain preserved.

1. CANDIDATE_FACTS.json is the initial wrong advisory mapping. Actual source index_path[0]was1for the first paragraph while its model content position was0. That probe found no useful leaf equality. It is not a cache regression or an adoption receipt. Do not correct it by subtracting an assumed constant.
2. MAPPED_CANDIDATE_FACTS.json independently matches SHA-2560b482ec0f5310e01c06762aabf0aae8daf5712470b3b5baddf4f8fcf9c7f4ae3. Its actual per-source-owner paragraph ordinal and preenumerated physical model indices produce20small/52large actual table literal reads with offered roots, both finite values and complete leaf equality in each original action. Scalar addresses and key/selector sets were stable, but numeric address equality alone was not treated as an Arc identity proof. This probe did not record the final structural domain.
3. IDENTITY_CANDIDATE_FACTS.json independently matches SHA-256cef2d35d4d22c4bbe7e462a515641d63b2c45b229ac02ff48b335da85400526d. It closes both omissions using actual Weak upgrade plus Arc::ptr_eq and the exact proposed structural-domain predicate.

The identity diagnostic's actual build child0, restored4309inputs and10wheel/payload members independently authenticate. Its wheel is583377b3daa9c0759a54962ef55ed48ce56b7dec9398ebe17c8d9ce0345b9a18and Native1c44b397f04a4e60cd3063abdc284419201f18b200e51c81ea37c5ed113b1ea3. Its producer manifest currently has no source mismatch. The mapped diagnostic's build/restoration/payloads likewise authenticate. None is a shipping timing producer.

The actual source appended a TLS Vec of Weak<BodyContent>. It assigns a previously issued token only after upgrade and Arc::ptr_eq against a live candidate. Expired roots cannot receive an old token through allocator address reuse. It retains no strong AST beyond the factory. observe.py:42resets that Weak table before each fresh sequence, and its post-action boundary resets it before semantic getters/save/reopen. This source and observer were inspected, rather than trusting the token label.

old-identity-diagnostic/METADATA_CORRECTION.json:1accurately corrects the copied BUILD scope sentence that originally denied retaining Weak beyond the factory. Original BUILD.json is preserved and independently matches its recorded SHA-25678840b1d5ac67d24cc5a7146ef8b33cec3277e7291acf7985958d3e6d5386a13. The correction names real Weak control metadata and possible COW effects. Source inspection in investigation14found no production in-place mutable access to the inner staged Arc, so this diagnostic does not silently replace an existing inner uniqueness/mutation branch. Full output agreement is separate supporting evidence. No RSS/elapsed conclusion is drawn from the instrumented build.

Independent inspection of all four exact original reply0/1/2sequences found:

| Size | Trials | Actual table reads in each action | Distinct live table-root tokens | Domain, both finite and complete leaf equality | Consecutive root/selector/full-key persistence |
|---|---|---|---|---|---|
|25|2|20|5|All20|All20for0to1and1to2|
|100|2|52|13|All52|All52for0to1and1to2|

No semantic getter, save or reopen was added between the original actions. Post-sequence assertions cover reply/parent IDs, author/text, anchors/issues, unchanged table contents and body text, then reopen. All four identity outputs and every ZIP member independently match the current ordinary normal14e78/e490outputs and recorded part hashes. The mapped outputs also matched those same ordinary controls. Full XML artifacts match their recorded hashes/lengths. This proves actual available persistent equal values in the closed domain, not a positive cache hit in an implementation that does not exist yet.

## Explicit model-domain contract

Root has chosen model-domain-only eligibility. This supersedes investigation14's broader sentence promising literal standalone reads for every nested/control/raw source origin. A root predicate alone cannot enforce that sentence: advisory shifts can select a coincidentally equal simple value for another source ancestry.

The precise contract is:

- Enumerate candidates only from an already offered immutable direct BodyContent::Table root that passes the structural predicate below. Enumerate no nested table, control or raw model selector, and perform no visual merged-grid traversal.
- A main-part TableCell Paragraph location remains advisory. Its source ancestry is not certified by an ordinal or model selector. A first miss always executes the complete original fixed standalone reader. If that result and the selected finite leaf are completely equal, the certificate proves that owned value only.
- Consequently, a paragraph whose actual source ancestry is nested/control/raw may incidentally obtain a value hit when the unchanged source path admits it and its exact complete XML equals an actual-reader-proved eligible value. This creates no source affiliation, nested-model support, edit permission or Source credit. Do not publish a literal-reader guarantee for all such source origins.
- Unknown/absent/invalid/ineligible model candidates retain literal reads. All original current-source, ownership, flow/location/marker, namespace, error and resource work remains authoritative before the value lookup.

No new source-ancestry classifier, parser, scan or whole-factory refusal provider is needed under this chosen wording. If a future contract demands literal reads by ancestry, it needs a separate proven source-domain discriminator. StoryOwnerSpan at `crates/rdocx/src/document.rs:9095` contains kind/index/range/fingerprint, not control ancestry. It must not be reinterpreted as that missing proof.

The actual approved model predicate is the exact one observed by the identity producer: table.extra_xml and table.content_controls empty, every row.extra_xml and row.content_controls empty, every cell.extra_xml empty, every cell.content entry CellContent::Paragraph, and optional CT_TcPr.grid_span/h_merge/v_merge all None. This declines nested/control/raw model contents and merged-cell routes. Other table/row/cell properties do not become value/source proofs. The predicate is one pure metadata walk per factory/root and adds no reader or whole-table equality.

The ordinary owned graph is defined by `crates/rdocx-oxml/src/table.rs:3366`, rows/cells at2890and direct CellContent at2388. A valid selector is exactly physical[row,cell,content]with a Paragraph at the final bounds-checked get. Source paragraph ordinals are never used as physical indices directly. Selector resolution failure is a miss, not a new error.

## Concrete state and snapshot arrangement

Use a private concrete selector enum for root `Paragraph` versus `Cell([usize;3])`. The root remains an existing immutable Arc<BodyContent>. One positive state has one Weak root and one private key enum:

```rust
enum StandaloneParagraphKeys {
    Paragraph(Vec<u8>),
    Cells(HashMap<[usize; 3], Vec<u8>>),
}
```

This preserves the existing single Vec key for a direct-body paragraph, avoiding a new HashMap allocation on every such certificate. Cells holds one complete XML Vec per valid physical selector in the current active table root. Variant/root/selector compatibility is checked. No second CT_P/table AST, root list, Source receipt or historical map is stored. The enum has two actual consumers and owns substantive state selection, not forwarding. No new trait, generic, module, file, dependency, public API or builder is required.

Extend the existing CommentParagraphRead snapshot at Native document.rs:32422. Keep one block Arc per table, then compact cell hints containing root index, row, cell, physical paragraph indices and the next source paragraph ordinal. Their strong references and index vectors are factory-local. Do not clone the table or retain a root Arc per selector when one per table suffices. No model walk/serialization/lock is repeated per paragraph.

Advance the source-owner-local paragraph ordinal for every main-part qualifying TableCell paragraph independently of finite eligibility or a successful comparison. Select a preenumerated physical paragraph index, not index_path[0], a decremented source number or a visual grid slot. Missing/skipped model roots may shift hints, but their safety remains actual-reader/finite/full-equality value proof. Do not upgrade advisory indexing to physical membership. Existing direct-body selection and its source traversal order remain unchanged.

The snapshot still follows complete comment_story_range_paragraphs and ownership checks. comment_source_markers still executes before each read at `crates/rdocx/src/comments.rs:3100`. Source extraction, namespace restoration, flow, offsets and all current error priority are computed afresh. The cache never substitutes a Part-child read or staged writer value for the first actual CT_P::from_xml_fragment.

## Qualification, hits and lifecycle closure

On miss, invoke the exact fixed standalone reader once. Propagate its actual error. Only lexical Field-free input, both actual returned and selected leaves passing the unchanged paragraph_has_finite_plain_properties at Native document.rs:8830, followed by complete CT_P equality, may install a positive key. Never use whole-table equality or inspect private sibling equality. Do not replace the finite policy with owner_facts_reusable_by_value, strip markers or broaden the cold invariance grammar. Marker/reference grammar extension remains unselected.

A hit requires exact complete XML bytes, a valid selector resolved on the live root, successful Weak upgrade and Arc::ptr_eq with that same root. A root Paragraph state accepts only its Paragraph selector. A Cells state accepts only its valid table selector. Text, length, hash, pointer number, independently equal root or selector alone is insufficient. Return the required owned leaf clone outside all locks. Root lookup and selector validity convey no Source authority.

One cell stores one active root. On root change/expiry, replace and drop the entire old state/map/container, including its capacity, before installing a new root's positive state. Do not keep a map per old root or clear a map while treating its historical capacity as released. For one root/selector, successful qualification replaces its key and drops the old Vec/capacity. Successful nonmatch removes that selector's key, retaining unrelated valid entries. Direct Paragraph keeps its existing one-key semantics.

Reader errors are not cached. For the same valid root they may leave its old unrelated successful key available. Root-invalid state can be discarded without suppressing the actual error. Every changed/malformed full key misses before it can affect a returned value. Construct replacement state outside the mutex and publish an intact state. Reader, finite checks, equality, callbacks and owned leaf clone stay outside locks. Poison recovery retains the same root/selector/key validation. Concurrent roots may evict optimization but never accumulate historical roots, and concurrent selectors may keep only entries belonging to the active root.

Investigation14's complete producer audit remains binding. New writer/from_source records start empty. Clone/fresh_context may share an initialized cell only while staged Arc remains identical. read_source_child at12211can replace staged after Arc::make_mut on the enclosing block at10602. Reset/detach that block's cell on replacement or removal, rather than relying on a stale Weak mismatch while old keys remain retained. The explicit source-only staged=None at28675also keeps/reset emptiness. The temporary staged seed at39520starts empty. New edited blocks at10435start empty. Reused qualified writer blocks, ordinary staging clones and qualified reopen carry retain only their same immutable root and exact-key state.

All direct live-model/table/source mutation and publication paths keep their existing guards and qualification. No new certificate is copied onto a new root merely because values or bytes compare equal. A future in-place staged-root mutable access must dissociate Weak and reset state before mutation. The current source audit found no production inner staged Arc::get_mut/make_mut/strong-count branch. The diagnostic Weak observation therefore closes the actual identity question without retaining a strong AST or adding a new production mutation policy.

## Required implementation and acceptance evidence

The factual receipt selects a viable mechanism, not its implementation. Required discriminators remain:

1. Two different cells/paragraphs under one table root first read for real, then both hit with two selector keys. Reject root-only lookup, wrong-selector/different-value, independent equal root and one-slot thrashing. Preserve direct paragraph's one Vec path and existing positive hits.
2. Model-domain exclusions, absent/misaligned hints, physical index versus paragraph ordinal, invalid selectors and key/root mismatch. Include incidental equal-value hits from shifted source ancestry under the explicit model-only contract. They must remain value-only and preserve actual Source/location/marker behavior. Do not require impossible literal-source exclusion in those tests.
3. Actual first reader and both finite checks before complete leaf equality. A private Field sibling must introduce no new whole-root/private-sibling comparison or source read. Unsupported selected leaves and malformed unrelated sources keep original errors/reads. Cold grammar and marker/reference policy stay unchanged.
4. Actual staged replacement/removal, source-only clear, block COW, same-root frame carry, new-root same-value replacement, live cell/source edits, clone/reopen/refusal/rollback and all releases. Old state/maps/keys and factory references must die at their declared boundaries.
5. Per-selector successful key replacement/nonmatch, error without negative caching, expired Weak, concurrent same/different roots and selectors, and poisoned locks. Bound retained entries by valid selectors of one active root. Observe actual capacities and root/key/cell counts, not only logical lengths.
6. Genuine installed predecessor positive table literal-reader calls after successful public semantics, then successor actual table hits and fewer literal calls in the exact original actions. Cold/first-read controls stay positive where actually required. Count real key copies and returned CT_P clones separately. No added getters between original actions may warm the measured state.
7. Same original input ZIPs, full parent/anchor/range/text and untouched-cell/held/fresh/refusal semantics, then complete output ZIP/member identity and reopen. Independent source/implementation review follows affected scoped Native/installed/lint checks over authenticated uninstrumented artifacts. Only a new ordinary final-source compiler-free qualification decides elapsed behavior.

Update the existing plan and HLD03/04/10before editing product. Explicitly replace the old one-key-per-block contract and investigation14literal-source fallback wording with this model-domain-only policy. Preserve the public API and actual fixed reader policy. A later change to reader policy or grammar reopens this proof.

The existing memory StandaloneValue omission endpoint must cover the complete enum, Cells map/nodes/capacity, root Weak, every full key and lazy/clone/carry/reset path. Direct Paragraph avoids extra map allocation but still pays actual enum/inline layout and shared-cell overhead. Key count is bounded by current valid selectors, not a new RSS cap. Observe actual positive table activation, one-active-root lifetime and absence of old-key/root history in the existing memory/calibration riders. No extra AST, summed-delta claim or policy relaxation is introduced.

There is no remaining factual or technical design blocker for Root to adopt this exact finite contract. Implementation correctness, actual cache mechanism results, reply elapsed sufficiency and all final memory/Office/typing/hosted/archive/full-union gates remain pending. Document.add_picture still has a separate unclosed failure. No status, ledger, baseline, commit, push, tag or publication action follows from this investigation. The pass ends here.
