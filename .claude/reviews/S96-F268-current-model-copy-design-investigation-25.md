# S96, current model copy investigation25

## Result and scope

No small, correctness-closed CT_P or BodyContent copy repair is selected from these producers. The broad copies have distinct owned consumers. Removing them by borrowing existing caches, changing their retained ownership, masking the live model temporarily or substituting XML edits would cross contracts that this investigation has not proved. This is an honest bounded result, not a claim that further optimization is impossible or that the remaining failures are noise.

Root31 and Support13 freeze reviewed21+24A. The latest ordinary result is76/79flat with all five linear and both407 protocols passing. Reply, Document.add_picture and Cell.add_picture remain failed rows. Earlier passing Document picture results do not waive its new failure. No quiet window is active, and other builds may run. This investigation neither adopts a candidate nor requests another unchanged timing run.

Read the current F268 plan/progress, Root31/Support13, investigation22, and the relevant Native source consumers. CLAUDE.md and the canonical workflow remain controlling. Branch is sprint/s96 at cec87f9501eb13d0d5d60572472b227483262620. Actual current source hashes match the corresponding f268-direct-owner-key-implementation-20261108-1/FINAL-inputs.json entries:

| File | SHA-256 |
|---|---|
| crates/rdocx/src/document.rs |9cbb239777eba5b9311838c3715d6fa8f9baa78713a880e31b74b1dd77f96e60|
| crates/rdocx/src/comments.rs |08d9f4f8a0160d87d96b809f8f6777fc1c522fb40374a30b2a3d466702a0ab13|
| crates/rdocx-oxml/src/text.rs |dd52458b751a823c91532bd35f69a3570da6bd809e315beb1e5a923c45259765|

These are checks of the inspected producer files, not another all4309-input verification.

## Older profile attribution and current ownership

Investigation22's five-path symbol packet predates21+24A. Its reply stack actually contains source_blocks_in, proved_document, BodyContent::clone, CT_P::clone and CT_RPr::clone at `.claude/scratch/S96-final-evidence/f268-five-remaining-symbol-diagnostic-20261108-1/reply-k2-attempt-1/profile.sample.txt:2884`. The Cell picture profile actually contains opaque_relationship_ids_from_serialized_story and CT_Document::clone at that packet's `cell-picture-k1-attempt-1/profile.sample.txt:425`. These are real older sampled producer chains. They do not quantify current21+24A clone cost or establish the latest three failures' cause. Root's new current three-path profiles were not available to this completed investigation.

The source producer ownership remains as follows:

| Producer | Owned consumer and boundary |
|---|---|
| clone_for_staging | Independent mutable transaction Document, including a DocumentModel clone at `crates/rdocx/src/document.rs:25130` |
| proved_document | Complete owned CT_Document, converting each actual read Cow into ownership at `crates/rdocx/src/document.rs:11365` |
| source_blocks_in | Mutates that owned document through replacement and validates it before returning a model/record pair at `crates/rdocx/src/document.rs:29128` |
| read_source_child | Actual returned item plus separate reader-cache and eligible writer-cache values at `crates/rdocx/src/document.rs:12447` |
| opaque_relationship_ids_from_serialized_story | Temporary mutable masked DocumentModel at `crates/rdocx/src/document.rs:25645`, subsequently validated and serialized |

These copies are not the certified-hit leaf return copies removed by20. Current CommentParagraphValue returns its existing ephemeral strong-root loan only after its established exact hit, at `crates/rdocx/src/document.rs:9409`. It does not make a mutable full-document reconstruction borrowable.

## Reply reconstruction and its smaller apparent opportunity

source_blocks_in first calls record.proved_document before computing changed prefix/suffix boundaries at `crates/rdocx/src/document.rs:29047`. proved_document verifies exact part/frame/partition context at `:11352`, reads children in original order at `:11354`, converts borrowed children into owned values at `:11365`, clones the shell and completes empty-namespace facts at `:11367` and `:11371`.

read_child can perform fresh parsing for Field/private classification at `crates/rdocx/src/document.rs:12472`, initialize a real shared read proof at `:12476`, or parse again for changed namespace context at `:12485`. Thus moving range refusal ahead of proved_document would skip actual reads and shared cache effects. Investigation5 already records that blocker. A staged writer model alone does not replace these reader obligations.

There is a real apparent dead-copy seam: a borrowed old child can be cloned at `:11365` and later removed by the splice at `:29128`. However, selecting which clones to omit requires the replacement boundaries that are currently calculated after all those reads and namespace completion. A deferred Cow collection would introduce another temporary owner buffer, change clone/drop interleaving, and need to keep all actual old reads, replacement reads and shared cache effects in their original sequence. It is not a drop-in move of an existing owned item. All kept children still need owned values for the final CT_Document.

The distinction matters for private Field. Its ordinary clone executes the explicit child enumeration machinery at `crates/rdocx-oxml/src/text.rs:2111` and `:2133`. It does not mint a receipt, but arbitrary Field clone work cannot silently become omitted work under a generic borrowed-child shortcut. A finite exclusion would need an actual complete domain proof and matching current-case counts, rather than inferring safety from visible paragraph text or clone-named samples. No such new finite gate is selected here. This possible future seam is not certified safe by this report.

At the replacement producer, read_source_child performs actual child parse, write/round-trip classification, shell parse and shell comparison before returning. It then retains a cloned parsed item in proved, optionally a separately cloned staged writer item, and returns the original item at `crates/rdocx/src/document.rs:12447`. Moving the original into either cache merely forces another copy for its returned consumer. Sharing the two caches through a new Arc representation changes retained cache ownership, proof/root identity and release behavior. It is outside the requested narrow repair. Unsupported normalization, private shapes and reader/writer provenance must keep their existing differences.

## Picture masking and staging blockers

Picture preparation runs current Field-model preparation before identifier canonicalization at `crates/rdocx/src/document.rs:28251`. The later opaque relationship inventory obtains current canonical source and its relationship set before cloning the complete masked model at `:25638`. It chooses collision-free sentinels at `:25648`, modifies authored drawings, hyperlinks and section references at `:25666`, and serializes the masked model at `:25712`. The final inventory subtracts only actual chosen sentinels and intersects current relationships at `:25730`.

canonical_source_from validates the complete supplied model and calls the existing checked serializer with that model and its entire content at `crates/rdocx/src/document.rs:29183`. Existing neighbor blocks are only offered reusable bytes. This preserves actual equality, resource checks and fallback writer errors. Replacing the temporary model with an XML-only sentinel edit would require a separate exact correspondence proof for all modeled/opaque references, namespace replay and preserved XML. It cannot be justified by the fact that the requested public action inserted one picture. Investigation17's raw_xml/namespace equality blockers remain unresolved and are not reused as a fast path.

Temporarily masking self.document in place and restoring its strings is also not equivalent to the current immutable helper. Mutable access invalidates the opening-clone marker at `crates/rdocx/src/document.rs:7986`. Serialization can fail, so restoration would need an exception/panic-safe owner, all original identity and cache behavior, and no changed first cause. No rollback owner or mutation protocol is selected. General Arc copy-on-write storage for DocumentModel would likewise change transaction and retained-root ownership, requiring a separate representation design and Q22 evidence.

The drawing source rewrite later uses source_blocks_in again at `crates/rdocx/src/document.rs:26415`, installing its complete owned result or falling back to the actual whole reader at `:26428`. That path inherits the reconstruction constraints above. It does not supply permission to omit a read or to keep a mutable borrowed child.

## Useful next evidence and handback

The already requested current profiles can establish which actual producers remain active after21+24A. If they justify revisiting the discarded-old-child seam, a source-bound diagnostic should count each old clone by original index, classify whether that exact index is eventually spliced out, and record Cow borrowed/owned/read-anew outcomes, complete source-context proof/cache fills and final refusal or success. Observe actual private Field enumeration separately. Keep semantic, first-error, full output ZIP/member/ID/relationship, Source and held/fresh/reopen equality ahead of copy-work assertions. Retain actual temporary capacities and release behavior. This describes evidence that could close a future proposal, not new execution or a demanded product change.

Zero new defects, zero smells and zero nitpicks are established. No finite model-copy candidate is approved. Current finite qualification,20loans,21scheduling and24Alookup remain untouched. No Source/admission/capture/event/attribute/resource read, equality requirement, owner generation, callback, publication order, threshold or memory policy is relaxed.

Only this report was written. No source/plan/test/progress edit, Cargo/build/test, product execution, new profile, timing or Git/network mutation occurred. All three latest elapsed misses and final memory/equivalence/calibration, Office, installed typing, hosted/archive and integrated-union obligations remain pending. Control returns to Root.
