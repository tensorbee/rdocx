# S96 MSRV58 comment identity proposal review

Read-only review of the private PRODUCT_CANDIDATE.patch in `/private/tmp/rdocx-s96-prepared-integration-readiness/MSRV70-note-fragment-comment-private-proposal-1`. The candidate is not adopted. Its duplicate exact move no-op hunk was excluded because canonical already has that repair. No product, test, plan or progress edits and no Cargo execution occurred. This is bounded proposal support, not a whole-feature peer or review-cap reset.

## Defects

### D1. The durable-ID owner lookup accepts the wrong paragraph

`/private/tmp/rdocx-s96-prepared-integration-readiness/MSRV70-note-fragment-comment-private-proposal-1/after-comments.rs:3896` changes IDS_REL lookup from the comment's terminal paragraph map to by_parent_para, which includes every paragraph. After an append, the old ID can remain on an earlier paragraph and still qualify. This changes the global owned-graph validator, including unrelated read, import and destructive-operation consumers.

The existing last-only proof is correct. Microsoft's CT_CommentId contract identifies the associated comment through its final paragraph, not any paragraph. The new lookup therefore accepts an invalid companion relationship instead of preserving the invariant. See [MS-DOCX CT_CommentId](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/9c360cd7-653f-4d82-82be-7bda2488c0c1). CT_CommentEx likewise keys its own row by the final paragraph. See [MS-DOCX CT_CommentEx](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/9660dacc-2ceb-4352-87d2-42ba1184f522).

F-X279c-design.md:96 and F-X279b-design.md:422 require imported commentsIds bytes to remain unchanged. That does not authorize relaxing their owner semantics. Retain the final-paragraph proof. A publisher must lawfully retain or transport the actual terminal identity while preserving the imported part, or refuse if it cannot do so under the existing contract. A new relaxed owner lookup is not an acceptable substitute.

### D2. Repairing one comment changes unrelated legacy definitions

`/private/tmp/rdocx-s96-prepared-integration-readiness/MSRV70-note-fragment-comment-private-proposal-1/after-comments.rs:3706` treats any missing final paragraph ID as sufficient reason to repair. The loop at3721-3728 then resizes and allocates IDs for every comment with that shape, regardless of which comment's content or final owner actually changed. The setter marks the whole comments model dirty at `/private/tmp/rdocx-s96-prepared-integration-readiness/MSRV70-note-fragment-comment-private-proposal-1/after-document.rs:16310`.

For example, retain legacy comment C0 without paraId and append a picture paragraph only to C1. C0 receives a new ID even though its source is untouched. An incoming primary source identical to the old legacy source can also satisfy this predicate without an owner change. The same all-comment loop refuses an unrelated empty definition at after-comments.rs3722. These are incidental source changes or refusals outside the actual selected owner.

F-X279d-design.md:168 explicitly preserves untouched imported definitions, and F-X279c-design.md:96 preserves untouched producer content. Limit identity reconciliation to the actual changed comment owner and the actual successful publication handback. Missing metadata in unrelated legacy comments is not authorization to repair them.

## Smells

0 beyond the defects above.

## Nitpicks

0.

## Existing implementation pieces

- comments.rs2402 supplies last_para_id. The nearby para_ids helper intentionally permits legacy parent-link lookup through earlier paragraphs. That documented parent-link tolerance does not apply to IDS_REL.
- resolve_comment_staged at comments.rs2106-2132 reuses or allocates a terminal ID only for its selected comment. It also creates/resolves thread state, so calling it wholesale would introduce unrelated behavior.
- reply_to_staged at comments.rs1916-1942 applies the same selected-parent terminal identity policy, with reply-specific effects. Its allocator/identity pieces can be reused without invoking reply authoring.
- The REF companion planner at field.rs15117 and15195-15213 qualifies source and replacement companions against their actual final paragraph. It handles Extended and Ids separately. Its whole import/remap behavior is not a drop-in append repair.

No existing end-to-end append identity helper was found. A narrow private reconciliation at the actual changed-owner producer can reuse these pieces. For the reported no-commentsIds case, a new terminal ID plus qualified lossless commentsExtended paraId/paraIdParent edits is distinct from weakening imported IDs ownership. Existing imported IDs require their original byte preservation and final-owner proof.

## Scoped writer and evidence limits

The candidate parses the actual incoming comments Source through source_comments and later retains that Source in comments_reader_source. Canonical document.rs28997-29045 uses source_model_comments_from_read for that actual reader path. Marking the staged model dirty does not by itself register another Field input or justify new source credit. This review found no separate source-backed credit defect solely from the dirty flag. Actual returned-model membership, metadata handling and shared-work writer checks still require the real scoped gate.

The proposal has no executed acceptance from this review. Existing MSRV58 failure, opaque companion preservation, changed-owner plus unrelated legacy comment, imported IDs final-owner refusal and actual scoped publication controls remain necessary discriminators after repair.
