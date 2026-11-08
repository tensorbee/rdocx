# S96 standalone value loans, design investigation20

## Disposition

The exact private owned-or-model-loaned return is technically feasible. Existing product consumers need only immutable paragraph access. Preserve actual parsed CT_P ownership on every miss, and replace only the deep CT_P clone on an already qualified exact hit with an ephemeral Arc-root plus validated selector loan. No standalone grammar, first reader, certificate qualification, key, Source or publication rule changes.

There is no unresolved semantic design blocker for the finite arrangement below. Root must explicitly adopt its representation, lifecycle and evidence obligations before implementation. This is not implementation permission, elapsed sufficiency or final memory acceptance. Zero existing defects, zero smells and zero nitpicks established. Only this report was written, with no source/plan edit, compiler, runtime, test or benchmark. Root remains sole product/Cargo writer. Prior design/support numbering remains unchanged.

## Actual ownership seam and consumers

The current reader validates live Weak-root identity, checked selector and complete XML key under the metadata lock, releases that lock, then returns paragraph.clone on a hit. A miss executes the literal CT_P::from_xml_fragment, performs unchanged lexical Field/both-finite/full-leaf-equality qualification and finally returns that actual parsed value. This is an ownership seam after the read proof, not a new reader provider. References: `crates/rdocx/src/document.rs:9353`, `crates/rdocx/src/document.rs:9374`, `crates/rdocx/src/document.rs:9390`, `crates/rdocx/src/document.rs:9479`.

The sole product consumer is CommentAnchorRead::from_prepared. Its tuple currently stores CT_P by value. project uses accepted_text for unmarked intervening paragraphs, while marked paragraphs still execute accepted_comment_source_projection over actual source XML. reply_boundary calls accepted_run_paths, accepted_run and reads run.content. None mutates, extracts or clones the stored CT_P. Marker/location/flow/source work remains separate and fresh. References: `crates/rdocx/src/comments.rs:3069`, `crates/rdocx/src/comments.rs:3095`, `crates/rdocx/src/comments.rs:3157`, `crates/rdocx/src/comments.rs:3172`, `crates/rdocx/src/comments.rs:3258`.

Both complete product lifetimes are closed. anchor_reply_to_parent explicitly drops CommentAnchorRead before reserving the reply id or publishing the anchored edit. Snapshot listing holds its staged source clone through construction and all projections, then returns only owned ranges/text/errors. The private read object is not installed in Document, returned through Python or retained in an anchor result. References: `crates/rdocx/src/comments.rs:2026`, `crates/rdocx/src/comments.rs:2034`, `crates/rdocx/src/comments.rs:3519`.

Actual Root28 mechanism evidence already records55 small and154 large owned CT_P hit clones, composed of35/102 direct and20/52 table hits in each original third-reply trial. This establishes a real predecessor copy operation. It does not establish a current optimized allocation total or a predicted elapsed gain. The read implementation at this seam is unchanged by the separate18/19 materialization/traversal changes. Reference: `.claude/scratch/S96-final-evidence/f268-native-story-frame-name-implementation-20261108-1/new-diagnostic/observation/OBSERVATION.json:7`.

## Minimal concrete representation and visibility

Use one concrete crate-private paragraph value owner in existing document.rs, with private two-case storage:

```rust
pub(crate) struct CommentParagraphValue {
    storage: CommentParagraphStorage,
}

enum CommentParagraphStorage {
    Owned(CT_P),
    Model {
        root: Arc<BodyContent>,
        selector: StandaloneParagraphSelector,
    },
}
```

Keep storage, selector and constructors private to document.rs. Comments needs only the value type and immutable access. A substantive std::ops::Deref<Target=CT_P> implementation matches the two storage forms directly. Owned borrows the actual reader result. Model resolves its private selector on its retained immutable root. Construct Model only at the existing positive-hit branch after a valid selected paragraph is proved. Clone the candidate root Arc and store that same selector outside every metadata lock. Do not retain BodyBlockScan, its bytes/caches, the writer vector or the certificate cell.

The existing selector resolver uses checked row/cell/content access and requires a Paragraph terminal. A constructor must establish this invariant before publishing a loan. Deref can then use a documented invariant assertion on that same resolver, because its private selector and immutable root cannot change. Do not expose arbitrary-root/selector construction, use unchecked indexing, retain a raw paragraph pointer, or create a self-referential borrow. References: `crates/rdocx/src/document.rs:9198`, `crates/rdocx/src/document.rs:9204`.

This is a substantive invariant/ownership owner, not a forwarding-only wrapper. It owns two actual storage forms and the validated immutable selection lifetime, with one local implementation of their common read access. The private enum makes invalid cross-form states unrepresentable. No custom trait, generic, dynamic dispatch, builder, module, file, stable API or public export is needed. One existing immutable access trait is sufficient. Do not add both Deref and AsRef solely for convenience.

Do not derive or add Clone, PartialEq, DerefMut, mutable root access or an owned-extraction convenience API. Derived equality could compare the entire BodyContent/Table and introduce private sibling Field work. A generic wrapper clone could reintroduce deep Owned copies or obscure the intended lifetime. Product consumers require neither. Tests compare explicit selected CT_P references.

Keep Owned inline. Boxing would add a new allocation on every miss and is not selected. If large_enum_variant is reported, a narrow documented allowance on the private storage enum can retain this explicit no-extra-allocation contract. Record actual size/alignment of CT_P, owner/storage enum and the CommentAnchorRead tuple. The inline discriminant/unused variant storage still has a cost in every tuple, including loan hits and omission mode. No layout improvement is assumed.

## Exact read, snapshot and private-state invariants

On every miss, move the actual parsed CT_P into Owned after the current qualification/cache updates. Even a successful first qualification returns Owned. Do not immediately replace that first reader result with a model loan. Changed/malformed input, absent/ineligible hints, independent equal roots and unsupported/private leaves keep the same reader and errors. Failed collection drops all earlier owned values/loans normally. Reference: `crates/rdocx/src/document.rs:9388`.

Only the existing exact hit creates Model. Root pointer identity is checked through live Weak upgrade, not scalar addresses or value equality. The selector has already resolved to a paragraph, and the full XML key matches. A concurrent key replacement or cache clear after releasing the lock does not invalidate the selected immutable root/value. The retained strong Arc keeps that exact snapshot alive without extending any cache proof. No new key or AST is installed on a hit.

The finite selected leaf contains no Field, bookmark, raw/control/revision/ruby/equation or tab path and only its existing qualified properties/text. Its immutable projections therefore cannot mint or mutate a private reader identity. Unsupported selected leaves remain Owned actual reads, preserving their separate Field source_owner_id and existing projection behavior. A table root may contain unsupported private siblings, but resolving the selected leaf performs only checked indexing. Never compare, clone, serialize, qualify or traverse those siblings to return a loan. References: `crates/rdocx/src/document.rs:8804`, `crates/rdocx/src/document.rs:8830`, `crates/rdocx/src/document.rs:86840`.

Advisory ordinals remain value storage hints. Model-domain-only incidental equal-value hits are still allowed under the adopted certificate contract and still confer no physical source ancestry, source membership, edit permission or Source credit. The exact fixed reader and full equality created the value proof previously. Changing the hit's ownership does not create another authority.

Clone/fresh_context carry, actual source-child staged replacement and source-only staged removal keep their existing behavior. New values retain an immutable root independently, so an already returned loan remains an old snapshot after block/root replacement. A new read must miss the old key/root as before. Production has no inner staged Arc get_mut/make_mut/strong-count branch on which this extra strong owner can change admission. Enclosing block/vector COW and all source guards remain. References: `crates/rdocx/src/document.rs:9855`, `crates/rdocx/src/document.rs:10796`, `crates/rdocx/src/document.rs:12407`, `crates/rdocx/src/document.rs:28890`.

## Lifetime and memory obligations

A direct loan retains its paragraph root. A table-cell loan retains the entire existing table root, including unrelated/private siblings, until that ephemeral value drops. Several cells can have several Arc owners of one allocation, never several copied table ASTs. This can delay release and adds atomic reference operations. Do not describe it as free memory or count each Arc as a copied root payload.

The factory currently holds candidate block/root ownership until it returns. The proposed values extend selected root ownership through CommentAnchorRead's projection lifetime. In the two actual product consumers the source Document remains alive throughout that interval, and reply drops read before mutation. Still require explicit retained-byte/capacity and release evidence, including held loan after factory/document drop, last-loan release, same-root sharing, new-root same-value replacement, partial collection failure and source-only clearing. Cache cells must continue to hold only Weak roots. A cache Weak can now upgrade while an independent returned loan legitimately owns the snapshot. Tests must wait for the last loan to drop before requiring root expiry, rather than weakening expiry assertions or mislabeling the loan as a retained cache.

The existing StandaloneValue omission endpoint must name returned loans and their ephemeral root lifetimes alongside keys, maps, capacities, Weak/control state and inline representation overhead. Omission performs literal reads and returns Owned without allocating/carrying certificates or loans. It must not disable hits while retaining model-loan state. No new retained-cache endpoint, second AST, root history or RSS ceiling is introduced. Final equivalence/calibration/peak/RSS work remains mandatory.

## Required honest tests and mechanism evidence

Adapt the actual production read return type and CommentAnchorRead tuple directly. Do not keep a test-only owned facade that clones dereferenced values and masks product ownership. Existing equality assertions should borrow explicitly, for example `let actual: &CT_P = &value`, and compare that selected leaf to the actual expected reader result. This can support a genuine predecessor discriminator without referencing a successor-only variant. Retain exact tests for actual first reads, malformed keys, key/root/selector mismatch, selected-key removal, private Field identities, source replacement, COW, poison and parallel roots/selectors.

Keep the existing read/hit/key-copy/owned-model-clone counter meanings. A hit must stop incrementing owned-clone count when that real CT_P.clone is removed. Add a separately named actual returned-loan counter if needed. Never relabel the old clone slot as a loan or constant success marker. Native predecessor semantics must succeed before positive owned-clone count fails a zero-copy expectation. Successor must show unchanged literal reads/hits/key copies and exact leaf results, zero hit CT_P clones and actual positive loans. No hidden clone in Deref, an adapter, tuple construction or test helper may make the counter pass.

Add direct-body and at least two table selectors, equal independent root, shifted incidental equal-value hint, nonmatching selectors, and a private Field sibling with no whole-root equality/copy. Misses must still be the actual first parsed Owned result. Verify that held loans keep their exact old leaf after source/model replacement, cache clearing and factory drop, while new reads qualify independently. Drop all loans and prove old roots/cells/keys release at the declared boundaries. Poison/concurrency must retain immutable validation and never parse or clone a paragraph under the metadata lock.

Exercise the full CommentAnchorRead consumers, not just the private getter. Preserve marker source projection, accepted range/text, parent/reply boundary, Field first causes, unrelated malformed-source refusal, callbacks, held/fresh behavior and rollback. Exact original actions must retain setup/getter order and complete semantic/reopen checks before mechanism assertions. Use identical inputs and compare complete ZIPs/every member. Observe actual Arc returns, owned copy payload, source/read work and retained unique roots separately. Counter diagnostics are not optimized allocation totals or elapsed measurements.

Explicit plan/HLD03/04/10 and memory endpoint adoption, scoped correctness/strict lint, source/artifact custody, installed controls and independent implementation review must precede a new ordinary qualified timing artifact. The separate18/19 repair evidence and its pending review are not approved by this investigation. All latest ordinary elapsed failures and final memory/Office/typing/hosted/archive/full-union gates remain pending. The pass ends here and returns control to Root.
