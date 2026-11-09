# F-X277b implementation, Root final pass 1

Reviewed the clean-Base worker diff fromd1da8b8e, nine paths with1386added and
49removed lines, against the operative A277-2/3/6 contract. Frozen patch SHA256
is `183d5a6ed9c068a15cd9b373e9e806105a7ff1caaf88913835a39fcf8e42edd2`.
Native document source SHA256 is
`69facdcfcda3ab1b34c15cafc346d2c12648d63fc828c78b6e13a27764f23684`.
Citations name that frozen worker source. Root reviewed independently of the
implementer. Verdict: one defect, zero smells, zero nitpicks.

## Defects

### D1, pending normal id is not occupied when adding special records

`crates/rdocx/src/document.rs:34294`
`crates/rdocx/src/document.rs:34313`
`crates/rdocx/src/document.rs:15397`

add_footnote correctly computes its legacy max-plus-one id before adding the
specials. However, missing_note_specials receives only the preexisting notes
and raw source. It does not consider that already selected normal id occupied.
The normal note is appended only afterward at
`crates/rdocx/src/document.rs:34335`.

For an imported part with one normal note at id-2 and no specials, the legacy
allocator selects-1. The helper then appends a separator at-1 and a continuation
separator at0, followed by the new normal note at-1. The analogous existing
normal id-1 case selects0 and collides with the added continuation separator.
Normal notes at these ids are explicitly supported by the integrated provider.
This introduces duplicate physical note ids within one authoring transaction.
Existing imported owners can keep their bytes while this new collision still
breaks note identity and package validity.

Preserve the approved legacy normal allocator and reserve its pending id when
deciding which optional specials can be appended. Do not change the returned
normal id, rewrite an existing owner or invent another special id. The approved
collision rule requires leaving the occupied conventional special id alone and
still creating the normal note. Add both adjacent negative-id cases with exact
raw preservation, unique resulting ids, returned normal classification and a
successful save/reopen. Current imported tests use either existing-1and0 or
unknown-7, so they miss the pending conventional-id collision. This finding is
source-traced, not a claimed executed historical RED.

## Smells

None found. The private maps represent two actual integer domains. The
concrete source-owner traversal handles paragraph, run, table, row, cell and
control owners without a new public visitor, trait, generic or module.

## Nitpicks

None.

## Not found

No additional defect found in the exact three hyperlink constructors, shared
preparation hook or typed source selection examined. Paragraph source_runs is
the terminal inventory, outer controls include direct Runs and retained
revision owners, and raw/comment-only inputs do not trigger style creation.
Style identity remains the id, imported bare links stay bare, and an existing
exact id is not rewritten. Output-only publication retains the existing actual
Field-normalization assign-back condition.

No additional defect found in the numbering emission and inverse reopen
mechanism examined. The permutation covers only main-document and numbering
writes. Saved live values restore after the captured write Result, before
propagation. New maps install only after complete flush success. Model and
body translations use one original id lookup, and separate actual reopen paths
retain imported preserved-id sets. Direct picture, normalizing save/to_bytes,
signing, ordinary output candidates and process reopen were traced separately.
The signature control requires exactly one retained signature, avoiding an
empty-list success.

The remaining spacing and special-record implementation uses explicit raw
children without empty Text or added run properties. Caller whitespace and
custom marks are preserved. The source-qualified special type check reuses
the integrated provider classifier rather than trusting the legacy typed
unknown fallback. This does not excuse D1's pending-owner omission.

## Evidence and exclusions

The retained provenance table distinguishes actual Base semantic failures,
two failed fixtures, corrected normalization Base RED and corrected output-mode
Base PASS. It preserves original logs. The latest scoped source reports12
integration and9unit passes,56affected provider cases and one existing marker
case, with all-target/all-feature compilation and warnings-denied Clippy,
formatting and prose checks. D1 is not covered by those passing cases.

The Python test additions are source and AST evidence only. Actual installed
Python invocation and typing, schema/native checks, C1 composition, output hash
attribution, full workspace, hosted CI, packaging and all final union riders
remain pending. Only A277-2/3/6 belong to this child. Existing A277a/C tests are
preservation controls and do not transfer their delivery credit. No source,
test, plan, HLD, baseline or delivery record was modified during this review.
