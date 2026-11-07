# S96 bounded comment-history identity reconciliation

Verdict: 1 defect, 0 smells, 0 nitpicks in the proposed test oracle. No product defect or runtime result is asserted. No source changes or Cargo were performed.

## Contract reconciliation

The raw cross-history equality is too strong for newly allocated comment paragraph identities. F-X279d-design.md:135-139 requires save/reopen preservation and the parent paraId link. It does not require independent edit histories to allocate identical fresh values. The actual publisher at crates/rdocx/src/comments.rs:3694-3758 preserves physical paragraph identities through the typed move, allocates a missing new final identity and rebinds its actual thread links. Canonical image relationship ordering and producer bytes remain separately checked at crates/rdocx/tests/regression_test.rs:3688-3732. The reported actual diff is confined to newly authored image paragraph identities2 and3 and their corresponding extended parent link. An alpha comparison over just these fresh identities is legitimate.

## D1: retain exact preexisting paragraph identities outside alpha equivalence

At crates/rdocx/tests/regression_test.rs:4607-4612 the proposed alpha map includes every primary paragraph ID. At :4626-4631 it consequently normalizes preexisting base identities as well as the newly allocated image identities. That can hide a future change to an existing durable base paragraph ID when both histories still save/reopen consistently. Repeated bytes within each history do not establish preservation against the shared original fixture. The original base comment is created at :3507-3525 before either history, so it supplies the actual baseline identity without a new fixture or policy.

Restrict the alpha map to newly introduced paragraph identities, or assert each shared original paragraph's exact identity in both outputs before normalization. Also establish the actual physical-owner ID map is unique before using it. Retain exact parent-link correspondence under the same fresh-ID map, the existing repeated to_bytes checks, exact reopen bytes, image order and unmodeled producer checks. No raw comparison of other package parts should change.

## Boundaries

The leading attribute boundary at :4629 is present in the reviewed source. This helper is specific to the canonical-prefix fixture. It does not qualify a general namespace-aware normalization of imported XML. The existing raw producer checks remain required. commentsIds absence or untouched imported bytes must remain outside the normalization. No new native, performance or final sprint acceptance follows.

Reviewed regression source SHA256: `e26cabd9b256721c7eb825cfec5252ff2bd4ef14456f4fdcb4ebe9d83fdeb8d5`.
