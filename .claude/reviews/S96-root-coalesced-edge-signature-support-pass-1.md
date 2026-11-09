# S96, coalesced edge signature, support pass 1

**Reviewed**: The new `unprotected_edges` contribution to the existing coalesced paragraph signature, its six local controls, unchanged caller exclusions, and the corrected unrelated-legacy comment fixture premise. This is a bounded Root repair review, not a fourth F279 implementation peer or a complete provider review.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Source identity

HEAD is `a5765aa00709d88604093bfe9d46ba62d5af397a`. Sampled full-file SHA-256 values are:

- `crates/rdocx/src/comparison.rs`: `3b875a8aff46669e41a5759eecb26b73d4c69894c2e7072a6038688763a9f6ca`.
- `crates/rdocx/tests/regression_test.rs`: `ca9aa87450ff5fa5c9b550b5846a8ead3d472fed927dea3f04f93daeb4200693`.

Other working changes in those files are outside this review. The previous speculative granular-path signature and unsplit-Text changes have been reverted. Their earlier source review does not establish that the previously failing public default-Run case was repaired. Root retained that actual failed attempt.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Examined behavior

The actual early equality caller at `crates/rdocx/src/comparison.rs:13581` now consumes the added content facts at line 13288. For the non-attributed Run path, unprotected XML whitespace at a physical Text boundary contributes its scalar position. Protected boundaries contribute nothing. This distinguishes actual protected-versus-unprotected leading and trailing content while allowing valid run partitions whose scalar content, effective edge protection and formatting agree. Byte offsets are used only to identify each Text's own edge interval. Accumulated positions count Unicode scalars, including the accented-text discriminator.

The new contribution uses the same four XML whitespace characters as the existing `has_edge_whitespace` policy at line 21153. It does not introduce Unicode-whitespace trimming. `uses_attributed_run_path` at line 13844 preserves the existing ignored-formatting, ignored-whitespace, Field, comment and text-box routing. The attributed signature contribution is empty under those policies, preserving their existing boundary-protection behavior.

The existing nonplain-run equality at line 13491 still compares complete barrier runs. The physical Field and control source checks immediately below it remain in place. The new numeric facts do not make raw aliases, opaque children, Fields or changed drawing owners eligible for publication or infer source identity. The shared signature remains an equality predicate. F279's approved paragraph coalescing and preserved larger-run behavior remain required by `.claude/plans/F-X279b-design.md:256`.

The six controls at `crates/rdocx/src/comparison.rs:30075` distinguish both leading and trailing unprotected edges, protected internal splits, a newly unprotected internal split, and byte-versus-scalar positioning. Each also checks the attributed ignored-formatting behavior. The original facade edge-space matrix remains unchanged and is still the decisive end-to-end regression.

The unrelated-legacy comment fixture now captures the actual authored paragraph ID, removes that exact retained attribute only inside the qualified selected owner's body, and reparses the resulting XML before the mutation at `crates/rdocx/tests/regression_test.rs:413972`. Its new prerequisite assertion proves actual ID absence. Post-mutation byte preservation and ID-absence assertions remain intact. This corrects fixture construction rather than weakening the publisher invariant.

Correctness, contract, introduced panic surfaces, OOXML whitespace and namespace handling, test discrimination and structural constraints produced no findings within this scope.

## Evidence and exclusions

No Cargo or runtime test ran in this review. Root's next combined batch, including the original default-Run public regression and the new six controls, is pending at this boundary. The prior actual failure remains retained. The separate text-box scan change, comment publisher implementation, all-sequence paint investigation, native oracles and final union gates are excluded. No complete F279, F272 or sprint acceptance is claimed.
