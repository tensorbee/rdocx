# S96 bounded comment, tab and run repair review

Independent source review only. This is not another whole-feature peer pass or a review-cap reset. No product source, tests, progress or plans were changed. No Cargo, product execution or performance measurement ran. Root owns the actual affected gates and the later full microscope.

## Scope and source

Reviewed the working-tree production changes in:

- `crates/rdocx-layout/src/paginator.rs`, SHA256 `92a42a9286231d3b2b147728dec7636f535340800e26213d22683d040d0544e2`.
- `crates/rdocx/src/comments.rs`, SHA256 `cfa1a5f63d35b016d2865c06d9c25f771bab74226e11a19975ac66fdbf0d5fc2`.
- `crates/rdocx/src/comparison.rs`, SHA256 `aab5bee8595dd71e8a09ce77f9d724068a71182133ca69a668accc8c75720990`.

Read CLAUDE.md, WORKFLOW.md, microscope instructions, the operative F-X279d repair and A1 scope, F-X279b P1/P2 and comparison preservation contract, and the affected source callers. The separate semantic revision candidate and fixture changes are outside this pass.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Boundaries checked

The tab metadata repair admits the exact legacy MultilingualText base already used by the real paint branch. The rich carrier remains excluded from this legacy metadata inference. It retains the existing field-source identity, segment boundaries, decimal policy and later field-width realignment. Source references: paginator.rs5767-5788,5940-5974,6055 and engine.rs10022,10172. Existing centred PAGE, TOC page-number and STYLEREF position assertions remain meaningful discriminators with unchanged expected positions.

The comment qualifier first validates definition/thread/marker ownership. It then applies the unchanged generic pairing to every noncomment marker before allowing independent comment pairing. Its stack token names the exact physical noncomment frame, rather than depth. Comments enclosing a complete nested bookmark remain valid. Crossing one bookmark boundary refuses. Leaving one frame and entering another at the same depth also refuses because the token differs. Source references: comments.rs537-593 and the unchanged pair_located_story_markers at399-532. Current tests inspect mixed, unmatched and duplicate-reference atomic refusal alongside independent crossing replies in separate story owners. No new typed interpretation of opaque XML was introduced.

The exact move no-op runs checked_movable_comment first, so unknown IDs, replies, incomplete source counts, ambiguous or cross-owner selected anchors are not bypassed. It requires equality with the actual current anchor and resolves the endpoint snapshots before returning. A changed destination still follows the existing staged move and destination preflight. No source rewrite, identifier allocation or commit occurs in the exact branch. Source references: comments.rs824-835,854-907 and document.rs30742-30763,31660-31689. This is an unchanged existing range, not permission to place a new unsupported range.

The unsplit Run comparison branch now preserves the complete CT_Text value, including its original space flag, when whitespace is not ignored. Word and Character splitting retain their existing edge-whitespace normalization. Run with ignore_whitespace retains the prior whitespace-fragment path. Existing whole-owner signatures and boundary checks still govern unchanged protected owners and run-indexed marker preservation. Source references: comparison.rs16308-16329,13869-13904 and F-X279b's run-granularity/preservation contract. No threshold, ignore flag, field-owner refusal or raw-source publication policy changed in this branch.

## Pending evidence

This report does not claim actual GREEN for the current source, native parity, installed binding behavior, performance or final-union acceptance. Root's affected runtime batch and final source qualification remain due. No claim is made about unrelated changed functions or the remaining semantic projection work.
