# S96 reply shared source, support pass 1

**Reviewed**: The frozen five-path reply increment over the recorded before files, not the entire concurrent sprint diff. Scope includes Native endpoint and refreshed-owner source sharing, regression tests and the existing F268/HLD03/HLD10 amendments. This independent microscope support pass does not reset feature review caps.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Authenticated boundary

/private/tmp/rdocx-s96-reply-shared-source-implementation-1/SOURCE_BOUNDARY.json has SHA256 6138cf32a05f62d3897acdeb2b20065b1c90290e47d428d7f3d16c98fce738be. The increment patch matches 743dcaf623a0e25c4b2918f2d85d6684a65820698a048fd3541786e3027c7eab. All five actual source pins, nine protected source/review pins, both retained binaries and all 13 recorded logs match. The final source hashes are:

- Native document.rs, a19c8dbaca0b15e65f882439633d1b4e7d17b97aa16699d2b45fd4dde3438e38
- Native comments.rs, 422fedcd7276b2f51bec83dca7445f79737351e6fdfe92741c916d8afbe6af17
- F268 plan, 33ff48896a750041ca659093b14267537c4684c0f64198f6da6b644a215e0f4a
- HLD03, 9a45191cde3ef1c100810d394e5e6e42ea4025a3a2d985459c623f9ac5042c61
- HLD10, d77d835ab14b1c62c2d0b28dc7314dc1eef207fab7a5bde54c12b9b43fa5ab0c

The operative F268 amendment approves only full-equal endpoint reuse and complete fresh discovery plus selected-owner qualification on the same actual source buffers. It retains all existing public APIs, staging, ID order, Source/resource admission and publication behavior.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Aspects checked

Correctness and contract: crates/rdocx/src/document.rs:30720 first completes the existing start endpoint qualification. Only full ContentLocation equality permits reuse of its actual item. This includes actual story, fingerprint, kind, index path and end state. Distinct endpoints still execute the independent old read. The enclosing story equality and nested-control branch at 30661 remain unchanged. Replacing an unnecessary same-paragraph CT_P clone with None at 30757 preserves both accepted-run boundary checks and the one-paragraph versus two-paragraph anchoring branches. Descending edit order and namespace closure remain unchanged.

The new crate-private refreshed_story_owner_span at document.rs:31516 prepares complete current sources once. It runs the same stories_in_sources implementation used by public stories at 30061 over every source before selecting the full kind/part/owner-index selector. It then uses the same story_source_owner_and_layout_in qualifier on those actual source buffers, including current fingerprint verification. It does not stop discovery at the first matching owner, trust the previous fingerprint, or grant a new Source/input allowance. The existing main-layout fallback and complete related discovery remain authoritative.

The actual consumer at crates/rdocx/src/comments.rs:2050 retains parent preparation, ID reservation and anchor_story_range before the refresh. Marker selection still uses the refreshed physical owner span, and actual parent/reply adjacency repair remains unchanged. Both outer transaction staging boundaries and comment graph publication remain. No owned public reply bridge is introduced.

Panics and errors: the one new expect at document.rs:30790 is guarded by the same full-span inequality that created Some(last), with no intervening mutation. The fresh source vector is produced by spanned_story_sources, whose first entry is always the main body. The shared discovery helper retains existing fallible full-owner parsing and fingerprint checks. Missing owner still becomes the same no-range error at the existing comments consumer. Errors are not swallowed or converted into a second attempted read.

OOXML: raw fragment extraction, qualified namespace closure, accepted paragraph parsing, anchored marker/reference payloads, descending span replacement and final source-part writing remain in their old implementations. Actual full package comparisons cover the effect of using one actual source rather than regenerating it. The increment introduces no parser grammar, root namespace policy, source normalization, serialized attribute or child-order change.

Structure: the helper performs substantive complete discovery and owner qualification, rather than forwarding to an old method. It introduces no public API, new type, generic, trait, module, crate, flag or retained cache. Test-only counters and forced legacy routes remain under cfg(test). They instrument real source-producer work and do not substitute a timing result.

## Evidence and test discrimination

The retained actual predecessor RED performs successful reply creation, exact package equality, anchor equality and reopen checks before failing the real source preparation observation [2,2] versus [1,1]. It is not an absent-API or setup failure. The retained RED binary hash is 1ead3c8c3da1c0dfe7d2eaa18bb0394dc2dd13fe2f493cbb53401a0e6d182872.

The seven final selected reply tests pass. The same-endpoint test at comments.rs:4925 requires [1,1] with complete old/new package and anchor equality. The distinct-endpoint test retains [2,1]. The existing four-mode differential at 5081 compares body/header/aliased current sources, unsaved paragraph state, simple and complex Field caches, raw tracked content, old second-flush behavior and final full package/graph equality. Atomic malformed unrelated-header, unknown parent and nonmetadata reference-payload refusals remain covered. Complete discovery is independently challenged by an actual referenced malformed header, preventing a selected-owner-only shortcut from passing.

The earlier full comments invocation remains recorded as exit101 with 41 passed, one fixture failure and one ignored native case. The failed fixture did reference the actual header, but its truncated-EOF form was accepted by legacy stories(). The corrected frozen test uses an actual unterminated attribute. The corrected selected successor is the seven-test actual exit0, not a relabelled full comments pass. The fixture-edit failure is also retained. Three additional affected filters pass, covering the Cell source tests, story-range scan differential and cross-owner bookmark identity/refusal. Default check and strict Native/binding all-target/all-feature lint plus format/prose/adapter/diff gates have actual exit0 receipts. The final retained Native binary is 79319b0d8e9215780ba6ac6bb9980c43840ac7a15be820b18c067801a0c0a314. Existing executions were authenticated without rerunning them in review.

## Limits and handback

The bounded source increment has no blocking finding. Root final remains separate. No timing or shipping-artifact improvement is proved by this review. Actual deciding reply timing, installed binding and full integrated suites, native, memory, archive and hosted final riders remain pending. The protected F273 frame and Native Cell work remains unchanged. This pass modified only this review file and ran no Cargo, tests, product imports, renders, measurements or native UI. No fix, status, ledger, baseline or review-cap change is made.
