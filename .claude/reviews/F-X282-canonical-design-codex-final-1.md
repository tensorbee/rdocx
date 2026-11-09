# F-X282 canonical design, Root targeted final1

REPAIR REQUIRED BEFORE APPROVAL. Two contract-closure defects, zero smells.
This is a targeted Root preapproval review of the direct-cell bridge, not a
fourth design peer, whole-plan approval or implementation review. Three design
peers remain used and zero remain. Other accepted preparation dispositions
stay intact. No product source changed.

Canonical draft SHA-256 `40971bc22036a8b47b92e8ead4fe0ff9fb4b12d9a96d2f9bfa57a5702ebda920`.
Current source traces use the preserved dirty canonical tree at
HEAD5820539cd277b417ae060e15010eb73bd1d00b01. Actual implementation-Base
revalidation remains mandatory. Root's bounded native probe used the existing
library SHA-256 `5e254d7051a256a327a80dc9ac53c3e85e4a6f496b236fd74201482d7a05724c`,
compiled0 and ran0. Its source, raw results and manifest are under
`/private/tmp/rdocx-f282-cache-atomicity-probe`. It tests existing APIs and a
counterfactual acquisition sequence, not the unimplemented new bridge.

## C282-1, mutable access invalidates before a failed candidate

`.claude/plans/F-X282-design.md:129` calls `table_mut` before the guard and
fallible closure. The claimed barrier at line140 preserves the paragraph but
not the existing cached observation. `crates/rdocx/src/document.rs:34755`
unconditionally calls `invalidate_layout`, whose implementation at line24947
clears both layout caches and the checked-body refusal memo. Therefore a
closure error occurs after those caches already changed. The existing checked
editor returns at its closure error before publication at
`crates/rdocx/src/document.rs:32871`.

Concrete probe: prime the deterministic layout Arc and confirm a repeated read
shares it. A checked cell paragraph closure that sets a temporary style then
returns a synthetic error leaves the original style and same layout Arc.
Merely obtaining `table_mut(0)` without changing any paragraph gives a different
layout Arc on the next read. See
`/private/tmp/rdocx-f282-cache-atomicity-probe/results.txt:1` through line3.
The proposed order therefore does not mirror the checked failure barrier.

Resolve and qualify the source without mutable access that invalidates live
state. Specify the callable order for candidate validation, stable target
reacquisition and successful publication. Cache and refusal-memo changes must
follow successful preparation. Add a failed-closure cache discriminator and a
valid successful edit control. Do not redefine failure atomicity or perform a
whole-document clone merely to hide the early mutation.

## C282-2, successful closure is not successful checked publication

`.claude/plans/F-X282-design.md:139` assigns the candidate immediately after
closure `Ok`, while line166 promises not to skip paragraph and run validation.
The existing checked editor still constructs a checked fragment and calls
`replace_story_content` after the closure at
`crates/rdocx/src/document.rs:32874` and line32881. The draft names no equivalent
candidate admission before the direct assignment. Its public hidden generic
bridge accepts an arbitrary Paragraph closure, not just a prevalidated enum.

Concrete existing behavior: a closure sets style id `Bad` plus U+0000 plus
`Style` and returns `Ok`. The checked editor returns an OPC XML1.0 character
error. An ordinary `Normal` style succeeds. See
`/private/tmp/rdocx-f282-cache-atomicity-probe/results.txt:4` and line5.
The paragraph setter itself accepts the string at
`crates/rdocx/src/paragraph.rs:1360`. Candidate clone plus closure success is
therefore not an adequate validation proof for the new direct route.

Name concrete local admission and publication calls, their applicability
predicate and an equivalent checked fallback when the candidate cannot be
qualified. Preserve closure single-execution and its two existing return types.
The negative character case and compatible ordinary-style case must reach the
actual new bridge. Preserve existing source, namespace, resource, relationship,
style and paragraph checks. A statement that validation remains is insufficient.
No new public refusal policy or narrowed closure API is approved here.

## Limits and next phase

These are design defects, not assertions that an unimplemented bridge already
corrupts product output. Prior reviews accepted the clone barrier without
following this access side effect and post-closure publication. That is a
cross-review coverage miss, including Root, rather than solely an author fault.
Root's separate remediation phase will assign the correction without another
design peer. All original performance criteria, no staging or reopen promises,
F26879/252/407 checks, saved bytes, hash, installed bindings, schema, native,
packaging, WASM and CI riders remain. Do not start implementation. End review
and return to orchestration.
