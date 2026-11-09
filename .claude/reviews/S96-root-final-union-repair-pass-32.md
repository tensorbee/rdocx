# S96, Root final repair review, pass32

Reviewed the frozen26/28/29namespace wave and the separately approved Q-S96-16
release carry on sprint/s96 at cec87f9501eb13d0d5d60572472b227483262620.
Actual independent Support14, SHA-256814fc9cabdb2282d64240622919daf02b4307811eceb8123e22be813bfcc4423, was consumed.
This is Root's own separate review, not another independent peer, a full sprint
review, completed release gate or publication approval.

## Defects

None.

## Smells

None.

## Nitpicks

None outstanding. The separate carry-review backlog table spacing is corrected.

## Source and contract assessment

Reauthenticated all4309current inputs against the final freeze. Exactly14tracked
paths differ from reviewedRoot31, with exactly2runtime Rust files. Namespace
patch3f1ae2172bdbaa631c00cbe323974980dd3042dd193d7368903c652f627bcc89
contains4paths/241lines. Freeze952ca7cd6a445d659552dcc712b3d889d76d26da746093bace56d62c5cafedaf
also binds the accepted release/carry metadata. Read the actual implementations,
adopted contracts, investigations26/28/29, ScopeCarryReview1, HLD04/12/14,
selected test assertions and actual receipts, then completed Support14.

At `crates/rdocx-oxml/src/text.rs:804`, the attribute iterator and exact URI/value
decoder remain before declaration discard. Invalid unused declarations still
refuse before later attributes. Non-namespace values become owned before
existing fallible namespace/name/duplicate checks. Owned decoder results and
retained plan storage remain owned, with no event borrow escaping an iteration.
No actual reader, capture, constructor, Source/Field/resource or emission step
is omitted or granted authority by this storage change.

At `crates/rdocx-oxml/src/numbering.rs:657`, validated prefix borrowing preserves
UTF8-before-URI failure order, exact prefix removal, encoded binding push,
Word comparison/alias push, inherited-Cow transition and returned ownership.
Only Word aliases gain owned strings at their existing push. Foreign prefixes
leave no discarded standalone copy. Pure allocation/drop interleaving changes
without callbacks or recoverable errors. Every scope value and bare alias
retains its original meaning. The separate properties helper is unchanged.

At `crates/rdocx-oxml/src/numbering.rs:628`, checked byte sizing creates exactly
NUL/prefix/NUL/URI bytes. Original formatting handles unavailable sizing.
No escaping, delimiter repair, char-count sizing or new Result/error occurs.
All five production callers keep their own validation/dedup/projection order.
The new literal-format/split test is independent of the shared encoder used by
the legacy scope matrix. Retained capacities for arbitrary short inputs may
change and are disclosed. No new helper/type/trait/generic/module/file/cache,
unsafe code, stable API, namespace language or Source allowance is introduced.

## Actual evidence and its limits

The three immediate predecessor/successor mechanisms each retain6original
public states,18pairs total, with exact input/full ZIP/member/staging/semantic/
reopen identity and complete decode/plan/scope/value/capacity observations.
The13Root component outcomes include decoder-Owned namespace3/32bytes,
retained rather than credited as removed work. Each artifact has actual build0,
4309restoration and10payload authentication. The initial generated bare-CR probe
remains failed, not a product or semantic pass.29four-append cardinality is
accurately labeled source bookkeeping, not independently measured copies.
Actual initial/final capacity equality on append-only buffers proves no logical
growth, without claiming allocator-call counts or hidden old formatter growth.

Final actual scoped evidence is23unique Rust PASS/0IGNORE,3crateall-target
check/test compilation/strict all-target/all-featureClippy0, and cheap gates0.
The ordinary debug wheelcdc576b0dc0c2bbc9c8fd53c4cd64b0cdf7f88ac4c797e2d41d01d31436c0184
and Native69877540f3ca5ff7f51d5d0f1d520ac24c4618216343a2cc97d857327d6c4f03
match all10actual installed members, with temporary hooks absent. Final focused
JUnit contains42PASS/0skip across namespace/Field/held/refusal/splice/picture/
reopen controls. Existing unrelated suites are not falsely counted as reruns.

Q-S96-16is fresh, actually user-approved and explicitly operative. Current
records, F275/F268/F282/F283plans, HLD12/14and release notes carry numerical
performance and unfinished memory/RSS/history/calibration to S96.1. F-X284review
precedes F-X285performance andF-X286memory work, all pending. S97and later IDs
are preserved. Nine actual affected release/capability/roadmap policy checks
pass. Published notes remain unchanged. No current capability owner is falsely
completed, Issues437/445remain open, and historical F-X268is not reset.

## Retained release gates and handback

Latest76/79before-final-wave timing remains a failed snapshot, never a new final
artifact guarantee. No further numerical pause/RSS qualification gates are
introduced before this approved release. Correctness, held/last-loan/resource/
Source/refusal/security, actual applicable native/oracle, installed runtime/
typing, hash/pixels, packages/archives/SemVer/hosted-source proof and final exact
main-SHAapproval remain. Full workspace39315is still active, not a passing full
union. Current feature completion, source availability, sprint close and registry
publication are separate later facts.

Zero findings in correctness, contract, panics, OOXML preservation, tests and
structure. This review writes only this report, changes no product/plan/progress
or tracker, and executes no runtime/Cargo/test/timing/Git mutation/publication.
The pass ends and returns control to release preparation.
