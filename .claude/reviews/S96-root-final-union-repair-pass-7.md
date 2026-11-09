# S96 integrated run-view and layout repair, Root pass7

Defects: 0. Smells: 0. Nitpicks: 0 in this bounded source repair. This is not
a full sprint review, feature completion or release acceptance.

## Source and independent support

The reviewed source is the working successor of H7
031d8e6eba6bf9ac2bd591ca7024dbb39db819be. Engine source is
c54f572a9ec62a5282f3c111425a978b8a31c72b12b69dfae1735ac7cc82558e,
and paginator source is
c4bd208658ae4303863c662bf2b4385360665bac74bf1c072ae5327d93df30d8.
The actual separate cross review is S96-H8-layout-cross-review-support-1.md.
Its author excludes its own SEQ and fixture proposals. Root separately
reviewed those proposals, both C16 getter repairs and all integration hunks.
No reviewer is impersonated and no feature peer counter is reset.

Run.font at crates/rdocx-py/src/run.rs:510 and Paragraph.runs at
crates/rdocx-py/src/paragraph.rs:991 retain typed and scope validation before
using the existing generation-indexed content-view capture. That capture
uses the same checked physical owner and removal, content, ancestor and
furniture scopes. It does not refresh story XML solely to capture a view.
Related typed reads and the legacy owner-free path retain their behavior.

At crates/rdocx-layout/src/engine.rs:4263 and :4334, the two checked entry
points prepare one layout-only revision view before source registration.
Borrowed model admission precedes recursion or cloning. The actual retained
reader registry reaches the existing checked revision factory. REF and SEQ
selection, provenance and rendering consume that prepared owner. Raw XML
remains authoritative, the caller is unchanged and failures propagate.
Revision-free input stays borrowed, and snapshot derivation reuses an owned
preparation clone. No public API, parser, allowance or budget is introduced.

The stored-display route retains the actual normalized Field owner before
replacing it with its cache runs. Its first cache run now emits the same
instruction-event markers as the ordinary Field route, before visible cache
text. The six-argument concrete helper has those two actual consumers.
Source ordinals and the existing snapshot remain the authority. The actual
RED diagnostic proved that the IF owner was previously absent at that point
while its14and15snapshot events were present. All temporary prints are removed.
See crates/rdocx-layout/src/engine.rs:12041 and :13194.

Generated stored Field text remains excluded from automatic prose hyphenation.
The neutral wrapping fallback applies only to pre-laid Body blocks without
Word paint. Word and furniture producers keep the missing-reflow refusal.
Synthetic drawing and field fixtures now carry their actual producer paint.
Multilingual tests inspect the original base metadata and preserve formatting,
source ranges, ordering, visible text, geometry and strict paint checks. Cache
staging retains both its byte and entry ceilings and actual FIFO behavior.

## Actual evidence

The first source repair fixed held font access but left fresh run collection
access expensive. Its exact counts stayed RED and both extended lifecycle
cases failed. The subsequent three-file C16 producer records actual build,
install and payload authentication success, all11original counter batches
with zero body scans, and both exact lifecycle cases passing. The count
artifact is0b85371d9fb05f83c8d28318d98c3d783e292e056780b88efe6e77219660a3cc.
Its clocks are instrumented diagnostics, not ordinary performance acceptance.

The actual complete layout library passes419cases with zero failures and
one unchanged ignored case. The prior15-failure hosted observation and
416-pass/3-failure local observation remain historical failures. The actual
current49-entry hash check is unchanged. Full all-target, all-feature
workspace Clippy passes with warnings denied. These observations retain
their original dirty-source identities and are not fabricated clean builds.

The changed layout archive was actually measured before metadata adoption:
461721compressed bytes,2530504normalized member bytes and15members. Only its
owned README row and script tuple change. Compression tolerance, all other
archive rows, speed claims and numeric gates are unchanged. Post-adoption
measurement and current verified package dry runs remain separate checks.

## Remaining acceptance

New exact-source normal CI, current shipping and private artifacts, installed
functional and typing qualification, quiet timing, memory and native riders
remain due where unperformed. H7 successes keep their exact source identity.
The same-owner physical-reference native exception remains unproved. Earlier
artifacts, full-suite passes and producer signatures are not relabeled as
current acceptance. No completion ledger, main merge, tag, reporter comment
or registry action is supplied by this review.
