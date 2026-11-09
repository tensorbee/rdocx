# F-X272 Drawing current maps, Root bounded final pass 1

Reviewed the seven-path820line incremental patch from the accepted publication
model snapshot, including five product files, HLD03 and progress. Parent
contract172904 remains authoritative. Verdict: zero defects, zero smells,
zero nitpicks in this bounded source review. No whole C1 completion follows.

Manifest SHA256 is
`5febb3d38395f12be69bda1af06b6fb8f2cc198e48e0407aad8a94f369323a54`.
Patch SHA256 is
`05dc05566e363f680df6cbb09b77e0b23dfa1adeeb9bdd1f4bd7ec0246512376`.
Citations refer to the frozen final snapshot, not the moving worker.

## Defects

None found in the bounded delta. Actual successful paragraph return gates the
publication commit at `crates/rdocx-oxml/src/document.rs:5710`. The commit now
walks actual retained physical WordOwner children at
`crates/rdocx-oxml/src/text.rs:2936`, rather than accepted-view TextWalk output.
Its actual source/current registry joins remain required for every selected
range. Visiting a retained owner does not itself prove its source role.

Run readers pass the real cursor through paragraph, hyperlink, SDT, revision and
Simple cache paths. Drawing Anchor capture uses the actual reader occurrence,
then the existing namespace insertion transcript and actual textbox body parser.
`crates/rdocx-oxml/src/document.rs:2551` retains initial_delivery from the input
itself and checks both map lengths. It does not infer first delivery from an
Original byte piece. The same registry and existing resource receipt remain.
Unproved capture maps remain unknown. Generated namespace bytes do not become
Original or create another allowance.

The body-level SDT producer at `crates/rdocx-oxml/src/document.rs:10834` forwards
the same capture map into the existing mapped factory. Its child cursor retains
the parent's delivery role. No byte search discovers origins. Known occurrence
ranges and actual unchanged-capture checks qualify the transport.

The outer parser commits only after complete success. Late drawing metadata
errors return before commit, and traversal of the returned model excludes
replaced Anchor candidates. Existing inline and alternate/VML paths do not gain
unproved completeness merely because this Anchor path is connected.

## Smells

None in the reviewed increment. Existing public reader behavior is retained
through its compatibility entrypoints. The private mapped readers perform the
same substantive parse with additional actual input context. No new trait,
module, generic, public API or resource equation is introduced.

## Nitpicks

None.

## Evidence and exclusions

The new controls distinguish actual initial versus derived input, five run
families, aliased/default Word namespaces and a foreign canonical w binding.
They inspect returned physical fields and exact publication ranges, require
outside opaque XML exclusion and assert no copy-work allowance is renewed.
Textbox table, body SDT, inline SDT and following-owner coverage distinguish a
lost nested capture map. Late invalid docPr and overwritten bodies test actual
success/retention boundaries. Deletion and move controls retain the ordinary
flat raw-child negative beside real Ordered-mode positive cases, without
changing the selector merely to make a fixture pass.

The packet reports18focused and113disjoint affected cases,131unique passes,
seven scoped gates and final metadata checks. Actual pre-batch product plus
only four new tests gives13passes/four failures. Controlled omission of the
body map gives17passes/one failure and is labelled counterfactual, not an
historical defect. Failed fixture premises and author compile errors remain
separate. Mechanical authentication joins all51source/config pins, copied binaries,
exact seven-path reconstruction and all131named outcomes. Its AUTH SHA256 is
`1b12db2e5f61344a07d0478c44f75ded64f445eb55477e4bf164943e46f95d51`.
The pre-batch RED reconstructs five actual prior product files plus exactly
four inserted tests. This is not a wholly historical repository execution.
Both restorations match their own saved states. They are not relabelled final
source: the RED restoration predates later source changes, and the omission
restoration differs from final document.rs by the retained formatting/test
layout diff. The actual final131passes and gates have their own stable source
pins. No stronger full environment or verbose rustc envelope is invented.

This review accepts the source mechanism only. Lexically changed capture maps,
alternate/VML entry contexts, changed Complex recipes, component transitions,
whole-root completeness, four remaining operations, Native wiring, nested Text
refresh correspondence and deferred/current/constructor obligations remain
open. The integrated child and final union riders are not qualified by this
lower increment. No new quota, first-input grant, fieldlessness inference,
provider completion or finalH claim follows. Root changed only this review
record during the pass and returns to orchestration.
