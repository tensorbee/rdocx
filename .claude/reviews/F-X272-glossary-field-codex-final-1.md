# F-X272 glossary and prepared Field increment, Root final pass 1

Reviewed the frozen nine-path1850line incremental patch under canonical
contract `a282cf2cbcbb28b33c340ba9073d845683c11aef86786c1cc78bd8445f58065d`.
Manifest SHA256 is
`e3d083b81ba1d4470c0885ae244843651fc9b9b137aaceead5f0469d46fd5172`.
Patch SHA256 is
`c82cdbe3a841c1fed3d414eade28b10ddde5671424bc716530832befe123f901`.
Verdict: zero defects, one smell and zero nitpicks. This bounded Root final
does not reset the exhausted full-feature peer cap.

## Defects

None found in the bounded implemented scope.

## Smells

### S1, duplicated glossary edit assembly policy

`crates/rdocx-oxml/src/glossary.rs:2058`
`crates/rdocx-oxml/src/glossary.rs:2116`
`crates/rdocx-oxml/src/glossary.rs:389`

The new apply_glossary_edits duplicates the existing structural edit algorithm:
sorting, capacity arithmetic, overlap validation, ordered gap copies and the
terminal tail copy. The old implementation remains on the real from_body_xml
constructor while glossary and doc-part publication use the new implementation,
including their ordinary untracked path. A future repair now needs two policy
edits to keep equivalent consumers consistent. Route the remaining constructor
through the same substantive edit engine and remove the redundant algorithm.
Preserve current byte results and existing validation/error order. No new
generic, trait, quota, validation policy or test carrier is needed.

## Nitpicks

None.

## Other aspects checked

Correctness, contract, panic handling, XML preservation/order, tests and receipt
ownership yield no additional bounded finding. The four getters borrow actual
stored buffers. SDT lookup uses absolute stored content indices and does not
override publication eligibility. Glossary edits preserve retained gaps,
pending insertion/deletion ordering and typed body event offsets. Immutable
invocation catalogs are shared across independent record sinks.

The ordinary clone follows the same substantive clone producer without a
receipt registry. Private Weak/occurrence identities retain no source or
model, and Field equality excludes them. Preparation preserves the raw bypass
and child-before-owner normalization order. Scalar events follow actual
FieldResultParts slices, exclude tabs/breaks and do not select typed caches as
parent scalar input. Direct changed-cache content and wrong-frozen live
receipts cannot acquire the old scalar role. Foreign/dead receipts confer no
authority while a real new invocation pair can still prove its own input.

Root inspected18scoped controls and49distinct existing results, with seven
unchanged-source scoped gates. Independent authentication verified both
nine-path snapshots,51retained source/config pins, the exact binary and full
nine-file patch reconstruction. The controlled guard omission failed its
actual wrong-role assertion. It is not historical product RED. Earlier fixture
mistakes, the introduced foreign/dead forwarding repair and the first Clippy
failure retain their separate labels and artifacts. No tests were rerun merely
to confirm the already applicable frozen results.

Parsed changed-cache transport, ordered requested/new-Text/reparse lineage,
nested accepted companions, five missing approved methods, actual Native
consumer wiring and deferred/constructor/current pre-growth closure remain
outside this bounded acceptance. Fact-copy and whole-operation qualification
remain mandatory. No provider, complete API family, native/installed oracle,
full sprint or release readiness is established. Root changed only this
review record during the pass. Return to separate remediation for S1.

## Consumer-name correction

Root's S1 prose and initial handoff incorrectly named from_body_xml as the
remaining old-engine caller. The cited glossary.rs389call is actually inside
replace_body_content_xml, starting at383. That mutable replacement API is the
consumer to route through the shared engine. from_body_xml is not the affected
caller. The duplicate-engine finding, cited code and remedy remain unchanged.
The author correctly identified the actual method before applying the cleanup.
