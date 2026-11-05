# F-X272 format and ordered increment, Root final pass 1

Reviewed the complete1755line six-path frozen increment against approved
parent plan7d21dfc4f68483d7a11715e2f467151c3c83239d451ac461f10bba94c37d6845.
Verdict: one defect, zero smells and zero nitpicks. No new peer allowance.
Manifest SHA256 is
`af9523303120dfde52f5d99230125fa320b1baa6b1ad4c01e44d9ff4c2017871`.
Final text SHA256 is
`279787d618151c77ea508edffacad540dd2801ae2a23a13ebad542469bec9d38`.

## Defects

### D1, requested Text authority crosses an unproved nested reparse

`crates/rdocx-oxml/src/text.rs:3101` in the frozen final snapshot resolves a
produced Text fact by the old AcceptedRunPath and content ordinal, then installs
its origin at3109. The fact was recorded before replace_ordered_result_slot.
For Revision/ContentControl paths that operation and the subsequent
normalize_field_owners_in at6173 may refresh a nested revision. The actual
revision refresh in `crates/rdocx-oxml/src/revision.rs:1020` serializes/reparses
and retains Field identities, without a returned Text correspondence transcript.
The lease is not restricted to direct stable paths, and its later length check
cannot establish that the newly found Text is the actual assigned occurrence.

An ordered result containing requested Text under an accepted revision reaches
this unproved boundary. This violates the required actual producer/clone/parser
relationship for input authority. It is not a demonstrated wrong-output runtime
failure. The author confirmed the missing proof from the frozen source and made
no remediation during this review. Existing six direct/hyperlink controls do
not discharge the nested case.

Repair must forward the actual nested producer/parser correspondence, or contain
this partial increment to the positively proved direct Run paths and direct
no-slot insertion while leaving nested coverage explicitly unfinished. Do not
infer authority from equal text, length, token or a reused path. Do not add a
new input refusal or claim whole ordered ownership from containment. Retain a
different nested revision/control discriminator beside direct positive cases.

## Evidence and other aspects

Root authenticated12changed source snapshots, all51final pins, the complete
patch, binary35c47def7eeb28df and142unique raw named passes. Seven scoped gates
are green. The JSON dedup omission has27authenticated mutant source pins,
actual binary/restored-source identities and one changed membership guard.
The actual test fails exit101on repeated same-operation alias input state.
This is a controlled omission, not historical product RED.

The real format_paragraph shares the substantive prepared kernel, preserves
occurrence indices through sorting, validates dependencies and callback
exclusivity before activation and clone, and deduplicates actual Values without
Template64. Final mapped transport has actual Vec/Normalized instantiations,
exact live writer/ledger matching, nested full demand and lexical restoration.
Its first zero-budget negative had no active Field metric and old Vec also
accepted. That remains a fixture error. The paired repaired positive/negative,
ordinary compilation fixes and initial Clippy failure remain separately recorded.

No additional correctness, panic, XML, test or structural finding was established
within this frozen partial scope. The four absent operations, nested/typed and
ordinary accepted Text coverage, committed Field publication ranges, Native
wiring, deferred/current/constructor closure and all final riders remain open.
Root modified only this review record and returns to separate remediation.
