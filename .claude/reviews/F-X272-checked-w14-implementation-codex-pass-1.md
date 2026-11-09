# F-X272, checked w14 implementation, Codex pass 1

**Reviewed**: Sealed checked-w14-source increment. Eight paths, 389 additions
and 49 deletions in the complete 568-line patch.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Bounded bridge is clean.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Contract and producer review

Read the repository workflow, microscope requirements, current progress and
canonical7be66404. Its plan SHA256 authenticates as
`f82b700835d68f7ddd052522f3b408b042959c67389e20eefffe1647ef529abb`.
The approved additive method, its two concrete consumers and unchanged metric
are the review boundary. Inspected immutable before/final snapshots and the
existing composition, admission and reader-origin helpers.

The private planner at `crates/rdocx-oxml/src/text.rs:1625` is shared by the
ordinary helper and scoped producer. It retains the existing conservative raw
w14-name selector, first actual Start-root handling, aliases, no-use decision
and already-declared policy, including a foreign w14 URI. Borrowed events avoid
materializing the root header before copy admission. The actual live
reader/input relationship supplies physical root-end coordinates, including
the UTF8 BOM displacement. It does not search for equivalent bytes or synthesize
a source identity.

The ordinary helper at `crates/rdocx-oxml/src/text.rs:1612` consumes that same
plan. Its explicit intentional output correction places a missing declaration
at the physical root end on BOM input. The previous logical offset inserted
three bytes early. This is not a no-output-change refactor. The retained actual
pre-fix discriminator compares prefixed output with the independently assembled
prefix plus correctly declared unprefixed output. The scope of the correction
is documented for separately labelled integration treatment. No other namespace
selection change is inferred.

The additive declare_source_w14_on_root method at
`crates/rdocx-oxml/src/document.rs:6155` validates operation identity and sticky
refusal before planning. A no-insertion result returns the same owned Source,
including its bytes, payload receipt, initial-delivery state and constructor
state. Actual insertion supplies precisely the selected before/after images and
generated declaration to existing checked composition. Constructor admission,
map/event relocation and applicable copied Field demand therefore precede the
new combined buffer growth. Generated declaration bytes gain no Original,
current, Word or ordinary input role. No payload prepayment or allowance is
renewed by this helper.

source_model_document consumes the actual writer result through the method at
`crates/rdocx-oxml/src/document.rs:5965`. No insertion keeps the existing later
admission, while insertion obtains its constructor facts from actual checked
composition. Field publication and escaped Text events remain attached to the
actual produced source. Native assembly at `crates/rdocx/src/document.rs:9821`
now declares on that Source before obtaining its separately paid legacy record.
The final head/root offsets use the actual produced length delta. The ordinary
None path still invokes the ordinary helper. No mutated legacy Vec is matched
back to a separately serialized source by equality.

The exact authorized template test successor now requires the real w14 main
Source handback. Its assertions retain legacy program/package parity and held
caller model/package bytes. Inspected the existing assembly/replay context.
Target-event replay, later writers and other source families are not silently
qualified by the shared planner.

## Authentication and evidence

Authenticated SOURCE_BOUNDARY.json SHA256
`b9fdef97cc0e8934cd1fde6dd1d62a82df57881dcefe13da9e1a8f6dc7db0b62`
and increment.patch SHA256
`81e51ed0b459a463f77bc37f75a3c263f8212ac6441e40e581731059b14eec31`
under `/private/tmp/rdocx-s96-root-c1-implementation/native-template-support/checked-w14-source-increment`.
All 16 before/final snapshot hashes match. Regenerated the complete unified
patch in memory and matched exact bytes. The late template-test predecessor
matches the accepted first-writer source, with the narrow authorization retained.
All ten final log/exit hashes authenticate. Seven product/HLD gate pins match
the final sealed source.

Authenticated the archived Native executable SHA256
`7fd92c97f40f5d888753392eed529a2495b0d81cfbeae67dd036530b8b73ec25`
and lower executable SHA256
`848fa871d4266bc9012092034174d1ba4caa3e38effe68a294c4c4572d7b29cd`.
Read the actual 50 Native template PASS names, five new scoped w14 PASS names
and 44 distinct affected lower names. Inspected new inputs and assertions,
including no-op buffer/payload identity, existing foreign binding, aliases/BOM,
generated insertion maps, foreign-operation refusal, unchanged admission state
on malformed preparation and sticky zero-input known-Field copy refusal.
The full-document control carries two actual events through selected paragraph
closure and TextWalk, including the Field cache and escaped CR/newline Text.

Authenticated the pre-fix BOM source, executable and raw log. Its actual one-test
failure exits101. This is an observed predecessor-helper defect in this increment,
not a historical claimBase RED. The owned-argument compile correction executed
no tests. The two event-count/value fixture corrections retain their actual
4PASS/1FAIL outcomes without inventing product fixes. Earlier affected-lower
observations are not relabeled as final-source runs.

Both scoped checks, final strict lower/Native all-target/all-feature Clippy and
final runtime cohorts have recorded exit0. Final Clippy completes in24.81s.
Formatting and adapter drift succeed. Initial prose exit1 remains recorded,
with the exact metadata-only correction and explicit successor prose exit0.
Final diff-successor evidence is separate. No review build or test was run.

## Not found and limits

No correctness, contract, panic, OOXML, testing or structural findings were found
in the bounded delta. The36th carrier method is the sole additive public API.
No public type, generic framework, parser replay, namespace policy, metric,
quota or transport exemption is introduced.

The separately executing v6 external client is not claimed by this report.
Final SemVer inventory and the intentionally labelled BOM-output integration
remain distinct obligations. Actual target-event replay, second serialization
and grammar consumer, related/comment writers, private stored graphs, model
cache/clone accounting, parsed-complex publication and the canonical hyperlink
neighbor remain pending as sealed. Whole C1, provider, installed/native and
final integrated acceptance are not declared.

Review ended after writing this canonical report. No product, progress, HLD or
plan was edited and no remediation occurred.
