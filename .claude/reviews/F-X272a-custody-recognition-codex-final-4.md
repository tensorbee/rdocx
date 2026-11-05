# F-X272a custody recognition, Root bounded final pass 4

Reviewed the 217-line v6 patch across two files, the inherited resolver paths,
canonical child contract748766 and the actual source-pinned evidence.
Verdict: zero defects, zero smells, zero nitpicks in this remediation.
D1 from Root final3 is closed locally. This is permission to prepare the
sprint-final-acceptance handoff, not final feature or native acceptance.

Source manifest SHA256 is
`acdf55a4507ce9f9353bc8d36053dbcc02f638ec0c3385a5c16494c3f72d7bec`.
Delta SHA256 is
`03a2134d560b74742a06aed146763cb299791a10f33af4f38703492efbf15ee5`.
Custody SHA256 is
`d2498decb09c4de32e880a9e6a00b82a6dc4ba9c3d314a170d4334e453d5ba9a`.
Citations refer to that frozen repository-relative source.

## Defects

None in the bounded remediation. The private actual note admission vector is
retained after qualification at `crates/rdocx/src/revision.rs:1454` and governs
semantic dispatch at `crates/rdocx/src/revision.rs:2459`, scalar conversion at
`crates/rdocx/src/revision.rs:2572`, shell cleanup and paragraph merging. Clearing
revision metadata is no longer the sole barrier against opaque payload changes.

The scalar exception at `crates/rdocx/src/revision.rs:1474` requires a recognized
Word scalar leaf directly under an admitted actual run. It does not transmit
semantic authority through an unknown or foreign ancestor. Existing raw scalar
representation therefore does not disable real deleted-text and instruction
restoration. The footnotes envelope and opaque Drawing/VML ancestors remain
traversable. Actual child-story roots can restart admission, rather than being
lost through a blanket raw-subtree shortcut. Namespace relocation still runs
independently of semantic writes.

## Smells

None in this delta. No public API, new module, trait, generic or model field is
introduced. The private tree-local optional vector keeps non-footnote callers
on their inherited path. Generic parsing, reset and snapshot reading are
unchanged. Property-prior reparsing only locates retained markup and does not
start a second unqualified resolution pass.

## Nitpicks

None.

## Evidence and exclusions

The new regression at `crates/rdocx/tests/integration_test.rs:38866` checks both
opaque literal delText and delInstrText through real note removal, accept,
reject and reopen. Plain text and complex-field instructions provide genuine
conversion positives. Canonical and aliased Word namespaces with local foreign
w bindings exercise expanded identity. Restored public note content and opaque
bytes are asserted. The actual old-v5 binary fails in the first plain case at
the reject-to-original guard. No historical outcome for its unexecuted variants
is inferred.

Mechanical authentication independently joins all57custody members,12source
snapshots, exact v5 delta and cumulative actual claim Base. Native document.rs
is unchanged from its returned finite lease. The actual final source has
56integration cases with zero exclusions or ignores,5unit,1regression and1CLI
passing. Six source-stable receipt groups support these executions, prose0,
all-targets/all-features scoped Clippy0 and workspace format0. No runtime result
is relabelled from a later lint-only source. Existing real textbox/table/control
selector and count positives and all Q29 cases remain in the batch.

The existing progress record maps all45A430criteria and14G430gates. Its earlier
held-Q29 entries are historical and superseded by its explicit later operative
acknowledgment. Q29 remains accepted locally within the approved missing-N
corruption-detection limit. No following-reference owner inference is restored.
The separate source trace does not establish a snapshot-text leak, so this
review does not invent one from the renderer defect.

Final integrated composition with the parent carry, installed bindings and held
handles, schema/SDK, actual Word-admitted native outputs, hash and full workspace
riders remain due. Q28's exact shared-reference native positive stays UNPROVED
under the approved one-release exception. No provider dependency is unblocked
by this review alone. Root wrote only this review record during the pass and
returns to orchestration for handoff preparation.
