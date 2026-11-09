# F-X277b implementation, Root final pass 2

Reviewed the pending-normal-id repair against pass1 and the unchanged approved
A277-2/3/6 contract. Patch SHA256 is
`df9e8d2608a383b15720aea9cfa7ab78cfa6b1df8e9bcd9746bcb6b51c2322d6`.
Final Native document SHA256 is
`5d9c8be9241d2806f8c57904603785db9cbf43558cf6e050965f56d4a2ec0ad0`.
Root authenticated all nine current freeze entries and the patch hash.
Verdict: zero defects, zero smells, zero nitpicks. D1 is closed locally.

## Defects

None. At `crates/rdocx/src/document.rs:15349`, the existing helper includes the
actual pending normal id in its occupied set before optional special creation.
add_footnote passes its already selected legacy id at
`crates/rdocx/src/document.rs:34324`. The smallest-free-at-least2 creator still
allocates after specials, so its None argument does not conceal a pending
reservation. Endnote creation supplies its actual selected id at
`crates/rdocx/src/document.rs:34552`. No allocator formula, special id, imported
type interpretation or raw owner is changed by this repair.

## Smells

None. The optional private argument distinguishes the two real allocation
orders. It adds no public signature, type, module or independent policy owner.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test or structure finding
in the changed repair. Prior style and numbering review remains applicable.
Their source is unchanged except line displacement from the note repair.
HLD03 states that the selected pending normal id is already occupied.

## Evidence and exclusions

The actual pre-repair regression compiled and executed all four combinations
of existing normal id -2 or -1 and omitted or explicit normal type. Its aggregate
failure reports duplicate ids and a special sharing the selected normal slot
for every combination. The source pin matches pass1's document. The corrected
source passes the same test. Assertions independently retain the returned
legacy id, original raw bytes, resulting normal classification, unique ids,
successful save/reopen and reopened imported bytes. This is actual introduced
regression RED/GREEN, distinct from the earlier Base failures and fixture
corrections.

The two affected note controls pass, as do actual all-target/all-feature check
and warnings-denied Clippy. Format and changed prose checks are recorded0.
Unchanged style, numbering, signing, output-mode, provider and marker cases
retain their identified prior observations. They were not rerun merely for
this review. The original failed attempts remain retained.

Accept this child for local sprint-final-acceptance preparation and independent
handoff, not final completion. Actual installed Python and typing, native and
schema observations, C1 integration, deterministic hash attribution, workspace,
hosted CI and packaging remain due at the final integrated source. No baseline,
delivery ledger or other child's acceptance is supplied here. Stop confirmation
review of this unchanged implementation.
