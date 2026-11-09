# F-X272a custody recognition, Root bounded final pass 3

Reviewed the combined314line v4 and395line final v5 deltas, their inherited
consumer paths, the current748766canonical contract and authenticated evidence.
Verdict: D1 remains open in non-metadata revision rendering. One defect,
zero smells, zero nitpicks. Q-A430-29's local implementation is accepted within
its explicit corruption-detection limit, but this is not feature completion.

Final source manifest SHA256 is
`a71992628cc90eb7c6a94bc2e0fd781c76bf0357d65040ba07db7d1b688ff0c2`.
Final v5 delta SHA256 is
`32095f1c3709654c39b8bc1d483e0642c6be3cffa89c4090884a3dd7f74713aa`.
Frozen revision.rs SHA256 is
`a2b27b0ae7c99773fe4d22a78e0ab52c27a4dd05f9be24556f18d547f541ce3c`.
Citations below refer to that repository-relative final snapshot.

## Defects

### D1, clearing revision metadata does not protect opaque payload from deleted-text conversion

`crates/rdocx/src/revision.rs:1458`
`crates/rdocx/src/revision.rs:2546`
`crates/rdocx/src/revision.rs:2712`

The qualification fold removes revision metadata outside admitted owners, but
does not make that admission boundary authoritative for other semantic writes.
The ordinary renderer still descends into every lexical child and propagates
convert_deleted_text. It renames Word delText and delInstrText using only the
namespace flag. An opaque subtree needs no revision metadata for this path.

A normal note can contain a real run with ordinary text plus an unknown foreign
child holding literal Word delText or delInstrText markup. That foreign child
is retained raw by the model. Removing the note tracks the real run as deletion.
Rejecting that deletion then reaches the opaque child with deleted-text
conversion enabled and renames its literal payload. The repair's metadata-only
filter cannot prevent this. The comparison may refuse at its reject-to-original
postcondition, or a missed residual check may publish altered raw bytes. Neither
is the promised supported removal and exact rejection restoration. This is a
source-derived counterexample, not an executed public failure in this review.

Make the same actual footnote admission boundary govern semantic rendering,
not only metadata enumeration. Preserve required namespace relocation of raw
payload and the actual admitted textbox/table/control positive paths. Do not
patch only one spelling or broaden the inherited generic parser policy as a
shortcut. The different discriminator should exercise note removal and reject
with both opaque literal spellings and a real deleted-text conversion positive.
Retain D1 and escalate the shared invariant to Root orchestration before repair,
rather than adding an unrelated example-specific predicate or a new peer pass.

## Smells

None found in the bounded combined delta.

## Nitpicks

None.

## Accepted portions and remaining evidence

The actual owner fold is linear in tree elements and uses existing physical
story roots, including real TextBox and TableCell roots. It avoids deriving an
empty structural owner from paragraph presence. The note-only metadata filter
now reaches unification, custody membership, removal counts, listing and the
actual resolver entrypoints. These metadata cases and their genuine structural
positives are repaired. The generic XmlTree parse/reset policy remains unchanged.

Q-A430-29 removes unmatched-B adjacency inference, including the no-part branch.
Actual reference coherence is checked independently. N-present owner, token,
bits, cardinality and counterpart checks remain, and fully decoded admitted
pair corruption refuses explicitly. Opaque or prefix-only lookalikes do not
gain custody. The formerly held history case now checks both resolution orders,
reopen, unrelated existing note6 and independent prior revision90. No new marker
grammar, input allowance or Q28 exception follows from the approved choice.

Mechanical support authenticated65v4and77v5custody files, exact incremental and
cumulative Base reconstruction and the one-line Native visibility lease. The
old-v3 two actual failures and one positive, old-v4 Q29 failure and introduced
malformed-pair regression remain distinct. All55integration cases run without
exclusion or ignore, alongside5unit,1reportedregression,1CLI and prose0.

Those runtime executions belong to their actual runtime source SHA. Root read
the exact later pure nested-if to let-chain diff and accepts its unchanged
guard semantics without relabelling those executions as final-SHA runs. Final
all-targets/all-features Clippy and workspace format exit0 are authenticated
separately. The additive receipt joins attest stable producer pins but contain
no independent before/after inventory dictionaries. No stronger receipt is
invented. Python AST remains syntax-only. Full installed, SDK, native, hash,
workspace and other final riders remain pending. Q28's one exact native
positive remains unproved, never passed.

No additional public API, panic, ordering or structure finding is recorded in
the examined combined delta. Source remains in the dirty child worker and both
Native leases are returned. No complete-feature or provider handoff follows.
Root modified only this review record and returns to orchestration for the
concrete shared-invariant remediation handoff.
