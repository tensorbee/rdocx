# F-X272, ordinary property implementation, Codex pass 1

**Reviewed**: Sealed ordinary-property-events increment, 11 paths, 1279 additions
and 99 deletions in the complete 2291-line patch. Seven lower source files,
three HLD files and the existing progress record.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Bounded increment is clean.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Contract and mechanism

Read repository guidance, workflow, microscope, current progress and the approved
Positive ordinary property events from actual readers contract. Canonical plan
SHA256 at preparation was
`8c3d710b1139729167ab6700a8e97dffa29e1ab6f0ab0c085bbc042e0b3386df`.
Reviewed immutable before/after snapshots rather than concurrent Native work.

The actual Empty pStyle and rStyle branches record their selected event bounds
only after the existing value getter succeeds at
`crates/rdocx-oxml/src/paragraph_properties.rs:918` and
`crates/rdocx-oxml/src/run_properties.rs:1008`. Their local output replaces the
prior selection and escapes only on successful property-reader return. Existing
attribute spelling, missing-value and unknown-entity behavior remain intact.
The substantive scoped readers at `crates/rdocx-oxml/src/numbering.rs:1397` and
`crates/rdocx-oxml/src/numbering.rs:1458` preserve the same parser and translate
logical offsets through the live reader/input origin, including BOM input.

Run and paragraph observations accompany the actual selected property model.
A later started property replaces its own fact even when the replacement has
no style. Ignored Empty rPr retains the existing model and fact. Successful
children do not publish early. Whole run/paragraph failure drops their local
output. Actual capture maps select copied images rather than a guessed enclosing
ordinary subtree or a recovered Original interval.

The projection sidecar at `crates/rdocx-oxml/src/text.rs:14586` keeps run and
control observations aligned with actual retained rows. Inspected direct runs,
hyperlink append, gap filling, complete-result control transfer and complex
projection. The splitter's real Outside parse supplies its property fact at
`crates/rdocx-oxml/src/text.rs:16231`. The actual replacement splice at
`crates/rdocx-oxml/src/text.rs:15524` selects the same rows as the returned model.
Bare/fallback and synthetic Field rows receive no ordinary fact. Control rows
moved into a Field result are removed from the outer ordinary rows.

Cell, row, table and SDT factories retain observations in their existing pending
companions until their own successful return. Inspected nested table/control
pushes, direct inline SDT Run, late retained-root validation and Option/raw
fallback. Body and whole-document companions stay local until whole success.
The actual later smartTag/customXml wrapper reader publishes only its own
successful result at `crates/rdocx-oxml/src/text.rs:17062`, after sticky-resource
checks and delivered Field commitment. Initial raw retention supplies no such
observation.

`crates/rdocx-oxml/src/document.rs:72` groups actual mapped intersections by
collector, unions overlapping observations and leaves generated mapping gaps
unproved. Existing positive Field ranges take priority. The public source
readers commit delivered Field ownership before publishing property facts.
No observation registers input, increases allowance, funds a copy or changes
valid-input acceptance. None-mode ordinary callers create no added observation
transcript. That last conclusion is source inspection, not an allocator
measurement. HLD03/04/10 describe these same bounded facts and exclusions.

## Evidence and custody

Authenticated BOUNDARY.json SHA256
`ebddbf5bc4afe869b0814c13d689b8ddcff4e5d0a99fb6cbf273f52e901e97cb`
and increment.patch SHA256
`240d868c8261f42ef5205de8e1ff0f2ef4075967456727ef0fa7bb424647738a`
under `/private/tmp/rdocx-s96-root-c1-implementation/word-grammar-carriers/ordinary-property-events-increment`.
All 22 before/after snapshot hashes match. Regenerated the complete unified
patch in memory and matched it byte for byte. Every boundary-listed evidence
file matches its recorded hash.

Read all 13 final property test assertions and actual PASS names. They cover
inner/outer overwrite, ignored Empty rPr, malformed attributes, later property
and owner errors, raw control fallback, aliases/BOM, nested retained containers,
hyperlinks, outside-run projection, Field priority and generated map holes.
The archived final executable matches FINAL_BINARY.json SHA256
`dc2ec9ccb8882157c9b0b5b0d668773ffb0ea18c95ac7701b56b3acad14da594`.

Read all seven related logs and checked 89 distinct PASS names, disjoint from
the final 13. Their receipt retains the earlier executable SHA256
`136ad5dcac412fbe50e25d7d915315a3e66cfcb7db295941e2a1f096c9e4c0ac`.
These earlier results are not relabeled as runs of the final executable. The
recorded final successor changes the single-Range Vec test expression for
Clippy. Final 13 tests ran after that repair. Strict lower/Native all-target,
all-feature Clippy completed successfully in 35.86 seconds. Formatting, prose
and adapter evidence is clean. No review build or test execution was performed.

Authenticated the controlled overwrite omission's four source/binary/log pins.
Its actual one-test failure exits 101 when a retired rPr observation is wrongly
retained. Restored GREEN text matches the final text snapshot exactly. This is
a controlled omission, not a historical product RED. The earlier preparation
string assertion stopped before source edits or Cargo. The unknown-entity
fixture failure reflects its incorrect premise about the existing lexical
getter and remains distinct from product regressions. Earlier and final Clippy
failures and repairs remain recorded rather than erased.

## Not found and limits

No correctness, contract, panic, OOXML, testing or structural findings were found
in this bounded delta. No new public signature, model field, global transaction,
parser, quota or resource exemption is introduced. Existing optional companions
are substantive consumers of the same readers and retain decisions.

This review does not certify whole ordinary XML, the Field complement, relevant
Native import consumers, relationship/drawing identity scanners or whole C1.
Revision, Ruby and drawing-specific property handbacks remain separate actual
producer obligations. Native lineage/writer migration, source reconciliation,
installed/native qualification and final integrated riders remain pending.
There is no provider or sprint completion claim.

Review ended after writing this report. Only this canonical review file was
written. No source, progress, plan or HLD edits and no remediation were performed.
