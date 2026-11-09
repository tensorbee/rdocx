# F-X280b implementation, Root final pass 2

Reviewed the bounded successor to pass1 against the unchanged A280-1/2/3/5/6
contract. Patch SHA256 is
`9af1d7514f657a82befc36c4e37c6bc437669bc879b240f4f114e073e245faea`.
Final Native document SHA256 is
`107837e1bcc4f89bc63ecd10857623bb982c4f7efe34d9aedb2483cfe9a89b50`.
Root authenticated all thirteen successor paths, the exact patch, copied GREEN
executable and actual review-cohort RED executable with its two source copies.
Verdict: zero defects, zero smells, zero nitpicks. D1 through D4 are closed
locally. This is not final integrated acceptance.

## Defects

None.

D1 is repaired at `crates/rdocx/src/document.rs:82854`. Actual picture and
supported Word shape/group namespace containers remain traversable for the
named drawing check. Unknown foreign ancestors remain opaque. A genuine
authored picture control changes only a:ext while preserving wp:extent, and
an independently authored WPS textbox exercises the same missed-container
invariant. Both have actual RED/GREEN observations and valid-base controls.

D2 is repaired at `crates/rdocx/src/document.rs:83526`. Every selected related
XML part still passes the lossless lexical check. Only actual story parts
enter story/style semantic consumers. The four UTF-16 encoding variants are
actually admitted by open, retain exact bytes through save/reopen, and now
validate both with and without the main custom relationship. Main/story open
semantics and the existing error channels are not replaced with a new decoder
or unchecked success route.

D3 is repaired at `crates/rdocx/src/document.rs:83486`. A deterministic map
groups actual normalized part identities and unions story eligibility across
all related edges before one scan. Custom-first and header-first orders,
uppercase and dot aliases, valid headers and a header-only qualified fault
are discriminating controls. Part-local relationship lookup remains in the
scanner. No global relationship-id shortcut is introduced.

D4 is discharged by
`crates/rdocx/tests/integration_test.rs:40116`. The actual clone publisher
refuses the ambiguous imported fragment before remapping or publication.
The test proves unchanged bytes, retained source/destination XML and locations,
fresh-location agreement and reopened state. An unrelated valid clone then
succeeds with the imported repeats intact. This is a new acceptance observation,
not a product repair or a claimed historical RED. The applicable fx211 raw
paragraph refusal remains separately covered.

## Smells

None. The production repair adds three concrete recognized container namespaces,
merges one existing part-role collection and scopes one existing semantic
consumer. No new public API, module, trait, generic or quota is introduced.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test or structure finding in
the bounded repair. Pass1's other examined OPC, lexical, note, finding-cap and
bookmark mechanisms are unchanged. The tests preserve foreign-opacity positives
and actual package-open controls, not just negative flag assertions.

## Actual evidence and remaining gates

The first review cohort executed D1, D2 and D3 against the retained pass1
production source and failed the required assertions. D4 passed its actual
public lifecycle control before a production change, as expected for missing
evidence. Its copied source, executable and raw log are retained. The WPS
discriminator has its own separately identified pre-repair artifact and
GREEN observation. The final affected Word cohort passes16tests. Scoped strict
all-target/all-feature Word Clippy, formatting, prose and whitespace checks
pass. The copied final integration executable is identified by SHA256
`9cebb4361b5e23158ec53ea53c2044038535e6340af251432f85d327e0b52245`.

Unchanged OPC, private reopen, CLI, PPT,56Issue430 and fx211 observations remain
bound to their actual earlier source and applicability. No unchanged whole
cohort was rerun solely for confirmation. Original broader failures, compiler
and fixture distinctions and the prepared-versus-actual first Base test hash
remain as documented. This repair's actual source copies do not retroactively
invent the missing original full Base snapshot.

Accept F280b for sprint-final-acceptance handoff after the local checklist is
updated. Keep its plan approved and all installed Python/typing, native,
SDK/schema, LibreOffice, C1 integration, hash/pixel, full workspace, packaging
and hosted checks pending. No delivery completion or publication follows.
Stop confirmation review of this unchanged bounded implementation.
