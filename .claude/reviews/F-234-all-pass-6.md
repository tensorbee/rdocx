# F-234, all aspects, pass 6

**Reviewed**: complete working-tree diff across 3 implementation and test files, 2,796 inserted lines and 101 deleted lines
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, the existing comparison regression gate is red

`crates/rdocx/tests/regression_test.rs:9090`

The approved modeled-property boundary changed bold-to-italic run formatting
from diagnostic-only behavior to a tracked `w:rPrChange`, but this existing
test still requires one diagnostic, no revisions, and byte-identical original
content. The complete `rdocx` regression binary therefore fails. The same stale
expectation remains in the matched-table test at
`crates/rdocx/tests/regression_test.rs:9344`, which still requires paragraph
and run diagnostics plus no revisions even though only the unsupported cell
property should remain diagnostic. `cargo test -p rdocx --test
regression_test` reports 257 passed, 2 failed, and 3 ignored.

### D2, new paragraph and table property owners adopt edited unmodelled XML

`crates/rdocx/src/comparison.rs:1630`

When the original paragraph has no `w:pPr` and the edited paragraph adds both
a supported modeled property and an unmodelled property child, `current` starts
from the edited owner. The raw-child diagnostic and restoration run only when
the original owner exists, so the tracked result silently adopts the edited
unowned child. The table path has the same absent-original-owner hole at
`crates/rdocx/src/comparison.rs:1908`, where edited `extra_xml` is retained and
the restoration branch is skipped. This violates the approved rule that
unsupported formatting is diagnostic or rejected and that unowned XML comes
from the original. The focused raw-property regression covers only `w:rPr`, so
neither case is gated.

### D3, public revision-resolution documentation still says main document only

`crates/rdocx/src/revision.rs:102`

`accept_all` and `reject_all` now resolve every indexed related story as well
as the main document, but both public method descriptions still promise only
main-document resolution. This directly contradicts the F-234 selector
contract and gives native callers an incorrect mutation scope.

## Smells

None.

## Nitpicks

None.

## Not found

The pass-5 diagnostic-provenance defect is fixed. Header, footer, comment,
normal footnote, endnote, and nested text-box diagnostics retain their concrete
story and child paths, while the existing main-body path remains
`body/paragraph[0]/run[0]`. Related parts are deduplicated before comparison
and revision resolution, so no duplicate story traversal was found.

The focused provenance tests, the run-property preservation test, all four
ordinary full-story differential tests, scoped related-story resolution, and
`cargo check -p rdocx --all-targets` pass. No additional panic, OOXML ordering
or namespace, differential-gate, or structural findings were found.
