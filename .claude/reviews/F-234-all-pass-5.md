# F-234, all aspects, pass 5

**Reviewed**: complete working-tree diff across 3 implementation and test files, 2,681 inserted lines and 95 deleted lines
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, related-story diagnostics lose their story identity

`crates/rdocx/src/comparison.rs:696`

The related-story path wraps a header, footer, comment, footnote, or endnote in
a synthetic document and delegates to `compare_body` without carrying the
story location into diagnostic construction. `compare_body` consequently uses
the hard-coded `body/...` location from `body_location`, so differences in
distinct stories can report the same location and callers cannot identify the
affected story. The new header regression confirms this incorrect public
result by expecting `body/paragraph[0]/run[0]` at
`crates/rdocx/tests/regression_test.rs:10717`. This violates the approved
contract's requirement for concrete story identities and child paths.

## Smells

None.

## Nitpicks

None.

## Not found

The pass-4 run-property defect is fixed. The current tracked run retains the
original unmodelled property child exactly, the prior `w:rPr` snapshot contains
only modeled original properties, and the edited unmodelled child is not
adopted. The focused regression proves one diagnostic and the tracked,
accepted, and rejected outcomes. It passes together with all four
`full_story_comparison` tests and `cargo check -p rdocx --all-targets`.

No other correctness, contract, panic, OOXML preservation or ordering, test
adequacy, or structural findings were found.
