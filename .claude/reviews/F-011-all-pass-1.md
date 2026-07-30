# F-011, all aspects, pass 1

**Reviewed**: working-tree diff, 3 files, 40 insertions and 5 deletions
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, testing specification citation names a missing section

`.claude/plans/F-011-design.md:19`

The plan cites `docs/hld/12-testing-strategy.md`, "Unit and property tests",
but that document has no section with that name. The truncation-pinning
requirement is under "New tests the extracted crates need". The invalid citation
prevents a reviewer or later worker from following the design contract to its
source without re-deriving the intended section.

## Smells

None.

## Nitpicks

None.

## Not found

No correctness, panic, OOXML, test-gate, or structural findings were found. The
tests cover all approved constructors with positive and negative fractional
inputs, and the temporary rounding mutation proves that each named gate detects
the prohibited behavior.
