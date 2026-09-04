# F-234, all aspects, pass 4

**Reviewed**: working-tree diff, 3 implementation and test files, 2,589 inserted lines and 95 deleted lines
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, a modeled run-format change adopts edited unmodelled run properties
`crates/rdocx/src/comparison.rs:2397`

When an aligned run changes a supported property such as bold and also changes
a foreign or otherwise unmodelled child of `w:rPr`, `compared_run_xml` starts
the current run from the edited source. `modeled_run_properties` excludes the
raw property children from the equality decision, but the changed branch does
not restore the original raw children before serializing the current run. The
tracked result therefore adopts edited unowned XML instead of preserving the
original bytes and reporting the unsupported difference. This violates the
approved raw-preservation boundary and is not covered by the current raw-story
test.

## Smells

None.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test, or structure findings
were found. The prompt-free Word 16.104 harness now validates the same 24-record
normalization used by the ordinary differential and mutation-sensitive gates.
