# F-234, all aspects, pass 1

**Reviewed**: working-tree diff, 3 implementation and test files, 1,831 inserted lines and 94 deleted lines
**Verdict**: 5 defects, 0 smells, 0 nitpicks

## Defects

### D1, text-box discovery is not namespace aware
`crates/rdocx/src/comparison.rs:652`

The scanner treats every lexical `:txbxContent` suffix as a Word text-box
owner. A producer extension such as `x:txbxContent` is therefore revised as
Word content instead of being preserved as unmodelled XML.

### D2, text-box handling rejects simultaneous host-story edits
`crates/rdocx/src/comparison.rs:609`

When a story contains a text box, the outer story must have a byte-identical
skeleton. A supported paragraph, field, table, or formatting edit beside that
text box therefore fails even though both changes are inside the approved
boundary.

### D3, inherited Word prefixes fail for comment and note owners
`crates/rdocx/src/comparison.rs:566`

Each isolated comment, footnote, or endnote fragment is reparsed without the
namespace declarations inherited from its part root. Normal package XML that
declares `xmlns:w` only on `w:comments`, `w:footnotes`, or `w:endnotes` fails
with an expected Word owner-root error.

### D4, complex-field result splitting assumes a fixed prefix and attribute spelling
`crates/rdocx/src/comparison.rs:2406`

The complex-field result scanner searches for literal `fldCharType` values and
literal `w:r` boundaries. Prefix aliases, alternate valid attribute prefixes,
or retained source formatting fail despite the prefix-tolerant contract.

### D5, the differential test does not implement its named predicate
`crates/rdocx/tests/regression_test.rs:10120`

The mutation-sensitivity test only checks that move wrapper substrings exist.
It does not compare normalized records and does not perturb kind, order,
story, field owner, formatting kind, or move-pair identity as required by the
approved gate.

## Smells

None.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test, or structure findings.
