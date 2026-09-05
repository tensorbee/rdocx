# F-X077, all, pass 1

**Reviewed**: Uncommitted working diff on `work/f-x077-codex`, 7 files and
1,574 changed lines, with 781 insertions and 793 deletions
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, XML declaration values accept references outside the declaration grammar

`crates/oxml-core/src/xml.rs:128`

The declaration validator normalizes references before checking `encoding` and
`standalone`. XML 1.0 declaration pseudo-attribute values use literal grammar,
so input such as `encoding="UT&#70;-8"` or `standalone="y&#101;s"` must be
rejected. Normalization turns those values into `UTF-8` and `yes`, which lets a
lexically invalid declaration pass the shared strict validator and all three
consumers.

## Smells

None.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test, or structure findings.
