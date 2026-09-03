# F-229, all aspects, pass 3

**Reviewed**: remediated uncommitted worker diff, 13 implementation files,
2,635 additions and 15 deletions
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, normalized group baselines are discarded before pagination
`crates/oxml-layout/src/line.rs:1252`

Line measurement treats a non-finite baseline as the established top-aligned
case and clamps every finite baseline to the group height. The conversion at
line 1302 nevertheless copies the original value into `LineItem`, and the
paginator uses that unnormalized value at
`crates/rdocx-layout/src/paginator.rs:2595`. A public `InlineItem::Group` with a
negative, oversized, or non-finite baseline therefore measures one line box
but is positioned against a different baseline. A non-finite value reaches the
page transform as a non-finite coordinate.

## Smells

None.

## Nitpicks

None.

## Not found

No additional findings in correctness, contract, panics, OOXML preservation,
tests, or structure. Pass 1 reusable context, oracle derivation, rendered-page
mutation, and exact-token findings are fixed. Pass 2 matrix, delimiter, and
traversal findings are fixed.
