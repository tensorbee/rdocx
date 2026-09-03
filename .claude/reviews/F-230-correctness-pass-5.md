# F-230, correctness, pass 5

**Reviewed**: complete remediated working diff against `53fbdd0`, including
untracked files and the four prior review records, 17 files, 5,164 additions
and 15 deletions. The implementation scope excluding prior review records is
13 files, 4,620 additions and 15 deletions.
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, an unbraced nested n-ary operand is left incomplete
`crates/rdocx/src/math.rs:1865`

The operand path parses one atom and its scripts, but it does not apply the
n-ary completion logic that the outer argument loop applies at
`crates/rdocx/src/math.rs:1860`. For supported input such as
`\sum_i^n \prod_j^m x`, the outer sum receives a product with an empty base,
then `x` remains a sibling of the sum. The same gap leaves `\limits` or
`\nolimits` on a nested n-ary operand to be parsed as an unrelated command.
Operand parsing needs one shared completion path for n-ary placement, scripts,
pre-scripts, and the recursive base.

## Smells

None.

## Nitpicks

None.

## Not found

Pass-4 D1 is closed for ordinary scripted and pre-scripted n-ary operands.
Pass-4 D2 is closed with one diagnostic per rejected operator attribute.
Pass-4 D3 is closed without weakening canonical adjacent-run normalization.
No other defect or smell was found in correctness, contract, panics, OOXML
handling, tests, or structure. The focused conversion suite and scoped clippy
remain green.
