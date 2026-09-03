# F-230, correctness, pass 4

**Reviewed**: complete remediated working diff against `53fbdd0`, including
untracked files and the three prior review records, 16 files, 4,987 additions
and 15 deletions. The implementation scope excluding prior review records is
13 files, 4,505 additions and 15 deletions.
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, a pre-scripted n-ary operand is still attached with post-script semantics
`crates/rdocx/src/math.rs:1881`

The pass-3 remediation parses scripts on an unbraced n-ary operand, but it
always calls `apply_scripts`. When the operand starts with the supported empty
base form, as in `\sum_i^n {}_j^k x`, that helper creates a post-sub-superscript
around an empty base. It does not consume `x` into a
`MathPreSubSuperscript`, unlike the same grammar at the top level. The n-ary
base therefore has the wrong tree and leaves `x` as a sibling.

### D2, contradictory ordinary operator attributes produce duplicate diagnostics
`crates/rdocx/src/math.rs:485`

When strict fence recognition rejects an endpoint, the generic token path first
reports each operator attribute as unsupported. It then reports an invalid
boolean or enum value for the same attribute at
`crates/rdocx/src/math.rs:507`. An endpoint such as
`stretchy="invalid"` now produces two diagnostics for one discarded format
fact. The stable loss contract requires one ordered diagnostic per lossy fact,
and the existing duplicate-suppression regression covers that invariant for
preserved OfficeMath content.

### D3, adjacent empty runs can bypass the new LaTeX loss diagnostic
`crates/rdocx/src/math.rs:1281`

Canonical normalization merges every adjacent pair of fully modeled default
runs before the LaTeX writer sees them. An empty run beside a nonempty run is
therefore absorbed into that run, so the empty-run check at
`crates/rdocx/src/math.rs:2733` never executes. A public argument containing
`MathRun::new("")` followed by `MathRun::new("x")` exports as `x` with no
diagnostic, even though the same empty run is correctly diagnosed when it is
the only expression.

## Smells

None.

## Nitpicks

None.

## Not found

The pass-3 fixes otherwise close the cited expanded-name, fence, separator,
direct-text, matrix-row preservation, reader-depth, accent, n-ary precedence,
delimiter grouping, canonical MathML, XML character, and installer mutation
gaps. No new panic, arithmetic overflow, OOXML child-order violation, raw XML
loss, dependency-family violation, runtime oracle dependency, trait, generic,
wrapper, feature flag, binding expansion, binary fixture, or unapproved file
was found. The exact Pandoc 3.10 differential, all focused conversion tests,
the 94-test workflow suite, scoped clippy, and the full `rdocx` crate suite
pass after the remediation.
