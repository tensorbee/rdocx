# F-230, correctness, pass 2

**Reviewed**: complete working diff against `53fbdd0`, including untracked files and the pass-1 review record, 14 files, 3,885 added lines, 15 removed lines, 3,900 changed lines. The implementation scope excluding the prior review record is 13 files, 3,734 added lines, 15 removed lines, 3,749 changed lines.
**Verdict**: 14 defects, 0 smells, 0 nitpicks

## Defects

### D1, limit-form n-ary recognition still ignores the operator expanded name
`crates/rdocx/src/math.rs:911`

The `munder`, `mover`, and `munderover` path decides that the first child is an
n-ary operator from its text alone. A foreign element or an `mi` containing
`∑` is therefore promoted to `MathNary` and can consume the following sibling
as its base without the required foreign-element diagnostic. Pass-1 D6 fixed
the multiscript marker and fence lookalikes, but not this n-ary lookalike.

### D2, MathML export accepts n-ary characters that it cannot import
`crates/rdocx/src/math.rs:1537`

The MathML writer emits any `MathNary::character` as `mo largeop="true"`, while
the reader recognizes only the seven characters in `is_nary_operator`. A
validly constructed single-scalar value such as `⊕` is exported with no
diagnostic, then reopens as ordinary run text instead of an n-ary expression.
This leaves pass-1 D1 remediated only for the hard-coded operators.

### D3, LaTeX limit-placement commands are silently discarded
`crates/rdocx/src/math.rs:1607`

Both `\limits` and `\nolimits` are consumed without setting
`MathNary::limit_location` or emitting a loss diagnostic. The two inputs
therefore normalize to the same tree even though F-228 has a typed
`LimitLocation` property and the conversion contract requires nonrepresentable
format facts to remain visible. Pass-1 D2 covered export-side layout losses but
not this import-side loss.

### D4, fenced-content splitting does not respect nested environment scope
`crates/rdocx/src/math.rs:2172`

The fence splitter tracks braces and nested `left` and `right` pairs but does
not track `begin` and `end` environments. A supported input such as
`\left(\begin{matrix}a,b&c\end{matrix}\right)` splits at the comma inside the
matrix and then fails to find the matrix close in the first fragment. Pass-1
D4 and D5 are fixed for nested delimiters and matrix-local separators, but the
composition of those two supported grammars still fails.

### D5, nested parsers can accept a prefix and silently drop the suffix
`crates/rdocx/src/math.rs:1847`

Matrix cells and fenced arguments call `parse_argument(None)` and absorb the
nested counters without requiring that the nested parser reached the end of
its slice. `parse_argument` stops on `\right` or `\end`, so
`\begin{matrix}a\right)b\end{matrix}` succeeds as a matrix containing only
`a`, with the unmatched delimiter and trailing `b` silently discarded. This is
an ambiguous recovery point with neither an error nor a diagnostic.

### D6, arbitrary scalar delimiters are not token-safe in canonical LaTeX
`crates/rdocx/src/math.rs:2590`

The MathML reader now accepts arbitrary scalar fence characters and the LaTeX
writer copies unrecognized scalar delimiters directly after `\left`,
`\middle`, or `\right`. An alphabetic opening delimiter such as `a` merges
with the command and following body into a command token such as `\leftax`.
A backslash delimiter is also emitted as a bare command introducer. These
trees export with no diagnostic and cannot be read back, so pass-1 D3 is not
fully remediated.

### D7, canonical LaTeX cannot preserve all accepted run text
`crates/rdocx/src/math.rs:2575`

`latex_escape` leaves square brackets and whitespace unchanged, while the
reader rejects `[` and `]` as atom delimiters and skips whitespace. A run such
as `[x]` produces LaTeX that the same module rejects, and MathML
`<mtext>a b</mtext>` becomes the different run `ab` after a LaTeX round trip.
Neither export reports that the accepted `MathRun` text is outside its
representable LaTeX subset.

### D8, escaped writer output can exceed the reader byte limit
`crates/rdocx/src/math.rs:148`

Tree validation caps raw text bytes but neither writer caps serialized bytes
after escaping. A run at the 512 KiB text boundary made from `&` expands past
the 1 MiB MathML input limit, and a run of backslashes expands much farther in
LaTeX. Both writers return nonempty output without a diagnostic that their own
readers reject immediately. The output-limit remediation for pass-1 D11 does
not cover escape expansion.

### D9, structural multiscript nodes still bypass loss diagnostics
`crates/rdocx/src/math.rs:769`

`mprescripts` is located directly and `none` is returned directly by
`mathml_optional_script`. Neither node passes through attribute, text, or child
diagnosis. Attributes or content on these structural nodes are therefore
silently discarded inside an otherwise accepted `mmultiscripts`. Pass-1 D7
added matrix and recognized-value checks, but these recognized nodes retain the
same attribute and structural blind spot.

### D10, one preserved OfficeMath loss can produce duplicate diagnostics
`crates/rdocx/src/math.rs:149`

The public writer first diagnoses `MathArgument::has_unsupported_content`,
which recursively includes descendant expressions. Each expression then
checks its recursive `has_unsupported_content` again at
`crates/rdocx/src/math.rs:1323`. Passing a complete fraction whose numerator
contains one raw preserved child therefore reports both an argument-level and
an expression-level loss for the same discarded construct. The test contract
requires one ordered diagnostic per lossy construct. Pass-1 D9 is fixed for a
bare preserved nested argument, but the complete-tree path now duplicates it.

### D11, the limit regression still omits reader-specific limit branches
`crates/rdocx/src/math.rs:3087`

The expanded test never exceeds the MathML byte limit, the LaTeX text limit,
or the LaTeX diagnostic limit. Deleting any one of those checks leaves this
test green, despite its name and the HLD claim that every limit is covered.
The added MathML depth and node cases plus LaTeX and output matrix cases do
directly remediate the branches named in pass-1 D11, but the declared
two-reader limit contract is still not proved.

### D12, the live oracle tolerates an unrecorded pre-script divergence
`crates/rdocx/src/math.rs:3286`

The 28-case gate explicitly accepts Pandoc converting a pre-sub-superscript to
a post-sub-superscript with an empty base plus a following run. The HLD records
only Pandoc `display`, `semantics`, and delimiter-scope divergences at
`docs/hld/12-testing-strategy.md:1304`. The differential-testing rule requires
every oracle disagreement to be classified, recorded, cited, and asserted.
This one is asserted in code but has no recorded decision or citation.

### D13, an existing executable bypasses the pinned archive digest
`scripts/install_pinned_pandoc.py:120`

When `prefix/bin/pandoc` already exists, installation accepts it solely from
the first version-output line and returns before checking whether the prefix is
populated or downloading the digest-pinned archive. A different executable
that prints `pandoc 3.10` therefore satisfies a successful installer run
without matching the reviewed bytes. This contradicts the exact pinned-oracle
contract and is not exercised by the installer regressions.

### D14, CI mutation coverage does not enforce the documented gate order
`scripts/test_sprint_workflow.py:1000`

The helper checks only that installation precedes the Pandoc gate. It does not
locate the full workspace suite, so moving the gate after that suite leaves the
mutation test green even though the HLD requires the exact differential before
the workspace suite at `docs/hld/15-build-and-toolchain.md:586`. Pass-1 D12 is
fixed for missing, no-op, conditional, failure-tolerant, and successful
short-circuit step mutations, but not for this documented ordering mutation.

## Smells

None.

## Nitpicks

None.

## Not found

Pass-1 D8 is fully remediated by selecting the first directly supported
`semantics` child. The new direct regressions also establish the intended
known-operator n-ary grouping, custom nonalphabetic separators, token-aware
nested delimiters, escaped ampersands and nested matrices, unsupported
environment consumption, direct bare-argument preservation reporting, broad
tree coverage, added matrix and depth limits, and the listed disabled or no-op
CI mutations.

No additional panic or arithmetic-overflow path was found within the declared
bounds. No OOXML schema-order change, prefix collision, raw OfficeMath
serialization change, dependency-family violation, runtime Pandoc dependency,
new trait, one-use generic, wrapper-only type, feature flag, or unapproved file
was found. The public API remains additive and documented, the production
dependency graph is unchanged, and the canonical MathML writer retains fixed
namespace and child order.
