# F-230, correctness, pass 1

**Reviewed**: complete working diff against `53fbdd0`, including untracked files, 13 files, 2,711 added lines, 15 removed lines, 2,726 changed lines
**Verdict**: 12 defects, 0 smells, 0 nitpicks

## Defects

### D1, MathML export cannot round-trip valid n-ary trees
`crates/rdocx/src/math.rs:1186`

A `MathNary` with no visible limits is emitted as a plain `mo` followed by its
base. The reader maps that `mo` to a `MathRun`, so even
`MathNary::new("∑", MathArgument::text("x"))` reopens as text rather than an
n-ary expression. When limits are present, the root reconstruction at
`crates/rdocx/src/math.rs:417` consumes only one following expression as the
base. A base containing two expressions is therefore split, with its second
expression moved outside the n-ary node.

### D2, typed OfficeMath properties are silently lost on export
`crates/rdocx/src/math.rs:1072`

The MathML writer does not inspect `alignment` on sub-superscript or
pre-sub-superscript values, and it does not inspect `hide_degree` on radicals.
The LaTeX writer also omits alignment for every script form, radical degree
visibility, delimiter growth, and n-ary growth and limit placement at
`crates/rdocx/src/math.rs:1898`. Both writers silently omit a nonempty n-ary
script when its hide flag is set. These public typed values cannot survive the
conversion, but the result contains no diagnostic, contrary to the facade
contract.

### D3, accepted delimiter forms fall outside the canonical round-trip subset
`crates/rdocx/src/math.rs:765`

`mfenced` accepts any scalar opening, closing, and separator characters. The
explicit-fence reader recognizes only a short hard-coded delimiter set at
`crates/rdocx/src/math.rs:836`, so an accepted `mfenced` value such as `+x-`
exports to explicit fences and reopens as ordinary runs. The LaTeX splitter at
`crates/rdocx/src/math.rs:1759` recognizes only comma and vertical bar, so an
accepted two-argument `mfenced` with `separators=";"` becomes one argument after
MathML to LaTeX to tree conversion. Neither path reports the loss.

### D4, nested LaTeX delimiters are not parsed by delimiter scope
`crates/rdocx/src/math.rs:1443`

The matching scan recognizes `\\left` and `\\right` by string prefix rather
than command token, while the subsequent separator scan tracks braces but not
nested left-right depth at `crates/rdocx/src/math.rs:1739` and
`crates/rdocx/src/math.rs:1759`. A supported expression such as
`\\left(\\left[a,b\\right],c\\right)` is split at the inner comma and fails to
parse. An unsupported command whose name starts with `left` or `right` can also
change scope instead of producing its declared loss diagnostic.

### D5, matrix and unsupported-environment scanning is not grammar-aware
`crates/rdocx/src/math.rs:1472`

Supported matrix bodies are located with the first textual closing marker and
split with raw `split("\\\\")` and `split('&')`. This breaks nested matrices and
the explicitly supported escaped ampersand command. For example, a cell
containing `a\\&b` is split into multiple cells. The unsupported-environment
branch returns immediately after its diagnostic without consuming the body or
matching `\\end`, so `\\begin{array}x\\end{array}` ultimately returns a trailing
input error instead of the promised stable lossy result.

### D6, MathML structural recognition uses local names without namespaces
`crates/rdocx/src/math.rs:613`

The `mprescripts` marker is selected by local name alone. Explicit fence
separators and endpoints do the same at `crates/rdocx/src/math.rs:815` and
`crates/rdocx/src/math.rs:836`. A foreign-namespace `mprescripts` or `mo` can
therefore control the normalized tree and avoid the foreign-element diagnostic.
This violates the expanded-name reader contract and changes meaning based on a
same-local-name lookalike.

### D7, MathML loss diagnostics have attribute and structural blind spots
`crates/rdocx/src/math.rs:449`

Attributes are treated as lossless merely because their local names appear in
an allowed list. A standalone `mo` with `largeop`, or an invalid `fence`,
`stretchy`, `form`, or accent value, is converted without representing or
diagnosing that attribute. Matrix rows and cells bypass attribute diagnosis
entirely at `crates/rdocx/src/math.rs:660`, and non-whitespace text directly on
an `mtr` is ignored. The HLD requires every unsupported safe format fact to be
reported rather than silently discarded.

### D8, `semantics` does not retain the first supported descendant
`crates/rdocx/src/math.rs:540`

The implementation selects the first element child without checking whether
it is supported. With `annotation` first and `mi` second, it diagnoses and
drops the annotation, ignores the supported `mi`, and returns no expression.
The approved design requires the first supported descendant at this sole
transparent boundary.

### D9, top-level argument preservation can disappear without a diagnostic
`crates/rdocx/src/math.rs:1006`

Both argument writers iterate only the typed expressions and never check
`MathArgument::has_unsupported_content`. A public nested argument obtained from
parsed OfficeMath can retain raw siblings in its own preservation sidecar.
Passing that argument directly to either conversion function drops those
siblings with no diagnostic. The LaTeX path has the same omission at
`crates/rdocx/src/math.rs:1845`.

### D10, the round-trip and live differential gates cover only a fraction of the declared tree
`crates/rdocx/src/math.rs:2034`

The sole round-trip source omits standalone runs, single scripts,
pre-sub-superscripts, lower and upper limits, and n-ary expressions, so the
failures above cannot reach `supported_equations_preserve...` at
`crates/rdocx/src/math.rs:2306`. The live Pandoc gate at
`crates/rdocx/src/math.rs:2329` uses only a fraction, a radical, and one
delimiter case. It does not establish the HLD claim that source-built cases
cover every supported MathML element, LaTeX command family, and normalization
rule.

### D11, the limit and perturbation regression does not prove its named contract
`crates/rdocx/src/math.rs:2228`

The limit test omits MathML depth and node caps, LaTeX matrix caps, and output
matrix caps. Its MathML text case exceeds the equal byte cap before the text cap
can be observed. The delimiter-scope perturbation at
`crates/rdocx/src/math.rs:2393` compares a fraction with unrelated parenthesized
text instead of comparing scoped and unscoped forms of the same expression, so
it remains green even if delimiter scope is ignored completely.

### D12, the claimed CI mutation coverage does not protect the Pandoc gate
`scripts/test_sprint_workflow.py:978`

The workflow regression checks only that selected strings occur and that the
two named steps appear in order. Adding `continue-on-error: true`, an `if:
false` condition, or a no-op run line to either Pandoc step is not rejected.
The adjacent HLD claims mutation coverage and direct failure propagation, but
this test does not establish either property.

## Smells

None.

## Nitpicks

None.

## Not found

No panic or arithmetic-overflow path was found within the declared parser and
writer bounds. No OOXML schema-order change, dependency-family violation,
runtime Pandoc dependency, new trait, one-use generic, wrapper-only type,
feature flag, or unapproved file was found. The public functions and result
types are re-exported and documented, the Pandoc release archive and executable
identity are pinned, and the installer rejects traversal and non-file archive
members.
