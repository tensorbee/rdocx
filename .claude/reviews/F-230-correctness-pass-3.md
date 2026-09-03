# F-230, correctness, pass 3

**Reviewed**: complete working diff against `53fbdd0`, including untracked files and both prior review records, 15 files, 4,453 added lines, 15 removed lines, 4,468 changed lines. The implementation scope excluding prior review records is 13 files, 4,129 added lines, 15 removed lines, 4,144 changed lines.
**Verdict**: 13 defects, 0 smells, 0 nitpicks

## Defects

### D1, special MathML tokens still bypass expanded-name and content diagnostics
`crates/rdocx/src/math.rs:510`

The n-ary branches read the operator child's text directly and never convert or
diagnose that child. An `mo` containing a supported operator plus an unsupported
attribute or nested element is promoted to `MathNary` with that content silently
dropped. The accent branch has a second form of the same gap at
`crates/rdocx/src/math.rs:933`: a parent `accent="true"` causes any one-character
second child, including a foreign-namespace element, to become the typed accent
without the required foreign-element diagnostic.

### D2, explicit fence recognition accepts contradictory layout attributes as lossless
`crates/rdocx/src/math.rs:1118`

An endpoint with `fence="false"`, `stretchy="true"`, and the expected `form` is
still classified as a fence by the alternate branch. Conversely, an endpoint
with `fence="true"`, `stretchy="false"`, or a positionally wrong but valid
`form` is converted to `MathDelimiter` without retaining or diagnosing those
facts. Canonical output then writes `stretchy="true"` and omits `form`, contrary
to the contract that nonrepresentable format attributes remain visible.

### D3, mixed explicit MathML separators are silently collapsed to the last character
`crates/rdocx/src/math.rs:1098`

Each separator closes an argument, but every separator overwrites the one
`separator` variable. A row using comma and vertical-bar separators therefore
keeps all arguments while normalizing the delimiter to only the final
character. Re-export changes the earlier separator without any diagnostic.

### D4, direct text inside `mfenced` disappears without a diagnostic
`crates/rdocx/src/math.rs:1003`

`parse_mathml_fenced` visits element children only, and `mfenced` is excluded
from the generic structural-text check at `crates/rdocx/src/math.rs:470`.
Non-whitespace text such as the `x` in `<mfenced>x<mi>y</mi></mfenced>` is
therefore discarded silently rather than rejected or reported as unsupported
safe content.

### D5, preserved matrix-row content is invisible to both exporters
`crates/rdocx/src/math.rs:1378`

The direct-preservation helper clears all rows before asking whether a matrix
has unsupported content. That removes every `MathMatrixRow` preservation
sidecar from the query. The writers then iterate only row cells at
`crates/rdocx/src/math.rs:1508` and `crates/rdocx/src/math.rs:2648`. A matrix
parsed from OfficeMath with one raw row child is serialized to either format
without the required loss diagnostic.

### D6, writer depth validation permits canonical output that both readers reject
`crates/rdocx/src/math.rs:1260`

Tree validation starts at depth zero and accepts depth exactly 64. The LaTeX
reader starts the root argument at depth one at
`crates/rdocx/src/math.rs:1689`, while the MathML reader also counts the
outer `math` element and every serialization wrapper. A tree containing exactly
64 nested radicals passes `validate_tree` and produces nonempty output, but the
LaTeX output reaches parser depth 65 and the MathML output reaches an XML stack
of 64 before its final radical. Both outputs are rejected by their own readers.
Nested matrix wrappers make the MathML mismatch occur at even shallower tree
depths.

### D7, MathML export silently accepts an accent it cannot import
`crates/rdocx/src/math.rs:1570`

The writer emits any `MathAccent::character` without checking the reader's
one-scalar requirement. A public tree with an accent such as `"ab"` returns
MathML with no diagnostic. Reimport reaches the scalar check at
`crates/rdocx/src/math.rs:945`, reports a loss, and drops the whole accent.

### D8, an unbraced script on an n-ary operand attaches to an empty sibling
`crates/rdocx/src/math.rs:1761`

After parsing the n-ary scripts, the parser takes only `parse_atom()` as the
n-ary base and does not apply scripts to that operand. For ordinary LaTeX such
as `\sum_i^n x_j`, the base becomes plain `x`. The following `_j` is then parsed
as a separate subscript whose base is an empty run, instead of a scripted `x`
inside the n-ary base. No error or diagnostic exposes the precedence change.

### D9, delimiter argument text can be reinterpreted as additional arguments
`crates/rdocx/src/math.rs:2726`

The LaTeX writer emits delimiter arguments without grouping them. The reader's
fence splitter treats every top-level comma or vertical bar as a separator at
`crates/rdocx/src/math.rs:2366`. A one-argument delimiter whose argument is the
run `a,b` therefore exports without diagnostics and reopens as a two-argument
delimiter. A vertical bar in run text has the same problem, and a raw comma can
also conflict with a separately emitted `\middle` separator.

### D10, an empty run silently vanishes in LaTeX
`crates/rdocx/src/math.rs:2581`

The run writer emits no token and no diagnostic when `run.text` is empty.
Reimporting that output produces an empty `MathArgument`, not the original
argument containing one empty `MathRun`. Empty run values are constructible on
the public F-228 tree, so this is an undeclared loss boundary.

### D11, canonical MathML depends on unnormalized adjacent run boundaries
`crates/rdocx/src/math.rs:1403`

Imports normalize adjacent compatible runs, but export walks the caller's tree
without applying the same normalization. Two semantically normalized inputs,
one containing runs `a` and `b` and one containing run `ab`, produce different
MathML token sequences even though both reopen as the latter tree. This
violates the canonical-output and adjacent-run normalization contract.

### D12, MathML export does not reject XML 1.0 forbidden characters
`crates/rdocx/src/math.rs:1652`

`xml_escape` escapes markup but performs no XML character validation. A public
run containing U+0000 is consequently returned inside purported canonical
MathML with no diagnostic, even though the result is not a well-formed MathML
document. The same unchecked path is used for operator, fence, separator, and
accent text.

### D13, the installer does not have the documented mutation coverage
`scripts/test_sprint_workflow.py:1007`

The installer contract test checks a few source substrings, and the behavioral
test at `scripts/test_sprint_workflow.py:1092` exercises only digest mismatch,
path traversal, and an occupied prefix. There is no installer mutation loop,
despite the HLD claim at `docs/hld/15-build-and-toolchain.md:589`. Removing the
platform rejection, archive-root rejection, non-file rejection, extracted-size
cap, expected-executable check, or runtime identity call leaves the current
regressions green. The CI step mutations are now covered, but they do not prove
the installer guards.

## Smells

None.

## Nitpicks

None.

## Not found

The direct pass-2 remediation cases are present for expanded-name n-ary heads,
unsupported n-ary export characters, `limits` and `nolimits` diagnostics,
nested environment scope, nested-parser suffix rejection, token-safe delimiter
spellings, bracket escaping and whitespace diagnostics, serialized byte caps,
multiscript marker diagnostics, duplicate-preservation suppression, every
previously omitted reader limit branch, recorded Pandoc divergences, occupied
installer prefixes, and CI gate ordering.

No additional panic or arithmetic-overflow path was found within the declared
input bounds. No OOXML schema-order or raw OfficeMath serialization change,
dependency-family violation, runtime Pandoc dependency, new trait, one-use
generic, wrapper-only type, feature flag, extra binding surface, binary fixture,
or unapproved file was found. The Pandoc version, release archive digest, and
live test path remain exact, and the live comparison remains structural rather
than byte-based.
