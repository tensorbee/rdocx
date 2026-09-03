# F-230, MathML and LaTeX conversion

**Status**: completed
**Sprint**: S65
**Size**: M
**Depends on**: F-228

## Problem

F-228 will provide one normalized OfficeMath expression tree, but the native
Word facade has no MathML or LaTeX boundary. Its current conversion modules are
facade-owned and private at `crates/rdocx/src/lib.rs:24`, while the public
surface re-exports concrete native values at `crates/rdocx/src/lib.rs:43`.
F-230 must import and export the supported equation subset without creating a
second equation model or silently flattening constructs that cannot round-trip.

MathML is XML with namespace and entity concerns. LaTeX is a compact recursive
grammar with grouping, commands, environments, scripts, and ambiguous tokens.
Both directions need finite input and recursion bounds, stable source-path
diagnostics, deterministic canonical output, and a structural differential
gate against a pinned independent converter.

## Spec reference

- `docs/hld/02-scope-and-non-goals.md`, "Beyond v1" and "Still non-goals, and still permanent".
- `docs/hld/03-architecture.md`, "Why these seams" and "Facade conventions".
- `docs/hld/04-opc-and-packaging.md`, the namespace-aware XML and raw-preservation rules.
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy" and the differential testing rules.
- `docs/hld/14-development-backlog.md`, "F-230, MathML and LaTeX conversion".

## Approach

Add one approved private `crates/rdocx/src/math.rs` module. Bind it directly to
the F-228 `MathExpression` and `MathArgument` values. Do not introduce a facade
wrapper around the expression tree and do not add a trait with one implementer.
Expose four free functions because Rust cannot add inherent methods to a type
owned by `rdocx-oxml`:

```rust
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MathConversionDiagnostic {
    pub path: String,
    pub message: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MathConversionResult<T> {
    pub value: T,
    pub diagnostics: Vec<MathConversionDiagnostic>,
}

pub fn equation_from_mathml(input: &str)
    -> Result<MathConversionResult<MathArgument>>;
pub fn equation_to_mathml(expression: &MathArgument)
    -> MathConversionResult<String>;
pub fn equation_from_latex(input: &str)
    -> Result<MathConversionResult<MathArgument>>;
pub fn equation_to_latex(expression: &MathArgument)
    -> MathConversionResult<String>;
```

Re-export the F-228 equation types and these functions from `rdocx`. Python,
WASM, and CLI surfaces remain unchanged. The native additions are additive on
the pre-1.0 facade.

The MathML reader uses the existing `quick-xml` dependency and expanded names.
It supports `math`, `mrow`, `mi`, `mn`, `mo`, `mtext`, `mfrac`, `msub`,
`msup`, `msubsup`, `mmultiscripts`, `msqrt`, `mroot`, `mtable`, `mtr`, `mtd`,
`munder`, `mover`, `munderover`, `mfenced`, and the bounded accent forms that
map to F-228. The writer emits one canonical MathML 3 namespace, stable
attribute order, and explicit grouping where needed. Unsupported safe elements
become ordered diagnostics and retain their supported descendants only when
doing so cannot change grouping or operator meaning.

The LaTeX reader is a local bounded recursive-descent parser. It supports text
and number atoms, braces, `\frac`, `\sqrt` with an optional degree, `_`, `^`,
pre-scripts, `\sum` and the declared n-ary operators, `\left` and `\right`,
accent commands, and `matrix`, `pmatrix`, and `bmatrix`. The writer emits one
canonical spelling with braces around every non-atomic argument. Unsupported
commands, environments, optional arguments, and ambiguous recovery points
produce stable byte-offset paths. A construct that cannot be represented
without semantic loss is not silently substituted.

Both readers enforce limits on input bytes, tokens or XML events, recursion
depth, expression nodes, matrix rows and columns, text bytes, and diagnostics.
Both writers enforce the same tree limits before allocating output. Successful
results normalize insignificant grouping, adjacent compatible runs, canonical
operator spelling, and matrix shape, so every comparison is over one F-228
tree rather than serialized bytes.

Use Pandoc 3.10 with its bundled texmath engine as the independent differential
oracle. The source-built cases are converted between LaTeX and MathML by Pandoc,
then parsed back through this module and compared as normalized F-228 trees.
The gate verifies the exact executable version, compares structure rather than
bytes, records all intentional divergences, and includes perturbations for
fraction order, script attachment, delimiter scope, matrix cell order, and a
dropped diagnostic. Keep Pandoc in test infrastructure only and run the
version-pinned differential explicitly. Add the approved
`scripts/install_pinned_pandoc.py` installer, pin the release archive and
digest, assert the executable identity, and run the live differential in CI.

## Rejected alternatives

- A second conversion AST would duplicate the F-228 model and make layout and
  conversion disagree about normalization.
- A wrapper around `MathExpression` would only forward and is forbidden by the
  structural rules.
- A conversion trait has one implementer today and is forbidden.
- Pulling a LaTeX or MathML converter into production would make an external
  oracle a published runtime dependency.
- Comparing XML or LaTeX bytes with the oracle would reject harmless prefix,
  whitespace, and canonical spelling differences rather than semantic defects.
- Treating internal import-export symmetry as differential evidence was
  rejected because it is only a round-trip test.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `mathml_supported_subset_maps_to_one_normalized_expression_tree` | Every supported MathML element and attribute maps to the exact F-228 structure. |
| unit | `latex_supported_subset_maps_to_one_normalized_expression_tree` | Commands, groups, scripts, delimiters, accents, and matrices attach deterministically. |
| regression | `unsupported_conversion_constructs_report_stable_paths_without_semantic_substitution` | Every rejected or lossy construct has one ordered diagnostic and no guessed replacement. |
| regression | `math_converters_reject_every_declared_input_tree_and_output_limit` | Byte, event, token, depth, node, matrix, text, and diagnostic caps fail closed. |
| round-trip | `supported_equations_preserve_their_normalized_tree_through_all_four_conversion_directions` | MathML and LaTeX import and export return the same normalized F-228 tree. |
| differential | `mathml_and_latex_conversion_matches_pinned_pandoc_texmath_trees` | Source-built cases agree structurally with Pandoc 3.10 in both directions. |
| regression | `conversion_differential_rejects_structure_scope_order_and_diagnostic_perturbations` | Mutated fractions, scripts, delimiters, matrices, and loss reporting fail the predicate. |

The **test gate** is the backlog's differential gate:
`mathml_and_latex_conversion_matches_pinned_pandoc_texmath_trees`. Tests live
inside the approved module so F-230 does not contend with F-229 for the existing
`rdocx` integration test binary. No binary fixture and no new test binary is
added.

## HLD impact

- `docs/hld/02-scope-and-non-goals.md`
- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- Parser and serialiser. Use expanded names for MathML, canonical namespace and
  child order on write, bounded LaTeX parsing, structural round trips, and
  stable loss diagnostics. OfficeMath raw preservation remains owned by F-228.
- Public API of a published crate. State the additive pre-1.0 impact, run
  rustdoc with warnings denied, run `cargo publish --dry-run -p rdocx`, and
  assert the archive remains below 10 MiB.
- New module or file. Obtain explicit approval for
  `crates/rdocx/src/math.rs` and `scripts/install_pinned_pandoc.py`. Add no new
  crate, test binary, production dependency, trait, generic parameter beyond
  the result container instantiated as both `MathArgument` and `String`,
  wrapper-only type, builder, or feature flag.
- External oracle. Verify Pandoc 3.10 exactly, keep it outside published crates,
  compare normalized trees rather than bytes, record intentional differences,
  and prove structural and diagnostic perturbations fail.

## Hash harness

Expected unchanged. Existing samples do not call the new conversion functions.
Any delta blocks the story and does not authorize a baseline update.

## Implementation checklist

- [x] Add the approved facade conversion module and public result surface.
- [x] Implement bounded namespace-aware MathML import and canonical export.
- [x] Implement bounded recursive-descent LaTeX import and canonical export.
- [x] Normalize both formats into the single F-228 expression tree.
- [x] Emit stable ordered diagnostics for every declared loss boundary.
- [x] Add source-built unit, round-trip, limit, and perturbation tests in the module.
- [x] Run the exact pinned Pandoc differential gate.
- [x] Pin and install Pandoc for the live CI differential gate.
- [x] Run archive, rustdoc, focused, and full verification gates.
- [x] Update exactly the listed HLD files.

## Open questions

Resolved for S65. The new native conversion module, free-function facade,
Presentation MathML contract, diagnostic handling for nonrepresentable
OfficeMath properties, and live Pandoc 3.10 texmath oracle are approved. Pandoc
remains pinned test infrastructure and never becomes a published runtime
dependency.
