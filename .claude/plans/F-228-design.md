# F-228, OfficeMath model and authoring

**Status**: approved
**Sprint**: S65
**Size**: L
**Depends on**: none

## Problem

The WordprocessingML grammar models paragraph runs and preserves every other
paragraph child as raw XML at `crates/rdocx-oxml/src/text.rs:1759`, but it has
no typed OfficeMath content. The public paragraph facade exposes runs and the
other direct paragraph items at `crates/rdocx/src/paragraph.rs:276` and
`crates/rdocx/src/paragraph.rs:929`, but an equation can only survive as an
opaque sibling. Callers cannot inspect, mutate, or author the OfficeMath subset
named by F-228.

The missing grammar includes inline and display equations, math runs,
fractions, scripts, radicals, matrices, limits, n-ary operators, delimiters,
accents, and their properties. It must distinguish modern `m:` OfficeMath from
legacy Equation Editor objects, preserve unsupported descendants at their
original boundaries, and write supported children in schema order without
moving ordinary Word runs or raw siblings.

## Spec reference

- `docs/hld/02-scope-and-non-goals.md`, "Beyond v1" and "Still non-goals, and still permanent".
- `docs/hld/03-architecture.md`, "Three families, one workspace", "The dependency rule", "What stays put", "Crate-level conventions", and "Facade conventions".
- `docs/hld/04-opc-and-packaging.md`, "The package" and the Word namespace, schema-order, and raw-preservation rules under package integrity.
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", "The hash harness", and "New tests the extracted crates need".
- `docs/hld/14-development-backlog.md`, "Milestone 22, Word depth" and "F-228, OfficeMath model and authoring".

## Approach

Add one approved `crates/rdocx-oxml/src/math.rs` module. Keep OfficeMath in the
WordprocessingML grammar owner and use the existing `quick-xml` pull parser,
fallible text decoding, namespace resolution, and raw subtree capture. Add the
OfficeMath namespace to `namespace.rs`, re-export the public model from
`rdocx-oxml`, and do not add a crate, dependency, feature flag, trait, generic
parameter, builder, or dynamic dispatch.

Use one normalized recursive model whose recursion is carried by `Vec`, not a
trait object:

```rust
pub struct CT_OMath {
    pub expressions: Vec<MathExpression>,
    // private slot-indexed source and raw-preservation state
}

pub struct CT_OMathPara {
    pub properties: MathParagraphProperties,
    pub equations: Vec<CT_OMath>,
    // private slot-indexed source and raw-preservation state
}

pub enum OfficeMath {
    Inline(CT_OMath),
    Display(CT_OMathPara),
}

pub enum MathExpression {
    Run(MathRun),
    Fraction(MathFraction),
    Subscript(MathScript),
    Superscript(MathScript),
    SubSuperscript(MathSubSuperscript),
    PreSubSuperscript(MathPreSubSuperscript),
    Radical(MathRadical),
    Matrix(MathMatrix),
    LowerLimit(MathLimit),
    UpperLimit(MathLimit),
    Nary(MathNary),
    Delimiter(MathDelimiter),
    Accent(MathAccent),
}

pub struct MathArgument {
    pub expressions: Vec<MathExpression>,
    // private slot-indexed raw children
}
```

Each construct has a concrete property struct only for values needed by
authoring, later layout, or conversion. These include fraction type, script
alignment, radical degree hiding, matrix rows and columns, limit placement,
n-ary character and limit visibility, delimiter characters and separator,
accent character, math-run text and formatting, and display justification.
The recommended scope also types the existing settings-owned `m:mathPr`
defaults because F-229 needs the configured math font, justification, margins,
spacing, and integral and n-ary limit placement. Unknown attributes and
children remain attached to their owning construct with the inherited namespace
bindings they require.

Teach `CT_P` to project `m:oMath` and `m:oMathPara` at direct run boundaries.
The sidecar records the run boundary and the number of raw and typed boundary
items that precede the equation, matching the existing paragraph preservation
model. Parsing accepts any prefix bound to the OfficeMath namespace. Writing
uses fixed `m:` names and emits `xmlns:m` only when the document contains typed
OfficeMath. A parsed construct that has not changed may reuse its exact source
bytes. A typed mutation emits one canonical supported subtree while retaining
every unsupported sibling in its original slot.

Teach `CT_Settings` to project the single schema-positioned `m:mathPr` subtree
through the same source-preserving replacement pattern. The facade exposes
read and set access through `Document`, while unrelated settings and a
relationship-resolved nonconventional settings part remain untouched.

Expose `ParagraphItemRef::Equation`, borrowed equation iteration, indexed
mutable access, and one `Paragraph::add_equation(OfficeMath)` insertion method.
Re-export the concrete equation model from `rdocx`. Provide core leaf and
argument constructors on the concrete values, but do not add a fluent facade
for every grammar production. Python, WASM, and CLI surfaces remain unchanged.
The native additions are additive on the pre-1.0 `rdocx` facade and the
published `rdocx-oxml` grammar crate.

## Rejected alternatives

- Keeping OfficeMath as raw XML would preserve bytes but fail typed inspection,
  mutation, authoring, layout, and conversion.
- Adding the model to `text.rs` would make an already large mixed grammar file
  harder to reason about and would violate the repository's local-readability
  test.
- A trait-based expression tree has only one implementation and is forbidden by
  the structural rules.
- A second facade-owned equation tree would force layout and conversion to
  translate between competing normalized models.
- Typing legacy Equation Editor or OLE payloads would cross the permanent legacy
  boundary and is outside F-228.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `every_supported_officemath_construct_writes_schema_order_and_reparses` | All named expressions, arguments, and property slots serialize in schema order and reparse structurally. |
| unit | `officemath_reader_accepts_aliases_and_writer_uses_fixed_math_prefix` | Expanded names decide typed meaning and canonical output uses `m:`. |
| regression | `unsupported_officemath_siblings_survive_typed_mutation_byte_for_byte` | Unknown attributes and subtrees remain at the same owning boundary after a supported edit. |
| regression | `legacy_equation_editor_objects_remain_unmodelled_raw_xml` | VML, OLE, and legacy equation payloads never become typed OfficeMath. |
| integration | `paragraph_items_keep_runs_equations_controls_and_raw_xml_in_source_order` | Inline and display equations retain exact paragraph ordering with ordinary content. |
| integration | `public_equation_authoring_saves_reopens_and_remains_mutable` | The native paragraph facade authors the complete supported corpus, saves, reopens, and mutates it through one model. |
| round-trip | `officemath_corpus_parses_mutates_saves_and_reopens_without_losing_supported_or_raw_siblings` | The complete source-built corpus preserves every supported expression and every opaque sibling after mutation and reopen. |

The **test gate** is the backlog's round-trip gate:
`officemath_corpus_parses_mutates_saves_and_reopens_without_losing_supported_or_raw_siblings`.
The fixture is built as source XML and source API calls inside existing unit and
integration targets. No binary fixture and no new integration test binary is
added.

## HLD impact

- `docs/hld/02-scope-and-non-goals.md`
- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`

## Risk routing

- Parser and serialiser. Read the packaging and Word model rules, prove
  prefix-tolerant reads, fixed-prefix schema-ordered writes, and byte-exact raw
  subtree preservation through typed mutation and reopen.
- Public API of published crates. State the additive pre-1.0 impact, run rustdoc
  with warnings denied, run `cargo publish --dry-run` for `rdocx-oxml` and
  `rdocx`, and assert both archives remain below 10 MiB.
- New module or file. Obtain explicit approval for
  `crates/rdocx-oxml/src/math.rs`. Add no new crate, test binary, trait, generic
  parameter, builder, wrapper-only type, dependency, or feature flag.

## Hash harness

Expected unchanged. Existing source-built samples contain no OfficeMath, and
`xmlns:m` is emitted only for documents that contain a typed equation. Any
existing sample delta blocks the story and does not authorize a baseline update.

## Implementation checklist

- [ ] Add the approved OfficeMath grammar module and namespace binding.
- [ ] Implement the single normalized expression tree and bounded properties.
- [ ] Project document-wide `m:mathPr` defaults through the existing settings owner.
- [ ] Parse and serialize every supported expression in schema order.
- [ ] Preserve unsupported attributes and children at their owning slots.
- [ ] Integrate inline and display equations into paragraph source order.
- [ ] Add the bounded native paragraph authoring and reader surface.
- [ ] Add the source-built round-trip corpus to existing test targets.
- [ ] Run focused parser, facade, archive, rustdoc, and full verification gates.
- [ ] Update exactly the listed HLD files.

## Open questions

Resolved for S65. The new `rdocx-oxml` math module, document-wide `m:mathPr`,
Transitional OfficeMath namespace boundary, direct concrete facade references,
and bounded constructor surface are approved. Strict OfficeMath and legacy
Equation Editor content remain raw.
