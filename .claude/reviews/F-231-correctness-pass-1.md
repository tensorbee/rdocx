# F-231, correctness, pass 1

**Reviewed**: complete working diff against `be9a49b`, 6 files, 1,262
additions and 22 deletions
**Verdict**: 7 defects, 0 smells, 0 nitpicks

## Defects

### D1, common valid TOC forms are rejected or cannot be represented
`crates/rdocx/src/field.rs:2978`

The final selector check rejects a bare `TOC`, although that valid field builds
from the built-in heading styles. The `\o` branch at
`crates/rdocx/src/field.rs:2927` also requires a range even though omitting the
range selects all outline levels. Likewise, `\f` always requires an identifier,
so the supported form that selects all TC entries is indistinguishable from no
`\f` selection in `TocField`. These inputs retain stale displays instead of
returning the structured TOC request required by the approved family boundary.

### D2, barcode switches use a different grammar from Word
`crates/rdocx/src/field.rs:3145`

The parser treats `\x` as a value-bearing switch, accepts arbitrary text for
`\p`, and has no `\c` branch. Word defines `\x` as a flag, limits `\p` to the
point-of-sale styles, and uses `\c` for the ITF14 case style. The `\q` branch at
`crates/rdocx/src/field.rs:3158` accepts `L`, `M`, `Q`, or `H`, while the Word
field accepts correction levels 0 through 3. Supported fields such as
`DISPLAYBARCODE value UPCA \x` therefore fall back, while unsupported switch
values can be returned as validated `BarcodeField` outcomes. The public
`fix_code: Option<String>` shape at `crates/rdocx/src/field.rs:112` also cannot
represent the actual flag contract faithfully.

### D3, formula percentage is implemented as binary remainder
`crates/rdocx/src/field.rs:3328`

The multiplicative parser consumes `%` exactly like `*` and `/`, then requires
a right operand and applies floating-point remainder. Word formula fields use
`%` as a postfix percentage operator. As a result, `= 50%` retains its stored
display instead of resolving to `0.5`, while the invalid infix expression
`= 5 % 2` resolves to `1`. This contradicts the pinned Word formula contract.

### D4, extended non-text fields bypass recursive argument resolution
`crates/rdocx/src/field.rs:2163`

Barcode evaluation reparses `instruction.raw` and never resolves the typed
`FieldArgument` values. Nested fields are deliberately absent from that raw
text while being retained in the recursive grammar at
`crates/rdocx-oxml/src/text.rs:4935`. A complex field such as
`DISPLAYBARCODE { MERGEFIELD Code } QR` evaluates the nested field for the
ordered result list, then sees no barcode value and retains the outer cache.
The approved approach explicitly requires the existing recursive argument
resolver for nested operands.

### D5, the public one-based merge counters accept zero
`crates/rdocx/src/field.rs:2108`

Story state copies both optional counters without validating their documented
one-based invariant. Supplying `Some(0)` therefore returns
`MailMergeControl::RecordNumber(0)` or `SequenceNumber(0)` at
`crates/rdocx/src/field.rs:2148`, and `NEXT` advances an invalid zero record to
one. Invalid explicit context should retain the stored display with a stable
diagnostic rather than produce a validated control outcome.

### D6, the public API impact is misclassified
`docs/hld/10-bindings-spec.md:494`

The HLD calls the new `FieldEvaluationContext` fields additive. Adding public
fields to an existing Rust struct breaks every external struct literal that
does not use update syntax, just as existing repository literals had to be
changed in this diff. The risk rider requires the pre-1.0 source impact to be
stated accurately. The context-field addition is a source break alongside the
new exhaustive enum variants.

### D7, the required fallback and cache-update matrix is not exercised
`crates/rdocx/tests/regression_test.rs:1155`

The cache regression contains one structured barcode plus one resolved formula.
It never updates a TOC, TC, mail-merge control, unavailable-context result, or
unsupported instruction. The focused unsupported assertions at
`crates/rdocx/src/field.rs:4453` check only the outcome variant and do not pin
the promised stable diagnostics or dirty and cache behavior. This leaves the
approved requirement for cache preservation across every structured and
unsupported result unproved, and it lets the TOC, barcode-switch, nested-value,
and zero-counter failures above remain green.

## Smells

None.

## Nitpicks

None.

## Not found

No additional panic or arithmetic-overflow path was found within the declared
token and nesting limits. No OOXML child-order, namespace-prefix, whitespace,
unmodelled-XML preservation, dependency-family, runtime-oracle, new trait,
generic, wrapper-only type, feature-flag, module, or file-creation issue was
found. The focused `rdocx` field unit tests and both extended-field regression
tests pass.
