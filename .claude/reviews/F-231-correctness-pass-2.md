# F-231, correctness, pass 2

**Reviewed**: complete remediated implementation diff against `be9a49b`, 6
files, 1,689 additions and 22 deletions, plus the pass-1 review record
**Verdict**: 8 defects, 0 smells, 0 nitpicks

## Defects

### D1, explicit TOC sources still include the bare-field heading default
`crates/rdocx/src/field.rs:3012`

Every TOC starts with built-in heading levels one through nine, and the `\t`,
`\f`, and `\u` branches never clear that default. Word applies the built-in
heading default to a bare TOC. An explicit source such as `TOC \f C` selects
matching TC entries unless `\o` is also present. The current outcome requests
both built-in headings and identifier `C`, so the subsequent F-232 rebuild will
include entries that the field did not select. The new `TOC \f` test at
`crates/rdocx/src/field.rs:4742` pins this incorrect combined selection.

### D2, TOC switch values are not normalized to Word semantics
`crates/rdocx/src/field.rs:3094`

The `\p` separator is copied without enforcing Word's one-character limit, so
`TOC \p "ab"` returns a validated request containing two characters instead
of applying the one supported separator or retaining the cache. Custom style
pairs at `crates/rdocx/src/field.rs:3188` are also split without trimming the
spaces allowed around list-separated pairs. A normal value such as
`"Heading 1,1, Appendix,2"` therefore requests a nonexistent style named
`" Appendix"`. Both paths produce structured outcomes that cannot rebuild the
same TOC as Word.

### D3, the CASE barcode alias does not receive ITF14 semantics
`crates/rdocx/src/field.rs:3298`

`CASE` is accepted as a barcode kind and is an alias of ITF14, but the `\c`
case-style branch accepts only `BarcodeKind::Itf14`. The payload validator at
`crates/rdocx/src/field.rs:3408` likewise applies the ITF14 digit constraints
only to that one enum variant. A CASE field with an invalid nonnumeric payload
is therefore reported as validated, while a valid CASE field using `\c EXT`
falls back as unsupported.

### D4, TC and barcode evaluators silently ignore extra positional operands
`crates/rdocx/src/field.rs:3236`

Barcode switch scanning jumps directly to the first token beginning with a
backslash, or to the end when none exists. It never checks tokens between the
required value and kind and that first switch. Consequently
`DISPLAYBARCODE value QR unexpected \t` returns a barcode outcome while
discarding `unexpected`. The same pattern at `crates/rdocx/src/field.rs:3120`
makes `TC "Entry" unexpected` a valid TC outcome. Unsupported syntax must
retain its instruction and cache with a stable diagnostic, not be silently
accepted.

### D5, the second instruction lexer rejects escapes accepted by the field grammar
`crates/rdocx/src/field.rs:2981`

The shared field lexer treats backslash-escaped quotes and backslashes inside a
quoted token as content, but this local lexer toggles quoting at every quote.
An instruction such as `TC "A\"B"` is validly parsed into the entry `A"B`,
then the local pass reports unclosed quoting and retains the stale display. It
also drops quoted-token metadata, so a quoted entry or barcode value beginning
with a backslash can be mistaken for the first switch. Re-parsing the raw text
with a different grammar makes valid fields depend on their spelling.

### D6, nested switch operands still bypass recursive resolution
`crates/rdocx/src/field.rs:3008`

Pass 1 added recursive resolution for TC and barcode positional operands, but
TOC still reparses only raw text, and every new family's switch values still
come from that raw token list. Nested fields are absent from `instruction.raw`.
Valid recursive forms such as `TOC \b { MERGEFIELD Scope }`,
`TC "Entry" \f { MERGEFIELD Kind }`, or a barcode `\s` value supplied by a
nested field therefore evaluate the nested child in document order and then
report that the outer switch has no argument. This remains contrary to the
approved requirement to reuse the recursive argument resolver for nested
operands.

### D7, formula support changes when equivalent whitespace is added
`crates/rdocx/src/field.rs:3759`

The generic instruction-shape check caps formula arguments at 256 before the
formula parser applies its documented 512-token bound. An expression made from
129 numbers and 128 operators is accepted when written without spaces because
the field lexer creates one argument, but the equivalent spaced expression
creates 257 arguments and falls back without reaching `FormulaParser`. Formula
syntax and the published resource boundary must not depend on insignificant
whitespace.

### D8, the declared differential and barcode-bound gates remain incomplete
`crates/rdocx/tests/regression_test.rs:1047`

The differential matrix contains only supported outcomes, although the
approved test explicitly requires exact ordered diagnostics as well. The
focused barcode cases at `crates/rdocx/src/field.rs:4763` exercise option enums
and two invalid option values, but never cross the 1,024-character value cap or
the height, scale, rotation, and colour bounds claimed in the HLD. The cache
matrix is broader now, but these missing cases leave D1 through D7 and several
declared limit branches able to regress while every named gate remains green.

## Smells

None.

## Nitpicks

None.

## Not found

Pass-1 D3 is closed by postfix percentage parsing and direct regressions.
Pass-1 D5 is closed by rejecting zero merge counters. Pass-1 D6 is closed by
recording both context-field and exhaustive-enum source breaks. The cache
matrix now exercises resolved text, all four structured outcome families,
pagination deferral, unavailable context, unsupported instructions, dirty
state, simple and complex fields, and retained XML scaffolding.

No additional panic or arithmetic-overflow path was found within the declared
limits. No OOXML child-order, namespace-prefix, whitespace-preservation,
unmodelled-XML, dependency-family, runtime-oracle, new trait, generic,
wrapper-only type, feature-flag, module, or file-creation issue was found. The
focused field unit tests, both extended-field regression tests, and
`cargo check -p rdocx --all-targets` pass.
