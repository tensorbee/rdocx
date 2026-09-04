# F-231, correctness, pass 3

**Reviewed**: complete revised working diff against `be9a49b`, 8 files, 2,087
additions and 115 deletions, plus the pass-1 and pass-2 review records
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, sequence-prefixed TOC page numbers cannot be represented
`crates/rdocx/src/field.rs:2299`

The TOC evaluator accepts `\d` and stores its separator, but it has no branch or
structured field for the corresponding `\s` sequence identifier. A valid field
such as `TOC \o "1-3" \s chapter \d ":"` therefore reaches the unsupported
switch fallback at `crates/rdocx/src/field.rs:2366`. Without the identifier, the
accepted `entry_page_separator` value cannot tell F-232 which sequence number
to put before each page number. This loses one of the page-number selections
that the approved native rebuild request must retain.

### D2, unformatted decimal formula results expose binary floating-point noise
`crates/rdocx/src/field.rs:3656`

Every non-integral formula result is returned through `f64::to_string()` with
no Word-compatible result normalization. For example, `= 0.1 + 0.2` resolves
to `0.30000000000000004` instead of Word's displayed decimal result `0.3`.
The differential matrix exercises only `50% + 0.5`, whose exact result hides
this behavior. A supported formula can therefore disagree with the pinned
Word result while every declared test remains green.

### D3, the required published-crate dry run does not verify
`.claude/plans/F-231-design.md:104`

The public API risk rider requires `cargo publish --dry-run -p rdocx` and a
verified archive below 10 MiB. Running the dirty-worktree equivalent packages
33 files at 2.9 MiB, 522.3 KiB compressed, but verification fails with 17
compile errors. The packaged crate resolves the published `rdocx-oxml` and
`rdocx-layout` versions, which do not provide the math APIs used by this source
tree. The size assertion passes, but the mandated publish check does not, so
the public-crate contract is not yet complete.

## Smells

None.

## Nitpicks

None.

## Not found

Pass-2 D1 through D8 are closed. Explicit TOC sources no longer inherit the
bare heading default. TOC separators and custom style whitespace are
validated. `CASE` uses ITF14 payload and case-style rules. Extra TC and barcode
operands are rejected. Extended fields use the shared escaped-quote grammar.
Nested positional and switch operands resolve recursively. Compact and spaced
formula forms share the 512-token limit. The differential matrix now contains
ordered fallback diagnostics, and the focused barcode tests cross every
declared value, height, scale, rotation, and colour boundary.

No additional correctness, contract, panic, arithmetic-overflow, OOXML child
order, namespace-prefix, whitespace-preservation, unmodelled-XML,
dependency-family, runtime-oracle, trait, generic, wrapper-only type,
feature-flag, module, or file-creation issue was found. The complete
`rdocx-oxml` suite passed with 367 unit tests and one doctest. All 20 focused
`rdocx` field unit tests, both extended-field regression tests, and
`cargo check -p rdocx --all-targets` passed.
