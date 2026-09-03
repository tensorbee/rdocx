# F-228, all aspects, pass 1

**Reviewed**: uncommitted worker diff, 11 files and 2,505 added or changed lines
including the untracked grammar module
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, recursive parsing has no depth ceiling

`crates/rdocx-oxml/src/math.rs:223`

Every nested argument calls `MathExpression::from_raw` recursively without a
depth budget. An untrusted document with deeply nested fractions or scripts can
overflow the thread stack instead of failing with a normal parse error.

### D2, modeled property containers discard unsupported content

`crates/rdocx-oxml/src/math.rs:529`

The fraction parser reads `m:fPr` values but does not retain the property
container. The writer constructs a fresh property container at line 559. This
drops unknown attributes and children owned by `m:fPr` after any typed equation
serialization. The same pattern appears in script, radical, matrix, n-ary,
delimiter, and accent properties.

### D3, math run text decoding is not fallible

`crates/rdocx-oxml/src/math.rs:425`

Math run parsing calls the shared loss-tolerant text helper. That helper maps an
undecodable text event to an empty string. The approved contract requires
fallible text decoding, so malformed encoded equation text can be published as
partial or empty content instead of rejecting the typed projection.

### D4, incomplete grammar productions become authored defaults

`crates/rdocx-oxml/src/math.rs:523`

The fraction parser accepts missing, duplicated, and out-of-order numerator or
denominator children, fills missing arguments with empty defaults, and the
writer then emits a different canonical equation. Equivalent unchecked
production parsing exists for scripts, radicals, matrices, limits, n-ary
operators, delimiters, and accents. Malformed productions must remain opaque at
their owning expression boundary rather than acquire invented typed meaning.

## Smells

None.

## Nitpicks

None.

## Not found

No additional contract, panic, paragraph-order, legacy-boundary, test-gate, or
structural findings were found beyond the defects above. The change introduces
no trait, generic public parameter, crate, dependency, feature flag, builder,
or dynamic dispatch.
