# F-011, Pin unit truncation behaviour

**Status**: completed
**Sprint**: S02
**Size**: S
**Depends on**: none

## Problem

The float constructors for `Length`, `Twips`, and `Emu` use Rust casts that
truncate toward zero, but their current tests use whole-unit inputs that would
also pass if a later extraction changed the code to rounding. Such a change
would shift document geometry and silently invalidate output comparisons.

## Spec reference

- `docs/hld/01-glossary.md`, "Units".
- `docs/hld/11-migration-plan.md`, "Preserve behaviour, do not improve it".
- `docs/hld/12-testing-strategy.md`, "New tests the extracted crates need".

## Approach

Add tests to the existing `length.rs` and `units.rs` test modules. Use finite
positive and negative values whose scaled results lie beyond the half-unit
boundary so `.round()` would produce a different integer. Cover every float
constructor for `Length`, `Twips`, and `Emu`, plus the exact integer constructors
named by the story. Do not change production conversion code.

## Rejected alternatives

- Replace casts with an explicit truncation helper. The current code is clear,
  and a wrapper would add production structure to a tests-only story.
- Use only positive values. Rust casts truncate toward zero, so negative cases
  are necessary to pin the full behavior.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| unit | `length_float_constructors_truncate_toward_zero` | Every `Length` constructor retains positive and negative truncation |
| unit | `twips_float_constructors_truncate_toward_zero` | Inch, centimetre, and point `Twips` constructors truncate rather than round |
| unit | `emu_float_constructors_truncate_toward_zero` | Inch, centimetre, and point `Emu` constructors truncate rather than round |

The backlog test gate is the set of pinning tests, which fail if truncation
becomes rounding.

## HLD impact

- `docs/hld/11-migration-plan.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- Unit conversion. Read the glossary and deliberately wrong behavior, preserve
  `as i64` or `as i32` truncation, run focused rdocx and rdocx-oxml tests, and
  require an unchanged harness.

## Hash harness

Expected to remain unchanged. This story adds tests only.

## Implementation checklist

- [x] Choose positive and negative fractional vectors that distinguish
      truncation from rounding.
- [x] Pin every `Length` constructor in its existing test module.
- [x] Pin every `Twips` and `Emu` constructor in the existing units test module.
- [x] Demonstrate the gates would fail under rounding, then restore production
      code unchanged.
- [x] Run focused unit tests and the hash harness.

## Open questions

None.
