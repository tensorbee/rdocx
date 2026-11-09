# S96, private release counter compiler review, final pass 3

**Reviewed**: Frozen five-file private overlay `1dafa26c79e3695c46b4ff9af0074f177f8e07cc98f5102a3eac01d9c4ce0ff9` and actual offline optimized native/binding library check.
**Verdict**: 1 defect, 0 smells, 0 nitpicks. Earlier final2 was source preparation only.

## Defects

### D2, release counter storage depends on a test-only Cell import

`/private/tmp/rdocx-s96-final-runner-recovery/release-counter-prototype/after/crates/rdocx/src/document.rs:7598`

The overlay makes the two F282 thread-local counters available in release, but their `Cell` type and constructor still rely on the test-only import at line11. Outside test configuration, the actual optimized library check fails with four E0425/E0433 errors at lines7598to7599. Qualify these two storage declarations with the real standard-library type or make only that required import available. Preserve the independent test-only failure-injection state.

The actual child exited101 in66.21seconds. The failed source and raw compiler diagnostics are frozen at `/private/tmp/rdocx-s96-private-diagnostic-typecheck-20261104/release-counter-current-base-check`. No runtime or final-source acceptance was attempted.

## Smells

None found in this bounded compiler review.

## Nitpicks

None found.

## Not found

Source custody and held-output preservation checks passed. No further type errors were reported before native compilation stopped. Binding compatibility remains unproved until the repaired native dependency allows its actual check to finish. This pass made no source changes.
