# S96, private release counter binding compiler review, final pass 5

**Reviewed**: Exact v2 source and actual optimized native and binding library check.
**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D3, native counter exports remain disabled for the release binding

`crates/rdocx/src/lib.rs:69`

The private overlay enables the native producers and Python consumers, but leaves the crate-root `take_reopen_work` and `take_staging_work` reexports behind `cfg(debug_assertions)`. The native library now checks, closing D2, but the actual release binding check fails with two E0425 errors at private `rdocx-py/src/document.rs:110` and line119. Include the two specific reexports in the private overlay. The ordinary shipping source and its conditional exports must remain unchanged.

Actual child exit101 in10.87seconds and unchanged inputs are recorded at `/private/tmp/rdocx-s96-private-diagnostic-typecheck-20261104/release-counter-v2-current-base-check`. The missing sixth source file is part of this producer-to-export-to-binding chain. Audit that complete chain before another check.

## Smells

None found in this bounded compiler review.

## Nitpicks

None found.

## Not found

D2's missing Cell type errors are absent. No further native type errors were reported. Compiler custody and held-output checks passed. No runtime or final-source acceptance occurred. This pass made no source changes and returns to separate remediation.
