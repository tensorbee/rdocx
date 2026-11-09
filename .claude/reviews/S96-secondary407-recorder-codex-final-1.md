# S96 secondary407 recorder, Root final1

Root read the complete273-line driver and314-line isolated controls. Driver
hash is596f8cf6a31ee77b0f8add88c14eb517b7588897c6949e539f5f8e3fe84ca223.
All six preparation artifact hashes and source-bound32saved controls match.
This reviews prepared measurement code, not a product benchmark. No product
was imported, built or timed, and no real clock measurement occurred.

Verdict: one numeric failure-reporting defect, zero smells and zero nitpicks.
Fix before acceptance use. All finite protocol mechanics otherwise match the
Root pin in current F-X282 plan a68fade45104aa88c60646e9e088fe64212574eebbec091890bb7e48fe639b30.

## D1, derived nonfinite elapsed value enters a failure receipt

In /private/tmp/rdocx-s96-final-runner-recovery/original407-secondary/record_secondary407.py:207
only start and end are tested for finiteness. At210their difference is appended
without checking that derived value. Two finite endpoints can overflow their
subtraction. The later scorer rejects it, but the retained raw report then
contains infinity and strict JSON publication fails instead of retaining a
readable incomplete failure record.

Root reproduced this using the actual run_secondary core with isolated fake
document, clock and postinspection functions. Endpoints -1e308and1e308 produce
an appended infinity. A subsequent synthetic error leaves an incomplete report
that json.dumps(..., allow_nan=False) cannot encode. No real clock or product
behavior is implicated by this synthetic failure-path observation.

Compute and validate elapsed together with the endpoints before appending the
sample. Preserve invalid endpoint/difference spelling in active_sample using
JSON-safe diagnostics, retain earlier valid samples and fail incomplete. Add an
actual core or guarded-main fake control for the derived-overflow case and
verify its failure output can be published. Do not substitute zero, alter valid
timing arithmetic or change the sampler/threshold. Preserve the pre-repair
source and controls so the reviewed state remains reproducible.

## Qualified preparation

The original factory and RPR are extracted from authenticated source without
executing its imports or primary benchmark. The matrix uses84fresh documents,
84single-call windows and42same-target warmups, with the specified parity order
and seven separate values per method/size/state. Handle priming, rich secondary
removal, mutation history, returned-value creation/drop and separate audit
fixtures are disclosed. Semantics and serialization run after each clock.
Original mean3 remains a separate unchanged script and is not pooled.

The guarded entrypoint checks artifact/source/release/interpreter identity before
product import and checks actual imported module locations afterward. Context
claims do not prove exhaustive inputs or monitor a quiet machine. Final actual
context, release3.12run, all raw observations and other native/memory/full-sprint
gates remain pending. No source repair occurred in this pass. This ends review
and returns the bounded defect to the preparation author.
