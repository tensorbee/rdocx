# S96 private COW controller and consumer, Root final pass 3

Root reviewed the complete residual D1 delta against v2, actual build receipt
vocabulary and the added direct receipt controls. Verdict: zero defects, zero
smells and zero nitpicks remaining. D1 is closed for controller preparation.
This is the existing bounded repair, not a history-source peer or product run.

The shared substantive verifier now requires build.source_before_sha256to
equal execution_identity.source_manifest_sha256. Both process preflight and
standalone receipt consumption use it. The positive mock includes the actual
field. Missing/wrong values and a genuinely changed non-Native source file
with a re-pinned child manifest are rejected. That discriminator retains the
same Native document and Git identity, so stale digest checks alone do not
explain rejection.

The previously accepted complete spec/identity, source/document, executable,
asset-role, optimized build/profile, PID, stack, cleanup and raw-protocol joins
remain unchanged. Controller and protocol bytes match v2. Final consumer is
`b59456db2eee75863e9777711fc76b30a0e470adcafb0a06a5be1528a81db929`.
Reported121controls are offline source-shaped/mock evidence. Independent
mechanical replay remains pending and no Rust or Native process ran.

Correctness, contract, panic, tests, structure and evidence boundaries yield
no new finding in this repair. Historical PID honesty still comes from the
authenticated execution producer, not a hash. Final-H compilation, real COW
branch observations, fixed-font/text evidence, RSS and shipping calibration
remain pending. Neutrality, natural ownership, transitive payload destruction
and bounded-history acceptance are not supplied. Root changed only this review.
