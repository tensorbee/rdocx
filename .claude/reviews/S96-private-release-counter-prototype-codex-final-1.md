# S96, private release-counter prototype, Root final1

Reviewed the five-file TEMP overlay,1355patch lines, SHA-256
`119bfc79b414a385993aaa2dd893940cd4a6a637d321ff9c95c51f4b6647f85b`,
and its exact Issue437 case2 count observer and controlled preparation packet.
This is preparation review. No product build, benchmark or runtime acceptance
has occurred. The ordinary final release artifact remains separate.

Verdict: one defect, zero smells and zero nitpicks.

## Defects

### D1, the recursive serializer retry omits the new role argument

`/private/tmp/rdocx-s96-final-runner-recovery/release-counter-prototype/after/crates/rdocx/src/document.rs:9703`

`MainPartBlocks::serialize` now takes four arguments, including the explicit
diagnostic role. The `None if reused` branch still calls
`Self::serialize(frame, content, None)` with three. This fails type checking
regardless of runtime eligibility. Forward the same actual `counter_role`
through this retry, preserving the original retry policy. Audit both external
and recursive calls when a private diagnostic signature changes. Syntax-only
Rustfmt parsing and the existing three-producer static check cannot detect
this missing argument. A discriminating source control should reject the
preserved defective retry and accept the repaired one. Actual type checking
remains mandatory later against the final reviewed SourceH variant.

## Not found

No additional concrete correctness, Result-propagation, XML policy, consumer
AST or structural finding was established in the reviewed preparation. Main
writers are observed at actual calls, with copied, assembled and proof-model
roles distinct. The two existing Python test consumers retain their original
operations and assertions. The Issue437 addon preserves the original script
after stripping observations and labels its instrumented timings nonqualifying.
The repeated fake controls do not constitute a release-path measurement.

## Remaining boundary

All final-source applicability, full producer audit, actual optimized variant
type checking, debug-disabled proof, ordinary-artifact byte/error parity and
actual counters remain pending. Existing counter arithmetic must not be
represented as independently overflow-qualified by the new map checks.
This report ends review and returns to orchestration for separate repair.
