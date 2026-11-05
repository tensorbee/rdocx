# S96 original PowerPoint recorder, Root final1

Root read the complete264-line recorder,157-line isolated controls and exact
original package script. Recorder hash is
7191b524bb67725dd2b0d445ec4aebdf2a814c7910c29555c713e0976e7dd811.
Root independently authenticated nine immutable, three authority and four
artifact hashes, including the source-bound saved23-control outcome. No product
was imported, built or timed. No original benchmark ran during this review.
The reviewed source is preserved under
/private/tmp/rdocx-s96-final-runner-recovery/pptx-original-controls/v1-snapshot.

Verdict: one failure-reporting defect, zero smells and zero nitpicks. Remediate
before acceptance use. Source-level sampler/fixture preservation is otherwise
consistent with this finite preparation scope, not real performance evidence.

## D1, rejected nonfinite observation cannot publish its failure receipt

record_original_controls.py:144 rejects nonfinite elapsed observations, but the
original samples remain in report. main catches that error at251and attempts
json.dump(..., allow_nan=False) at256. Serialization then raises outside the
handler after opening and partially writing the destination. The result is
truncated invalid JSON, losing the promised readable failure receipt and raw
partial observations.

Root reproduced this through actual main with an isolated fake workload that
records NaN then raises the sample error, and an in-memory destination. No files
or product work were involved. main raises the JSON compliance ValueError, and
the captured output cannot be parsed as JSON. This is a recorder failure-path
test, not a product or benchmark RED.

Represent invalid numeric observations explicitly in the failure record without
emitting nonstandard JSON or inventing zero measurements. Preserve their raw
diagnostic spelling, earlier valid samples, actual error and incomplete/nonzero
status. Stage complete JSON encoding before opening the output destination so
encoding errors do not leave a truncated receipt. Add an isolated main-level
control for this failure path. Existing finite sample/mean/threshold arithmetic
and immutable originals must remain unchanged.

## Preserved scope

The AST observer wraps only the already-computed original elapsed operand.
Stripping it recovers the original AST, retaining enclosing10-call shape and
five-call package timers. It does not insert individual clocks. Original four
package rows and stdout precede the exact matching10-slide/600-payload addon.
The strict Q14 comparison uses the matching payload class in milliseconds.
Plain/media shape profiles and package remain three independent observations.

Real Python3.12, final release/source/artifact/installed module linkage and
complete build-input proof remain mandatory. Context assertions do not certify
those external facts. Instrumentation bookkeeping occurs outside each timer
but may affect between-window state, as the readiness note discloses. The
separate248catalogue, original407, memory, Markdown and final sprint riders
remain pending. No source fix occurred in this review. This ends the pass and
returns remediation to the preparation author.
