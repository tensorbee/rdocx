# S96 private history connection, Root final pass 2

Reviewed the428line v2source delta, finite root-role/semantic helpers, comparer,
source/mock controls and independent connection peer2. Before source is
547866189c910bcc8f15ad87ea3680c4a60b81337691b61d12b8bf8829d5363b.
After source is8975569861ac78ac647a970a7292d365154cf20807eaeb37df8df1ef693a38b9.
Root authenticated packet/input hashes and exact complete source reconstruction.

Verdict: two defects, zero smells and zero nitpicks. These are offline evidence
consumer defects, not demonstrated product or Native COW failures. No archive
installation, compiler, build or product execution occurred. This final pass
preserves peer2finding identities and does not reset prior phase caps.

## D1, complete Normal policy is not required

`/private/tmp/rdocx-s96-final-runner-recovery/memory-reconstruction/history-runtime-connection-v2-20261104/compare.py:24`

The consumer checks mode and equality across lanes without requiring the full
fifteen-false Normal policy vector/report. Root independently reproduced the
peer's in-memory counterexample: set the same omission flag true in both
otherwise passing records. The consumer still returns semantic_equal=true.
The actual Rust driver correctly asserts Normal. Repair the finite consumer and
its both-lanes-wrong controls, not product policy installation.

## D2, required terminal saved-package success can be absent

`/private/tmp/rdocx-s96-final-runner-recovery/memory-reconstruction/history-runtime-connection-v2-20261104/compare.py:31`

Root independently reproduced both cases: both terminal_saved ZIP observations
set to None, or both otherwise successful ZIP observations containing the same
content_types Err. Each still returns semantic_equal=true. The terminal phase
must contain the successful actual saved-package/part/relationship transcript
and successful content-type writer outcome. None remains valid only for the
specified clone/reopen phases. The source emits those actual saved bytes and
records errors explicitly. Reject missing/error records in the comparer and
add symmetric-absence/error controls without calling to_bytes again.

## Not found and remaining boundary

No additional source ownership, drop-order, disabled-Weak gating, policy
installation, frozen font lifecycle, complete mandatory-field projection or
pointer-provenance defect was established in this increment. The direct scalar
COW seam is unchanged. Weak construction is gated before argument evaluation.
Role snapshots use actual fields and per-observation identity. They do not
prove moved/transitive payload destruction.

Meaningful COW counts and branches remain visible. Pre-control Weak-observer
comparison is limited to0through N. The enabled-only historical strong control
leaves later allocation neutrality explicitly unqualified. Terminal writer
diagnostics add disclosed transient work and are not deciding RSS/timing.
Actual compiler/runtime, fresh activation, failure calibration, complete
attribution, final SourceH and bounded-history acceptance remain pending.

This pass ends here. Separate remediation may repair D1/D2, followed by the
remaining third connection peer and Root final. No fourth peer or earlier
bridge-review reset is authorized.
