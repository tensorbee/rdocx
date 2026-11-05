# S96 additional text-history consumer, Root final pass 1

Root read the330line consumer, its finite receipt contract and synthetic
controls. This is a bounded Root follow-up, not another history peer pass.

Verdict: one evidence defect D2, zero smells and zero nitpicks. No product,
archive, compiler, native process or source fix ran during this review.
The unchanged v4 font helper and earlier text-source D1 remain separate.

## D2, contradictory lifecycle observations pass validation

At consume.py:154, owner snapshots validate liveness vector shapes without
linking them to watched_total. At consume.py:215, memory lifecycle counts are
checked only for shape, final zero and upper bounds. The corresponding current,
clone-live and reopen-live snapshots are not cross-checked.

Root independently exercised the actual validate_records function in memory
with the source-shaped positive synthetic builder. All three contradictory
rows were accepted: current changed to all-zero while terminal_saved_semantic
return liveness stayed nonzero, clone_live changed to zero while its snapshot
stayed nonzero, and an after-call snapshot claiming999999live allocations in
each category while final watched_total remained2in each category.

These are inconsistent observations from the same watcher, not an arbitrary
retention threshold or a payload-destruction inference. Bind memory current
to terminal_saved_semantic_return, clone_live to clone_semantic_return and
reopen_live to reopen_semantic_return. All snapshot liveness counts must fit
the final accumulated watcher totals. Actual per-node activation counts must
fit the snapshot's live watched-node count where that invariant applies. Add
meaningful independent negative controls and retain the positive/no-Weak and
existing phase/custody controls. Do not guess equality at phases lacking an
actual source snapshot or silently fabricate drop observations.

## Inspected controls and limits

The implementation checks actual exit before log content, exact named PASS,
policy, seven/eight-field semantic schemas, all29identifier fields, public
count chains, package/relationship invariants, separate query ranges and
workload-only creation. It retains COW branch differences with null neutrality.
Offline mock controls and receipt failures are explicitly synthetic.

No further correctness, contract, panic, XML or structural finding was
established in the bounded read. Mechanical support authenticated all inputs
and replayed261controls in a separate directory with no product execution.
Those controls did not include D2. Actual final-H compilation, runtime, history
qualification, COW calibration, payload attribution and fixed-font/RSS riders
remain pending. The review ends here for separate remediation.
