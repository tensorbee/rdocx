# F-X283 A283-14, GLM trial Codex review 1

The one-story GLM candidate is not accepted yet. Review covers its four allowed
paths against52db2dd422924970534d04c8c1bf8e140e555cde, plus the Root-owned
A283-14 plan addition. No product, CI workflow or controller change is present.
The actual CLI finished0in1041.571seconds. Its19synthetic checks and8targeted
tests are recorded in the actual session stream. Root inspected source and
assertions rather than treating their aggregate count as acceptance.

## Defects

D1. The new atomic writer deletes a pre-existing partial file after exclusive
creation refuses it. The finally block executes even if temp.open fails because
the path already exists. This can destroy another unfinished report or request,
and also affects the existing benchmark report writer. Cleanup must remove only
a temporary file this invocation actually created. Add a preservation control
for an existing partial path and a cleanup control after failure following
successful creation. Citation: scripts/benchmark_f283.py:45 and :49 in the
OpenCode worktree. Root's in-memory failure probe observes one unlink call after
exclusive open raises FileExistsError. The first mock lacked a string suffix and failed during fixture setup. After correcting that setup, the stated unlink observation reproduced. No existing file was deleted by review.

D2. Transfer authentication and payload consumption use separate file reads.
checked_file hashes a path, then prepare_dispatch rereads it to obtain the text.
If it changes between those reads, output can contain bytes whose digest differs
from config_sha and the supplied approval hash. Read one byte snapshot, hash
that snapshot and decode the same snapshot. Retain path eligibility checks and
prove that changed supplied bytes cannot be emitted under an earlier digest.
Citation: scripts/benchmark_f283.py:486 and :487 in the OpenCode worktree.

## Smells and nitpicks

No additional smells or nitpicks established. Positive decimal-string IDs match
the existing CI transfer predicate, while bool IDs are correctly refused. The
plan's positive-integer wording should be clarified as received integer values
or decimal-string representations, without changing that existing contract.
This is a Root prompt precision issue, not a new implementation defect.

## Evidence and remaining gates

Exact transfer preservation, typed booleans, eight input keys and the complete
ASCII request byte boundary are covered by finite local tests. The model itself
caught and corrected three issues in self-review, including nondiscriminating
negative tests and a self-test option regression. Existing workload script
literals remain unchanged in this candidate. High effort was configured and
requested, but backend application is not independently established.

This is the first targeted trial review, not a new F283 peer phase or a reset
of its used peers. A focused author repair of D1 and D2 follows separately.
Final hosted CI, Linux benchmarks, final union riders and release acceptance
remain pending. Review changed only this review record and ends here.
