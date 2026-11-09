# F-X283 A283-14, GLM trial Codex review 2

The stopped first repair finished0in377.062seconds. Review is limited to the
same A283-14 story. No new peer cycle or story is started.

## Defects

D1 is closed. Exclusive creation now precedes the cleanup-ownership flag, and
only an owned temporary is removed (scripts/benchmark_f283.py:46 and :52).
The finite preservation, link-failure cleanup and exact serialization controls
exercise that behavior (scripts/test_sprint_workflow.py:1876).

D2's production repair is correct, but its regression remains incomplete.
The same byte snapshot is now hashed and decoded (scripts/benchmark_f283.py:493).
The new test changes the file before invocation and supplies an earlier hash
(scripts/test_sprint_workflow.py:1902). The original double-read implementation
also rejects that input at its first hash check. Thus this test does not catch
a return of the reported changed-between-reads behavior. Add a deterministic
control which changes the data observed after an initial successful identity
check, then asserts that any emitted bundle hashes to config_sha. A controlled
reader or identity-check substitution is sufficient, with no timing-sensitive
thread or real race. Prove this control fails with the former double-read
behavior and passes with the current source. Retain the ordinary stale-digest
and positive controls. This continues D2's evidence obligation, not a third
production defect.

## Smells and nitpicks

No additional smells or nitpicks established. The first repair encountered an
invalid new success fixture which reused the preserved partial path. GLM's own
test run exposed it, and it corrected the fixture before reporting green.
That routine author test iteration is not counted as a Root review finding.

## Stop

A second focused repair may strengthen only this regression and progress notes.
No further production change is requested. All final CI, Linux and union gates
remain pending. Review changed only this record and ends here.
