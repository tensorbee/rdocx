# S96 private text-history controller, Root final pass 1

Reviewed the new private controller and its separate D1 remediation after
actual independent passes1and2. Final run_text_child.py SHA256 is
`d0c488dce84fc03b85a78731ad1fef40bbaf397d9fc6b97a47d4e12a25c335bc`.
Verdict: zero remaining defects, zero smells and zero nitpicks. This controller
review does not add a history-source peer or reset any existing review cap.

Independent D1 correctly identified reopening a mutable spec path to label
previously parsed inputs. V2 requires the caller's reviewed spec digest,
reads and authenticates the raw bytes once, preserves them exclusively, then
parses that same buffer. Its receipt binds those bytes before preflight. The
original path may change later without changing the preserved execution spec.
V1 and its finding remain unchanged.

Root exercised12offline failure controls against the actual Python functions,
with process boundaries mocked. They cover wrong-pin refusal before preflight,
spec-path drift, existing output refusal, consumer rejection after successful
child completion, exact inventory encoding, timeout kill/reap and cleanup
exceptions. All pass. They are not real Native or child-process observations.

The finite controller permits only enabled/no-Weak text workers at100or400
edits, checks complete source/build/profile/binary/protocol identity before
spawning, monitors its direct child, and cleans it up in finally. Process
completion remains distinct from semantic consumer eligibility. Operational
limits produce incomplete diagnostics, never product acceptance. The caller
must inspect controller and consumer outcomes before any later pair or count.

Correctness, contract, panic, structure and evidence boundaries have no
remaining finding in this controller. Actual filled final-H spec, optimized
build, real process execution, observer calibration and all canonical font,
RSS and release obligations remain pending. No product or archived diagnostic
source changed in this review, and no Native process ran.
