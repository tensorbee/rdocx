# S96 private COW controller and consumer, Root final pass 2

Reviewed the complete three-file D1 remediation against preserved v1 and the
actual build receipt format. Verdict: one residual D1 defect, zero smells and
zero nitpicks. This is the same bounded controller repair, not a history peer.

The shared verifier correctly closes the missing spec/identity, executable,
asset-role, inventory/document, profile, PID, stack and cleanup joins from
final1. Direct re-pinned receipt controls now cover these paths. One complete
source-to-build join is still absent.

## D1 residual, child source inventory is not bound to the actual build inventory

`consume_cow.py:190` through207validates successful stable build flags, Git
base, Native document digest and artifact/profile, but does not compare the
build's complete source inventory digest with execution_identity's
source_manifest_sha256. `controls.py:94` explicitly permits a positive build
receipt with no source_before_sha256field at all.

The existing actual diagnostic build receipt already has source_before_sha256.
Git SourceH and the Native document digest do not identify every private
instrumented build input. Another file can differ between build and launch
while the source remains stable during each phase. A re-pinned child manifest
then passes without being the manifest used to compile the artifact.

Require the actual build source_before_sha256to equal the complete manifest
digest used by the child, using the same exact inventory bytes as the reviewed
build envelope. This preserves existing actual receipt vocabulary. Add direct
missing and mismatched build-manifest controls, including a valid re-pinned
child inventory containing a changed non-Native-document source file. Retain
the successful source-stable build and all other repaired joins. This asks for
evidence association, not independent proof of a historical process or a new
product acceptance threshold.

## Other boundaries

The finite command, raw three-lane protocol, artificial ownership labels,
full semantic comparisons, strict JSON, direct-child cleanup and separation
of process success from consumer eligibility have no additional established
finding. The shared verifier is substantive and serves both real consumers
without a new module or framework. Reported118controls remain offline-only.
No Rust build, Native matrix, final-H, font/text history, RSS, neutrality or
transitive payload result is claimed. Root changed only this review record.
