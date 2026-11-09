# F-X283, Codex final review 2

**Reviewed**: resumed CI material on worker b25f8abe, five changed or new source files, 758 additions and 24 removals. Frozen manifest 951642c510eef0284b9cfe4f78d046ff33f1ca7b33812be0a102890cd99a56cb and complete diff b5ddaf05f0ba6dc92753648b4641bec1dc5da384d80a6fbe335e3d8af17fb460. All seven source/plan/progress and eleven retained check hashes were authenticated. This is independent Root review of new material, without a fourth implementation peer or a counter reset.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. No code changed in this pass.

## Defects

### D1, skipping the required CI gate can qualify an unvalidated commit

`.github/workflows/ci.yml:1313` skips the required CI gate whenever a dispatch claims benchmark mode. The required status context remains CI gate at `docs/hld/15-build-and-toolchain.md:937`. The condition does not depend on successful benchmark admission or any normal check. GitHub treats a conditionally skipped job as successful for required-check purposes, as documented in [GitHub status checks](https://docs.github.com/en/pull-requests/reference/status-checks). Thus this route can emit a passing required context without running the normal validation it represents.

The separate strict-SemVer attempt exclusion does not repair that branch-protection consumer. The plan's benchmark-only exception itself needs correction. Preserve the ordinary required gate and its jobs. The smaller normal route is an optional benchmark job alongside the one final normal CI invocation, without special normal-job skips or a new SemVer exclusion. The source and authenticated release artifacts can still be pinned exactly. This is a technical CI-routing correction, not permission to weaken a gate or change repository settings.

## Smells

None found.

## Nitpicks

None found.

## Not found

No additional finding in benchmark AST/timer order, retained input geometry, artifact identity checks, XML preservation or structural rules. The runner remains the already reviewed 9d944b03 source. The 23 affected local tests, 19 synthetic runner controls and 15 selector controls are actual local observations, not hosted CI or product timing. Their current selector expectations do not test GitHub's required-check interpretation.

Authentic producer/interpreter proof, transfer-size admission, actual final Linux workloads and all original union gates remain pending. No dispatch, release, native or publication acceptance is claimed. End this review and return to orchestration for separate design reconciliation and remediation.
