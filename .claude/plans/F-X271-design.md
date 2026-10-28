# F-X271, Restore pinned viewer downloads and planning policy

**Status**: approved
**Sprint**: S96
**Size**: M
**Depends on**: none

## Problem

`scripts/install_pinned_libreoffice.py:21` retrieves the Linux oracle from a stable release URL that now returns 404. `.github/workflows/ci.yml:308` does the same for macOS. The prior CI failures are retrieval failures, not completed product tests. `scripts/test_sprint_workflow.py:9024` still expects the old future sprint placement set after the authorized roadmap shift.

## Spec reference

- `docs/hld/15-build-and-toolchain.md`, "Toolchain pinning", "Deterministic rendering" and "CI job matrix".
- `docs/hld/12-testing-strategy.md`, external oracle and deterministic rendering requirements.
- `docs/hld/14-development-backlog.md`, "F-X271, Restore pinned viewer downloads and planning policy".
- `docs/sprints/SPRINT_PLAN.md`, "Sprint S96, Contribution and issue repair", CI blockers and combined gate.

## Approach

Change retrieval to the official old-release archive for the same 26.2.5.2 binaries. Keep the independently retained SHA-256 pins and exact executable identity. Linux archive name becomes `LibreOffice_26.2.5.2_Linux_x86-64_deb.tar.gz`, under `https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.5.2/deb/x86_64/`. macOS uses the corresponding `mac/aarch64/LibreOffice_26.2.5.2_MacOS_aarch64.dmg`. Authenticate actual downloaded bytes before adopting either URL. A HEAD200 is not proof.

Retain the existing streaming download and Linux extraction bounds, traversal/non-file refusal, required package checks and absent installation-prefix guard. Add a macOS download size bound before mounting while preserving TLS, readonly mounting, checksum and executable identity checks. Keep installation authority platform-specific. Do not run Linux installation on this macOS host. Hosted Linux and macOS outcomes remain required.

Update only the existing roadmap policy's expected future set from 97/98/99 to 99/100/101. Keep its duplicate, missing-story, dependency and cross-record consistency assertions. Add focused cases to the existing `scripts/test_sprint_workflow.py` carrier. No new source file, module, dependency, feature flag or public API is proposed.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| Linux oracle identity | `download_archive(Path) -> None`, `safe_extract(Path, Path) -> Path`, `verify_soffice(Path) -> None`, `install() -> None` | Existing Linux installer used by Test, MSRV, Word/presentation fidelity | Retrieve, hash, bounded extract, install, version, actual gate |
| macOS oracle identity | Existing rpptx installed-Python CI step | Existing DMG checksum and exact soffice identity | Retrieve, bound, hash, readonly mount, installed suite |
| Sprint placement | `test_m23_m24_roadmap_has_no_duplicate_or_dangling_story` | Same HLD/backlog/roadmap inputs, revised authorized set only | Current plan and final completion metadata |

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| Linux bytes | SHA-256 `2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`, max 224 MiB | Retained source installer at Base `24af95d3`, official TDF26.2.5.2 archive | Actual bytes pending, wrong hash/oversize must refuse |
| Linux extraction | At most 256 members, at most 256 MiB extracted, expected root and core/impress packages, regular files/directories only | Existing `safe_extract`, installer lines 73 to 123 | Traversal, link, missing package, count and size boundary controls |
| macOS bytes | SHA-256 `c99fb4fe574437fc4cb820a4ca15271bca325920861f7139858b36d7f9df78ad`, proposed 320 MiB download cap | Retained CI source at Base `24af95d3`, official TDF26.2.5.2 archive | Actual bytes and macOS installed outcome pending |
| executable | Exact `LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb` | Existing independently pinned Linux/CI identities | Version mismatch refuses before use |
| roadmap | Story placements exactly once, future audited whole sprints 99/100/101 | Authorized S96 planning diff and unchanged audited F240 to 310 set | Current-Base stale-set RED, updated set GREEN, mutation controls |

## Dependency readiness

No formal dependency. Tool restoration is the execution prerequisite for the final S96 oracle union. Actual external CI is final acceptance debt, not a locally passed provider gate. F-X272/F-X273 may implement their local regressions after this tooling source is reviewed, without claiming the final oracle passed.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A1 | Official Linux old-release archive | Actual bytes equal retained checksum and correct root/packages | Installer, Linux CI | Focused archive retrieval contract case | Real Linux install and exact version |
| A2 | Official macOS archive | Actual bytes equal retained checksum, bounded retrieval, exact executable | rpptx Python macOS CI | Existing workflow contract plus cap assertion | Actual hosted macOS runtime gate |
| A3 | Wrong digest/version, oversize, traversal, non-file, missing package | Refusal before installation/use, no escape writes | Installer and workflow | Existing failure controls in `test_sprint_workflow.py` | Integrated policy suite |
| A4 | Current revised roadmap | Unique complete placements accepted, duplicate/missing/dependency mutations rejected | Sprint policy | `test_m23_m24_roadmap_has_no_duplicate_or_dangling_story` | Final-state metadata policies |
| A5 | Actual integrated source | Previously retrieval-blocked jobs execute and succeed with exact pins | Test, MSRV, Word fidelity, rpptx Python, CI gate | None, hosted-only evidence | New source-bound run/job IDs required, preserve old failures |

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | Linux and macOS actual downloads | Different platform hashes, same exact release/build identity | Changing URL to latest or trusting HEAD |
| C2 | Good archive versus digest/version mismatch and each extraction bound | Only authenticated safe inputs usable | Disabling integrity or bounds to get CI green |
| C3 | New roadmap set versus duplicate/missing/dangling stories | Valid plan accepted, mutations fail | Deleting policy assertions |

## Rejected alternatives

- Upgrade LibreOffice, disable TLS/checksums or skip fidelity. Each changes the oracle contract.
- Copy donor CI or archive claims. Neither qualifies the integrated source.

## Test plan

**Test gate**: regression, verbatim from the backlog.

Local scoped: run `python3 -m unittest scripts.test_sprint_workflow.SprintWorkflowTests.test_workspace_viewer_jobs_install_pinned_libreoffice scripts.test_sprint_workflow.SprintWorkflowTests.test_pinned_libreoffice_installer_enforces_runtime_guards scripts.test_sprint_workflow.SprintWorkflowTests.test_m23_m24_roadmap_has_no_duplicate_or_dangling_story`, plus the new archive-route assertion in that class. Run the stale-roadmap case on current Base for RED, then the repaired case for GREEN. Exercise new retrieval assertions on old source for RED without calling a network failure semantic product evidence. Existing security controls are retained controls, not invented RED claims.

Final sprint: full `python3 -m unittest scripts.test_sprint_workflow`, actual authenticated supported-target installs and named hosted jobs, then `/verify --full` and the S96 union. If the new source must be hosted first, use only the validated external-CI-source-availability route. No completion ledger or dependency unblocking on a source push.

## HLD impact

- `docs/hld/15-build-and-toolchain.md`

## Risk routing

External oracle comparison adds exact version/build/input provenance and actual supported-target installation. Completion/status policy adds final affected metadata checks. Shared tooling consumers add Linux/macOS route checks. No product parser, unit conversion, public binding or crate graph change. Final hash/pixel union still applies to S96.

## Hash harness

Expected unchanged,49 hashes and 7 deterministic pixel cases. This tooling/policy story authorizes no product baseline delta.

## Implementation checklist

- [ ] Authenticate official archive bytes against both retained platform hashes.
- [ ] Repair URLs without changing oracle identity or Linux bounds.
- [ ] Add macOS retrieval bound and preserve readonly mount checks.
- [ ] Repair only the outdated roadmap expectation, retain negative controls.
- [ ] Run unique scoped checks and independent zero-finding microscope.
- [ ] Update listed HLD intent before final gate.
- [ ] Retain actual hosted acceptance as pending until observed.

## Open questions

None requiring a user product choice. Archive authentication and supported-target CI are unavailable observations, not questions to waive.

## Contract closure

Existing callable installer signatures remain unchanged. Proposed 320 MiB macOS bound exceeds the historical 297407265 byte HEAD length and must be checked against actual bytes. No source is approved or started. Exact existing scoped test names are recorded above. Actual archive authentication is the first implementation checklist item and must pass before either replacement URL is adopted in source. Claiming may precede this read-only authentication step.

## Intake identity

Read-only GitHub refresh on 2026-10-28 found no body or contribution-head changes from the S96 intake. Original payloads and comments are retained under `/private/tmp/rdocx-s96-run-refresh`.

- Record 427, updated `2026-10-27T20:59:51Z`, body SHA-256 `5cef6bc8e0440076c8da5c46778ee22595c208f6698f367d2fbe9635208a46f4`.
  Contribution head `cc3473d0c05659167847a941e0c7eb2bbd7217f5`.
- Record 428, updated `2026-10-27T21:22:09Z`, body SHA-256 `635213f1f93940b66c2856d18a2d60caaf246d94419a40db90a1ec3b4e5aa5ba`.
  Contribution head `87f101bc8e688bd93d358b72bc69c28732e9dad6`.

Design closure on 2026-10-28: consumer map, observable A-IDs, discriminating cases, public signatures, imported states, HLD impact and local/final evidence were checked against current source. Q-S96-01 is answered in F-X272. Approvals follow provider contract order. No material user decision remains. Pending native/hosted/installed observations are retained final acceptance debt, not recorded passes.
