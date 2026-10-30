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
| Linux bytes | SHA-256 `2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e`, max 224 MiB | Retained source installer at Base `24af95d3`, official TDF26.2.5.2 archive | Authenticated 218075885 bytes with retained hash, runtime mismatch/oversize controls pass |
| Linux extraction | At most 256 members, at most 256 MiB extracted, expected root and core/impress packages, regular files/directories only | Existing `safe_extract`, installer lines 73 to 123 | Traversal, link, missing package, count and size boundary controls |
| macOS bytes | SHA-256 `c99fb4fe574437fc4cb820a4ca15271bca325920861f7139858b36d7f9df78ad`, proposed 320 MiB download cap | Retained CI source at Base `24af95d3`, official TDF26.2.5.2 archive | Authenticated 297407265 bytes with retained hash, actual hosted installed outcome pending |
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

- [x] Authenticate official archive bytes against both retained platform hashes.
- [x] Repair URLs without changing oracle identity or Linux bounds.
- [x] Add macOS retrieval bound and preserve readonly mount checks.
- [x] Repair only the outdated roadmap expectation, retain negative controls.
- [x] Run unique scoped checks and independent zero-finding microscope.
- [x] Update listed HLD intent before final gate.
- [x] Implement and independently review A7 public diagnostic retention without changing its deadline or stages.
- [ ] Obtain actual final-source hosted acceptance, retaining pending status until observed.

## Open questions

No unresolved user product choice. Both archive byte identities are authenticated. Supported-target installation, actual hosted CI and final sprint qualification remain pending observations, not questions to waive.

## Contract closure

Existing callable installer signatures remain unchanged. The implemented 320 MiB macOS bound exceeds the authenticated 297407265 byte archive size. Both replacement URLs were adopted only after actual bytes matched the retained hashes. Local implementation and the named scoped checks pass. The first microscope found no implementation defects and requested this S1 readiness-text correction. Independent follow-up microscope pass 2 reports 0 defects and 0 smells. Actual supported-target installation, hosted CI and final integrated sprint qualification remain pending.

## Intake identity

Read-only GitHub refresh on 2026-10-28 found no body or contribution-head changes from the S96 intake. Original payloads and comments are retained under `/private/tmp/rdocx-s96-run-refresh`.

- Record 427, updated `2026-10-27T20:59:51Z`, body SHA-256 `5cef6bc8e0440076c8da5c46778ee22595c208f6698f367d2fbe9635208a46f4`.
  Contribution head `cc3473d0c05659167847a941e0c7eb2bbd7217f5`.
- Record 428, updated `2026-10-27T21:22:09Z`, body SHA-256 `635213f1f93940b66c2856d18a2d60caaf246d94419a40db90a1ec3b4e5aa5ba`.
  Contribution head `87f101bc8e688bd93d358b72bc69c28732e9dad6`.

Design closure on 2026-10-28: consumer map, observable A-IDs, discriminating cases, public signatures, imported states, HLD impact and local/final evidence were checked against current source. Q-S96-01 is answered in F-X272. Approvals follow provider contract order. No material user decision remains. Pending native/hosted/installed observations are retained final acceptance debt, not recorded passes.

## Local implementation evidence

Worker Base `376a8cec`. Actual focused baseline child exit 1: retrieval-route and roadmap cases failed, runtime safety control passed. The same three named cases pass after the patch, child exit 0, 3 tests. Complete transcripts are `/private/tmp/rdocx-s96-run-refresh/fx271-local-red.log` and `fx271-local-green.log`. Runtime negative controls cover download hash/size, archive traversal/type/member/expanded-size/missing-package and exact executable identity. New macOS mutations reject unbounded retrieval, wrong checksum and writable mounting.

Root authenticated both actual official archives under `/private/tmp/rdocx-s96-archive-auth`, with exact source/final URLs and receipts in `evidence.json`. Linux is 218075885 bytes and matches the retained hash. Its 46 members total 218024484 declared bytes, have the expected root, no unsafe/non-file entries and both required packages. Actual `safe_extract` succeeded with 42 Debian packages, without executing or installing them. macOS is 297407265 bytes and matches its retained hash. Grok attempts produced no usable authentication evidence and none is credited.

Local `/verify --scoped F-X271` observations: the three named focused tests, `cargo fmt --all --check`, `python3 scripts/prose_check.py`, `python3 scripts/sync_agent_skills.py --check` and `git diff --check` pass. No Rust crate changed, so changed-crate lint is inapplicable. Independent microscope pass 2 is clean after the separate S1 documentation remediation. Actual hosted Linux/macOS installation and Test/MSRV/Word fidelity/rpptx Python/CI gate remain pending, as do the final full suite, hash49/pixels7, native/oracle and S96 installed/typing/packaging union. No final acceptance or dependency completion is claimed.


## Final gate remediation amendment

The actual integrated gate at a859c19f found four failures in the existing Python job contract because scripts/test_sprint_workflow.py1362 retained the superseded stable macOS URL. Replace that one literal with the already authenticated old-release route. Keep the digest, exact version, conditional matrix, failure-propagation and mutation assertions unchanged. The four actual failures are retained as RED. Run those four cases after the correction and the complete policy suite with its required process and network access at the final boundary. The original seven sandbox/process/DNS errors remain failed setup observations, not policy passes.

The same source failed cargo deny check on RUSTSEC-2026-0414 through pulldown-cmark0.13.4. Amend this existing gate-repair story with the minimum security patch: pin the existing facade dependency to exactly0.13.5, update only its lock entry and update the existing HLD15pin. This introduces no dependency edge, feature, public API, source file or parser option. Default features remain disabled, html remains selected and runtime options remain TABLES plus STRIKETHROUGH. No advisory allowlist or gate bypass is permitted.

Primary sources are the actual cargo-deny advisory, upstream release tag48a2767b5b8ab0a4893e80d942e510065889fc6b at https://github.com/pulldown-cmark/pulldown-cmark/releases/tag/v0.13.5 and its0.13.4comparison. The official registry archive is157218bytes, SHA256278e9268c93b9e498a2da46c34d90ba93a3255eb99009646398f34bc73ab36eb, with Rust minimum1.71.1. The observed advisory concerns ENABLE_WIKILINKS, which this importer does not enable. The actual old public importer control returned successfully. Do not claim an exploitable public-importer RED or additional advisory identities from this observation. Actual supply-chain refusal is the RED.

The consumer is the existing native Markdown-to-HTML importer in crates/rdocx/src/html.rs, inherited by Rust, Python, CLI and WASM. Existing import/projection/diagnostic/limit/template/round-trip controls must remain passing. The final union qualifies the new dependency through fresh affected native and installed suites, WASM and hosted MSRV, README and verified archives, supply-chain and49hashes/7pixels. Expected existing output is unchanged. Investigate every delta before adoption. Keep the security dependency increment separately labelled from the stale test-literal correction.

Root designer approved this bounded remediation after actual gate evidence and independent source design review by fx271_microscope. That review approved the scope, not unwritten implementation. The normal independent implementation microscope remains required. This amendment changes no other story or sprint acceptance row, and all existing final debt remains pending.


## Pinned presentation corpus transport recovery

Root adopts this technical amendment on2026-10-30after reading the actual hosted
read-timeout failure, bounded preparation and independent technical proposal.
F-X271remains approved/in-progress. All50manifest URLs, digests, producers, exact
oracle pins, streaming, cache verification, unmanifested-file refusal and final
whole-directory verification remain unchanged. No new file/module/dependency,
public option, endpoint fallback or missing-subject skip is authorized.

A6adds one consumer: scripts/fetch_pptx_corpus.py fetch, used by Test/MSRV and
presentation fidelity CI, must recover at most twice after a classified transient
GET/open/read failure, for at most3attempts with1and2second delays. Keep timeout60
per operation, without claiming a total subject deadline. Report subject, actual
error and attempt count, including exhausted failure. HTTP429/500/502/503/504,
actual timeout/reset/temporary connection errors, temporary DNS and raised
IncompleteRead/RemoteDisconnected qualify. Check HTTPError before URLError.
Unwrap only exception-valued URLError reasons with a cycle guard. Certificate and
generic SSL failures, permanent DNS/HTTP, string reasons and unknown failures are
terminal. Do not infer recovery from exception messages or arbitrary cause chains.

Catch retryable failures only at urlopen/response.read boundaries. Local open,
write, close, digest, replace and cleanup failures are terminal even if their
exception class resembles a transport timeout. Close failed responses before
retry. Discard partial bytes, never append them. Silent EOF with wrong bytes is a
terminal digest mismatch, not an invented network retry. Publish only a complete
checksum-matching temporary with the existing atomic replace. Preserve an old
destination on every failure. Create the invocation-owned temporary exclusively
and unlink only after ownership is established. A stale or colliding foreign
.download must be refused and left untouched. No broad cleanup is added.

Implement in the existing fetcher and scripts/test_sprint_workflow.py. HLD15is
already the sole HLD impact. Extend its current retrieval intent before verify.
Finite local tests in the existing carrier cover initial transient then success,
partial read then complete retry,3attempt exhaustion, allowed/terminal status and
TLS/DNS/reason controls, checksum/silent-EOF refusal, transient-looking local write
or publication failure, cache no-network and foreign temporary custody. Mock sleep
and responses, with no network. Capture actual old-source recovery RED before
source edits and fixed GREEN, preserving the same semantic assertion. Retain raw
commands/source/outcomes. Scoped policies/prose/diff and independent microscope
are required. Full policy and actual supported-host50file retrieval/fidelity remain
final integrated debt. No push, dispatch, native proof or completion follows this
local transport correction.


## Remaining public authoring CI timeout diagnostics and repair

The user explicitly requires the failed CI to be fixed in this sprint before publication. Extend existing F-X271 verification confidence scope with A7. Preserve all prior installer, transport, advisory and final acceptance obligations. This amendment authorizes bounded implementation after root adoption and the canonical writer handoff. It does not authorize publication, a CI dispatch, a timeout waiver or a claim that diagnostic repair fixes the observed timeout.

### Actual trigger and known boundary

The immutable sprint run40812739735 attempt1 at1431cb408dc57a24353bf826933408b0ae2ab8ec fails MSRV1.93 test public_authoring_conformance_harness::sanitized_public_authoring_fixture_passes_every_conformance_stage. The integration carrier invokes python3 scripts/docx_authoring_conformance.py --public. That route executes cargo run --quiet --offline -- OUTPUT in an invocation-owned fresh public consumer project, inherits the recorded environment with the existing target/docx-authoring-conformance fallback, captures both streams and sets timeout300. The log reports running over60seconds at04:07:24 and generic timeout at04:11:24. Ordinary Test passes the same public case after roughly210seconds, but this is not a same-temperature comparison. Test CI first runs the public harness standalone before its full suite, warming the dedicated conformance target. MSRV runs only the full suite. Neither YAML job sets CARGO_TARGET_DIR. The current script has no timeout repair relative to that source. This asymmetry is a source fact, not proof of a compile delay or target deadlock.

The catch at script main1334-1336 discards TimeoutExpired command, timeout and partial captured streams. Thus compile, Cargo lock wait and public consumer execution cannot be distinguished retrospectively. The offline flag prevents normal online dependency fetching but does not identify the waiting operation. This is the existing public authoring consumer, not F275's new SemVer generated project. Raw logs, source pins and the published-main comparison are retained in s96-ci-failure-triage-current-v1. The published-main installer404 and Python expectation failures are separately mapped to existing F271/F272 repairs and are not reclassified as this timeout.

### Exact initial implementation contract

Repair the missing diagnostics only at the known public consumer Cargo subprocess in existing scripts/docx_authoring_conformance.py. The existing nonzero-child branch also discards both streams and must retain the same public command/phase/result context while remaining a deciding failure. On TimeoutExpired, report the actual call-site phase, exact exception command/argv as data, actual timeout value and all available partial stdout/stderr. Distinguish bytes, text, empty values and None. Preserve non-UTF8 bytes through an exact escaped representation rather than a replacement-only decode. Diagnostics must allow the captured bytes to be reconstructed and must not execute or reinterpret the command or infer a wait reason from an exception message. Retain exact argv, cwd and effective inherited target/toolchain/profile facts, distinguishing an unspecified inherited value from an observed one. Bind actual CI host/compiler selection in the new hosted result rather than guessing it from a directory name. Do not dump unrelated environment variables or credentials.

Add three small stderr checkpoints to the existing generated PUBLIC_CONSUMER_SOURCE at actual main entry, immediately before its first deterministic render and immediately before its second deterministic render. Preserve the two render calls and every public output requirement. These call-site markers establish observed client progress without adding a product API or process framework. If no entry marker was captured, report that no client entry was observed, not that compilation or a lock wait was proven. A captured render marker narrows the actual program phase but does not itself prove why that phase waited. Retain markers with the child diagnostics on nonzero/timeout outcomes. Do not change private generators or their output.

Carry the diagnostic details out of the public TemporaryDirectory lifetime before normal cleanup removes the temporary source. A small local formatter and a catch at the actual public subprocess are sufficient. Reuse existing ConformanceError/main nonzero handling or the existing exception path. Do not add a new command mode, receipt schema, generic subprocess framework, process-tree manager, dependency, file/module, feature flag or global logging abstraction. Keep unrelated private/unclassified timeout behavior and the private-artifact non-disclosure boundary unchanged.

Preserve the actual300second timeout, subprocess argv, cwd, environment/profile semantics, fresh public source boundary, capture behavior and every conformance stage initially. Timeout remains a deciding nonzero failure. No silent retry, successful timeout disposition, stage bypass, skip, target relocation, resource/concurrency change or threshold increase is authorized by the missing diagnostics. Ordinary nonzero child exit and missing evidence remain failures.

### A7 consumer and discriminators

A7 consumer: the existing public authoring conformance integration test in crates/rdocx/tests/integration_test.rs and its actual Test/MSRV CI callers. Local evidence uses the existing HarnessSelfTests in scripts/docx_authoring_conformance.py, with no new test entrypoint or binary.

Before repair, capture a genuine synthetic old-source RED for the same diagnostic requirement. Mock the known public child raising TimeoutExpired, without Cargo, network, oracle or native execution. Cover mixed bytes/text fields, non-UTF8 bytes, empty bytes/text and None, list/tuple command representation and the actual timeout value. Assert phase/argv/timeout/partial streams survive in the deciding failure and no mock command is rerun. Verify the unchanged300second subprocess call and public invocation inputs. Add controls that a nonzero child retains its actual return code, command and both streams, and incomplete public output cannot pass. Discriminate the three unique ordered public source checkpoints while retaining both deterministic render calls. Verify that marker absence is not reported as a proven build/lock diagnosis. Keep the normal integration carrier enabled and its success requirement intact. Preserve the prior failed hosted observation and each source-bound local outcome.

Implement diagnostics in the existing script and its existing HarnessSelfTests. The existing Rust integration carrier may be adjusted only if needed to retain/display the actual diagnostic while preserving its public invocation and success contract. ci.yml changes require a later observed cause that earns the specific change. Do not preemptively alter workflow timeouts, jobs, caches, target directories or stage selection. HLD15 remains the sole HLD impact, documenting failure attribution and the unchanged requirement for actual host qualification.

### Actual cause and completion gate

A local synthetic diagnostic GREEN proves retention only. It neither repairs nor qualifies the old300second MSRV failure. The actual source-qualified MSRV host run must retain enough child output/phase evidence to establish build, lock or program waiting and drive the concrete final repair. Root then records the bounded observed-cause correction in this same F-ID before its source change. If the new run passes, preserve its exact source/toolchain/profile/child result and verify all required public conformance evidence rather than relabelling the old failure as success.

A7 closes only with a clean actual final-source MSRV public-conformance result and the required clean final CI union. Existing F271 native retrieval, F272 comparison/installed, F273 native/hash/full and new F275 compatibility/consumer gates remain separately deciding. Independent microscope, applicable scoped script policies/prose/adapters/diff checks and final integrated verification remain required. No source push, skipped old failure, old .5 wheel, build-only rehearsal or diagnostic-only GREEN substitutes for clean final-source CI before publication.


## Renewed current-source checkpoint evidence

A7local diagnostics pass12existing harness self-tests and the existing CI/privacy wiring test, with actual synthetic old-source failures retained. F271all-pass5 reports zero defects and smells. This qualifies diagnostic retention only. The original MSRV timeout remains failed, its runtime cause is unproved, and actual final-source public conformance and the complete CI union remain pending. Existing installer/retry review and input proof remain limited to their applicable unchanged scope.

The existing consumed source-availability handoff establishes its historical integration fact. A renewed canonical checkpoint must authenticate fresh clean reviewed source, workflow bytes, remote/ref/SHA and ancestry. Existing F271/F275progress and S96run state record the expanded pending native/signing/CI and build-only wheels gates before any push. Old1431results do not qualify current.17/.15source. No new worker handoff or completion is inferred, and native Word, full/hash/pixel/installed/WASM/oracle acceptance remains due.
