# F-X275, Prepare Word 0.17.0 and shared/PowerPoint 0.15.0 with strict SemVer gates

**Status**: approved
**Sprint**: S96
**Size**: L
**Depends on**: F-X270

Root adopts this amended contract on2026-10-30 after the user explicitly selected Prepare0.17.0/0.15.0. Root read the complete proposal, actual compatibility diagnostics, native tool contract and independent callable-contract review. Its C1 incremental reuse interface and C2 tracked native pin owner are incorporated. The earlier .5 preparation remains historical evidence and cannot qualify these new identities. Completed F-X267/F-X270 and published history remain unchanged. Preparation approval permits implementation, never publication or yank.

## Problem

Issue429 establishes that published oxml-opc0.14.4 changed a compatible-line public Vec contract to Arc and changed setter inference. Identical external clients compile against checksum-verified0.14.3 and fail against0.14.4. The pinned cargo-semver-checks0.51.0 comparison nevertheless returned success, with229checks passed and31skipped. Latest-only API checking would miss the earlier contract too. A tool-only gate cannot enforce this repository's compatibility promise.

A Vec-only OPC0.14.5 would break different published callers that require the Arc contract and minimum OPC0.14.4. Published rdocx0.16.4 and rpptx0.14.4 are actual counterexamples. New shared nominal types also leak through Word reexports and public signatures, so retaining Word0.16.x while moving shared types to0.15.x would hide another compatibility break. Prepare Word0.17.0 and shared/PowerPoint0.15.0. Do not label the contradictory old contracts universally compatible or silently ship the existing .5 preparation.

The actual reporter's published jubarte-redlines0.13.0 and0.14.0 graphs are narrower: each directly requires OPC0.14.2 by caret and has no discovered high-level minimum-.4 family dependency. Both genuinely select.4 in a frozen current registry fixture and.3 when only.4 is hypothetically yanked. Default host lib/bin checks fail with.4 on the expected API errors and pass with.3. Existing.4 locks remain.4 under that hypothetical yank. This supports a targeted recovery candidate, not a completed old-line registry fix, universal backport, runtime result or cargo-install claim.

## Spec reference

- docs/hld/10-bindings-spec.md, Packaging and CI, public nominal dependency reach and installed identities.
- docs/hld/14-development-backlog.md, current F-X275 scope only, preserving F-X267/F-X270 history.
- docs/hld/15-build-and-toolchain.md, Packaging, Release process and CI job matrix.
- .claude/commands/verify.md and .claude/commands/release.md, exact-source evidence reuse and separately approved publication.
- .claude/commands/release-notes.md, authenticated family-only claims and contribution inventory.
- docs/sprints/SPRINT_PLAN.md, S96 acceptance including429and430.
- Authenticated issue429-intake-v1, issue429-jubarte-graph-readonly-v1, issue429-jubarte-resolver-diagnostic-v1 and issue429-jubarte-compile-diagnostic-v1 under the existing TEMP evidence root.
- Pinned issue429-semver-tool-contract-readonly-v1 plus issue429-semver-ci-intent-supplement-v1. Their proposed interfaces are not already implemented commands. The supplement supersedes blanket absence/strict-less wording for ordinary CI.

## Approach

Prepare exactly seven Word Rust packages and the rdocx Python distribution at0.17.0, and exactly fifteen shared/PowerPoint Rust packages and the rpptx Python distribution at0.15.0. Derive membership from existing RELEASE_CRATES rather than adding a competing code allowlist. The Rust registry set is twenty libraries and the two binary-only CLIs. Binding/WASM crates are outside the Rust API registry comparison set. Keep rpptx-wasm0.12.1 and npm publication outside this preparation. Word WASM identity follows its existing inherited version contract and retains its actual compile gate.

Inspect current manifest carriers, workspace pins, lockfile versions, binding metadata, README examples and measurements, policy assertions and current HLD wording. Change only current identities. Never blanket-replace historical literals. Preserve published changelog bytes from `## v0.16.4` through EOF byte-for-byte against24af95d3, including all older sections. Existing .4 inventory, migration/history and mutation fixtures remain historical controls. Earlier .5 archives and hosted observations are preparation evidence only and cannot qualify .17/.15 artifacts.

Implement one release-semver subcommand inside existing scripts/sprint_workflow.py, with policy controls in existing scripts/test_sprint_workflow.py. Extend existing ci.yml and wheels.yml, canonical verify/release/release-note guidance and current HLD sections. Do not add a tracked source file, module, crate, test binary, API trait or feature flag. Regenerate adapters after command changes. If a process clarification requires WORKFLOW.md or another canonical command change, root first records that exact impact and its trigger rather than silently expanding the approved file list.

### API gate and baseline selection

The stable exact callable interfaces are:

    python3 scripts/sprint_workflow.py release-semver {all|TAG} --target HOST_TRIPLE --tool VERIFIED_TOOL_PATH --output DIR [--reuse-evidence DIR ...]
    python3 scripts/sprint_workflow.py release-semver {all|TAG} --check-evidence DIR [DIR ...] [--refresh-registry]
    python3 scripts/sprint_workflow.py release-semver TAG --check-evidence DIR [DIR ...] --release-candidate

Execution and evidence-consumption modes are mutually exclusive. --reuse-evidence is an execute-only repeatable option taking exactly one directory per occurrence. It is forbidden with --check-evidence or --release-candidate. Execution writes a new output directory and leaves earlier packets immutable. Registry applicability is refreshed before acquisition. Each retained row must authenticate exact applicable source/dependency identity, genuine native host/target, tool archive/executable, effective feature sides, baseline archive/checksum, generated locks/configuration, compiler-canary source/locks/results and raw log/artifact bytes. Acquire only missing or invalidated rows. A newer failed attempt blocks older green reuse. Evidence-only --refresh-registry refreshes inventory without executing. Without that flag evidence mode checks the frozen recorded inventory, with freshness imposed by its deciding caller. Execute mode requires target/tool/output together and produces compatibility evidence, not publication authorization. Evidence mode forbids execute-only options. --release-candidate is legal only for one explicit family tag in evidence mode, implies a fresh selected-family registry/candidate preflight and never repeats the API suite. Invalid modes, mismatched families/tags, missing rows or malformed evidence fail before mutation. This command does not publish, yank, tag, post messages or alter source versions.

For each crate, freeze the complete public version inventory and select every stable published baseline numerically at or below the candidate within the same Cargo compatibility boundary, including yanked historical versions and an equal published candidate. For these0.y families that means the same y. Every selected archive must match its published checksum. A docs-only CI run at an already published version must actually compare that equal archive and all earlier applicable versions. It must not fail merely because the candidate is occupied, skip the equal row, select latest-only, omit an unavailable baseline or replace it with a Git sprint base. New registry baselines missing from reused evidence make that evidence incomplete for the refreshed inventory.

The first unpublished0.15.0/0.17.0 has explicit empty-compatible-line applicability. It is not a claim that the old contracts remain compatible. Preserve previous-line .3 last-compatible and .4 accidental-divergence migration witnesses separately, reviewed intentional differences, public dependency reach and actual new-contract compile canaries. cargo-semver-checks treats the initial minor movement as tool-level Major and can skip breaking queries. Exit0 on such a witness is not a clean API-diff inventory. After publication, equal-version baseline coverage becomes mandatory and the initial-line exemption ends.

Only publication validation requires selected-family crates.io/PyPI candidates to be absent. It additionally requires the candidate to be strictly greater in SemVer precedence than every stable published version in its selected compatibility line, including yanked versions and versions above the candidate. API comparison rows still select only versions at or below the candidate. Build metadata does not affect precedence. A newer incompatible minor elsewhere does not prohibit a deliberate old-line patch. Publishing shared/PowerPoint first must not make its already occupied version a duplicate Word candidate. Existing unified-release-family validates source carriers, not registry occupancy. Keep fresh selected-family registry, owner/publisher and remote-tag checks at /release independently deciding, with one authenticated preflight receipt reusable by consumers where applicable.

### Pinned tool, native views and compile canaries

Use cargo-semver-checks0.51.0 with one concrete native pin constant mapping in existing scripts/sprint_workflow.py, copied from authenticated tool-contract native-assets.json. For exactly the three supported native targets, this mapping owns the official archive URL/member and archive/executable SHA256. Workflows acquire against this tracked policy, verify both hashes and retain actual acquisition provenance. The wrapper verifies --tool against the same mapping. A version string or TEMP machine-local path is not integrity policy, and a pin change invalidates affected rows. The required actual native targets are x86_64-unknown-linux-gnu, aarch64-apple-darwin and x86_64-pc-windows-msvc. A cross-target Mac invocation is not a Windows native qualification. Do not use an unpinned latest installer, unsupported binary fallback or missing-host skip.

For each selected library/baseline/native target, execute explicit default and all-feature views separately using --baseline-root on the checksum-verified extracted archive and the matching --target. Bare invocation's feature heuristic is insufficient. Record actual feature applicability on both sides, generated dependency locks/configuration, query completion and skips. A missing explicitly requested feature cannot be silently ignored as coverage. Preserve the existing applicable no-default, target/package, CLI and WASM gates because default/all do not prove every subset or every distributed target.

The exact stable tool call is `TOOL check-release --manifest-path crates/CRATE/Cargo.toml --package CRATE --baseline-root VERIFIED_ARCHIVE_DIR --target HOST_TRIPLE --default-features --color never --verbose`, repeated with --all-features. Do not invent stable --json/--offline/--locked/--jobs tool options. Capture raw output and emit the wrapper's own receipt. Tool-generated baseline locks must be retained and qualified rather than assumed to equal workspace Cargo.lock. Baseline build/rustdoc/registry/tool failure, timeout or incomplete output is a failed/incomplete row, never permission to drop it. No release-type override or lint suppression may preserve a patch number.

The compiler layer retains real external map/insert/remove/getter and Vec.into setter-inference clients, the Arc contract client and a Word nominal public-type consumer. Known-compatible.3 must compile the old Vec clients. Known-broken.4 must fail at the expected API sites/errors and pass its actual Arc contract client, not fail for unrelated setup/MSRV/network reasons. The prepared new minor must compile its declared Arc/nominal contracts. Later compatible patches must keep those contracts compiling. Do not require old Vec clients to compile against the intentionally incompatible new minor. Existing CLI behavior, binding signatures/defaults and behavioral/native review remain necessary because neither compiler canaries nor the API tool cover every compatibility promise.

### Native artifact retention and parser closure

Keep the exact approved tool argv. The authenticated0.51.0generator disables caches
for local projects, including both current and extracted baseline roots. It loads
current before baseline and may overwrite a shared rustdoc output. A source-derived
slug omits source path/content and side role, so equal slug text is not source proof.
Use a controlled recorded Cargo environment without CARGO_TARGET_DIR or another
shared build.target-dir override. Preserve semantic flags and dependency/configuration
facts. Require actual metadata-derived current/baseline generated project, target
and rustdoc paths not to alias, including equal-version rows. Preserve both sides'
generated Cargo.toml, Cargo.lock, lib.rs and actual library-target JSON as immutable
bytes before a subsequent invocation. Reject incomplete or aliased retention rather
than using a watcher, unsupported flag or later regenerated substitute. A separately
labelled locked/offline Cargo metadata observation may authenticate retained graphs
only with source/locks/configuration unchanged, without a fresh solve.

Parse both raw streams: statuses and verbose query results are on stderr, detailed
failure sections on stdout. Require exact crate/version phase order, one completed
Checking/Starting/Checked/Summary/Finished sequence, expected unique query IDs and
consistent counts. A successful exit with warnings, truncation or unqualified
required-query skips is not complete coverage. Freeze the pinned query catalog and
effective override inputs to explain selected/skipped sets. Do not assume every
skip is invalid or hardcode historical229/31counts. Initial empty-compatible-line
applicability remains separate from old-line migration tool output and required
compiler canaries. The exact primary generator/parser reports under the existing
TEMP evidence root close these technical interfaces, not any native execution.

The publication ordering discriminator includes registry0.17.0,yanked0.17.2and
0.18.0. Candidate0.17.1gets an API row for0.17.0but publication refuses because
0.17.2exists. Candidate0.17.3passes ordering because the incompatible0.18line does
not block it. Missing or unavailable required baseline/archive/native evidence
retains its deciding failure.

The exact pinned runtime query catalog has260IDs:192Deny/10Warn Major,
23Deny/4Warn/31Allow Minor. The historical229/31false-green used an explicit
assumed Patch release type. The approved future argv has no override. Under
upstream's inferred same-line0.y patch/equal Minor classification, default policy
selects202Major queries and skips58Minor queries. Old-line0.y minor migration is
tool Major and may select0/skip260. Derive selected/skipped ID sets from the pinned
catalog and actual classification, never universal count constants.

Own the exact260query tuples in existing sprint_workflow.py alongside the pinned
tool policy, derived from authenticated catalog4f1f5ae93d0c2e325eea77d11aa7db842adf6b506b7c3a17b444e95198243ce4.
No Rust/RON parser or separate tracked catalog file is added. The wrapper's current
no-user-suppression policy refuses nonempty SemVer lint override tables in current
package/workspace and baseline package/workspace inputs. Bare inheritance markers
are allowed only with no overrides. Malformed metadata, unknown IDs or explicit
false inheritance refuse. Baseline override inspection is repository policy, not
upstream policy merging: upstream uses only current selection's override stack.
Preserve built-in Allow/Warn defaults. A triggered WARN/FAIL blocks clean evidence.
Future approved override support needs separate exact precedence/coverage design.

### Evidence and workflow consumption

A compatibility receipt binds schema, scope/family/candidate membership, exact relevant source tree and file hashes, plan/review identity, Rust/rustdoc/host/target, tool/archive/executable hashes, Cargo config/flags and current/generated locks, effective features on both sides, raw inventory and archive identities, exact argv/environment/UTC/status, raw logs and rustdoc/cache/artifact hashes, and canary sources/locks/results. Record every applicability row, skipped query and failure. A JSON label or zero aggregate exit without those required rows is not proof. Evidence validation authenticates producer CI job/artifact/source identities and referenced bytes.

Ordinary CI requires the three native jobs and fail-closed ci-gate dependencies. API-affecting changes cannot be filtered away, and absent/stale evidence cannot satisfy docs-only reuse. /verify --full consumes the exact-source native union once. Build-only wheels rehearsal and tag ancestry consume authenticated evidence before publication can start. /release validates applicable H/R/M reuse and refreshes selected-family mutable preflights without repeating the native API suite on a tag. Relevant source/tool/feature/config/dependency/baseline drift requires only the missing or invalidated rows. A latest failed attempt blocks reuse of an older green one.

The native matrix covers those three host APIs only. It does not qualify Linux musl/aarch64, Mac x86 or WASM APIs. Existing builds, packaging, no-default and WASM checks remain due and their limit stays explicit. Current hosted deciding-job counts must be re-derived after adding SemVer jobs and routing, not copied from the old27/31-job rehearsal. Preserve the26 unique payload inventory, with aggregate duplicates excluded as new subjects and four publication jobs skipped during build-only preparation.

### Archive and deciding-history closure

Baseline acquisition uses a private16MiB compressed,128MiB whole decompressed-tar
and4096logical-member budget. These exceed the retained22historical metric peaks
of4,904,112compressed bytes,54,306,876member bytes and60members. They do not prove
every historical archive fits. Excess is incomplete required acquisition, never a
baseline omission, API verdict or public document limit. Authenticate compressed
registry SHA256before gzip/tar interpretation. Retain actual raw bytes/failure,
with bounded partial-prefix identity on excess. Decompress to an owned bounded tar,
validate complete inventory/exact crate-version root and manifest, and copy regular
nonsparse files/directories exclusively. Refuse links/specials, traversal/alias
names, backslash/colon, duplicate paths, file-parent conflicts, native collisions
and nonzero hidden tar tail. Require actual resolved containment and exact lengths.
Do not apply archive owners/timestamps or use extractall. Publish only a complete
owned source bundle into an absent destination and clean only its owned stage on
failure. Preserve immutable source, original-lock disposition and file inventories.
Tool-generated locks belong to a separate authenticated working copy.

An arbitrary old reuse directory cannot prove that a newer failed directory was
not omitted. Local exploratory packets therefore qualify only their supplied
history. Deciding CI/full verification/wheels/release must supply a complete finite
authenticated attempt catalog and cutoff from existing sprint verification state
or actual qualified GitHub run/attempt/job/artifact authority. The catalog is an
immutable evidence snapshot, not a second maintained tracker or service. Preserve
actual producer identities, raw observations, artifact-download hashes and original
reused row references. A label, path or self-written success cannot authenticate
producer authority. Evidence-only mode validates its frozen supplied authority,
without silently adding GitHub network access or interpreting refresh-registry as
refresh-all-authorities. Its caller acquires fresh required authority separately.

Within one hosted run, actual run_attempt orders retries. Across runs, initial
run_number cannot order a later rerun of an older run. Use qualified actual attempt
start/completion facts and explicit acquisition-round predecessor edges. Refuse
incomparable histories instead of choosing green. Later failed/canceled/incomplete
or required-skipped attempts block older successful rows for the same relevant
inputs, including failed jobs without surviving artifacts. A later actual success
may resolve that failure. Deciding callers refresh authority at their required
cutoff. Frozen proof claims no knowledge of later events. Preserve exact H/R/M
producers and relevant tree/input equality. These duties fit the approved modes,
existing runstate and workflow consumers without a new CLI option.

### Old-line recovery and public communication

Do not publish a Vec OPC0.14.5 in isolation, overwrite .4 bytes/tags or claim the new minor alone fixes users constrained below0.15. The validated narrow recovery is deliberate OPC0.14.3 selection for the two exact published jubarte sources. The eight actual resolver observations and four default-host lib/bin type checks establish their bounded selection/source compatibility. They do not establish runtime output, installation, every dependency graph or old lock security. Packaged --locked graphs retain.2, existing.4 locks retain.4 and exact/minimum-.4 high-level users need separate deliberate recovery.

Prepare an exact, reviewable optional yank packet for ONLY oxml-opc0.14.4, with fresh registry state, the faithful hypothetical-only-yank results, affected exact/minimum-.4 consumers and explicit existing-lock limitations. A real yank requires separate exact-action approval and is outside /run-sprint publication authority. No maintained .14 backport is promised or implemented by this plan. If no safe old-line package can satisfy both conflicting contracts, document that evidence and the narrow verified workaround rather than inventing a universal patch. This mitigation packet is a held candidate, not executed registry remediation.

Retain the original release-note security constraint: describe pulldown-cmark0.13.5 and the advisory relevance accurately, without claiming that the disabled wikilink path was exploitable through the public importer. The current dependency and local layout changes require their genuinely affected fresh checks. Do not import donor archive measurements or classify an adapted PR as an unchanged merge. Preserve the established compression tolerance in archive qualification, the exact source/helper measurement identities and independent applicability assessment for any retained native/render proof.

Rebuild contribution inventories from previous published family .4 tags through reviewed S96 source. Retain423 reporter set2374,424/425/426 reporter hadim and427/428 contributor hadim with their qualified outcomes. Add429 reporter arthrod to the affected migration/release-policy inventory and430 reporter lawrenceakka to Word notes only after actual qualified fixes. The eventual Word inventory has eight candidate records, not automatic inclusion before their acceptance. The shared/PowerPoint429 outcome must describe the actual minor/API-policy consequence and honest limits, without relabelling old S95.4 COW work as a new performance feature.

The user explicitly authorized a429 reporter apology/status, conditioned on fixes and validation before posting. Keep it unsent until the concrete implemented repair and deciding validations pass, and describe exactly any prepared-versus-published and workaround-versus-registry disposition. No430 comment is authorized. Other release-bound notifications follow the canonical post-publication ceremony and separate final approval. No reporter-facing message is part of this preparation execution.

## Consumer map

| Rule | Entry points and outputs | Owner and policy | Lifecycle |
|---|---|---|---|
| Family identity | Exact RELEASE_CRATES, Cargo/lock/Python/CLI/WASM carriers | Existing family validator plus selected version policy | Source, verified archives, installed import, registry preflight |
| API compatibility | Existing script subcommand and twenty native libraries | Complete compatible inventories, pinned tool, effective features and compiler canaries | Native acquisition, raw artifact proof, exact-source evidence reuse |
| CLI applicability | Two binary-only release members | Explicit no-library-tool row plus existing CLI behavior/version gates | Native build, archive, host execution |
| CI decision | ci.yml native jobs and ci-gate | Failed, missing or unexpectedly skipped required rows block | PR/push, full verification and source-qualified artifact consumers |
| Publication | wheels.yml and canonical /release | Selected family, authenticated compatibility proof, fresh mutable preflights and separate main-SHA approval | Sprint R, clean main M, tag, registry and release body |
| Old-line recovery | Exact jubarte graphs, optional OPC.4 yank packet | Bounded real diagnostics, no universal compatibility or automatic registry action | Prepared candidate, separately approved future action, existing-lock limits |
| Notes and credit | Existing CHANGELOG/release-note validators | Authentic issue/PR outcomes, immutable published suffix, family separation | Review/render, held429 reply, post-publication notifications |
| Product acceptance | F-X271to274 including F-X272430 amendment | Their current contracts and final composition riders | Native/cache/reopen/render, installed, schema and full union |

## Behavior contracts and evidence

No public authoring method, XML policy or product feature flag is added here. F-X272to274 retain those contracts. Exact new family identities, stable wrapper argument modes, numeric baseline selection, native applicability, fail-closed outcomes, proof lifecycle and no-mutation boundary are specified above. Primary tool/source/download pins are in the authenticated tool-contract packet. Registry absence and yank state are mutable observations, not release approvals. Initial new-minor applicability, tool query skips and missing full API-diff coverage must remain visible.

## Dependency readiness

F-X270 is done in closed S95.4 and remains unchanged. F-X271to274 final S96 acceptance, F-X272430 extension and all retained native/schema/installed/hosted riders remain due. They are final acceptance dependencies, not an invented start cycle. Existing .5 local/hosted observations remain historical preparation and are not silently transferred to new versions or gate inputs. Root is sole canonical writer. Amend the current plan and dependent current wording before any new gate/version implementation.

## Acceptance criteria

| ID | Input/setup | Required observation | Local evidence | Final evidence |
|---|---|---|---|---|
| A1 | Exact7/15Rust and2Python candidates | All current carriers/pins/lock/project identities consistently.17/.15, no historical rewrite, rpptx-wasm.12.1 retained | Existing version-family mutation controls retargeted with honest old-source failures | Fresh selected-family registry/tag preflight later |
| A2 | Published suffix from v0.16.4 at24af95d3 | Byte-identical suffix, exactly one new meaningful section per family | Existing suffix hash and mutation refusals | Independent family-note review and exact check/render |
| A3 | Qualified423to430 inventory and final reviewed range | Every included record linked and specifically credited, every exclusion explained, no false unchanged-PR or old-line repair claim | Existing inventory tests and reviewed notification list | Fresh final intake and held communication boundary |
| A4 | Final-source22locally patched archives | Correct members/version/README metadata, each below10MiB | Existing inventory/size controls | One integrated verified packaging/README qualification |
| A5 | New source-qualified workflows and downloaded payloads | All re-derived deciding jobs pass, four publication jobs skipped,26unique payloads authenticated | Existing build-only routing/inventory mutation controls plus SemVer dependency controls | Actual exact-SHA native/hosted job and artifact receipts |
| A6 | Fresh host CLI/wheels and full product inputs | Exact.17/.15version/import identities, existing3.9/3.12priority and3.12typed/stub requirements, all product riders retained | Existing selected installed/ABI/current-metadata controls | One actual final installed/CLI/full/native union |
| A7 | Ordinary compatibility CI at initial and later equal/patch versions | Complete numeric compatible/yanked/equal baseline and effective-feature/native matrix, honest initial-line applicability | Meaningful policy negatives in existing test carrier | Genuine pinned-tool native union plus compiler canaries and independent review |
| A8 | False-green tool and real external clients | Expected.3Vec pass/.4API-specific fail, new-minor Arc/Word nominal clients pass, no unrelated failure substituted | Captured historical compiler/tool diagnostic plus source-specific canaries | Actual new-source native canary rows and API/behavior review |
| A9 | Missing/stale/altered proof or newly required baseline | Evidence reuse and ci-gate refuse before publication, equal occupied CI remains valid while selected occupied release refuses | Mutated source/tool/log/lock/feature/native/baseline/mode controls | Authenticated CI artifact union and fresh release-candidate consumer |
| A10 | Exact jubarte recovery and optional only-OPC.4 yank candidate | Bounded actual resolver/type-check evidence preserved, exact/minimum-.4 and existing-lock limits explicit, no universal.14patch claim | Root-authenticated eight resolver/four compile observations | Reviewed concrete mitigation packet, any registry action separately held |
| A11 | Reviewed sprint R and later main M | Honest H/R/M reuse, canonical close before separate selected-family release approval, shared family before Word | Existing closure/evidence policies plus exact SemVer consumer | Canonical close and publication preflight, no release during sprint |

## Test matrix

| Discriminator | Required result | Wrong implementation rejected |
|---|---|---|
| Twenty libraries versus two CLIs | Required API rows versus explicit applicability plus CLI gates | Claiming binary tool checks or omitting a library |
| Numeric10versus9, prerelease and yanked inventories | Stable compatible numeric selection with every historical row | String sort, latest-only, skipped yanked contract |
| Published equal candidate in CI versus publication | Equal API row required in CI, occupied selected release refused | Rejecting docs-only main or bypassing equal-version checks |
| Initial minor before/after publication | Explicit migration applicability becomes real equal-baseline coverage | Permanent empty-line exemption or migration exit0 as universal compatibility |
| Tool false-green versus external type failure | Compiler failure decides | Repeating the known0.51.0blind spot |
| Missing host/feature/baseline/raw file or changed source/tool/lock | Fail closed, retain actual failure | Receipt label alone, cross-target substitution or stale proof |
| New registry baseline after evidence | Require only missing affected row | Invented execution or redundant whole-suite rerun |
| Shared already published, Word absent | Word selected-family preflight remains eligible | Treating other-family occupancy as duplicate Word release |
| Invalid CLI mode/all+release-candidate | Reject before acquisition/mutation | Ambiguous execution and publication intent |
| Incremental execute reuse | Repeated --reuse-evidence parses, altered prior bytes and newer failed rows refuse, one missing or invalidated row reacquires, unaffected exact row retains provenance | Silently rerunning all rows or reusing stale evidence |
| Hypothetical only.4yank and existing.4lock | Fresh reporter graph.3, existing lock.4, exact/minimum-.4 limit visible | Universal rollback or old-family patch claim |
| One stale carrier, modified published suffix or wrong artifact identity | Existing validators refuse | Partial bump, historical rewrite or old artifacts as new acceptance |

## Rejected alternatives

- Revise completed F-X270 or overwrite published.4 history.
- Ship current.5 COW carriers as though the old compatible line were preserved.
- Publish Vec-only OPC.14.5 or claim new.15 alone repairs below.15 consumers.
- Retain Word.16 while changing its publicly reachable shared nominal types.
- Tool-only/latest-only SemVer proof, forced release type, waived lints or unsupported feature skips.
- Universal candidate-absence checks in ordinary CI, or skipping the equal published baseline.
- New tracked gate script/module/allowlist or a second sprint/release tracker.
- Rerun unchanged native/product suites for paperwork or rerun the API union on a tag.

## Test plan

**Test gate**: release regression.

Local scoped work uses existing scripts/test_sprint_workflow.py. Batch current carrier, immutable suffix, family contribution and finite SemVer/mode/evidence/CI-routing discriminators without building twenty libraries inside each unit test. Capture actual prior-source failures honestly. Keep the genuine.4versus.3compiler/tool mismatch as an authenticated negative, not a rerun on invented fixtures. Run both current family validators, release-note check/render, prose, adapters and diff checks after their inputs change. Independent microscope must cover source, workflow gating, raw output parsing, evidence reuse, baseline selection and exact callable CLI modes.

Host downloaded-artifact qualification retains exactly42Word and44PowerPoint priority calls for each of Python3.9and3.12,172calls in total, plus actual3.12typing/stubs. Recompute their inventory hashes from current source. Authenticate the existing12CLI,12wheel and2sdist unique payloads, with exact name/version/platform/ABI and full metadata. These priority calls do not replace the separately required complete installed suites, equipped CLI chains, native observations or conformance. Retain the existing manual-build-only dispatch, mismatched-carrier, unpacked dry-run inventory and family release-note mutation controls in the existing test carrier.

The final integrated union runs genuine API acquisition once for the required native/feature/baseline matrix and new-contract canaries, actual22verified packages/README examples, current hosted build-only/artifact qualification and every existing S96full/risk/native/installed/typing/WASM/hash/pixel/supply-chain rider. Re-derived deciding-job counts replace old fixed27/31claims. Evidence consumption by verify, CI aggregation, wheels and publication does not repeat unchanged API execution. Mutable release preflights remain fresh. Version and workflow changes invalidate their affected identities, archive bytes and hosted policy observations. No stale completed or failed evidence is adopted as current acceptance.

## HLD impact

- docs/hld/10-bindings-spec.md, current family identity, public nominal dependency reach, packaging and CI compatibility proof.
- docs/hld/14-development-backlog.md, current F-X275 title/scope only, preserving completed history.
- docs/hld/15-build-and-toolchain.md, strict SemVer release policy, actual native CI limits, selected-family preflight and H/R/M evidence reuse.

## Risk routing

Release scripting, CI and supply-chain changes require meaningful fail-closed policy/negative controls, pinned executable/archive integrity, actual native-host execution and exact publication dependencies. Public API/compatibility changes require all compatible/yanked/equal baselines, effective feature views, stable compile canaries and independent behavioral review. Version/binding packaging changes require current carrier/pin/lock/README inspection, verified archives below10MiB, fresh installed3.9/3.12identity and3.12typed/stub qualification plus existing WASM checks. Contribution acceptance retains every issue row and authentic credit. Final metadata adoption reruns affected current-state policies only. No rendering baseline change or new source file is authorized.

## Hash harness

Preparation predicts no output delta in49hashes/7deterministic pixels. Keep each separately reviewed S96behavioral delta and its real producer. Unexplained deltas block acceptance. Version-only proof transfer requires actual consumer/input assessment, not a blanket old-green label.

## Source availability and publication boundary

Use the consumed F-X271 external-ci-source-availability handoff only within its actual validated scope. Before a new source push, authenticate the reviewed clean sprint source, current workflow/version/note/gate bytes, local/microscope evidence and exact remote/ref/ancestry. Preserve named final gates pending. If the actual handoff does not cover this expanded source, resolve that canonical checkpoint conflict before pushing. An old exception is not new authority. Build-only dispatch creates no tag, registry version, release or notification. Final /close-sprint alone merges main and creates sNN. /release alone creates family tags/publication after separate exact-main approval. An optional OPC.4yank remains a separately approved exact action.

## Implementation checklist

- [x] Root adopts the amended contract and reconciles current sprint/backlog/HLD/run-state wording while preserving old evidence.
- [x] Implement strict SemVer modes, complete baselines, pinned tools and external canaries in existing carriers.
- [x] Wire deciding native CI, full verification, wheels and publication evidence consumers with fail-closed controls.
- [x] Align current.17/.15carriers/notes/credits and preserve published suffix/history.
- [x] Prepare the bounded old-line mitigation packet and retain the user's429reply hold.
- [x] Obtain clean independent microscope and affected scoped gates.
- [ ] Authenticate actual new-source native/hosted/artifact and final integrated union evidence.
- [ ] Reconcile final metadata once and complete preparation, all publication/yank actions still separately held.

## Open questions

Q-S96-03is answered. On2026-10-30the user explicitly selected Prepare0.17.0/0.15.0 after reviewing the conflicting published Vec/Arc contracts and the verified0.14.3reporter workaround. Prepare those new minor families under this approved contract. The narrow old-line workaround and separately held optional only-OPC.4yank candidate retain their actual limits. No universal compatible .14patch is promised.

Real yank/publication approvals are later exact-action decisions and are not answered by elapsed time or this preparation approval. Native API/tool compatibility with Rust1.97.1, final product acceptance, source availability scope and fresh registry state are evidence obligations, not user choices to waive. No430comment is authorized. No new public authoring/API/performance contract is proposed beyond the existing separately approved product stories.

## Contract closure

For adoption, root checked every exact CLI mode, derived7/15and20/2membership, equal/older/yanked/initial-line baseline applicability, effective features, genuine native target scope, expected canary outcomes, evidence provenance/reuse, current/history carrier split and selected-family publication boundary against actual existing callers. The approved bridge must not call a nonexistent command or invent a stable upstream flag. This amendment materially expands capacity beyond the prior .5 metadata work and invalidates the old sprint estimate. The independent closure review required C1 execution reuse and C2 a tracked pin owner. Both are specified above and root verified their callable modes and existing carriers. Approval records current intent and readiness requirements, never successful execution.


## Isolated baseline execution and portable retained proof

Root adopts this technical clarification after authenticated generator-source
inspection. --output is immutable proof storage, not an execution workspace.
The tool runs Cargo metadata on the original baseline before its later placeholder
workspace isolation. Use a fresh invocation-owned system temporary working root
outside the current Cargo workspace and any discovered ancestor Cargo.toml.
Do not patch the published manifest to suppress ancestor membership errors.
Preserve actual current cwd/semantic configuration and the approved exact argv.
The current generated root may remain under the actual workspace target. Verify
current/baseline project, target and JSON paths really differ before retaining
both sides, including equal-version cases. Copy required bytes before owned cleanup,
and never delete a shared/preexisting current target path on a slug match.

Keep raw absolute producer paths unchanged as data. Portable consumers authenticate
retained relative files plus component/platform-aware role-prefix/source-identity
bindings, not existence of vanished original paths or execution of retained
manifests. Reject role conflicts/aliases, unsafe retained paths and missing hashes.
Record auxiliary metadata honestly with unchanged generated locks/configuration.
Empty-compatible-line applicability does not invent tool check-release/rustdoc
rows, and --version proves only executable identity. Native host/compiler canaries
and any separately performed migration witnesses keep their actual scope.


## Signed hosted proof, independent trust policy and deciding authority

This closes the existing evidence contract before implementation in the existing
script, tests and workflows. It adds no public CLI, tracked file, service, release
payload or product API. Root separately reviews and adopts this append. None of
the preparation packets below is production integration or hosted acceptance.

### Independently owned root and verifier policy

The approved public default trust snapshot has SHA256
65ca537f6ed8a47fd0e560c421baa1f6c1efb8b25fc200d8c5c02c0e92eb2b9c.
Root actually acquired its two JSON objects using the independently trusted local
gh2.97installation, default embedded official TUF anchors, explicit github.com
hostname, removed token variables and owned empty TEMP configuration/cache/state.
The failed DNS predecessor remains failed. This approves exact snapshot bytes as
future policy input, not an attestation, native result or publication approval.
The Homebrew producer executable hash remains a local observation, not a universal
verifier pin or execution of the vendor Mac executable.

Embed the exact approved root JSONL bytes as a bytes literal alongside their hash
in existing scripts/sprint_workflow.py. Verify that literal's hash before writing
an exclusively created owner-only root file beneath the consumer's owned TEMP,
then recheck its actual bytes before gh. No new tracked trust file is added. Own
the concrete gh2.97.0 three-host mapping in that same existing script. Copy official archive URL, exact executable member,
archive/executable hashes and sizes from the root-authenticated
fx275-gh297-verifier-pins-v1/proposed-policy-constant.py. Its records are Linux
x86_64-unknown-linux-gnu, Mac aarch64-apple-darwin and Windows
x86_64-pc-windows-msvc. Root authenticated all three archives against the official
checksum manifest AND release API digests, and exact inner executable bytes.
No vendor executable has thereby run. Select the verifier from independent deciding
consumer policy and actual consumer host, never a hash supplied by the evidence
packet. Resolve the local gh executable through the deciding caller's prepared
PATH and require the selected vendor executable hash. A different installed binary
refuses rather than silently replacing policy. Authorized setup prepares the
verified vendor tool outside frozen consumption. Verify both archive and executable before use and retain real process
identity/argv/input/output facts. Unsupported hosts, absent policy and pin drift fail.

A packet may transport the approved root copy because its bytes must match the
independent tracked hash. It cannot choose its own CA or verifier. A root rotation
requires deliberate policy review, retaining historical verification authorities.
If an authorized acquisition refreshes default TUF material, differing bytes refuse
until reviewed. Evidence-only consumption never fetches a replacement root. Scope
verifier child configuration/cache/state/data to owned TEMP, remove credentials
unneeded by local verification and select github.com explicitly through supported
host configuration. Do not invent an --offline or --run-attempt flag.

### Exact certificate-bound local verification

The deciding consumer runs its pinned native gh, rather than trusting a saved
verified.json or workflow-supplied authenticated flag:

    gh attestation verify PACKET -R tensorbee/rdocx --bundle BUNDLE --custom-trusted-root APPROVED_ROOT --signer-workflow tensorbee/rdocx/.github/workflows/ci.yml --signer-digest SIGNER_SHA --source-digest SOURCE_SHA --source-ref PRODUCER_REF --cert-oidc-issuer https://token.actions.githubusercontent.com --deny-self-hosted-runners --format json

Require genuine exit0, no launch/timeout failure, exact input hashes/argv and direct
raw stdout binding. Support exact gh2.97/sigstore-go1.2.2 verified-output fields.
Select one unique full signed subject AND certificate source/workflow/run context,
not the first digest match. The sole subject is the exact sealed packet SHA256.
A foreign same-packet attestation may precede the matching result. Duplicate full
matches, missing required certificate fields or foreign-only results refuse.

Use verified signature.certificate issuer, subjectAlternativeName, buildSignerURI/
Digest, runnerEnvironment, sourceRepositoryURI/Digest/Ref/Identifier, source owner
URI/Identifier, buildConfigURI/Digest, buildTrigger and runInvocationURI. Enforce
https://token.actions.githubusercontent.com, tensorbee/rdocx immutable repository
ID1164240333and owner ID263075749, independently expected source/signer/config
identities, github-hosted and exact
https://github.com/tensorbee/rdocx/actions/runs/RUN_ID/attempts/RUN_ATTEMPT.
Do not substitute certificateIssuer for the OIDC issuer extension. Keep source,
signer and initiating workflow revisions separate when they actually differ.
The expected producer ref belongs to the authenticated producer, not silently to a
later consumer ref. H/R/M reuse still requires actual relevant-input equivalence.

Require nonempty verified signed Tlog/TimestampAuthority observations with valid
zoned timestamps and retained source identities. CurrentTime or predicate-only
timestamps cannot replace them. Impose no invented max-age and do not manually
remove historical authorities. The trusted verifier checks cryptographic temporal
validity under the approved snapshot. The SLSA workflow/source/immutable IDs/event/
runner/builder/invocation predicate only corroborates the already selected
certificate. Workflow-controlled predicate fields never supply missing authority.

### Concrete public-repository workflow and permission path

In existing ci.yml, use three named native producer jobs with fixed host/target
policy and contents:read. They execute the required source-qualified matrix/canaries
or acquire only missing/invalidated rows under the approved reuse rules. Each freezes
one packet before upload and exposes the pinned upload action's artifact ID/digest,
packet digest and exact input identity as its own named job outputs. Three explicit
job IDs avoid ambiguous same-key matrix output aggregation. Always preserve failed
raw diagnostics, without signing a failed/incomplete subject as clean proof.

A separate signing matrix in the same existing workflow needs all three native
producer jobs. It alone receives contents:read, id-token:write and attestations:write.
Its admitted events are a reviewed push to main or maintainer-authorized
workflow_dispatch on the exact reviewed owned-repository sprint/source ref. The
latter is required before sprint close and remains subject to the existing source
availability checkpoint. The normal manual build-only path uses owner_compatibility
false. Preserve all current required dispatch inputs and source checks, including
source_sha, source_ref, workflow_sha, runner_label and config_sha. The existing
owner-only job/gate route is unchanged. Its runner_label/config_sha do not become
SemVer native-host authority. The owner_compatibility true external corpus mode
and supply-chain schedule do not qualify as this hosted native union.

Before signing, validate actual GitHub repository IDs, workflow/event/ref/head SHA,
run/attempt, reviewed workflow revision and the complete required native producer
job/conclusion/runner union against authenticated API/control-plane facts. For a
dispatch, supplied source/ref/workflow inputs must match those actual facts and the
reviewed checkout. A string input or actor name alone is not admission authority.
Do not dispatch a base-repository ref then claim a different checkout SHA as its
certificate source. An admitted fork contribution must first be reviewed and made
available on an owned exact ref under the existing authorized source checkpoint.

Each signing matrix entry selects one fixed native producer job's explicit artifact
ID/digest/output identity. Download exactly that artifact ID, verify the wrapper
bytes and parent run/head/attempt binding, safely inventory its payload and require
the exact one expected packet/digest. Reject ambiguous names, merged payloads and
ambient user-supplied subject paths. Execute no downloaded native code. Attest that
one unchanged packet with actions/attest-build-provenance at the existing pinned
4d101475d8b20a2381f78447822ac1eab6504dd8. Retain its original bundle outside the
signed subject. Separate invocations keep each statement a sole packet subject.
Upload packet and bundle as diagnostic proof, preserving actual artifact IDs,
wrapper digests and authenticated native-producer associations.

A signing certificate identifies its signer workflow/run/attempt, not numeric job
ID, native architecture, test success or latest history. The trusted reviewed
signing workflow and deciding API checks bind its packet to the actual native
producer jobs. All current-run workflow/job conclusions are checked after those
jobs finish, not inferred prospectively while the signer itself runs.

pull_request from forks receives no id-token:write or attestations:write, and
unadmitted PR source never gains them through pull_request_target, workflow_run or
an automatic privilege promotion. Unsigned PR diagnostics are explicitly
nonqualifying. Until a reviewed source has its authorized hosted union, the deciding
SemVer/full/wheels/release requirement remains unsatisfied. Do not count an expected
signing skip or a diagnostic green as final gate success. Same-repository PR status
also cannot invent portable signature authority. An existing applicable qualified
union may be consumed only under the exact-source/history rules, never merely
because a maintainer clicked approval.

ci-gate has explicit required native/signing/evidence dependencies and validates
the resulting union fail-closed. Actual event routing, missing/failed/canceled/
required-skipped jobs and missing proof refuse at deciding boundaries. Re-derive
job counts after implementation. Diagnostic tooling packets/bundles are outside
the unchanged26product payload inventory. Build-only rehearsal still skips the four
publication jobs. No source push, dispatch, signing or publication is authorized by
this technical append itself beyond the canonical workflow's existing authority.

### Concrete independent catalog owner and unchanged CLI

Use the existing CURRENT_SPRINT parser and state_path/load/save owners in
scripts/sprint_workflow.py. Add one semver_authority context to that existing
.claude/scratch/SNN-run.json. The state owner is resolved from the canonical
CURRENT_SPRINT, never from an evidence/reuse directory, packet field or environment
pointer selected by the evidence. State and catalog paths must be real owned paths
outside every supplied evidence/reuse tree, with no symlink/alias route into one.
The catalog is an immutable acquisition snapshot, not a second tracker.

A concrete private _collect_release_semver_authority(sprint, scope) in the existing
script performs the actual trusted deciding-controller/API acquisition. It derives
the validated family/source/config/input/policy scope from the trusted checkout and
caller policy, verifies real repository/workflow/event/producer identities, collects
all required API pages and attempts/jobs/artifacts without success/artifact-only
filters, and freezes raw request/response/status/time/hash facts with its cutoff.
The independently run controller establishes that origin. Saved JSON, labels and an
artifact-provided authenticated flag cannot substitute for this actual acquisition.

The controller writes the exact catalog SHA256/path, expected source/config/policy/
input identity, selected scope, immutable repository IDs, allowed workflow/event
context, complete run/attempt/job/artifact bindings and cutoff into semver_authority.
Preserve all existing runstate feature/review/verification/worker facts. Do not
reset the sprint or record feature/full-gate acceptance as a side effect. On failed,
incomplete or changing acquisition, block the new context and retain the genuine
failure. Do not silently fall back to an older context or cutoff.

The existing release-semver --check-evidence handler calls a private
_load_release_semver_authority(sprint, scope) through that fixed canonical state
owner before deciding evidence. It authenticates state/context scope and the
independently expected catalog digest, every raw/cutoff/job/artifact/policy binding
and supplied packet association. Evidence directories may contain matching catalog
copies for transport, but can neither select the authority state nor set its
expected digest. Missing, mismatched or incomplete independent context refuses.
No new catalog CLI argument, import command, trust Boolean or public API is added.

Canonical /verify and /release controllers call the private live collector first
at their required freshness boundary, then invoke the unchanged evidence CLI.
Hosted ci-gate and wheels do the same in their independently owned trusted job
checkout, with only the required contents:read/actions:read collection permissions.
If that checkout lacks local scratch state, initialize the valid EXISTING runstate
schema before adding the context. Never import a packet's state, force-reset an
existing run or invent acceptance. Local scratch portability is therefore not
assumed: each fresh deciding hosted/release consumer establishes its own authority
origin and exact expected digest by actual collection before frozen consumption.
No extra collector-signing job or new tracked carrier is needed.

Pure offline --check-evidence uses only that already established independent local
context and frozen catalog/cutoff. It performs no hidden API/TUF refresh. An
arbitrary supplied directory on a machine without such trusted state cannot become
deciding proof and honestly refuses. A new deciding machine first runs its normal
authorized live collector. Signature verification still uses the independent root,
local pinned verifier and original bundle. This separates portable signed packet
identity from current complete/latest-attempt authority without weakening either.

### Complete attempt catalog, offline scope and remaining proof

A signed green packet cannot establish absence of a newer failed attempt. The
independent deciding caller supplies a finite complete authenticated authority
catalog and cutoff, using existing verification state plus actual GitHub run,
attempt-qualified job and artifact observations. Enumerate the admitted workflow
run inventory and all pages without success-only or surviving-artifact filters.
Include actual retries of older-created runs and every required native/signing job,
including failures without downloadable artifacts. Authenticate scope/relevance,
repository/workflow/event/source identity and original artifact binding before
ordering. An artifact's self-written catalog cannot authenticate this completeness.
If pagination, cutoff or relevance cannot be established, retain incomplete status.

Within one run use actual run_attempt. Across runs use qualified actual intervals
or authorized acquisition-round predecessor edges, not initial run_number. Cycles,
ties/overlap without a qualified ordering edge, incomparable maxima and latest
failed/canceled/incomplete/required-skipped attempts block older green for the same
relevant inputs. A genuine later success may resolve that failure. Preserve exact
producer identities and failed raw observations. A frozen catalog claims no
knowledge beyond its authenticated cutoff. Fresh deciding callers acquire current
authority separately, including actual conclusions after signing completed.

Evidence-only mode validates supplied frozen catalogs, local packets/bundles,
independent policy roots/verifiers and all referenced bytes without hidden API,
TUF, canary or API-tool execution. Its existing --refresh-registry option remains
registry-only. Full verify, wheels and release independently acquire their required
fresh authority and mutable owner/registry/tag facts at their deciding boundaries,
then consume applicable proof without rerunning unchanged native rows. Signature
verification itself must genuinely run on the trusted consumer or reuse an already
applicable independently authenticated consumer observation, never saved JSON alone.

Local helpers/prototypes and official source/pin acquisition are preparation only.
Still unresolved as actual proof are three-host vendor verifier execution, genuine
packet signing and certificate/timestamp JSON, permission/event/source enforcement,
artifact-to-native-job binding, complete/latest attempt authority and the final
source-qualified native/API/canary/hosted/full union. Missing signing capability,
root policy, required certificate field or host is failed/pending proof, not a gate
waiver. Existing native/product/hash/installed/WASM/hosted riders and separate final
release approval remain unchanged.


## Renewed canonical source availability preparation

Local implementation is reviewed by F-X275all-pass4. The current-minor carrier packet and full-gate packet retain their actual local observations, including46selected policy tests and affected follow-up checks. These checks implement and exercise local paths. Actual three-host API/compiler execution, signatures, complete latest hosted authority, verified new archives, installed consumers and final union acceptance remain pending.

The validated F271external-ci-source-availability handoff was consumed at integration a13b12927af21290ba0ea9eb512cc7f2639b9102. Reuse only that historical consumed-marker fact. The renewed checkpoint uses the existing run-sprint canonical-source rule, with fresh reviewed source, local evidence, remote/ref/SHA/workflow bytes and consumed-integration ancestry recorded in existing progress and S96run state before any push. It does not reuse1431source qualification, fabricate an F275worker handoff or change the process rule.

Expanded pending gates are the existing selected Test, MSRV, Word/presentation fidelity, Python, WASM, hash, package, lint, docs, prose and supply-chain CI union, all three native SemVer producers, all three signing jobs and the deciding CI gate. Completed successful exact-source CI must precede the build-only wheels rehearsal and its new completed-CI/SemVer consumer. Owner-only, schedule and fork skips do not supply that union. The normal dispatch retains owner_compatibility=false and exact source/ref/workflow commit identities. Required unused owner runner/config inputs remain nonqualifying sentinels.

All implementation and scoped-review source must be frozen and clean before the narrow nonforce canonical sprint-ref push. F271A7diagnostics require their own clean review and local checks first. Every actual external run and attempt remains pending until observed. No ledger completion, dependency unblocking, baseline update, main merge, tag, reporter comment, yank or publication follows from local readiness.
