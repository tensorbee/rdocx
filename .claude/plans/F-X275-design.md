# F-X275, Prepare Word 0.16.5 and PowerPoint 0.14.5 patch releases

**Status**: approved
**Sprint**: S96
**Size**: L
**Depends on**: F-X270

Root adopted and approved this bounded contract on2026-10-29under the user's explicit post-S96release instruction. F-X270is done and unchanged. No material product choice remains. Current raw registry and contribution evidence was independently authenticated by root. Canonical design validation and contract closure apply before implementation.

## Problem

F-X270 is completed in closed S95.4. Its0.16.4/0.14.4 packages and notes are published, so they cannot represent the newly requested post-S96 releases. Cargo.toml34 and55to78, both Python pyproject.toml7 and current README/policy carriers still select those published versions. Fresh2026-10-29 public registry receipts show all22Rust and2Python .4 entries occupied, all24exact .5 queries absent, and both .4 origin tags at24af95d3 with no .5 tags. All current release project metadata reports .4. Canonical release.md forbids duplicate versions, and design.md forbids revising a done F-ID.

## Spec reference

- docs/hld/10-bindings-spec.md, Packaging and CI.
- docs/hld/15-build-and-toolchain.md, Packaging and Release process.
- docs/hld/14-development-backlog.md, new F-X275 story after admission.
- .claude/commands/release.md, Family contract, Preconditions and Final approval.
- .claude/commands/release-notes.md, Evidence and Validate and review.
- docs/sprints/SPRINT_PLAN.md, S96 combined gate and each issue423to426 acceptance row.

## Approach

Prepare0.16.5 across the exact7Word Rust packages and Word Python distribution, and0.14.5 across the exact15shared/PowerPoint Rust packages and PowerPoint Python distribution. These are patch increments within the user's explicitly requested0.16.x/0.14.x lines. This is a semver choice, not a request for0.17.0/0.15.0. S96 product compatibility and refusal decisions stay in their existing approved provider plans. No new product API, XML policy, dependencies, feature flags or runtime algorithm is authorized here.

Update current manifest carriers, workspace internal pins, lockfile versions, binding metadata, README install examples and measured archive tables, current policy assertions and rdocx WASM inherited version contracts. Keep rpptx-wasm at0.12.1 and npm publication outside scope. Inspect each literal, never perform blanket replacement. Historical published notes, version mutation fixtures, historical recovery proofs, F-X270 and its delivery entries remain immutable. Existing s954 named tests that actually assert current carriers must be renamed or retargeted as current assertions. Their .4 inventory/history controls remain historical and passing. Add the new .5 acceptance controls to the existing test carrier without adding a test binary or source file.

Add two meaningful family sections to CHANGELOG.md via the approved release-note ceremony, preserving the entire suffix from `## v0.16.4` byte-for-byte to published24af95d3. Word notes cover qualified S96 PAGE formats, header/footer clearance, TOC and bounded preserved-revision comparison, including tested refusal/workaround limits and qualified native observations. Security wording accurately states the pulldown-cmark0.13.5 dependency update and advisory relevance, without claiming the disabled wikilink path was exploitable through the public importer. PowerPoint notes describe the actual packaging/aligned shared-family outcome. Do not import Word fixes or old S95.4 Arc migration as new PowerPoint product fixes. State honestly when no new family product capability or fix exists. Retain relevant compatibility limits with historical attribution where useful.

Rebuild selected-family contribution inventories from previous family .4 tags at24af95d3 through final reviewed S96 source. Public refresh authenticates Issue423 reporter set2374, Issues424/425/426 reporter hadim, and PR427/428 contributor hadim. Word inventory includes all6records with specific outcomes and hardened-equivalent adoption, not unchanged PR merges. Do not carry S95.4 records forward as new included contributions. PowerPoint inventory is empty unless the actual reviewed commit range establishes a real included record. CI restoration and roadmap bookkeeping are internal and do not become user-facing fixes. Record each candidate exclusion and specific planned post-publication comment in the existing release evidence, without posting it now. Refresh PR heads/body/comments before final notes freeze.

## Consumer map

| Rule | Entry points and outputs | Provider/caller policy | Lifecycle |
|---|---|---|---|
| Rust family versions | Existing manifest and lock carriers, validate_unified_release_family | Exact7/15allowlists, existing internal pins | Source, verified crate, registry preflight |
| Python identity | Both pyproject.toml and binding crate metadata | Exact distribution, wheel/sdist metadata and import identity | Build, download, install3.9/3.12, typed use |
| CLI identity | Existing rd/pptx CLI env package version | Six target archives per family | Build, validate archive, host --version |
| Notes and credit | release-notes check/render | Family-only claims and authenticated input records | Review, render, future release-body equality |
| Product acceptance | Existing F-X271to274 consumers | Their approved acceptance contracts remain binding | Native/cache/reopen/render and composed cases |
| Publication | release command after close | Clean main exact reviewed tree, mutable preflights, separate final approval | Sprint R, main M, eventual immutable family tag |

## Behavior contracts and evidence

No public signatures, numeric conversions or XML placement behavior are added by preparation. Those contracts stay in F-X272to274. Version identity is0.16.5/0.14.5 exactly, with no suffix or second Python-only tag. Primary facts are exact registry JSON URLs, Git origin refs and published .4 changelog at24af95d3. Read-only receipts are under /private/tmp/rdocx-s96-run-refresh/release-prep, with URLs, statuses, UTC datetimes and SHA256. Absence is a mutable preparation observation and must be freshly rechecked at release.

## Dependency readiness

F-X270 is done in BACKLOG and closed S95.4 and supplies the unchanged machinery. Its evidence is historical, not new S96 artifact acceptance. F-X271to274 are integrated but final S96 acceptance remains pending. They are final acceptance dependencies of this release preparation, not formal start dependencies that prevent routine carrier/note work. No reviewed state is changed to done to unlock this story. Root is sole writer. The outstanding Word layout remediation must pass its independent microscope before any source freeze. Final supply-chain, native Word, installed, hosted and other unpassed criteria stay named blockers.

## Acceptance criteria

| ID | Input/setup | Expected outcome | Consumers | Local evidence | Final evidence |
|---|---|---|---|---|---|
| A1 | Exact22Rust and2Python candidates and remote refs | All .4 occupied, all .5 absent now, current carriers consistently .5 | Cargo/Python/CLI | Actual receipts, carrier discriminator failing at old Base then passing after edits | Fresh mutable release preflight later |
| A2 | Published .4 changelog suffix at24af95d3 | Bytes unchanged, exactly one complete new section per family | Notes users | Historical suffix hash assertion and mutation refusal | Independent note review and exact check/render |
| A3 | Reviewed .4toS96range and six Word records | Every included link/outcome/authenticated handle in rendered Word notes, exact family separation and limitations | Users/reporters | Existing carrier inventory controls extended for S96, .4history retained | Final fresh record/head refresh, planned notification list, no comments now |
| A4 | Actual final-source22locally patched crate archives | Correct version/member bytes/README metadata, each below10MiB, existing compression tolerance retained | Cargo users | Current policy and inventory controls | One integrated packaging/README qualification, retained actual archives |
| A5 | Reviewed clean sprint source and manual wheels workflow |27deciding jobs succeed,4publication jobs skipped,26unique payloads authenticated | Both artifact families | Existing build-only and exact inventory mutation controls | Actual SHA-bound31job receipt and downloaded12CLI/12wheel/2sdist inventory |
| A6 | Downloaded host CLI and both host wheels | Exact .5 --version and installed metadata/import identity,3.9/3.12priority calls and3.12typing/stubs pass | CLI/Python | Existing selected metadata/ABI/install controls | Fresh artifact installs,42Word+44PowerPoint priorities per interpreter,172calls and strict typed/stub qualification |
| A7 | Final S96union and every product criterion | Complete applicable full union, actual native/source/composition acceptance, fresh affected dependency/security gates | All affected consumers | Source-bound existing local regressions and independent microscope | One deciding integrated full union, old failed attempts retained, missing evidence blocks |
| A8 | Completed preparation, reviewed sprint R and later main M | Honest applicable H/R/M evidence reuse, exact close provenance, both releases held for separate exact-main approval | Release maintainer | Existing closure/evidence metadata policies | Canonical close then release preflight, shared family before Word |

## Test matrix

| Case | Expected difference | Wrong implementation rejected |
|---|---|---|
| Published .4 versus absent .5 | .4refused as candidate, .5selected | Republish occupied manifests or invent a larger minor line |
| One stale manifest/pin/project | Family validation refuses | Partial rename or Python/Rust drift |
| New Word row absent or miscredited | Inventory check refuses | Credit everyone only to hadim, omit set2374 or call adapted PR unchanged |
| Word fix copied into PowerPoint notes | Family review refuses | Reusing identical notes for both families |
| Old .4section edited | Byte-integrity check refuses | Rewriting published history to include S96 |
| Old artifact/wrong SHA/ABI/target/duplicate member | Exact validators refuse | Adopting S95.4artifacts or aggregate copies as new subjects |
| Dirty/revised source or incomplete hosted/native gate | Freeze/completion refuses | Treating a successful dispatch or local suite as full acceptance |

## Rejected alternatives

- Revise F-X270. Completed plans and delivery history are immutable.
- Reuse .4 or jump to .17/.15. Registry occupancy and user-requested lines determine .5.
- Rerun all historical S95.4native/performance/RSS suites for metadata alone. Reuse only applicable evidence and rerun genuinely invalidated inputs.
- Add release machinery, workflow, source module or second tracker. Existing paths suffice.

## Test plan

**Test gate**: release regression.

Local scoped: extend existing scripts/test_sprint_workflow.py with current .5family and .4suffix integrity/inventory discriminators. Batch current version-family, metadata, historical-note and family-only inventory checks. Capture actual Base failures honestly. Reuse existing test_manual_wheel_dispatch_is_build_only, test_unified_release_family_rejects_mismatched_version_carriers, test_release_package_inventory_uses_unpacked_dry_run_packages and release-note mutation controls. Run both unified-release-family validators, both release-notes --check/--render, prose, adapters and diff checks. No new product regression is claimed for version metadata.

Final integrated union: actual22package verified dryrun/README qualification once at final source, full canonical S96 verify union and current-state metadata policies, actual build-only hosted run and downloaded-artifact qualification. Fresh full installed S96 suites remain separately required where already earned by product stories. Host downloaded priority/typing checks do not replace equipped full suites, CLI chains, native observations or conformance. Reuse applicable existing union evidence with exact producer and input identity. Version changes affect CLI/package/import/WASM identity and package bytes, so those checks must be fresh. Rendering/native evidence may transfer only through explicit independent source-consumer assessment. New pulldown-cmark and local layout changes require their genuinely affected fresh checks. Do not reuse an older pass over a latest failed full attempt.

## HLD impact

- docs/hld/10-bindings-spec.md, Packaging current family identities and immutable previous published .4 identity.

- docs/hld/14-development-backlog.md, new F-X275 only and genuinely current prepared-version wording. Preserve historical F-X270 story.
- docs/hld/15-build-and-toolchain.md, Packaging and Release process current family versions/previous published .4 identity and H/R/M proof reuse. Workflow wins over stale prose calling for redundant full verification at main.

## Risk routing

Release scripting/version strings adds full carrier/pin/lock/README inspection and separate final approval. Published API packaging adds actual verified .crate size/inventory. Binding metadata adds fresh installed3.9/3.12identity and3.12typed use plus existing integrated WASM gate. Contributed issue/PR acceptance adds exact input pins, six Word acceptance links and authenticated credits. Completion/status/evidence adoption adds final affected metadata policies. Product algorithm/parser/layout/oracle riders remain the existing S96union, not new redundant preparation suites. No baseline change is authorized here.

## Hash harness

Preparation expects no output delta in49hashes/7deterministic pixels. Preserve every separately reviewed S96product delta and original producer. Trace version consumers before declaring prior native/render evidence applicable. Unexplained deltas block acceptance.

## Source availability and build-only rehearsal

Use the already consumed F-X271 external-ci-source-availability handoff for the canonical reviewed sprint source push. Its real workflow changes and actual local/independent evidence provide the canonical checkpoint route. Version and note edits must be in the reviewed clean source before that checkpoint. Authenticate the updated integrated clean sprint ref, remote, exact SHA/ancestry and applicable focused checks/microscopes, retaining all pending final gates. Do not infer that a historical F-X270 exception grants new authority. Once that source is available through F-X271, execute the release preparation's mandated manual wheels.yml build-only rehearsal at the same reviewed versioned SHA, respecting all existing restrictions. No separate source push exception or fabricated worker handoff is proposed. If the F-X271 consumed handoff's actual validated scope does not cover the final integrated source push, stop at that exact canonical checkpoint conflict and resolve it within the authorized workflow before pushing. Source availability does not complete F-X271 or unblock any dependency. No force push, main merge, tag, registry upload, GitHub release, publisher change or notification occurs during preparation. Actual CI and build-only receipts remain pending until observed.

Recompute workflow and42/44priority inventory hashes from current source rather than adopting historical hashes blindly. Current workflow structure has12wheel+2sdist+12CLI+1aggregate deciding jobs and4publication jobs. Authenticate26unique payloads, duplicates in aggregate archives do not add subjects. Validate name/version/platform/ABI and full metadata before fresh host installs. Final source edits invalidate affected observations.

## Implementation checklist

- [x] Admit story in existing HLD/backlog/current sprint/plan/runstate and approve bounded contract.
- [x] Capture actual old-source carrier/history/inventory discriminators.
- [x] Align current carriers and notes, preserve .4history and old controls.
- [x] Update listed HLD current intent and obtain clean independent microscope.
- [ ] Complete source checkpoint and actual hosted rehearsal/artifact qualification.
- [ ] Accept actual final integrated union, explicit reuse and all four product contracts.
- [ ] Reconcile final metadata once and complete preparation, publication still held.

## Open questions

No material user choice remains. Current instruction requests both retained version lines at end of S96. Fresh occupancy selects .5. Routine mechanical choices can be recorded by root under that authority. The existing F-X271 consumed source checkpoint remains the route for source availability, subject to its canonical constraints. Preparation approval is not final publication approval. Missing native/hosted/full evidence remains work or a blocker, never a question to waive.

## Contract closure

No new callable product operation, numeric contract or XML placement. Exact family identity, per-record credits, artifact inventory, installed evidence, historic byte preservation and lifecycle publication boundary are specified. Approval records scope and routine metadata choices, not gate execution. Capacity must explicitly add bounded release preparation to the already extended sprint instead of retaining the old four-story-only cap as current truth.
