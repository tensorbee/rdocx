# F-X275 registry pacing implementation support review

Date: 2026-11-08. Mode: independent bounded support review only.

## Disposition

Zero defects, zero smells and zero nitpicks in the five-file working increment.
This is a support observation over the hashes below, not a new formal microscope
lineage, review-counter reset, feature completion or release approval. Root owns
formal reconciliation. Product Rust repairs are outside this reviewed lease.

## Correctness and refusal

`scripts/sprint_workflow.py:2223` installs pacing as an actual urllib request
processor. Every HTTPS crates.io or www.crates.io `/api/` request waits1.05seconds
before transport. The two existing synchronous clients install that processor
at lines2247and3840. Redirected HTTPS API requests reenter the opener and therefore
receive their own delay, including entry from a non-API URL. This fixes request
spacing rather than merely delaying one outer batch. Non-API static.crates.io,
GitHub and PyPI requests are outside the condition.

The native client's existing HTTPS-only redirect guard remains present at
`scripts/sprint_workflow.py:2236`. HTTP errors retain their bodies and status at
line2252 and still refuse after receipt recording. No success retry, endpoint
substitution, error swallowing, alias or API-status waiver is introduced.
Publication verification retains its existing redirect behavior, exact package
versions and Python filename/hash comparisons. Its changed opener adds pacing
without bypassing those integrity checks.

## Routing and identity

`.github/workflows/ci.yml:956` and line1043 serialize Linux, Mac and Windows native
producers. Each successor retains its prior repository/event/mode condition and
adds `always()`, so a failed predecessor does not suppress another required native
observation. All-three signing dependencies at line1129 and deciding native/sign
results at lines1380to1383 remain intact. This does not turn an earlier producer
failure into a green union or introduce a canceling shared concurrency group.

`.claude/plans/F-X275-design.md:882` separately requires prior attempts to be
terminal before the orchestrator starts a coherent successor. The code does not
pretend to coordinate unrelated external API callers. HLD15:734 records this
bounded route accurately.

The policy-script, workflow and plan changes invalidate relevant existing source
and policy identities through the unchanged `_semver_source` at
`scripts/sprint_workflow.py:2095` and `_semver_policy` at line2113. Old raw403,
failed or incomparable attempts are not reclassified. Fresh actual final-source
hosted execution remains due under plan:884. No change relaxes signed packet,
artifact, baseline, target, complete-inventory or latest-attempt authority.

## Test mechanism and observed evidence

The new transport control at `scripts/test_sprint_workflow.py:345` retains the
real urllib opener/redirect/error-processing path and substitutes only the HTTPS
transport plus clock. It observes transport request-start times, API-to-API and
non-API-to-API redirects, a403body/status that must still refuse, a subsequent
request, and unchanged non-API timing. It does not infer pacing merely from the
presence of a sleep call.

The routing control at line386 rejects removed dependency edges and removal of
`always()`. The adapted publication integrity test at line8176 drives actual
opener redirects and retains the existing wrong-PyPI-file discriminator.
The reviewed fixture source and tests do not change acceptance thresholds or
turn a failed acquisition into a compatibility result.

Root's retained `REGISTRY_PACING_GREEN.json` reports actual exit0 for the four
focused methods. Its policy-script hash matches this review's exact source.
The associated log records four tests and OK. This reviewer read those receipts
and did not execute or rerun tests, external APIs, CI, Cargo or product runtime.
Root's prose/adapter/diff outcomes are not independently re-executed here.

## Review custody

HEAD: `e430071d75ed33e8831752d644aec3c0a865d3a3`.

Working increment SHA256: `148bb8b0fc88de02a471a19bfa3c63cdb75e96ceba3bb33bb1f13006bec82515`.

- `scripts/sprint_workflow.py`: `90de54ac6e1de30793b630f3f19a8caefdc999a2b7def22b202552e977d246d4`.
- `scripts/test_sprint_workflow.py`: `0ce2f67bb28fd336047e02617462e73bb625fdee79875a545f936b353e3707dd`.
- `.github/workflows/ci.yml`: `6807b48f662714dde2f03441ca642c5ed47ce7a6a8c3a35da4314904ad995079`.
- `docs/hld/15-build-and-toolchain.md`: `2f2129e861d8785d23a940a337adf40ae8584ab3a153f4e31192763607eb170f`.
- `.claude/plans/F-X275-design.md`: `b938cf9e6499024d1093462fe365230c5b9ee71c13d94ebb4c0aa643d0ad7dc7`.

Retained focused receipt SHA256: `f694cad342283977e2be0aeffcdb2a95aac07e7d818d4543f0c9c7fff3955c3f`.

The five reviewed files were stable across report preparation. Only this report
was written outside the worktree. No fixes, status edits, dispatch, publication,
new gate or review-cap mutation follows from this support pass.
