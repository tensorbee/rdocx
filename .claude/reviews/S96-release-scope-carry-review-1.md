# S96 release scope carry review 1

Date: 2026-11-08. Read-only metadata review of the adopted Q-S96-16 amendment on
`sprint/s96`, against HEAD `cec87f9501eb13d0d5d60572472b227483262620` and the
current dirty working tree. This pass writes only this report. No policy tests,
product runtime, Cargo, timings, network requests or publication actions ran.

## Disposition

Zero blocking findings in the scope amendment, successor ownership, capability
ownership or published-history categories. One minor Markdown rendering finding
is recorded below. Numerical performance and unperformed memory qualification
are carried under the accepted amendment. They are not additional S96 release
gates and no failed or unperformed result is recast as passing.

## Finding

1. Minor, backlog table continuity. `docs/sprints/BACKLOG.md:866` inserts a blank
   line between the existing F-X283 row and the new F-X284 row at line 867. The
   three successor rows consequently sit outside their Markdown table header.
   Remove that blank line so the pending S96.1 rows render in the existing table.
   This is a presentation correction, not a change to their parsed status or a
   release acceptance defect.

## Satisfied metadata checks

- Q-S96-16 is fresh relative to the committed predecessor. A read-only
  `git grep` found no Q-S96-16 occurrence at HEAD. The recorded previous occupied
  range ends at Q-S96-15 in `.claude/plans/F-X272a-design.md:580` and its explicit
  historical mapping at line 615. The adopted decision therefore does not reuse
  Q-S96-08 through Q-S96-15 or reopen their meanings.
- `docs/sprints/CURRENT_SPRINT.md:9` states the user's release-best decision and
  bounded implementation freeze. Lines 11 to 15 retain the pre-final-wave
  76/79 flat result, five linear and both Issue407 passes, three flat misses,
  numerical and unfinished qualification carry, and current safety/release
  gates. The F-X282 and F-X283 DoD rows at lines 133 and 134 explicitly preserve
  non-numerical correctness and leave Issues437/445 open.
- `.claude/plans/F-X275-design.md:859` and
  `.claude/plans/F-X268-design.md:1256` contain matching operative amendments.
  They explicitly override earlier numerical or final RSS/history/calibration
  prerequisites for this release. The retained gates include held handles,
  last-loan lifetime, Source/resource limits, refusal/security, hash/pixel,
  applicable native/oracle proof, installed runtime/typing, strict SemVer,
  verified packages and actual hosted-source proof. Separate exact-main-SHA
  approval is retained. No optional yank, completion or publication is approved.
- `.claude/plans/F-X282-design.md:830` and
  `.claude/plans/F-X283-design.md:1310` carry numerical acceptance unchanged to
  F-X285 after F-X284 review, with unfinished memory qualification assigned to
  F-X286. Saved bytes, activation, ownership and refusal/resource obligations
  remain current. Their earlier numerical rows remain successor contracts,
  rather than secretly weakened thresholds.
- `docs/hld/12-testing-strategy.md:3` acknowledges the same split without changing
  original protocols. `docs/hld/14-development-backlog.md:8953` also makes the
  F-X282 numeric carry explicit. Its F-X284, F-X285 and F-X286 entries at lines
  9041, 9051 and 9061 assign review first, unchanged numerical protocols second,
  and concrete A7/Q18/Q22 memory qualification. The memory entry preserves
  omission endpoints, mandatory identifiers/provenance, Native/Python consumers,
  calibration, fresh alternating pairs, raw units, equal-terminal history and
  last-owner release. It calls the 20k6 fixture reconstructed and the interrupted
  smoke unqualified. None becomes a present measured RSS claim.
- `docs/sprints/SPRINT_PLAN.md:3531` inserts S96.1 immediately before the original
  S97 heading at line 3543. Lines 3537 to 3539 match the pending M/L/L stories in
  `docs/sprints/BACKLOG.md:867`. F-X285 and F-X286 depend on the opening F-X284
  review in HLD14, with approved designs required before implementation. S97 and
  later numbering is preserved. Fractional sprint IDs are supported by the
  existing parser at `scripts/sprint_workflow.py:76` and line 91.
- `docs/sprints/BACKLOG.md:948` maps the carried F-X268, A282-1 to A282-3 and
  A283-1 to A283-4 numerical qualification to F-X285 while retaining
  non-numerical obligations. `docs/sprints/SPRINT_TRACKER.md:1272` records planned
  successor work without actual-day, completion or velocity credit. Historical
  F-X268 completion is explicitly retained.
- `.claude/scratch/S96-run.json:1404` records
  `F-X275.accepted_scope_amendment`, decision Q-S96-16, canonical paths,
  F-X284/285/286 ownership and S96.1. Its `status_or_completion_changed` value is
  false at line 1416. The metadata amendment does not pretend current feature
  verification, sprint closure or publication has happened.

## Capability ownership

No capability-owner row needs transfer for this carry. The canonical matrix
starts at `docs/hld/02-scope-and-non-goals.md:223`, with its owner column at line
266. It has no F-X268, F-X282, F-X283, F-X285 or F-X286 occurrence. Numerical
qualification carry therefore does not create a missing owner for an existing
matrix row or justify marking a capability complete. The applicable existing
classification and live-owner checks are
`test_modern_docx_capability_matrix_has_closed_classifications_and_evidence`
at `scripts/test_sprint_workflow.py:10268` and
`test_every_incomplete_modern_docx_row_has_one_live_owner` at line 10351.

## Release notes and immutable history

`CHANGELOG.md:46` adds the Word limitations section before contributors. Lines
48 and 50 identify the snapshot as preceding the final namespace wave, disclose
the three misses and unfinished memory/history/calibration, make no 6MiB or
68MiB claim, and leave Issues437/445 open. These are honest disclosed limitations,
not a new passed final-artifact claim. The words incomplete and unfinished do
not match the placeholder expression at `scripts/sprint_workflow.py:144`.

Read-only byte hashing found the complete published suffix beginning
`## v0.16.4` unchanged from HEAD, with SHA-256
`50b3e43c22f44f276537eea506b1e40136945fb8f9228d64a4bf50f5d4c244f8`, matching
the existing historical guard at `scripts/test_sprint_workflow.py:920`.
The selected PowerPoint section is also byte-identical to HEAD.

Read-only extraction of the selected-family links and contributor names matches
the exact inventories at `scripts/test_sprint_workflow.py:927`:

- Word issues 423,424,425,426,429,430,432,433,434,435,436,437,445 and PRs
  427,428,439,440,441,442,443,444. Contributors remain set2374, hadim, arthrod and
  lawrenceakka, including the previously added Issue433 credit.
- PowerPoint issues 429,431,433,435,436,437 and PRs 438,439,440,442,443.
  Contributors remain hadim and arthrod. The shared accessor text explicitly
  avoids a measured PowerPoint performance gain at `CHANGELOG.md:69`.

This inspection is not a policy-test execution. Root retains the existing
affected checks, notably
`test_s96_release_notes_preserve_published_s954_history`,
`test_s96_release_notes_keep_exact_family_contribution_inventory`,
`test_release_notes_require_complete_reviewed_changelog_sections`, and the
capability/roadmap checks already present in `scripts/test_sprint_workflow.py`.
No new test or gate is proposed by this review.

## Remaining disposition

Correct the single backlog table break, then consume the actual affected policy
results and continue the existing retained correctness/release sequence. Do not
schedule another 79/79 or RSS qualification packet as a condition introduced by
this audit. The pending successor stories require review and approved designs
in S96.1. Current feature completion, final integrated acceptance, sprint close
and exact-SHA publication approval remain separate facts, not outcomes of this
metadata pass.
