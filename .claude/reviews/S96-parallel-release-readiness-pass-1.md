# S96 parallel release readiness, pass 1

Read-only audit on `sprint/s96`, HEAD
`cec87f9501eb13d0d5d60572472b227483262620`, with the existing dirty integrated
tree preserved. Root retains the sole product/Cargo writer lease. This pass
identifies independent preparation work while Root implements performance
repairs. It is not a new final verification, source freeze or completion gate.

Read AGENTS.md, CLAUDE.md, the canonical workflow, current sprint/run state,
F-X275 plan/progress and the affected release-note, archive, memory and F-X278
reviews. No Cargo, test, benchmark, oracle, native application, network,
publication or external message ran. Only this review was written.

## Defects

Zero product defects established in this bounded readiness audit.

## Smells

Zero product smells established.

## Nitpicks

Zero nitpicks.

## Work ready now

The new Issue433 follow-up has an authenticated additional contributor outcome.
The preserved primary comment identifies set2374 and provides the numbering-level
reproduction at
`.claude/scratch/S96-final-evidence/f278-issue433-numbering-comment-20261108-1/issue433.json:1`.
The current Word credit at `CHANGELOG.md:48` credits set2374 only for Issue423.
The separate hadim issue/PR credit at `CHANGELOG.md:49` should remain intact.
F-X278's bounded review records the exact new case and actual local/installed
evidence at `.claude/reviews/F-X278-codex-final-3.md:38` and `:56`. This supports
a small prepared-note amendment without claiming all Issue433/native acceptance.

Suggested replacement for the existing Word credit sentence:

> [@set2374](https://github.com/set2374) reported section PAGE formatting in
> [#423](https://github.com/tensorbee/rdocx/issues/423) and supplied the
> numbering-level first-line indentation reproduction in
> [#433](https://github.com/tensorbee/rdocx/issues/433).

This adds no contributor name or contribution record to either family set.
Do not add the Word reproduction to the shared/PowerPoint credit. Preserve
the published suffix beginning at `CHANGELOG.md:84`. Existing focused checks
are sufficient for this metadata increment, without a new test file or product
suite. Their actual assertions are at `scripts/test_sprint_workflow.py:920`
and `:927`:

```bash
python3 -m unittest scripts.test_sprint_workflow.SprintWorkflowTests.test_s96_release_notes_preserve_published_s954_history scripts.test_sprint_workflow.SprintWorkflowTests.test_s96_release_notes_keep_exact_family_contribution_inventory
python3 scripts/prose_check.py
git diff --check
```

These commands are recommendations, not executions by this reviewer. Root may
apply the edit under the existing sole-writer lease. No reporter reply follows:
the held communication conditions remain at `.claude/plans/F-X275-design.md:191`.

## Existing proof and current limits

The release implementation has a clean bounded all-aspect review, with the
feature explicitly in-progress at `.claude/reviews/F-X275-all-pass-5.md:5`.
That review distinguishes actual measured archive intent from compilation-
verified packaging at `:29` and `:33`, and preserves the native/hosted/installed
debt at `:53`. Later archive support similarly authenticates actual measurements
without claiming final hosted acceptance at
`.claude/reviews/S96-root-archive-metadata-support-1.md:9` and `:11`.

Canonical run state records genuine H11 normal CI and signed native-union
success: 24 successes and two expected optional skips for exact HEAD
`cec87f95`, at `.claude/scratch/S96-run.json:4048`. Its scope explicitly excludes
later dirty repairs at `:4056`. This is stronger than the older failed H9
checkpoint and must not be reported as if no hosted success ever existed.

Present custody is distinct. This audit checked the literal H11 handoff path
named at `.claude/scratch/S96-run.json:4053` and checkpoint path at `:4068`.
Both are absent from the current filesystem. The older archive directory named
at `.claude/reviews/F-X275-final-archive-intent-support-6.md:33` is also absent.
The historical observations remain recorded, but this pass cannot reauthenticate
their missing raw packets. Any deciding reuse must reacquire actual authenticated
custody and complete attempt authority, or consume fresh final-source proof.
The contract requires raw artifact/source provenance and later-attempt handling
at `.claude/plans/F-X275-design.md:141` and `:153`. No network refresh was needed
to establish these local custody facts.

Current archive constants remain declared measurements at
`scripts/readme_doctests.py:405`, `:408` and `:410`. They are not new measurements
of ongoing Native, layout or OXML repairs. Refresh genuinely affected package
inputs after the source settles, retain applicable unaffected evidence only
with custody/source checks, and keep the existing 64-byte tolerance at `:394`.
F-X275 A4 still requires the integrated verified packaging/README gate at
`.claude/plans/F-X275-design.md:221`. Binding/support crates with `publish=false`
at `crates/rdocx-py/Cargo.toml:6` and
`crates/oxml-py-support/Cargo.toml:6` instead affect installed wheel evidence.

## Prioritized remaining preparation

1. Apply the bounded Word contributor credit above and run only its existing
   focused metadata checks. Preserve family separation and published history.
2. Prepare a source-to-evidence applicability inventory for the final freeze,
   including missing H11/archive custody and the truly affected package rows.
   Use the existing run state and plan, not another maintained tracker. Do not
   publish, push or launch the build-only rehearsal during active repairs.
3. Close the existing memory reconstruction contract before implementation.
   `.claude/plans/F-X268-design.md:452` still specifies a pre-cache commit pair.
   The proposed same-final-source omission registry is explicitly not an existing
   overlay at `.claude/reviews/S96-memory-omission-endpoint-design-investigation-2.md:5`.
   Its borrowed-result and binding obligations remain real at `:9`. The prepared
   20,000-by-six fixture has only structural checking, with reopen and RSS pending
   in `.claude/scratch/S96-final-evidence/memory-final-source-preparation-20261108-1/FIXTURE.json:32`.
   An explicit existing-plan amendment, full equivalence, calibration, actual
   per-layer and combined RSS, and activated N/4N history are still required.
4. Once the performance repairs are reviewed and the final source is frozen,
   execute or consume applicable final archive, installed Python3.9/3.12,
   strict typing/stubs, CLI/WASM, native/schema/oracle, hash/pixel and full union
   proof. These are final-union checks, not additional per-worker runs under
   `.claude/commands/verify.md:17`. The current-source normal CI and build-only
   26-payload rehearsal order remains at `.claude/plans/F-X275-design.md:594`.

All four latest performance misses remain Root's active implementation work,
not waived by this audit. Issue430's narrow unproved native exception and all
other applicable native obligations remain distinct at
`docs/sprints/CURRENT_SPRINT.md:66`. Current installed/native/full gates remain
unchecked in the canonical definition of done at `:111` through `:124`.
Issue429's bounded recovery guidance remains distinct from a universal old-line
patch or registry action at `.claude/plans/F-X275-design.md:166` and `:168`.
Chrome PyPI login is only user-reported readiness at
`.claude/scratch/F-X275-progress.md:1371`. It supplies no publication approval.

This audit changes no status, ownership, ledger, baseline, commit, push, tag,
registry state or reporter communication. Return control to Root.
