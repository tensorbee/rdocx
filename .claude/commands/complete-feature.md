---
description: Finish a feature. Updates every tracker, the HLD, and commits to the sprint branch.
---

# /complete-feature F-XXX [--prepare]

Close out an F-ID. Every tracker is updated in one commit so the delivery record
cannot drift from the code.

`--prepare` is the parallel-work variant a worker runs in its own worktree. It
stops short of the sprint ledgers and writes a handoff instead. See "Prepare
mode" below.

## Steps

1. **Check the preconditions.**
   - The F-ID is `in-progress`.
   - `.claude/plans/F-XXX-design.md` exists and every checklist item is ticked.
   - The latest `.claude/reviews/F-XXX-*` pass reports **zero defects and zero
     smells**.
   - `/verify` passes. Not `--fast`.

   If any fails, say which and stop.

2. **Confirm the test gate exists and is real.** Find the test named in the
   design plan. Run it. Then confirm it would fail if the implementation were
   reverted. A gate that passes against unmodified code is not a gate.

3. **Update the HLD.** Read the plan's `## HLD impact` and update **exactly
   those files**, no others. Replace stale prose with current reality. These
   documents describe what is true now, not a history of changes.

   If the implementation contradicted a spec section that the plan did not
   list, stop. Either the spec is wrong and the plan should have said so, or the
   implementation drifted.

4. **Append to `docs/sprints/AS_BUILT.md`** using the template in that file.
   The `**Hash harness**` field is mandatory for M1 through M6, recording either
   "unchanged" or the delta and its justification.

5. **Append to `docs/sprints/SPRINT_TRACKER.md`**: one row with the estimate and
   the actual.

6. **Set `done`** in `docs/sprints/BACKLOG.md` and
   `docs/sprints/CURRENT_SPRINT.md`, and regenerate the AUTOGEN counts.

7. **Set the design plan status** to `completed`.

8. **Delete** `.claude/scratch/F-XXX-progress.md` if present. It is in-flight
   memory and its job is over.

9. **Commit** to the sprint branch:

   ```
   F-XXX, short title

   One paragraph: what was built, why, and any non-obvious choices.

   Tests, <the gate plus any others>
   Harness, unchanged | <delta and justification>
   ```

   **No `Co-Authored-By` trailer.**

10. **Report** what changed, and the next pending F-ID in the sprint.

## Prepare mode

`--prepare` runs in a worker worktree on `work/<fid-lower>-<agent>`. The worker
owns its feature. It does not own the sprint's totals, and two workers appending
to `AS_BUILT.md` in parallel is a conflict in the one file that must not have
one.

Do steps 1 through 3 and step 7 normally. Then instead of steps 4 to 6 and 9:

1. **Write `.claude/handoffs/F-XXX-ready.md`**, exactly these fields:

   ```markdown
   # F-XXX ready for integration

   **F-ID**: F-XXX
   **Owner**: claude | codex
   **Branch**: work/f-xxx-<agent>
   **Worktree**: <path>
   **Base**: <the sha the claim recorded>
   **Head**: <the sha this commit will be>
   **Design plan**: .claude/plans/F-XXX-design.md
   **Microscope**: .claude/reviews/F-XXX-<aspect>-pass-N.md, 0 defects, 0 smells
   **Verify**: <scope>, pass
   **Hash harness**: unchanged | <delta and its justification>
   **Test gate**: <test name>, pass
   ```

2. **Validate it before committing.** A handoff that does not validate is not
   a handoff:

   ```bash
   python3 scripts/sprint_workflow.py validate-handoff \
     .claude/handoffs/F-XXX-ready.md --fid F-XXX
   ```

3. **Commit feature-local work only** to the worker branch: the code, the
   tests, the design plan, the review files, the HLD files the plan listed, and
   the handoff. Use the normal message format.

4. **Do not** touch `AS_BUILT.md`, `SPRINT_TRACKER.md`, `BACKLOG.md` or
   `CURRENT_SPRINT.md`. F-XXX stays `in-progress` with its `Owner`.

5. **Do not** delete `.claude/scratch/F-XXX-progress.md`. Integration has not
   happened, so the in-flight memory is still live.

6. **Do not push.**

Report the branch, the head sha, and the exact integration command:

```text
/integrate-feature F-XXX work/<fid-lower>-<agent>
```

### Release preparation

A release F-ID whose gate requires real publication has one narrow exception.
Its worker may prepare the reviewed release machinery before the irreversible
approval and publication checklist items are complete. All implementation
items, workflow tests, `/verify`, the hash harness and `/microscope` must still
pass. Leave the plan `approved`, record the handoff test gate as
`deferred to /release vX.Y.Z`, and integrate it in batch as `reviewed` rather
than `completed`.

After the integrated sprint passes `/verify --full` and `/sprint-review`,
`/release` obtains the separate final approval and proves the external gate.
Only then tick the remaining checklist items, set the plan to completed and
write the delivery ledgers. This exception does not apply to a dry-run-only
test gate or to any non-release F-ID.

## Refused situations

- **Any precondition in step 1 fails.** Name it and stop.
- **The latest review still has defects or smells.** Fix them and re-run
  `/microscope`.
- **The test gate does not fail against reverted code.** The gate is wrong. Fix
  it before completing.
- **Updating an HLD file the design plan did not list.** Either the plan was
  incomplete or the implementation drifted. Resolve it, do not paper over it.
- **Pushing, merging or tagging.** This command commits to the sprint branch
  only. Only `/close-sprint` touches `main`.
- **Completing an F-ID whose hash-harness delta was never declared.**
- **`--prepare` from the canonical sprint worktree.** It exists for a claimed
  worker branch. Run the normal command instead.
- **A `--prepare` handoff that records a `--fast` verify.** `--fast` is the
  inner loop. `validate-handoff` refuses it, and so should you.
- **Using release preparation to claim publication succeeded.** The F-ID stays
  reviewed until `/release` verifies the registry and GitHub release.
