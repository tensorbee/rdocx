---
description: Integrate a prepared worker branch into the canonical sprint branch. Validates the handoff first and never resolves a semantic conflict on its own.
---

# /integrate-feature F-XXX work/<fid-lower>-<agent> [--batch]

Bring one worker's finished F-ID onto the sprint branch. Run this **only from
the canonical sprint worktree**, after the worker has run
`/complete-feature F-XXX --prepare`.

`--batch` is the mode `/run-sprint` uses. It creates the local F-ID commit and
defers the sprint ledgers, the `done` status, consolidated verification and the
push to the end of the sprint run.

## Preconditions

1. **Clean tree on `sprint/sNN`**, matching `docs/sprints/CURRENT_SPRINT.md`.

2. **F-XXX is `in-progress`** in `CURRENT_SPRINT.md` and `BACKLOG.md`, with an
   `Owner` matching the agent suffix on the named branch.

3. **The branch name matches `work/<fid-lower>-<agent>`** and its merge base is
   an ancestor of the current sprint branch. A worker that rebased onto
   something else is not integrable without a decision.

4. **The handoff validates.** Extract `.claude/handoffs/F-XXX-ready.md` from the
   worker branch to a temporary path and run:

   ```bash
   git show work/<fid-lower>-<agent>:.claude/handoffs/F-XXX-ready.md > <tmp>
   python3 scripts/sprint_workflow.py validate-handoff <tmp> --fid F-XXX
   ```

   Then confirm what the script cannot: its `Base` is the sha the claim
   recorded, its `Head` is an ancestor of the branch tip, and **no change to
   `crates/` appears after that head**. Evidence recorded before further edits
   is not evidence.

5. **The branch actually contains** the design plan, the implementation, the
   tests, and a `.claude/reviews/F-XXX-*` pass reporting zero defects and zero
   smells.

6. **Read the branch diff before merging.** Stop for anything unrelated to
   F-XXX, any edit to another active F-ID's files, and any worker edit to
   `AS_BUILT.md` or `SPRINT_TRACKER.md`. Those totals belong to the integrator.

## Integration

1. `git merge --squash work/<fid-lower>-<agent>`. Do not commit yet.

2. **Resolve no semantic conflict automatically.** A conflict between two
   sprint features is reconciled against **both approved design plans**, then
   re-reviewed with `/microscope F-XXX --working`, and the reconciliation is
   recorded. Any other conflict stops for a human decision.

3. **Outside `--batch`, run `/verify`.** The full gate, or the scoped gate the
   design plan documented. The worker's own result is evidence, not a
   substitute. Two independently correct features can be jointly wrong, and this
   is the first moment anything can see that.

4. **Outside `--batch`, follow `/complete-feature F-XXX` from step 3 onward.**
   Update the HLD files the plan listed, append `AS_BUILT.md` and
   `SPRINT_TRACKER.md`, set `done` in `BACKLOG.md` and `CURRENT_SPRINT.md`,
   clear the `Owner` cell, and regenerate the AUTOGEN counts.

5. **Consume the handoff.** Its durable facts, the harness result and the test
   gate, move into the AS_BUILT entry. Then delete
   `.claude/handoffs/F-XXX-ready.md` from the staged integration. A handoff left
   behind blocks `close-preflight`, which is deliberate.

6. **Run the generated-skill drift check** if anything under
   `.claude/commands/` or `.claude/skills/` changed:

   ```bash
   python3 scripts/sync_agent_skills.py --check
   ```

7. **Commit once** using the `/complete-feature` message format.

8. **Outside `--batch`, push `sprint/sNN`** and nothing else.

9. Record the result:

   ```bash
   python3 scripts/sprint_workflow.py mark-feature SNN F-XXX reviewed \
     --head <worker-head-sha> --handoff consumed \
     --integration-commit <new-sha>
   ```

## What `--batch` changes

1. Commit only feature-local work: code, tests, the design plan, the review
   files, the HLD files the plan listed, and the consumed handoff.
2. Leave F-XXX `in-progress` with its `Owner` intact. Do not append `AS_BUILT.md`
   or `SPRINT_TRACKER.md`.
3. **Do not claim consolidated verification** in the commit message or anywhere
   else. It has not run yet. Write `Harness, deferred to sprint gate` when the
   worker's harness result has not been re-established on the integrated tree.
4. Do not push. `/run-sprint` pushes once, after its gate and its review loop.

## Refused situations

- **Integrating without a validating handoff.** No exceptions.
- **Auto-resolving a semantic conflict.** Two plausible merges of a parser and
  a renderer produce a file that compiles and renders the wrong thing.
- **Deleting the worker branch or worktree.** Report both as retained cleanup
  targets. `/close-sprint` removes them after the integrated sprint passes its
  gates and is pushed.
- **Pushing in `--batch`.**
- **Merging to `main` or tagging.** `/close-sprint` owns `main` and sprint
  tags. `/release` owns release tags.
