---
description: Claim an F-ID for parallel work. Creates its isolated branch and worktree, and records the claim in one commit.
---

# /claim-feature F-XXX {claude|codex} <worktree-path>

Assign one story to one worker so two sessions cannot write the same tree. Run
this **only from the canonical sprint worktree**, never from a worker's.

`/run-sprint` performs claims itself and does not ask separately. Invoke this by
hand when you are driving a parallel sprint manually.

## Preconditions

Check all of these before touching anything. If one fails, name it and stop
without altering sprint state, branches or worktrees.

1. **The branch is `sprint/sNN`** matching the heading in
   `docs/sprints/CURRENT_SPRINT.md`.

2. **The tree is clean, with one exception.** The approved design plan for
   F-XXX may be present and uncommitted, because a manual `/design` leaves it
   there. Refuse every other staged, unstaged or untracked path. Without this
   exception the claim would be impossible, since it needs both an approved plan
   and a clean tree.

3. **F-XXX is `pending`** in both `docs/sprints/CURRENT_SPRINT.md` and
   `docs/sprints/BACKLOG.md`.

4. **`.claude/plans/F-XXX-design.md` exists and is `approved`**, with non-empty
   `## Spec reference`, `## Test plan`, `## Risk routing` and `## HLD impact`
   sections. A whole-document spec citation is a failure here, the same as it is
   in `/start-feature`.

5. **The owner is exactly `claude` or `codex`.**

6. **The worktree path is given explicitly and does not exist.** Do not infer a
   path and do not delete one. Put it outside the repository, as a sibling
   directory.

7. **`work/<fid-lower>-<agent>` does not exist**, locally or on the remote.

## The claim

1. Set F-XXX to `in-progress` in `CURRENT_SPRINT.md` and `BACKLOG.md`.
2. Write the agent name into the `Owner` cell of the current-sprint wave row.
   Unclaimed rows keep `-`.
3. Regenerate the AUTOGEN counts in `BACKLOG.md`.
4. Stage the claim plus the approved design plan, nothing else, and commit to
   the sprint branch:

   ```
   F-XXX, claim for parallel work

   Assign F-XXX to <agent> in an isolated worktree.

   Tests, not applicable
   Harness, unchanged
   ```

5. Create `work/<fid-lower>-<agent>` from the resulting sprint-branch HEAD, so
   every worker starts from the same immutable base.
6. `git worktree add <worktree-path> work/<fid-lower>-<agent>`.
7. Record the claim:

   ```bash
   python3 scripts/sprint_workflow.py mark-feature SNN F-XXX approved \
     --owner <agent> --branch work/<fid-lower>-<agent> \
     --worktree <worktree-path> --base <sprint-head-sha>
   ```

8. Do not push the worker branch unless asked.

## Report

The branch, the worktree path, the claim commit, the design plan path, and the
exact command the worker runs first:

```text
/start-feature F-XXX --claimed
```

## Refused situations

- **Running from a worker worktree.** Claims are issued from the canonical one.
- **Unrelated dirt in the tree.** The design-plan exception is the only one.
- **An unapproved or draft design plan.** Run `/design F-XXX` to completion.
- **Reusing an existing branch or worktree.** A stale worktree from a previous
  attempt is not reused. Completed workers are removed by `/close-sprint` after
  the integrated sprint is pushed. Any other stale worker requires an explicit
  recovery decision.
- **Claiming an F-ID whose dependencies are not `done`.** `/design` already
  refused that. If the state changed since, stop.
