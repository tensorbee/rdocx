---
description: Bring README.md and docs/hld/ back into line with what was actually built. Grounded in plans, reviews and AS_BUILT. No invention, no changelogs.
---

# /realign-docs [--section NN] [--dry-run]

Repair accumulated documentation drift in bulk. `/complete-feature` updates the
HLD files one story's plan listed, which is right per story and insufficient
over a sprint. Decisions taken in review, alternatives rejected during
implementation, and scope that shifted mid-story do not appear in any single
plan's `## HLD impact` list.

Run at a sprint close, at a milestone boundary, or after `/sprint-review`
surfaced a `docs` finding.

**`--dry-run` reports the intended edits and changes nothing.** Prefer it first.

## The two hard rules

**1. Grounded, never invented.** Every sentence you write or change must trace
to a citable source: a design plan, a review file, an AS_BUILT entry, or the
code itself. If you cannot cite it, you do not know it, and a confident
paragraph about behaviour nobody implemented is worse than the stale one it
replaced.

**2. Current state, never history.** These documents describe what is true now.
Do not write "F-042 changed this from X to Y". Describe Y. Do not add a
"Changes in this milestone" heading. `git log` and `docs/sprints/AS_BUILT.md`
hold history, and a changelog inside an authoritative document is the first
sign it is rotting into one.

## Steps

1. **Establish the window.** Everything since the last realignment, or since a
   named tag. Report the commit range and the F-IDs it contains.

2. **Read the ground truth, in this order:**
   - `.claude/plans/F-XXX-design.md` for each F-ID in the window, especially
     `## Approach`, `## Rejected alternatives` and `## HLD impact`.
   - `.claude/reviews/` for the same F-IDs, plus any `SNN-sprint-review-*`.
     Review files hold decisions that exist nowhere else, because a finding
     that was refuted with evidence is a decision.
   - `docs/sprints/AS_BUILT.md` entries in the window, especially the
     "Notes for future sessions" and the hash-harness fields.
   - The code, last. It settles any disagreement between the above.

3. **Run `/audit-spec`** and fold its findings in. This command is the actor
   for the drift that command reports.

4. **Build the edit list before editing.** For each target file: what is stale,
   what replaces it, and the citation. Report the list. Under `--dry-run`, stop
   here.

5. **Apply, respecting precedence.** `docs/hld/README.md` sets it: the lower
   number wins on scope and intent, the higher number wins on mechanism. If a
   story changed which crate owns something, `03-architecture.md` changes **and**
   the mechanism document changes. Updating one of them creates the next
   contradiction.

6. **Realign `README.md` separately.** It is the front door and its audience is
   someone who has never read the HLD. Check the crate list, the feature list
   against `docs/hld/02-scope-and-non-goals.md`, and every version string.

7. **Do not touch `14-development-backlog.md` story status.** That is
   `/complete-feature` and `/sync-status` territory. Fixing a story's *shape*,
   its size or its test gate, is in scope. Fixing its *status* is not.

8. **Verify and commit.**

   ```bash
   python3 scripts/prose_check.py
   ```

   Then one commit, separate from any code change:

   ```
   docs, realign the spec set to as-built through <range>

   One paragraph: what drifted, and the sources it was realigned against.

   Tests, not applicable
   Harness, unchanged
   ```

## What this command does not do

- **It does not change code.** If realignment reveals that the code is wrong
  rather than the document, that is a finding and a backlog story, not an edit.
- **It does not resolve open questions.** An undecided thing stays in
  `13-risks-and-open-questions.md`. Writing a decision into a mechanism
  document because the prose reads better is how an unmade decision becomes
  invisible.
- **It does not rewrite AS_BUILT.** That log is append-only.
- **It does not soften a "deliberately wrong" entry.** `apply_tint_shade` and
  the truncating unit constructors are described as they behave, with the
  reason they are held.
- **It does not add aspirations.** Future intent belongs in a backlog story.

## Refused situations

- **Writing a claim you cannot cite.** Delete the sentence instead.
- **Adding a changelog or "what changed" section to any `docs/hld/` file.**
- **Editing a document to match a design plan the code contradicts.** The code
  wins on what is true. Then decide whether the code is what you wanted.
- **Folding a documentation realignment into a feature commit.**
