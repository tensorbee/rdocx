# Development workflow

This document is the canonical reference for how features land in this
codebase. It is opinionated and load-bearing. **When in doubt, this file wins
over any other document**, including `CLAUDE.md` and everything in `docs/hld/`.
Those describe the product. This describes the process.

## Atomic feature rhythm

Every F-ID flows through the same five steps:

```
1. /design F-XXX               Write the design plan, get aligned
2. /start-feature F-XXX        Mark in-progress, create test stubs
3. ...implement...             Code plus tests
4. /verify                     The gate must pass
5. /complete-feature F-XXX     Update tracking, commit
```

Required before completion:

```
4a. /microscope F-XXX --working   Review the implementation diff
```

`/microscope` is not optional and not skippable. Its exit condition is zero
defects and zero smells, iterating in numbered passes until it reaches that.

## Parallel work

One F-ID at a time in one worktree is the normal case. When several independent
stories are ready at once, they run in parallel worktrees instead, and the
rhythm gains three steps:

```
/claim-feature F-XXX claude ../rdocx-fxxx   Branch and worktree, one commit
...the normal rhythm, inside that worktree...
/complete-feature F-XXX --prepare           Handoff, no sprint ledgers
/integrate-feature F-XXX work/f-xxx-claude  Back onto the sprint branch
```

`/run-sprint` drives all of this for a whole sprint, including the design phase
that has to come first. The rules that make it safe:

- **Only one agent may write a given worktree.** Two agents never implement the
  same F-ID.
- **Workers do not write the sprint's totals.** `AS_BUILT.md` and
  `SPRINT_TRACKER.md` belong to the integrator. Two parallel appends to a single
  ledger is a conflict in the one file that must not have one.
- **The hash-harness baseline is exclusive.** There is one baseline. Two
  stories re-recording it in parallel produces a delta nobody can attribute.
- **Verification that matters happens once, over the integrated result.** A
  worker's `/verify` is evidence, not a substitute. Two features that are
  individually correct and jointly wrong are invisible until they are merged.
- **A worker's claims are validated, not trusted.** `/complete-feature
  --prepare` writes `.claude/handoffs/F-XXX-ready.md`, and `/integrate-feature`
  refuses a branch whose handoff does not validate.
- **Never resolve a semantic conflict automatically.** Reconcile against both
  approved design plans, then re-review.

Worker branches and worktrees are **retained** through integration, full
verification and sprint review. After `/close-sprint` has merged and pushed
both `main` and the sprint tag, it removes the clean worktrees and local
branches for completed workers recorded in the sprint state. It never removes
an uncommitted worktree, a carried worker or an unrelated worktree.

## The command surface

Every command is defined in `.claude/commands/<name>.md`, and that file is the
contract. This table exists so you can find the right one, not so you can skip
reading it.

| Command | Does | Writes code |
|---|---|---|
| `/design` | Write the design plan. `--draft` defers the questions | no |
| `/start-feature` | Mark in-progress, create test stubs | stubs only |
| `/implement-feature` | Build against the plan and the stubs | yes |
| `/microscope` | Adversarially review one F-ID's diff | no |
| `/verify` | The gate | no |
| `/complete-feature` | Trackers, HLD, commit. `--prepare` for a worker | no |
| `/claim-feature` | Branch and worktree for parallel work | no |
| `/integrate-feature` | Bring a worker branch onto the sprint branch | no |
| `/run-sprint` | The whole sprint, design through review loop | yes |
| `/sync-sprint` | Open a sprint, create its branch | no |
| `/sprint-review` | Review the integrated sprint delta | no |
| `/release` | **The only command that creates `v*` release tags or starts publication** | no |
| `/close-sprint` | **The only command that merges to `main` or creates `sNN` tags** | no |
| `/sync-status` | Audit the trackers against each other | `--fix` only |
| `/audit-spec` | Audit the spec set against the code | no |
| `/realign-docs` | Repair accumulated documentation drift | no |
| `/spec-bump` | Version and tag the spec set | no |
| `/impact` | Trace what a proposed change would touch | no |
| `/study-crate` | Build understanding of one crate | no |
| `/differential` | Run the corpus against python-docx and python-pptx | no |
| `/regen-fixtures` | Rebuild generated artefacts | no |

The reference material commands cite lives in `.claude/skills/`:
`voice-rules`, `hld-discipline`, `risk-routing`, `differential-testing`.

## Two agents, one workflow

Claude and Codex follow the same rules. `.claude/commands/` and
`.claude/skills/` are canonical for both, and `AGENTS.md` is the entry point
for any agent that does not read `CLAUDE.md`.

Codex discovers repository skills under `.agents/skills/<name>/SKILL.md`. Those
files are **generated adapters**, each pointing back at its canonical source
with that source's SHA-256. They are never hand-edited, and a copy of the
workflow is never checked in, because a drifted copy is worse than no Codex
support: the two agents would follow different rules while believing they
agreed.

```bash
python3 scripts/sync_agent_skills.py           # regenerate after any edit
python3 scripts/sync_agent_skills.py --check   # what /verify step 6 runs
```

Adding a command means adding `.claude/commands/<name>.md` with a `description`
in its frontmatter, then regenerating. The description is required, since it is
what both hosts show in their command list.

An agent claims work as `claude` or `codex`, which is what the `Owner` column
in `CURRENT_SPRINT.md` records and what the `work/<fid-lower>-<agent>` branch
name carries.

## Sprint cadence

- Sprints are about 2 weeks of focused work.
- The sprint clock starts at the first `/start-feature` of that sprint, not at a
  fixed calendar date.
- Each sprint carries 3 to 7 F-IDs from the active milestone.
- Long milestones span several sprints. M7 has 4, M10 has 4, M1 has 2.
- 13 milestones and 36 sprints to v1 per `docs/hld/14-development-backlog.md`.

## Where things live

| Artifact | Location | Updated by | Lifetime |
|----------|----------|------------|----------|
| Spec set | `docs/hld/00-vision.md` through `15-build-and-toolchain.md` | Rare, when scope changes | Permanent |
| Story definitions | `docs/hld/14-development-backlog.md` | Hand-curated | Permanent |
| Backlog status | `docs/sprints/BACKLOG.md` | `/complete-feature`, `/sync-status` | Live |
| Sprint roadmap | `docs/sprints/SPRINT_PLAN.md` | Hand-curated, rarely changes | Live |
| Active sprint | `docs/sprints/CURRENT_SPRINT.md` | `/sync-sprint SNN` | Replaced per sprint |
| Velocity log | `docs/sprints/SPRINT_TRACKER.md` | `/complete-feature`, `/close-sprint` | Append and update |
| Completion log | `docs/sprints/AS_BUILT.md` | `/complete-feature` | Append-only |
| Design plans | `.claude/plans/F-XXX-design.md` | `/design` | Tracked in git |
| Review findings | `.claude/reviews/F-XXX-<aspect>-pass-N.md` | `/microscope` | Tracked in git |
| Sprint reviews | `.claude/reviews/SNN-sprint-review-pass-N.md` | `/sprint-review` | Tracked in git |
| Worker handoffs | `.claude/handoffs/F-XXX-ready.md` | `/complete-feature --prepare`, consumed by `/integrate-feature` | Tracked, then deleted on integration |
| In-flight notes | `.claude/scratch/F-XXX-progress.md` | Hand-edited between sessions | Gitignored |
| Sprint run state | `.claude/scratch/SNN-run.json` | `/run-sprint`, `scripts/sprint_workflow.py` | Gitignored |
| Command definitions | `.claude/commands/*.md` | Hand-curated | Permanent |
| Reference skills | `.claude/skills/*.md` | Hand-curated | Permanent |
| Codex adapters | `.agents/skills/<name>/` | **Generated.** `scripts/sync_agent_skills.py` | Tracked, never hand-edited |
| Agent guidance | `CLAUDE.md`, `AGENTS.md` | Hand-curated | Permanent |
| Settings | `.claude/settings.json` (project), `.claude/settings.local.json` (gitignored) | Rare | Live |

**There is exactly one of each.** Never create a second backlog, tracker, design
plan or review record.

## The design plan is a machine-consumed contract

`/design` writes `.claude/plans/F-XXX-design.md` with these sections, and later
commands execute against them:

| Section | Consumed by |
|---|---|
| `## Spec reference` | `/start-feature`, which **refuses a whole-document citation.** Name sections |
| `## Test plan` | `/start-feature` step 3, which creates the stubs |
| `## Implementation checklist` | The implementing session |
| `## HLD impact` | `/complete-feature`, which updates **exactly those files** |
| `## Open questions` | `/design`, which asks them before writing code |

This is what stops the specification rotting. A story that changes behaviour the
HLD describes must name the section, and completion updates it.

## Voice rules

Enforced by `/verify` over tracked Markdown under `docs/`, `.claude/plans/`,
`.claude/reviews/`, `.claude/commands/`, `.claude/skills/`, over `CLAUDE.md`,
`AGENTS.md` and this file, and over commit messages.

- **No em-dash.** Use a comma, a hyphen, or rewrite the sentence.
- **No en-dash.** Write `M1 to M6`, not a dashed range.
- **No semicolon in prose.** Use a full stop or a comma.

Prose only. Code, identifiers and code comments are exempt. The full rule,
including what the scanner already exempts, is in
`.claude/skills/voice-rules.md`.

The reason is consistency of voice across a document set that several sessions
write into over many months.

## Test taxonomy

Six categories, defined in `docs/hld/12-testing-strategy.md`. Every design plan
picks the applicable ones and names exactly one as the story's test gate.

`unit`, `integration`, `regression`, `round-trip`, `golden`, `differential`.

**No binary fixture files.** Fixtures are constructed in code, including image
headers with precomputed CRCs. The deck corpus is the one exception and lives
outside the published crates.

Regression tests are **named as sentences describing the failure they prevent**,
so a reintroduction is obvious from the test name rather than from a diff.

## The hash harness

From M1 until M6 closes, **every PR gates on the output-stability harness.**

- An unexplained delta blocks the merge.
- An intentional behavioural change lands as its own labelled commit, with the
  expected delta stated in the message and reviewed.
- Never fold a behaviour change into a file move.

This exists because the extraction changes unit conversion and text-shaping
inputs, which alter output without failing to compile. The existing 320 tests
cannot see that class of defect.

## Git workflow

Per-sprint branches off `main`, named `sprint/sNN`. Every F-ID commit lands on
the active sprint branch, never directly on `main`.

Parallel work adds `work/<fid-lower>-<agent>` branches, cut from the sprint
branch head at claim time and squashed back by `/integrate-feature`. They are
retained until `/close-sprint` has pushed the integrated sprint, then removed
locally. They are never pushed unless asked.

- `/sync-sprint SNN` creates the branch off the latest `main`.
- `/claim-feature` cuts a worker branch from the sprint branch head.
- `/complete-feature` commits to the sprint branch.
- `/integrate-feature` squashes a worker branch onto the sprint branch.
- `/release vX.Y.Z` tags an already reviewed sprint SHA and starts publication.
- `/close-sprint SNN --next SMM` validates readiness, merges to `main` with an
  explicit merge commit, creates the annotated `sNN` tag, pushes both, removes
  completed worker worktrees and local branches, then runs `/sync-sprint` for
  the next sprint.

Only `/close-sprint` may touch `main` or create an `sNN` sprint tag. Only
`/release` may create or push a `v*` release tag or start crates.io
publication. `/spec-bump` may create a local `spec-v*` tag but never pushes it.
No command crosses those namespaces implicitly.

Commit message format, set by `/complete-feature`:

```
F-XXX, short title

One paragraph: what was built, why, and any non-obvious choices.

Tests, summary line
Harness, unchanged | expected delta and its justification
```

**No `Co-Authored-By` trailer, ever.**

Push to the sprint branch only when `/verify` passes locally.

## Sub-IDs when a story splits

If an F-ID grows into several natural chunks, letter-suffix the children:
`F-064` becomes `F-064a`, `F-064b`, `F-064c`. Each child gets its own design
plan and AS_BUILT entry. The parent closes only when every child closes. Update
both `docs/hld/14-development-backlog.md` and `docs/sprints/BACKLOG.md`.

Two stories are already sized `XL` and are expected to split: F-064, the
DrawingML text model, and F-098, shape text layout.

## Resuming mid-F-ID

1. Read `.claude/plans/F-XXX-design.md`, which is the contract.
2. Read `.claude/scratch/F-XXX-progress.md`, which is the in-flight memory.
3. Re-run `/verify --fast` to establish the current state.
4. Continue.

Never resume from the diff alone.

## Escalation triggers

| Signal | Response |
|--------|----------|
| An F-ID consistently exceeds 2x its estimate | Split into `F-XXXa`/`b`. Update both backlogs |
| `/verify` exceeds 10 minutes on changed files | Investigate the slow check, add a `--fast` path |
| The hash harness shows an unexplained delta | Stop. Do not proceed until it is explained |
| Corpus round-trip failures appear with no recent parser change | A corpus deck changed, or a dependency did. Find the root cause |
| A sprint exceeds 4 calendar weeks with no completion | Replan. Too many F-IDs in flight, or the scope is wrong |
| The same F-ID restarted 3 times from scratch | The design plan was wrong. Reset to `/design` with explicit questions |
| Velocity diverges from the plan by more than 30 percent over 3 sprints | Replan the remaining milestones rather than absorbing it |

## When this file changes

A process change is an ADR-sized event.

1. Note the trigger in the next AS_BUILT entry's "Notes for future sessions".
2. Update this file.
3. Update the affected `.claude/commands/*.md` and `.claude/skills/*.md`.
4. Run `python3 scripts/sync_agent_skills.py` and commit the regenerated
   `.agents/skills/` alongside. `/verify` fails otherwise.
5. Commit the workflow change **separately** from any feature change, so the
   rationale is legible in `git log`.
