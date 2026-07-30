# Repository guidance for coding agents

This repository uses one development workflow across every agent and human
contributor. Before changing code or tracked documentation, read `CLAUDE.md` and
`.claude/WORKFLOW.md`. **The workflow file wins on process questions.**

## Where the workflow lives

`.claude/commands/<name>.md` defines every command, and `.claude/skills/<name>.md`
holds the reference material those commands cite. Both are canonical for every
agent, whatever the directory is called.

Codex discovers the same set under `.agents/skills/<name>/SKILL.md`. Those are
**generated adapters** that point back at the canonical file with its SHA-256.
Never hand-edit one, and never treat one as the source. Regenerate with:

```bash
python3 scripts/sync_agent_skills.py           # after editing any command or skill
python3 scripts/sync_agent_skills.py --check   # drift gate, part of /verify
```

A generated adapter that has drifted from its source means the two agents are
following different rules while believing they agree, which is why the check is
a gate and not a warning.

## Shared state

- Treat `docs/sprints/CURRENT_SPRINT.md`, `BACKLOG.md`, `SPRINT_PLAN.md`,
  `SPRINT_TRACKER.md` and `AS_BUILT.md` as the shared delivery record.
- Treat `.claude/plans/` and `.claude/reviews/` as agent-neutral tracked
  artifacts despite the directory name.
- Treat `.claude/scratch/F-XXX-progress.md` as the local in-flight record. Read
  it when resuming an in-progress F-ID, and update it before handing work over.
- Treat `.claude/handoffs/F-XXX-ready.md` as the **structured** handoff a worker
  writes for the integrator. `/complete-feature --prepare` writes it,
  `scripts/sprint_workflow.py validate-handoff` checks it, and
  `/integrate-feature` consumes and deletes it.
- **Never create a second sprint tracker, backlog, design plan or review
  record.** There is exactly one of each.

## Non-negotiable rules

- **The hash harness gates every PR in M1 through M6.** An unexplained output
  delta blocks the merge. Every intentional behavioural change is its own
  labelled commit with the expected delta stated and reviewed. Never fold a
  behaviour change into a file move.
- **Rendering baselines use deterministic font mode.** Never record a baseline
  against system fonts.
- **`oxml-*` crates must not depend on `rdocx-*` or `rpptx-*`**, with the single
  documented `oxml-drawing -> rdocx-oxml` exception for the `Theme` adapter.
- **Preserve unmodelled XML verbatim.** Parse only what you render.
- **Respect schema child order.** OOXML uses `xsd:sequence`, and a violation
  makes PowerPoint refuse the file rather than warn.
- Do not use em dashes or prose semicolons in tracked Markdown or commit
  messages.
- Do not run `cargo clean`. Iterate scoped.
- **Do not commit or push unless the invoked workflow explicitly includes that
  action.** Never add an agent co-author trailer.
- Only `/close-sprint` may merge to `main` or create an `sNN` sprint tag.
- Only `/release` may create or push a `v*` release tag or start crates.io
  publication. It requires a separate final approval at the reviewed SHA.

## Commands

```bash
cargo check -p <crate> --all-targets        # default while iterating
cargo test -p <crate> [--test <file>]
cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo fmt --all --check
cargo test -p oxml-layout --no-default-features
cargo check --target wasm32-unknown-unknown -p rdocx-wasm -p rpptx-wasm

python3 scripts/hash_harness.py --check
python3 scripts/prose_check.py
python3 scripts/sync_agent_skills.py --check
python3 scripts/sprint_workflow.py status
```

Rules:

- Default to `cargo check -p <crate>` while iterating, not `build` or `test`.
- The `--exclude` on the binding crates is required. `pyo3/extension-module`
  makes a test binary fail to link.
- Adding a file under a `tests/` directory adds another binary to link. Add a
  module to the existing integration entrypoint instead.

The gate is `/verify`, defined in `.claude/commands/verify.md`. The list above is
the inner loop, not a second source of truth.

## Structural rules

A developer should be able to answer "what does this do?" from one file, without
following indirection to find out which code actually runs.

**The test for any new construct**: does it reduce the number of cases a reader
must consider, or increase the number of places they must look?

- No new trait unless two implementers exist **today**. A test double counts
  only if it exists and is used.
- No new generic parameter unless instantiated two ways **today**.
- No `Box<dyn>` or `Arc<dyn>` where the concrete type is statically known.
- No wrapper that only forwards. No builder for a struct with fewer than four
  fields.
- No new feature flag without a named consumer.
- No new crate, module or file without asking first.
- Smallest diff that solves the problem. No speculative extensibility, no "for
  future use" parameters.

When you introduce a trait, generic or crate, state in one sentence which
existing second implementer justifies it. If the sentence names a hypothetical
future one, do not introduce it.

The converse failure is equally real here: a sprawling function that trips a
complexity lint needs *more* structure, not less.

## Feature workflow

Use the same lifecycle for every F-ID: design, start, implement, microscope,
verify, complete.

Before resuming or modifying an F-ID, confirm all of:

1. The current branch and `git status`.
2. The F-ID status in `CURRENT_SPRINT.md` and `BACKLOG.md`.
3. The design plan and its spec references, test plan and HLD impact.
4. The latest microscope review and progress notes, when present.

The design plan is a contract, not a suggestion. Its `## HLD impact` section is
a file list that `/complete-feature` executes against.

## Coordination

- Only one agent may write to a worktree at a time.
- Two agents must not implement the same F-ID concurrently.
- Alternating agents use the sprint worktree and leave an explicit progress
  checkpoint in `.claude/scratch/F-XXX-progress.md` before handing over.
- Parallel implementation uses different F-IDs and different worktrees, claimed
  through `/claim-feature F-XXX {claude|codex} <path>`. The agent name goes into
  the `Owner` column and into the `work/<fid-lower>-<agent>` branch name.
- A worker owns its feature and not the sprint's totals. Never append to
  `AS_BUILT.md` or `SPRINT_TRACKER.md` from a worker branch. The integrator does
  that once.
- The hash-harness baseline is exclusive. One story per sprint wave may move it.
- Verification that decides anything runs once over the integrated result, not
  per worker.
- Worker branches and worktrees remain available through sprint verification
  and review. `/close-sprint` removes clean completed workers after pushing
  `main` and the sprint tag.

## Review tasks

If a prompt begins with `REVIEW TASK`, or you were invoked through
`/microscope`, `/sprint-review` or `/audit-spec`, **you are auditing. You are
not changing code.** This overrides the normal expectation that you produce a
working diff.

1. **Do not modify any file outside `.claude/reviews/`.** No source edits, no
   fixes, no incidental cleanups, no formatting. If you find a bug, record it as
   a finding and move on. A review that patched the code cannot tell anyone what
   state the code was in.
2. Every finding must carry a `path:line` citation. If you cannot cite it,
   delete it rather than softening it.
3. Reporting zero findings in a category is a valid and expected result. Say so
   explicitly. Do not manufacture findings to fill a section.
4. Stop when the required output file is written. Do not begin remediating.

Remediation is the opposite mode: a normal F-ID, the normal lifecycle, a working
diff with tests. The two must not happen in the same session.
