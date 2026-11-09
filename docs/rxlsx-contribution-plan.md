# rxlsx contribution and parallel sprint guide

**Date**: 2026-10-11
**Status**: Maintainer-approved direction and implementation guide.
**Discussion**: [Issue 281](https://github.com/tensorbee/rdocx/issues/281).

This document records the agreed direction and the changes needed before a
contributor starts the Excel stream. It is not a second backlog, sprint tracker
or F-ID design plan. The canonical workflow and delivery records remain under
`.claude/` and `docs/sprints/`. The workflow changes below are proposed work,
not commands or capabilities that already exist.

## 1. Decision to proceed

The maintainer approves proceeding with **rxlsx**. The intended name is rxlsx,
not rsxlsx or rsxls. The first deliverable is a preservation-safe workbook
reader and editor. The longer-term goals are local pivot refresh, a bounded
Power Query implementation and Office Scripts-compatible automation.

The contributor's Google Sheets review workflow gives the first cut a concrete
purpose: read and change cell text, edit a shared-formula column and retain
reviewer comments, including metadata that ordinary spreadsheet tools may
discard. A consistent CLI and Python interface across the Office families is
also part of the intended value.

This approves the programme direction. It does not claim that no competing
library supports the required features. Complete F-184's ecosystem assessment,
decision record and preserve/model/execute capability matrix before its
dependent implementation stories start. Compare maintained alternatives,
including reuse behind a consistent facade. Record why this implementation is
worth maintaining. Amend the existing non-goal to permit a separate rxlsx
family while keeping oxml-sml limited to chart workbooks.

Do not mark F-184 or S103 done merely because this document exists.

## 2. First-cut design and reusable groundwork

### Existing components

| Component | Reusable capability | Work still needed |
|---|---|---|
| `oxml-sml` | Deterministic one-sheet XLSX generation, strings, numbers, number formats and chart ranges | General workbook reading and editing remain separate |
| `oxml-opc` | Package parts, relationships, content types and bounded reads | Spreadsheet ownership and genuine streaming memory guarantees |
| `oxml-core` | XML helpers and raw subtree preservation | Preservation rules for each modeled workbook part |
| `oxml-drawing`, `oxml-media`, `oxml-chart` | Themes, drawings, images, ChartML and chart rendering | Worksheet anchors and workbook integration |
| `oxml-layout`, `oxml-pdf` | Fonts, shaping, positioned pages, PDF and raster output | Spreadsheet grid layout, print settings, scaling and pagination |

The small oxml-sml workbook has recorded Excel and LibreOffice acceptance.
That evidence applies to chart workbooks, not arbitrary Excel compatibility.

### Architecture and scope

Use a separate `rxlsx-oxml` family for typed workbook parts and `rxlsx` for the
owned facade, on `oxml-opc`. Keep one workbook model across Rust, CLI, Python
and later automation. Agree the Rust read/edit/save contract before exposing
it through bindings. This document approves those two crate boundaries in
principle. Their actual introduction still needs a story design and checks.

The first cut reads sheets, typed cell values, formulas and stored results,
shared strings, formats and styles, and merged ranges. It edits cell values and
formulas, applies styles, adds sheets, changes widths and heights, sets freeze
panes and auto-filters, and saves atomically to a new destination.

Chart authoring and row/column insertion or deletion are deferred. Structural
edits affect references, names, merges, comments, tables, validation and
conditional formatting, so they need their own contract and tests.

### Preservation

Retain unsupported parts and their relationships, including comments,
threaded comments, `commentsmeta*`, validation, conditional formatting, names,
extensions and advanced analytical content.

An untouched worksheet keeps its exact XML bytes. The ZIP archive itself need
not be identical because packaging metadata or compression may change. Edited
XML must retain unsupported attributes and children, their required namespace
bindings and schema child order. Existing style and shared-string indices stay
stable. Appending a deduplicated style must retain every component it references.

### Formula edits and stale results

Formula text and its stored result are separate values. Without an evaluator,
rxlsx must not report a stored result as newly calculated.

Changing an input also invalidates results. For example, if B1 contains
`A1*2`, changing A1 makes B1's stored result stale even though its formula text
did not change. Until dependency analysis exists, conservatively report formula
caches as potentially stale after value or formula edits. Drop the edited
formula's cache, request recalculation on load and remove obsolete calculation
chain parts, relationships and content-type declarations consistently. Do not
apply these changes to a no-op open/save.

Shared formulas require reference translation. If the master in B2 contains
`A2*2`, moving it to B3 requires `A3*2`. Absolute references such as `$A$2`
remain fixed, while mixed references follow their respective row/column rules.
Changing one member must leave the other members' effective formulas unchanged.
Handle single-member groups and non-contiguous remaining members explicitly.
Unsupported rewrites must fail without partial mutation.

### Atomic saves and memory

Save to a new destination and leave the source untouched. Failure must not
publish a partial workbook. Define destination replacement behavior in the
writer design rather than assuming that atomicity permits overwriting input.

The current OPC package owner stores parts in memory. Start with explicit size
limits. Pull-parsing worksheet XML alone does not satisfy F-187/F-188's streaming
contracts. Their large-workbook gates require measured memory evidence.

## 3. Parallel streams and responsibilities

The canonical roadmap retains S91.1 publication, S92 and S93 repairs,
S94 Python/CLI parity, S95 and S95.1 new repairs and native interoperability,
S95.2 through S95.4 Word features, S97 fillable PDF forms,
S99 through S101 Word features and the S102 external Word benchmark. The
conditional Excel programme starts at S103. S96 is the current contribution repair wave, and S98 reserves deferred native/bibliography recovery. An independent contributor stream
remains a workflow proposal until the extensions below are implemented.
Both use separate branches and checkouts. Sprint numbers do not dictate merge
order, but story dependencies remain mandatory.

Keep one canonical set of record paths. On `sprint/s91`, CURRENT_SPRINT describes
S91. On `sprint/s103`, the same path describes S103. Neither stream creates an
alternate tracker. Each branch is internally consistent, although the other
stream's unpublished progress may not yet appear there.

The contributor owns Excel designs, implementation, tests and feature reviews.
No separate maintainer approval is required for design decisions within agreed
scope. The contributor may add, split or reorder Excel stories, updating the
canonical HLD story definitions, BACKLOG and SPRINT_PLAN together.

Reserve an unused F-ID range before contributor allocation. Split existing
stories using child IDs, such as F-185a and F-185b. Every story needs a size,
dependencies, acceptance criteria and one named test gate. A parent completes
only when its children satisfy the full parent contract. Completed stories are
immutable, and new requirements become follow-up stories.

Coordinate changes to shared infrastructure, other Office families, programme
scope, release policy and rendering baselines. When the maintainer joins Excel
implementation, claim different F-IDs in separate worktrees. Record the person
separately from the coding agent, since both participants may use Claude.

The maintainer owns final integration, AS_BUILT entries, SPRINT_TRACKER entries,
integrated done statuses, sprint close and publication. Planned story edits by
the contributor are permitted. Delivery accounting remains with the integrator.

## 4. Workflow implementation checklist

Make these changes as a separate reviewed workflow feature before Excel
implementation. Assign its F-ID in the canonical backlog through the normal
lifecycle. Do not execute it directly from this guide.

- [ ] Update WORKFLOW and AGENTS with stream ownership, contributor permissions,
      branch-context rules and integration responsibilities.
- [ ] Extend `/sync-sprint` to open an independent stream in a separate worktree
      without closing the other stream. Preserve existing completion statuses.
- [ ] Extend `/claim-feature` with contributor identity and remote-checkout
      support. Claim only a designed story whose dependencies are complete.
- [ ] Make `/start-feature` and `/implement-feature` consistently accept assigned
      worker branches. The latter currently assumes a sprint branch.
- [ ] Extend `/complete-feature --prepare` with an explicit contributor push and
      PR path. Keep all verification and review requirements.
- [ ] Extend `/integrate-feature` to fetch and validate an exact remote PR head
      and its handoff without requiring a contributor's filesystem path to
      exist on the maintainer's machine.
- [ ] Update `/run-sprint`, `/sync-status` and `/close-sprint` for independent
      streams and out-of-order closes.
- [ ] Update `scripts/sprint_workflow.py` and its existing test entrypoint to
      enforce the new contracts. Do not rely on prose alone.
- [ ] Update affected canonical skills, then regenerate `.agents/skills/`
      with `scripts/sync_agent_skills.py` and pass its drift check.

`/design` and `/microscope` retain their design and review roles. Do not remove
dependency checks or weaken reviews to accommodate remote work. Update their
instructions only where the contributor context requires clarification.

### Record reconciliation and sprint close

Merge ledger records by F-ID. Preserve completion entries from both streams,
detect conflicting story changes, avoid duplicate entries and recompute
generated counts. Never resolve a ledger conflict by accepting an entire file
from one branch. Never synthesize missing verification or review evidence.

When incorporating main into an active sprint, retain that branch's active
CURRENT_SPRINT context. After a sprint closes, main retains the closing sprint's
record until another sprint is explicitly opened there. Other active worktrees
keep their own context.

Only `/close-sprint` merges into main. Serialize sprint closes. The next closing
branch incorporates the latest main, reconciles shared changes and re-establishes
its integrated verification and sprint review. S103 may close before S91 if its
dependencies and acceptance gates pass. Shared baseline updates are exclusive.

## 5. Proposed Excel story and sprint realignment

This table is a proposal to apply to the canonical roadmap, not another live
delivery record. Recalculate story sizes, dependencies and sprint capacity
before assigning implementation. Boundaries are capabilities, not calendar
promises. Add sprints if the reviewed sizing does not fit, rather than weakening
gates or silently overloading a sprint.

The current canonical schedule is below. The earlier reader/editor acceleration
proposal still requires child-story design and a capacity decision. This guide
allocates no additional IDs and does not supersede SPRINT_PLAN.md.

| Sprint | Canonical planned delivery |
|---|---|
| S103 | F-184 go/no-go, F-204 corpus, F-185 workbook and worksheet core |
| S104 | F-186 strings/styles/formats, F-189 formula parser, F-205 tables and structured references |
| S105 | F-206 advanced worksheet objects |
| S106 | F-187 reader and F-188 writer, including their full streaming contracts |
| S107 | F-190 calculation, F-191 charts, F-192 validation and conditional formatting |
| S108 | F-193 pivots, F-207 recalculation, F-208 slicers and Data Model boundary |
| S109 | F-209 Power Query package and bounded M language |
| S110 | F-210 Power Query execution and allowed connectors |
| S111 | F-211 Office Scripts surface and F-212 sandboxed execution |
| S112 | F-194 rendering and F-195 distribution and advanced lifecycle gate |

Split F-187/F-188 into initial read/write capability and genuine streaming work.
Split F-190, F-194 and F-195 into early bounded capability and later completeness.
Split F-204 into first-cut and advanced corpus coverage. Give structural edits a
separate story. Allocate child and new IDs in canonical records, not here.

Revisit every downstream dependency after splitting. An early capability may
depend on a delivered child rather than the entire parent, but that must be an
explicit contract change. Do not count the minimal editor as completing the
original full streaming or distribution contracts.

Consider an earlier reader/editor release only after its complete API,
bindings and acceptance checks pass. Add a release preparation story and use the
existing separate publication approval. Later rendering and advanced lifecycle
capabilities have their own reviewed boundaries. Neither sprint close nor this
guide authorizes publication.

Keep Power Pivot/OLAP, proprietary connectors, VBA/XLM, custom functions, Python
cells and Microsoft-hosted script services as preservation and diagnostic
boundaries unless separately approved. Reuse the chart engine when chart work
arrives, and retain deterministic font mode for rendering evidence.

## 6. Contributor and maintainer runbook

The following sequence applies after the workflow changes are implemented.

1. Maintainer completes the workflow feature, records F-184, realigns canonical
   stories and opens the Excel sprint branch from an agreed reviewed main SHA.
2. Contributor designs a ready story through `/design F-XXX`. Resolve questions
   within agreed scope independently. Coordinate cross-stream changes.
3. Maintainer claims the designed story and publishes its assignment, branch
   and immutable base. Contributor fetches it into their own checkout.
4. Contributor runs `/start-feature F-XXX --claimed`, then
   `/implement-feature F-XXX`.
5. Contributor fixes findings and repeats applicable checks and microscope
   passes until there are zero defects and zero smells. Review passes never
   modify source code themselves.
6. Contributor runs `/complete-feature F-XXX --prepare`, pushes through the
   revised contributor path and opens a PR against the assigned sprint branch.
7. Maintainer validates the exact PR head and handoff, integrates the feature,
   verifies the combined result and records completion through the canonical
   integration workflow.
8. Contributor starts dependent stories from integrated prerequisites, not from
   an unverified PR's apparent completion.
9. Maintainer closes each ready sprint through full verification and sprint
   review, reconciles PRs and credits the contributor. The contributor starts
   the next Excel sprint from the updated main.

`--claimed` skips tracker updates already made by the claim and creates the
feature's failing stubs. `--prepare` packages a verified contribution and
handoff without writing shared completion logs. `/implement-feature` invokes
verification and microscope at the end, but does not invoke start or complete.

A prepared PR is not delivered work. Mark each story done after verified
integration, and mark the sprint complete only after its closing gates. Keep
Issue 281 open while its programme discussion remains outstanding.

## 7. Validation and next actions

Use meaningful workflow tests for simultaneous S91/S103 contexts, remote claims
and handoffs, duplicate IDs, contributor story additions, completed-status
preservation, ledger conflicts, S103 closing before S91, and evidence becoming
stale after source changes. Confirm one stream cannot mutate or delete the
other stream's active worker state.

Workbook acceptance covers Google Sheets comments and metadata, shared-formula
master/member edits, stale dependent caches, stable styles and strings,
unchanged part payloads, unknown XML retention and atomic failure. Construct
ordinary fixtures in code. Any fetched corpus uses pinned checksums and clear
licences. Record Excel and LibreOffice identities and expected results
separately. Do not claim unperformed viewer acceptance.

Execute in this order:

1. Implement and review the workflow changes.
2. Complete F-184's decision evidence and canonical scope updates.
3. Realign and size the Excel stories and release boundaries.
4. Open the parallel Excel stream and reserve contributor IDs.
5. Send the contributor the agreed stories and usable workflow instructions.
6. Integrate contributions and close sprints with the normal gates.

Saving this document performs none of those lifecycle transitions. It makes no
GitHub comment, commit, merge, tag or publication.
