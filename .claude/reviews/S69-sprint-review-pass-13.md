# S69 sprint review, pass 13

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`1c626f801c316573eb0893e9d57571d4ba60ce2a` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 90 files and 8,237 changed
lines, comprising 7,110 additions and 1,127 deletions. The 20 touched crates
are `oxml-chart`, `oxml-cli-support`, `oxml-core`, `oxml-drawing`,
`oxml-layout`, `oxml-media`, `oxml-opc`, `oxml-pdf`, `oxml-sml`, `rdocx`,
`rdocx-oxml`, `rdocx-py`, `rdocx-wasm`, `rpptx`, `rpptx-chart`, `rpptx-cli`,
`rpptx-layout`, `rpptx-oxml`, `rpptx-render`, and `rpptx-wasm`
**Pass authority**: pass 13 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## F-X079 published-state boundary

- The immutable annotated `rpptx-v0.10.0` tag peels to reviewed SHA
  `1e409c553b950eb8029e3e78e39ff775f18ba3ab`, matching the completion record
  and every current-state HLD owner (`docs/sprints/AS_BUILT.md:11844`,
  `docs/hld/03-architecture.md:841`, `docs/hld/10-bindings-spec.md:1146`,
  `docs/hld/14-development-backlog.md:3739`,
  `docs/hld/15-build-and-toolchain.md:330`). A read-only remote tag fetch
  independently peeled to the same SHA.
- Hosted publish run `33984024736` completed successfully at the tagged SHA.
  The release record links that run and the resulting GitHub release
  (`docs/sprints/AS_BUILT.md:11847`). A read-only GitHub query independently
  returned the same successful run, tag, and release URL.
- All 15 selected crates are present and unyanked at 0.10.0 under sole owner
  `mantissaman`. `rpptx-wasm@0.10.0` and `rdocx@0.13.0` remain absent. These
  independent crates.io observations agree with the recorded publication,
  owner, and exclusion evidence (`docs/sprints/AS_BUILT.md:11864`).
- The published GitHub release body is 2,119 bytes and hashes to
  `af97bb5020b8cdaa5f7982bea55e471f89483e6fc3fe692929bb29a55199c43f`,
  byte-identical to a fresh render from the reviewed changelog. The recorded
  digest agrees (`docs/sprints/AS_BUILT.md:11868`), and the release notes remain
  restricted to the selected family (`CHANGELOG.md:7`, `CHANGELOG.md:26`,
  `CHANGELOG.md:35`).
- The empty external contribution inventory remains supported by the selected
  history and the reviewed notes. No notification was due or posted
  (`CHANGELOG.md:51`, `docs/sprints/AS_BUILT.md:11872`).
- F-X079 is consistently completed in its design, sprint state, backlog,
  active sprint, tracker, and completion log (`.claude/plans/F-X079-design.md:3`,
  `.claude/scratch/S69-run.json:46`, `docs/sprints/BACKLOG.md:534`,
  `docs/sprints/CURRENT_SPRINT.md:37`, `docs/sprints/SPRINT_TRACKER.md:394`,
  `docs/sprints/AS_BUILT.md:11838`). The plan's five-file HLD impact matches the
  published-state updates (`.claude/plans/F-X079-design.md:84`).

## F-238 claim-only boundary

- The claim changes only the two delivery trackers. Both mark F-238
  in-progress, the current sprint names `codex` as owner, and the regenerated
  backlog counts show one M22 story in progress (`docs/sprints/BACKLOG.md:40`,
  `docs/sprints/BACKLOG.md:444`, `docs/sprints/CURRENT_SPRINT.md:38`).
- The approved plan names completed F-236, F-X077, and F-X079 dependencies
  (`.claude/plans/F-238-design.md:3`, `.claude/plans/F-238-design.md:6`). The
  sprint state records the isolated `work/f-238-codex` branch, explicit
  worktree, owner, and base at the claim commit
  (`.claude/scratch/S69-run.json:3`). The worktree and branch both resolve to
  that base before implementation begins.
- No F-238 completion row or AS_BUILT entry exists yet, which is correct for a
  claim-only transition. The velocity tracker ends at the completed F-X079 row
  (`docs/sprints/SPRINT_TRACKER.md:394`), while F-238 remains in-progress in
  the live trackers cited above.

## Prior finding closure

- Pass 11 B1 remains closed. The state records successful full verification at
  both the reviewed release SHA and the current claim-only HEAD, with the
  unchanged harness (`.claude/scratch/S69-run.json:219`,
  `.claude/scratch/S69-run.json:225`).
- Pass 11 B2 remains closed. The selected-family notes describe the shared
  validator policy without claiming the stable consumer refactor
  (`CHANGELOG.md:18`, `CHANGELOG.md:26`).
- Pass 11 S1 remains closed. The HLD identifies v0.12.0 as the latest published
  stable family and distinguishes its shared 0.9.0 archives from current source
  pins to published 0.10.0 (`docs/hld/03-architecture.md:850`,
  `docs/hld/10-bindings-spec.md:1149`,
  `docs/hld/15-build-and-toolchain.md:446`).
- Pass 8 and pass 9 S1 remain closed. The sprint rows retain dependency order,
  and the CI documentation does not claim the independent package job feeds the
  aggregate filtered-job gate (`docs/sprints/CURRENT_SPRINT.md:34`,
  `.claude/plans/F-X080-design.md:20`, `.github/workflows/ci.yml:656`).
- Passes 1 through 7 remain closed. F-X079 finalization and the F-238 claim do
  not alter the resource preflight, quote-aware CSS scanner, escaped-syntax
  rejection, nested loss diagnostics, PNG and JPEG boundary, shared lexical
  mapping, or differential oracle (`crates/rdocx/src/html.rs:548`,
  `crates/rdocx/src/html.rs:1444`, `crates/rdocx/src/html.rs:1488`,
  `crates/rdocx/src/html.rs:1786`,
  `crates/rdocx/tests/integration_test.rs:102`).
- The F-X080 repairs remain intact. The package job still enumerates all 24
  TTFs and six legal files (`.github/workflows/ci.yml:533`,
  `.github/workflows/ci.yml:567`). The Pandoc installer retains its digest,
  bounds, exact two aliases, and fail-closed member policy
  (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:30`,
  `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:88`). The Python adapter still maps the new
  native variants exhaustively to the existing class
  (`crates/rdocx-py/src/lib.rs:66`).

## Milestone gate

The M22 gate requires one representative modern document to author and render
equations, rebuild fields and a table of contents, perform advanced merge and
comparison, inventory embedded content, and round-trip its modern package
variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold. The shared 0.10.0 publication is
complete, but F-238 is only claimed and F-X078 remains pending
(`docs/sprints/CURRENT_SPRINT.md:38`, `docs/sprints/CURRENT_SPRINT.md:39`). The
representative modern package round-trip and stable 0.13.0 publication therefore
remain future gates, not defects in this scheduled dependency prefix.

## Not found

- **Interaction**: zero interaction findings were found among F-X077, F-239,
  F-X080, the published F-X079 boundary, and the F-238 claim.
- **Duplication**: zero new duplicate helpers were found. The post-pass-12 delta
  contains release records and claim ledgers only.
- **Layering**: zero dependency-direction findings were found. No post-release
  source or manifest change exists, and no `oxml-*` crate gained a format edge.
- **Harness**: zero unexplained output deltas were found. Exact-current-HEAD full
  verification records an unchanged result
  (`.claude/scratch/S69-run.json:225`).
- **Gate**: zero unexpected completed-prefix gate failures were found. The
  published release evidence and exact-HEAD full verification are current.
- **Docs and ledgers**: zero inconsistencies were found. Published state, tag
  target, versions, selected family, contribution inventory, story status,
  ownership, and claim state agree across their authorities.
- **Dependencies**: zero new external dependency or feature findings were
  found. The only post-pass-12 dependency event is F-238 becoming eligible
  after F-X079 completion.
- **Surface**: zero unplanned public API findings were found. The post-release
  record and claim commits add no source surface.

Focused evidence passed the four selected release metadata, notes, and publish
workflow regressions, deterministic release-note validation and byte comparison,
the Python generic error-class test, all five MHTML unit tests, the shared XML
lexical matrix, and prefix `git diff --check`. Read-only external checks verified
the successful tagged publish run, remote annotated tag target, GitHub release
body, all 15 live unyanked registry entries and their sole owner, plus both
required package absences. The canonical state records `/verify --full` green
at exact HEAD with all 49 hashes unchanged
(`.claude/scratch/S69-run.json:225`).
