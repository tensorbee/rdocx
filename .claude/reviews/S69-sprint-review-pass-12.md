# S69 sprint review, pass 12

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`03e7a82a4957d7f15b228afd69a12033c6dce438` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 89 files and 8,001 changed
lines, comprising 6,878 additions and 1,123 deletions. The 20 touched crates
are `oxml-chart`, `oxml-cli-support`, `oxml-core`, `oxml-drawing`,
`oxml-layout`, `oxml-media`, `oxml-opc`, `oxml-pdf`, `oxml-sml`, `rdocx`,
`rdocx-oxml`, `rdocx-py`, `rdocx-wasm`, `rpptx`, `rpptx-chart`, `rpptx-cli`,
`rpptx-layout`, `rpptx-oxml`, `rpptx-render`, and `rpptx-wasm`
**Pass authority**: pass 12 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Pass 11 remediation

- **Pass 11 B1 is closed.** The sprint state records a successful full
  verification at the exact reviewed HEAD with an unchanged harness
  (`.claude/scratch/S69-run.json:203`). This supports the checked exact-SHA
  verification item in the release design (`.claude/plans/F-X079-design.md:124`)
  and satisfies the release precondition that the latest full verification
  cover current HEAD (`.claude/commands/release.md:61`).
- **Pass 11 B2 is closed.** The rendered `rpptx-v0.10.0` notes no longer claim
  that the selected release replaces the three stable Word consumers. They
  describe the additive validator in `oxml-core`, baseline-aware groups in
  `oxml-layout`, and glossary constants in `oxml-opc`
  (`CHANGELOG.md:11`, `CHANGELOG.md:18`, `CHANGELOG.md:20`,
  `CHANGELOG.md:23`). Those are exactly the three selected-package source
  changes since `rpptx-v0.9.0`, whose implementations remain at
  `crates/oxml-core/src/xml.rs:13`, `crates/oxml-layout/src/line.rs:99`,
  `crates/oxml-opc/src/content_types.rs:20`, and
  `crates/oxml-opc/src/relationship.rs:57`. Stable Word remains explicitly
  outside this release's authority (`CHANGELOG.md:46`).
- **Pass 11 S1 is closed.** The three current-state HLD boundaries now identify
  v0.12.0 as the latest published exact seven-package stable family, name the
  immutable tag target, retain the published archives' shared 0.9.0
  requirements, and distinguish the current shared 0.10.0 source pins
  (`docs/hld/03-architecture.md:850`, `docs/hld/10-bindings-spec.md:1146`,
  `docs/hld/15-build-and-toolchain.md:447`). The tag target and published
  package boundary agree with the completion record
  (`docs/sprints/AS_BUILT.md:11176`).

## F-X079 release preparation

- All 15 publishable incubating manifests, their workspace pins, and their
  lock records are at 0.10.0. The stable workspace and stable pins remain at
  0.12.0 (`Cargo.toml:34`, `Cargo.toml:55`, `Cargo.toml:71`). The unpublished
  `rpptx-wasm` preparation member is also at 0.10.0 and remains outside the
  workspace dependency allowlist (`crates/rpptx-wasm/Cargo.toml:5`,
  `scripts/test_sprint_workflow.py:5476`).
- The publication workflow preflights the 0.12.0 stable and 0.10.0 incubating
  families, verifies the published shared family only for a stable tag, and
  publishes the exact 15 incubating packages in dependency order
  (`.github/workflows/publish.yml:23`, `.github/workflows/publish.yml:26`,
  `.github/workflows/publish.yml:78`). Real publication, tag creation, and
  notification verification remain behind the separate approval boundary
  (`.claude/plans/F-X079-design.md:126`).
- The selected-family history contains only the shared changes recorded in the
  release notes. The contribution inventory is therefore empty, and the notes
  contain no issue or pull request links (`CHANGELOG.md:51`). The deterministic
  release-note check and render both pass at current HEAD, and the focused
  regression pins the intended selected-family claims
  (`scripts/test_sprint_workflow.py:5585`).
- The CI and source version literals agree with the prepared families. The WASM
  package gate retains stable 0.12.0 for rdocx and incubating 0.10.0 for rpptx,
  while the release-regression job runs the complete metadata module
  (`.github/workflows/ci.yml:363`, `.github/workflows/ci.yml:366`).

## CI repair closure

- The F-X080 package repair remains intact. The dedicated job enumerates all 24
  TTFs and all six legal files before archive verification and the 10 MiB bound
  (`.github/workflows/ci.yml:518`, `.github/workflows/ci.yml:538`,
  `.github/workflows/ci.yml:567`, `.github/workflows/ci.yml:582`). The package
  inventory regression remains mutation-sensitive for the bundled source
  assets (`scripts/test_sprint_workflow.py:1110`).
- The Pandoc installer retains its authenticated digest, 40 MiB download cap,
  256-member cap, 160 MiB extracted cap, exact two skipped aliases, root and
  path containment checks, and fail-closed member policy
  (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:30`,
  `scripts/install_pinned_pandoc.py:63`,
  `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:88`).
- The Python adapter still maps both MHTML and invalid embedded-mutation
  failures to the existing generic `RdocxError` class through an exhaustive
  match (`crates/rdocx-py/src/lib.rs:66`).
- Pass 9 S1 remains closed. The independent package job is not described as a
  dependency of the aggregate filtered-job status, and `ci-gate.needs` still
  contains only the detector and filtered jobs
  (`.claude/plans/F-X080-design.md:20`, `.github/workflows/ci.yml:656`).

## Earlier finding closure

- Pass 8 S1 remains closed. The active sprint rows keep F-X077 and F-239 before
  F-X080, F-X080 before F-X079, and F-X079 before its consumers
  (`docs/sprints/CURRENT_SPRINT.md:34`).
- Pass 6 B1 remains closed. The quote-aware CSS URL scanner retains parentheses
  inside quoted URLs, rejects trailing syntax, and advances beyond the complete
  function (`crates/rdocx/src/html.rs:1444`,
  `crates/rdocx/src/html.rs:1460`, `crates/rdocx/src/html.rs:1501`).
- Pass 5 B1 remains closed. Escaped resource syntax is decoded and rejected
  conservatively before literal resource extraction
  (`crates/rdocx/src/html.rs:1488`).
- Pass 4 B1 and B2 remain closed. The CSS identifier and delimiter matrix still
  covers the escaped collision forms, and the Python adapter retains its
  exhaustive existing public error surface (`crates/rdocx/src/html.rs:3513`,
  `crates/rdocx-py/src/lib.rs:66`).
- Pass 3 B1 remains closed. MHTML loss traversal still distinguishes shapes,
  other drawings, linked images, unresolved images, and supported embedded
  images (`crates/rdocx/src/html.rs:588`).
- Pass 1 B1, pass 2 B1, pass 1 B2, pass 1 B3, and pass 1 B4 remain closed.
  Resource selectors, PNG and JPEG limits, the shared mutation-sensitive Word
  predicate, and source-ordered nested loss traversal retain their reviewed
  forms (`crates/rdocx/src/html.rs:1786`, `crates/rdocx/src/html.rs:1824`,
  `crates/rdocx/tests/integration_test.rs:102`,
  `crates/rdocx/src/html.rs:548`). The F-239 completion record still attributes
  the corrected common-input Word oracle to its exact verified SHA
  (`docs/sprints/AS_BUILT.md:11779`).

## Milestone gate

The M22 gate requires one representative modern document to author and render
equations, rebuild fields and a table of contents, perform advanced merge and
comparison, inventory embedded content, and round-trip its modern package
variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold. F-X079 is reviewed but remains
in-progress until its separately approved release and verification completes.
F-238 and F-X078 remain pending (`docs/sprints/CURRENT_SPRINT.md:37`,
`docs/sprints/CURRENT_SPRINT.md:38`, `docs/sprints/CURRENT_SPRINT.md:39`). The
completed technical prefix and current release preparation retain their focused
gates, but the representative modern package round trip and both publication
boundaries are deliberately unfinished.

## Not found

- **Interaction**: zero interaction findings were found among F-X077, F-239,
  F-X080, and F-X079. The release carriers publish the shared validator without
  changing the completed Word consumers, and F-X080 covers the F-239 Python
  error addition.
- **Duplication**: zero sprint-local duplicate lexical, resource, package,
  release-family, or extraction helpers were found.
- **Layering**: zero dependency-direction findings were found. Metadata reports
  no `oxml-*` package dependency on `rdocx-*`, `rpptx-*`, `rdocx`, or `rpptx`.
- **Harness**: zero unexplained output deltas were found. The exact-HEAD state
  records an unchanged full gate, and the focused rerun reports all 49 entries
  match (`.claude/scratch/S69-run.json:203`).
- **Gate**: zero unexpected completed-prefix gate failures were found. The
  incomplete M22 clauses belong to the pending release and package stories
  named above.
- **Docs and ledgers**: zero documentation or ledger inconsistencies were
  found. Story status, ownership, dependency order, HLD impact, release-family
  boundary, and recorded verification agree.
- **Dependencies**: zero new external dependency or feature findings were
  found. Manifest and lockfile changes are release-version carriers only.
- **Surface**: zero unplanned public API findings were found. The additive
  shared validator is required by F-X077, and the earlier shared layout and OPC
  additions are correctly included in the selected-family release notes.

Focused evidence passed the eight selected workflow, release, version,
inventory, Pandoc, and authority regressions, deterministic
`rpptx-v0.10.0` release-note check and render, `cargo metadata --no-deps`, the
strict shared XML lexical regression, all five MHTML unit tests, the Python
generic error-class test, the 49-entry hash harness, skill-adapter drift, and
prefix `git diff --check`. The recorded `/verify --full` at exact HEAD supplies
the complete workspace, packaging, asset, binding, WASM, dependency,
supply-chain, oracle, and platform evidence and was not redundantly rerun in
this pass.
