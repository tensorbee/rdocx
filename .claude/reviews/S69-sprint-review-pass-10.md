# S69 sprint review, pass 10

**Reviewed**: full integrated dependency prefix on `sprint/s69` at
`0fed61fb052903dbb095d3bb2ceb53a75df54534` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 46 files and 7,256 changed
lines, comprising 6,288 additions and 968 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`, `rdocx-py`
**Pass authority**: pass 10 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Pass 8 and pass 9 closure

- **Pass 8 S1 is closed.** The active wave now lists F-X077 and F-239 before
  their F-X080 dependent, F-X080 before F-X079, F-X079 before F-238, and all
  prerequisites before F-X078 (`docs/sprints/CURRENT_SPRINT.md:34`). The
  sequencing prose agrees with the table and the approved dependencies
  (`docs/sprints/CURRENT_SPRINT.md:43`, `.claude/plans/F-X080-design.md:6`).
- **Pass 9 S1 is closed.** The F-X080 problem statement now says the three
  persistent failures leave the overall hosted workflow and its
  release-readiness checks red (`.claude/plans/F-X080-design.md:20`). It no
  longer claims the unfiltered package job feeds the aggregate status job. The
  tracked workflow keeps `package-oxml-layout` independent and limits
  `ci-gate.needs` to the detector and nine filtered jobs
  (`.github/workflows/ci.yml:518`, `.github/workflows/ci.yml:656`). The HLD
  describes that topology consistently (`docs/hld/12-testing-strategy.md:1877`,
  `docs/hld/12-testing-strategy.md:1902`).

## F-X080 integration

- The package job compares an explicit sorted inventory of all 24 TTFs and all
  six licence and notice files before verified packaging
  (`.github/workflows/ci.yml:533`, `.github/workflows/ci.yml:538`,
  `.github/workflows/ci.yml:567`). The independent regression derives the
  expected inventory from bundled source assets and rejects removal of every
  repaired Noto entry (`scripts/test_sprint_workflow.py:544`,
  `scripts/test_sprint_workflow.py:1110`).
- The Pandoc installer retains the 40 MiB download and 256-member limits, uses
  the reviewed 160 MiB extracted ceiling, and admits only the exact two alias
  name and target pairs (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`, `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:30`). Every member contributes to the
  bounds and passes the root and path checks before the two aliases can be
  skipped. Directories remain accepted and every other unsupported member type
  rejects (`scripts/install_pinned_pandoc.py:63`,
  `scripts/install_pinned_pandoc.py:68`,
  `scripts/install_pinned_pandoc.py:71`,
  `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:85`). Runtime regressions cover the exact
  aliases, wrong names and targets, hardlinks, devices, and FIFO entries
  (`scripts/test_sprint_workflow.py:1130`,
  `scripts/test_sprint_workflow.py:1168`,
  `scripts/test_sprint_workflow.py:1185`).
- The exhaustive Python adapter maps MHTML and invalid embedded-mutation
  failures to the existing `RdocxError` class without adding a method or
  exception class (`crates/rdocx-py/src/lib.rs:66`). Its exact class regression
  covers both variants (`crates/rdocx-py/src/lib.rs:137`).
- The approved HLD impact names exactly the three documents changed by F-X080
  (`.claude/plans/F-X080-design.md:88`). They describe the current package
  inventory, bounded authenticated archive policy, Python mapping, and release
  sequencing (`docs/hld/12-testing-strategy.md:1877`,
  `docs/hld/12-testing-strategy.md:1888`,
  `docs/hld/14-development-backlog.md:3747`,
  `docs/hld/15-build-and-toolchain.md:150`,
  `docs/hld/15-build-and-toolchain.md:580`).
- F-X080 is consistently completed in the backlog, active sprint, tracker,
  completion log, and sprint state (`docs/sprints/BACKLOG.md:535`,
  `docs/sprints/CURRENT_SPRINT.md:36`,
  `docs/sprints/SPRINT_TRACKER.md:393`,
  `docs/sprints/AS_BUILT.md:11791`). F-X079 depends on that completed repair
  before release preparation (`.claude/plans/F-X079-design.md:6`).

## Earlier finding closure

- **F-X080 microscope D1 remains closed.** The plan and HLD distinguish
  accepted directories from rejected unsupported member types, matching the
  extraction branches (`.claude/plans/F-X080-design.md:47`,
  `scripts/install_pinned_pandoc.py:85`).
- **Sprint pass 6 B1 remains closed.** The quote-aware CSS URL extractor
  preserves inner parentheses, rejects trailing syntax, and advances past the
  complete function (`crates/rdocx/src/html.rs:1444`,
  `crates/rdocx/src/html.rs:1460`, `crates/rdocx/src/html.rs:1501`).
- **Sprint pass 5 B1 remains closed.** Escaped CSS resource syntax is detected
  and rejected conservatively before literal extraction
  (`crates/rdocx/src/html.rs:1488`).
- **Sprint pass 4 B2 remains closed.** The Python adapter compiles exhaustively
  and preserves its existing public error surface
  (`crates/rdocx-py/src/lib.rs:76`).
- **Sprint pass 3 B1 remains closed.** The nested loss walk distinguishes
  shapes, other drawings, linked images, unresolved images, and supported
  embedded images (`crates/rdocx/src/html.rs:597`).
- **Sprint pass 1 B1, pass 2 B1, pass 1 B2, pass 1 B3, and pass 1 B4 remain
  closed.** Resource selectors, PNG and JPEG restrictions, the shared
  mutation-sensitive Word predicate, and source-ordered nested loss handling
  retain passing focused evidence (`crates/rdocx/src/html.rs:1786`,
  `crates/rdocx/src/html.rs:1824`,
  `crates/rdocx/tests/integration_test.rs:102`,
  `crates/rdocx/src/html.rs:548`).
- **Sprint pass 2 S1 remains closed.** The F-239 completion record attributes
  the corrected common-input Word oracle and full verification to
  `7fde4033b7cdf17f7c6e309dfccf7d1b9a6b1d44`
  (`docs/sprints/AS_BUILT.md:11779`).

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold at this dependency-prefix checkpoint.
F-X079, F-238, and F-X078 remain pending in dependency order
(`docs/sprints/CURRENT_SPRINT.md:37`, `docs/sprints/CURRENT_SPRINT.md:38`,
`docs/sprints/CURRENT_SPRINT.md:39`). Within the completed prefix, F-X077's
shared validator, F-239's bounded MHTML conversion, and F-X080's locally
reconstructed release-readiness paths all pass. The latest recorded full
verification covers the completed implementation at
`26f5ff5c5e68ab2f0b5118753600ad95fdc7ba23`. The only later changes are the two
documentation corrections independently reviewed here.

## Not found

- **Interaction**: zero code interaction findings were found among F-X077,
  F-239, and F-X080. F-X080 covers the Python consumer of F-239's new error and
  gates F-X079 before it publishes F-X077's shared API.
- **Duplication**: zero sprint-local duplicate lexical, inventory, extraction,
  or resource helpers were found.
- **Layering**: zero dependency-direction findings were found. No manifest or
  lockfile changed, and no `oxml-*` crate gained a dependency on `rdocx-*` or
  `rpptx-*`.
- **Harness**: zero unexplained output deltas were found. The independent check
  reports all 49 entries match, consistent with all three completion records
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`,
  `docs/sprints/AS_BUILT.md:11831`).
- **Gate**: zero unexpected completed-prefix gate failures were found. The
  remaining M22 clauses belong to the three pending stories named above.
- **Docs and ledgers**: zero inconsistencies were found. Story status,
  ownership, size, dependency order, HLD impact, and completion evidence agree.
- **Dependencies**: zero new dependency, feature, version, or release-allowlist
  findings were found.
- **Surface**: zero unplanned public API findings were found. F-X080 changes
  only CI, bounded installer policy, tests, and their approved documentation.

Focused evidence passed all 96 sprint-workflow regressions with one expected
skip, `cargo check -p rdocx-py --all-targets`, the Python generic error-class
test, focused `rdocx-py` Clippy, verified `oxml-layout` packaging with the exact
24-font and six-legal-file inventory in a 4,603,304-byte archive, all five MHTML
unit tests, both ordinary MHTML integration tests, the shared XML unit matrix,
all 42 glossary tests, the embedded owner mapping test, all five package-story
lexical tests, the 49-entry hash harness, skill-adapter drift, prose, and `git
diff --check`. The Linux-only authenticated Pandoc download and the ignored
Microsoft Word regeneration test were not rerun in this pass. Their integrated
evidence is recorded in the completion log (`docs/sprints/AS_BUILT.md:11822`).
