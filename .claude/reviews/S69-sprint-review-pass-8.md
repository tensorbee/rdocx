# S69 sprint review, pass 8

**Reviewed**: full integrated dependency prefix on `sprint/s69` at
`26f5ff5c5e68ab2f0b5118753600ad95fdc7ba23` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 44 files and 6,924 changed
lines, comprising 5,956 additions and 968 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`, `rdocx-py`
**Pass authority**: pass 8 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 0 blocking, 1 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

### S1, the wave table contradicts its dependency-order claim
`docs/sprints/CURRENT_SPRINT.md:35`
`docs/sprints/CURRENT_SPRINT.md:38`
`docs/sprints/CURRENT_SPRINT.md:41`
`docs/sprints/CURRENT_SPRINT.md:44`
`.claude/plans/F-X080-design.md:6`

The canonical sprint record says its rows are listed in dependency order, but
it places F-X080 before F-239 even though F-X080 explicitly depends on F-239.
The sequencing prose also says F-X080 runs after F-239 has settled its binding
surface. The state machine and completion statuses are correct, so this does
not invalidate the integrated code, but the delivery record is internally
contradictory. Reorder the completed rows so F-239 precedes F-X080 while
retaining every status and dependency.

## Nice-to-have

None.

## F-X080 integration

- The CI package job compares an explicit sorted inventory containing all 24
  TTFs and all six licence and notice files before verified packaging
  (`.github/workflows/ci.yml:518`, `.github/workflows/ci.yml:538`,
  `.github/workflows/ci.yml:567`). The mutation-sensitive regression derives
  the source-asset inventory independently and fails when any repaired Noto
  entry is removed (`scripts/test_sprint_workflow.py:544`,
  `scripts/test_sprint_workflow.py:1110`).
- The Pandoc installer keeps the 40 MiB download and 256-member bounds, raises
  only the extracted ceiling to 160 MiB, and admits only the exact two reviewed
  symlink name and target pairs (`scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`, `scripts/install_pinned_pandoc.py:30`).
  The extraction loop counts skipped entries toward both bounds, checks their
  in-root resolved paths before skipping them, accepts directories, and rejects
  every other non-file member (`scripts/install_pinned_pandoc.py:63`,
  `scripts/install_pinned_pandoc.py:71`, `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:85`). The regression covers exact aliases,
  wrong names, wrong targets, hardlinks, devices, and FIFO entries
  (`scripts/test_sprint_workflow.py:1130`,
  `scripts/test_sprint_workflow.py:1168`,
  `scripts/test_sprint_workflow.py:1185`).
- F-X080 introduces no Python method or exception class. The exhaustive adapter
  maps both MHTML and invalid embedded-mutation failures to the established
  `RdocxError` class (`crates/rdocx-py/src/lib.rs:66`), and the exact class
  regression covers both variants (`crates/rdocx-py/src/lib.rs:137`).
- The design's HLD impact names exactly the three mechanism and test documents
  updated by the implementation (`.claude/plans/F-X080-design.md:91`). The HLD
  records the exact inventory and bounded alias policy as current behavior
  (`docs/hld/12-testing-strategy.md:1877`,
  `docs/hld/12-testing-strategy.md:1888`,
  `docs/hld/15-build-and-toolchain.md:152`,
  `docs/hld/15-build-and-toolchain.md:580`).
- The completed F-X080 entry agrees across the backlog, current sprint,
  tracker, and completion log apart from S1's row position
  (`docs/sprints/BACKLOG.md:535`, `docs/sprints/CURRENT_SPRINT.md:35`,
  `docs/sprints/SPRINT_TRACKER.md:393`, `docs/sprints/AS_BUILT.md:11791`). The
  release story now depends on this completed repair before preparation begins
  (`.claude/plans/F-X079-design.md:6`).

## Prior finding closure

- **F-X080 microscope D1 is closed.** The amended plan and all three HLD files
  distinguish accepted directories from rejected unsupported non-file member
  types (`.claude/plans/F-X080-design.md:47`,
  `scripts/install_pinned_pandoc.py:85`).
- **Sprint pass 6 B1 remains closed.** The quote-aware CSS URL extractor
  preserves inner parentheses, rejects trailing syntax, advances past the
  complete function, and is shared by ordinary and quoted import URL forms
  (`crates/rdocx/src/html.rs:1444`, `crates/rdocx/src/html.rs:1460`,
  `crates/rdocx/src/html.rs:1501`).
- **Sprint pass 5 B1 remains closed.** Escaped CSS resource syntax is detected
  and conservatively rejected before literal extraction
  (`crates/rdocx/src/html.rs:1488`).
- **Sprint pass 4 B2 remains closed.** The Python adapter compiles exhaustively
  and maps both previously missing native variants to its existing generic
  class (`crates/rdocx-py/src/lib.rs:76`). F-X080 adds release-regression
  coverage around that repaired interaction.
- **Sprint pass 3 B1 remains closed.** The nested loss walk distinguishes
  shapes, other drawings, linked images, unresolved images, and supported
  embedded images (`crates/rdocx/src/html.rs:597`).
- **Sprint pass 1 B1, pass 2 B1, pass 1 B2, pass 1 B3, and pass 1 B4 remain
  closed.** Resource selectors, PNG and JPEG restrictions, the shared
  mutation-sensitive Word predicate, and source-ordered nested loss handling
  retain their passing focused evidence (`crates/rdocx/src/html.rs:1786`,
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
The modern package-variant story F-238 and both release stories remain pending
(`docs/sprints/CURRENT_SPRINT.md:36`, `docs/sprints/CURRENT_SPRINT.md:37`,
`docs/sprints/CURRENT_SPRINT.md:39`). Within the completed prefix, F-X077's
shared validator, F-239's bounded MHTML conversion, and F-X080's locally
reconstructed release-readiness gates all pass. The sprint state records a full
verification at the reviewed exact HEAD with the hash harness unchanged.

## Not found

- **Interaction**: zero code interaction findings were found among F-X077,
  F-239, and F-X080. F-X080 specifically tests the Python consumer of F-239's
  MHTML error and the release path for F-X077's new published shared API.
- **Duplication**: zero sprint-local duplicate lexical, inventory, or extraction
  helpers were found.
- **Layering**: zero dependency-direction violations were found. No manifest
  changed, and no `oxml-*` crate gained a dependency on `rdocx-*` or `rpptx-*`.
- **Harness**: zero unexplained output deltas were found. The independent check
  reports all 49 entries match, consistent with all three completion records
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`,
  `docs/sprints/AS_BUILT.md:11831`).
- **Docs**: zero HLD impact findings were found. S1 is limited to ordering in
  the active delivery record.
- **Dependencies**: zero new dependency, feature, version, or release-allowlist
  findings were found. The integrated prefix changes no manifest or lockfile.
- **Surface**: zero unplanned public API findings were found. F-X080 changes
  only CI, installer policy, tests, and their approved documentation.

Focused evidence passed all 96 sprint-workflow regressions with one expected
skip, `cargo check -p rdocx-py --all-targets`, the Python generic error-class
test, focused `rdocx-py` Clippy, verified `oxml-layout` packaging with the exact
24-font and six-legal-file inventory in a 4,603,300-byte archive, all five MHTML
unit tests, both ordinary MHTML integration tests, the shared XML unit matrix,
all 42 glossary tests, the embedded owner mapping test, all five package-story
lexical tests, the 49-entry hash harness, skill-adapter drift, prose, and `git
diff --check`. The Linux-only authenticated Pandoc download and the ignored
Microsoft Word regeneration test were not rerun in this pass. Their exact-SHA
integrated evidence is recorded in the sprint state and completion log
(`docs/sprints/AS_BUILT.md:11822`).
