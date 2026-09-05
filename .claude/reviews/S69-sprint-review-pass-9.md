# S69 sprint review, pass 9

**Reviewed**: full integrated dependency prefix on `sprint/s69` at
`e885b9454f181da2cfae31c77b043898697d7766` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 45 files and 7,080 changed
lines, comprising 6,112 additions and 968 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`, `rdocx-py`
**Pass authority**: pass 9 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 0 blocking, 1 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

### S1, the F-X080 plan overstates which failures feed the aggregate CI job
`.claude/plans/F-X080-design.md:20`
`.github/workflows/ci.yml:518`
`.github/workflows/ci.yml:656`
`.github/workflows/ci.yml:659`
`docs/hld/12-testing-strategy.md:1902`

The plan says the package, Pandoc test, and Python binding failures all block
the aggregate `CI gate`. The workflow's `ci-gate.needs` intentionally contains
the change detector and nine filtered jobs, including the test and Python jobs,
but not the always-running `package-oxml-layout` job. A package failure makes
the overall workflow red and remains independently covered by the release dry
run, but it does not make the required `CI gate` job fail. This does not weaken
the implemented inventory repair or its regression. Correct the problem
statement to distinguish overall hosted workflow readiness from the aggregate
filtered-job status, without changing the reviewed gate topology.

## Nice-to-have

None.

## Pass 8 remediation

- **Pass 8 S1 is closed.** The active wave now lists F-X077 and F-239 before
  their F-X080 dependent, F-X080 before F-X079, F-X079 before F-238, and all of
  those prerequisites before F-X078 (`docs/sprints/CURRENT_SPRINT.md:34`). The
  sequencing explanation agrees with that order
  (`docs/sprints/CURRENT_SPRINT.md:43`). No status, owner, size, or dependency
  changed in the ordering-only remediation.

## F-X080 integration

- The package job compares an explicit sorted inventory of all 24 TTFs and all
  six licence and notice files before verified packaging
  (`.github/workflows/ci.yml:518`, `.github/workflows/ci.yml:538`,
  `.github/workflows/ci.yml:567`). The independent workflow regression derives
  the expected inventory from the source assets and is mutation-sensitive for
  every repaired Noto entry (`scripts/test_sprint_workflow.py:544`,
  `scripts/test_sprint_workflow.py:1110`). The pass-9 reconstruction found an
  exact inventory match and verified a 4,603,301-byte archive below 10 MiB.
- The Pandoc installer retains the 40 MiB download and 256-member limits, uses
  the reviewed 160 MiB extracted ceiling, and admits only the exact two alias
  name and target pairs (`scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:30`). Every member contributes to the
  member and extracted-size bounds, its resolved path stays in the fresh
  destination, the two aliases are skipped without materialization,
  directories remain accepted, and every other unsupported member type is
  rejected (`scripts/install_pinned_pandoc.py:63`,
  `scripts/install_pinned_pandoc.py:74`,
  `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:85`,
  `scripts/install_pinned_pandoc.py:88`). Runtime regressions cover the exact
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
  inventory, Pandoc bounds and aliases, Python mapping, and release sequencing
  (`docs/hld/12-testing-strategy.md:1877`,
  `docs/hld/12-testing-strategy.md:1888`,
  `docs/hld/15-build-and-toolchain.md:150`,
  `docs/hld/15-build-and-toolchain.md:580`).
- F-X080 is consistently completed in the backlog, current sprint, tracker,
  completion log, and run state (`docs/sprints/BACKLOG.md:535`,
  `docs/sprints/CURRENT_SPRINT.md:36`, `docs/sprints/SPRINT_TRACKER.md:393`,
  `docs/sprints/AS_BUILT.md:11791`). F-X079 now depends on that completed repair
  before release preparation (`.claude/plans/F-X079-design.md:6`).

## Prior finding closure

- **F-X080 microscope D1 remains closed.** The plan and HLD distinguish
  accepted directories from rejected unsupported member types, matching the
  extraction branches (`.claude/plans/F-X080-design.md:45`,
  `scripts/install_pinned_pandoc.py:85`).
- **Sprint pass 6 B1 remains closed.** The quote-aware CSS URL extractor
  preserves inner parentheses, rejects trailing syntax, and advances past the
  complete function (`crates/rdocx/src/html.rs:1444`,
  `crates/rdocx/src/html.rs:1460`, `crates/rdocx/src/html.rs:1501`). The matrix
  covers the quoted URL and import prefix collisions and asserts the complete
  extracted URL (`crates/rdocx/src/html.rs:3535`,
  `crates/rdocx/src/html.rs:3573`).
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
F-X079, F-238, and F-X078 remain pending in their declared order
(`docs/sprints/CURRENT_SPRINT.md:37`, `docs/sprints/CURRENT_SPRINT.md:38`,
`docs/sprints/CURRENT_SPRINT.md:39`). Within the completed prefix, F-X077's
shared validator, F-239's bounded MHTML conversion, and F-X080's locally
reconstructed release-readiness paths all pass. The latest full verification
is recorded at `26f5ff5c5e68ab2f0b5118753600ad95fdc7ba23`. The only later implementation
change is the one-line dependency-order documentation correction reviewed here.

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
- **Docs and ledgers**: zero inconsistencies were found beyond S1. Story status,
  ownership, size, dependency, HLD impact, and completion evidence otherwise
  agree.
- **Dependencies**: zero new dependency, feature, version, or release-allowlist
  findings were found.
- **Surface**: zero unplanned public API findings were found. F-X080 changes
  only CI, bounded installer policy, tests, and their approved documentation.

Focused evidence passed all 96 sprint-workflow regressions with one expected
skip, `cargo check -p rdocx-py --all-targets`, the Python generic error-class
test, focused `rdocx-py` Clippy, verified `oxml-layout` packaging with the exact
24-font and six-legal-file inventory, all five MHTML unit tests, both ordinary
MHTML integration tests, the shared XML unit matrix, all 42 glossary tests,
the embedded owner mapping test, all five package-story lexical tests, the
49-entry hash harness, skill-adapter drift, prose, and `git diff --check`. The
Linux-only authenticated Pandoc download and the ignored Microsoft Word
regeneration test were not rerun in this pass. Their integrated evidence is
recorded in the completion log (`docs/sprints/AS_BUILT.md:11822`).
