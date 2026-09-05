# S69 sprint review, pass 11

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`d4b3437363483644e88ce0100015a36b9739f243` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 88 files and 7,797 changed
lines, comprising 6,695 additions and 1,102 deletions. The 20 touched crates
are `oxml-chart`, `oxml-cli-support`, `oxml-core`, `oxml-drawing`,
`oxml-layout`, `oxml-media`, `oxml-opc`, `oxml-pdf`, `oxml-sml`, `rdocx`,
`rdocx-oxml`, `rdocx-py`, `rdocx-wasm`, `rpptx`, `rpptx-chart`, `rpptx-cli`,
`rpptx-layout`, `rpptx-oxml`, `rpptx-render`, and `rpptx-wasm`
**Pass authority**: pass 11 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 2 blocking, 1 should-fix, 0 nice-to-have

## Blocking

### B1, exact-HEAD full verification is not recorded

`.claude/scratch/S69-run.json:191`
`.claude/plans/F-X079-design.md:124`

The implementing session reports that full verification passed at the reviewed
`d4b3437` SHA, and the focused evidence below is consistent with that report.
The canonical sprint state still ends its verification history at
`16a06965c30d23265e4e5ae613800598b4926d3b`, before the 49-file F-X079
integration. This also makes the checked exact-SHA verification item in the
F-X079 plan unsupported by the delivery record. `/release` requires the latest
recorded full verification to cover the current HEAD
(`.claude/commands/release.md:61`). Record the successful full gate and its
unchanged 49-entry harness result at `d4b3437`, or rerun and record it at the
post-remediation release candidate SHA, before release approval.

### B2, the incubating release notes claim a stable-only consumer change

`CHANGELOG.md:28`

The `rpptx-v0.10.0` Fixed section says this release replaces three independent
lexical checks while retaining each consumer's failure surface. Those three
consumer replacements are in `rdocx-oxml` and `rdocx`, outside the exact
15-package incubating family. The calls are at
`crates/rdocx-oxml/src/glossary.rs:481`,
`crates/rdocx/src/embedded.rs:1098`, and `crates/rdocx/src/field.rs:7953`.
Only the new shared validator in `oxml-core` belongs to this release
(`crates/oxml-core/src/xml.rs:37`). The release command requires notes to cover
only the selected family (`.claude/commands/release.md:51`). Rewrite the Fixed
claim to describe only behavior delivered by the selected packages, while
leaving the stable consumer adoption for the later stable release notes.

## Should-fix

### S1, the HLD still describes v0.12.0 as unpublished

`docs/hld/03-architecture.md:853`
`docs/hld/10-bindings-spec.md:1152`
`docs/hld/15-build-and-toolchain.md:448`

All three F-X079 HLD sections that describe the release graph still say the
latest complete stable family is 0.11.1 and that v0.12.0 has not completed.
The completion ledger records the exact seven-package v0.12.0 publication and
GitHub release as complete (`docs/sprints/AS_BUILT.md:11176`,
`docs/sprints/AS_BUILT.md:11181`). Because F-X079 edits these same current-state
sections and its approved HLD impact names all three files
(`.claude/plans/F-X079-design.md:84`), leaving the stale stable boundary makes
the prepared dependency graph internally contradictory. Update each section to
state that v0.12.0 is the latest published stable family while preserving its
published shared 0.9.0 archive boundary and the current source pin to prepared
0.10.0.

## Nice-to-have

None.

## F-X079 release preparation

- The exact 15 selected manifests and workspace pins are at 0.10.0, while
  `rpptx-wasm` and the binding crates remain unpublished. The package set and
  exclusions match the release contract (`Cargo.toml:55`,
  `.claude/commands/release.md:31`).
- The incubating publish path lists exactly the selected packages in dependency
  order and retains bare failure-propagating publish commands
  (`.github/workflows/publish.yml:78`). Stable and incubating tag predicates
  remain disjoint (`.github/workflows/publish.yml:61`,
  `.github/workflows/publish.yml:78`).
- The package inventory still names all 24 bundled TTFs and six legal files,
  and the package archive path retains the 10 MiB bound
  (`.github/workflows/ci.yml:533`, `.github/workflows/ci.yml:567`,
  `.github/workflows/ci.yml:582`).
- The selected-family history since `rpptx-v0.9.0` contains the internally
  authored F-229 layout primitive, F-237 package constants, and F-X077 shared
  validator. Their completion records name no external contribution
  (`docs/sprints/AS_BUILT.md:11277`, `docs/sprints/AS_BUILT.md:11653`,
  `docs/sprints/AS_BUILT.md:11702`). The empty contribution inventory at
  `CHANGELOG.md:55` is therefore supported. No notification is due unless the
  corrected selected-family notes add an externally sourced outcome.
- F-X079 remains reviewed in sprint state and in-progress in both delivery
  trackers, with completed F-X077 and F-X080 dependencies
  (`.claude/scratch/S69-run.json:43`, `docs/sprints/CURRENT_SPRINT.md:34`,
  `docs/sprints/CURRENT_SPRINT.md:37`, `docs/sprints/BACKLOG.md:534`). This is
  the correct prepublication lifecycle state.

## Prior finding closure

- Pass 8 S1 remains closed. The active sprint rows retain dependency order from
  F-X077 and F-239 through F-X080 and F-X079
  (`docs/sprints/CURRENT_SPRINT.md:34`).
- Pass 9 S1 remains closed. The F-X080 plan no longer claims the independent
  package job feeds the aggregate gate, and the workflow still limits
  `ci-gate.needs` to the detector and filtered jobs
  (`.claude/plans/F-X080-design.md:20`, `.github/workflows/ci.yml:656`).
- Passes 1 through 7 remain closed. The resource preflight, quote-aware CSS
  scanner, escaped-syntax rejection, nested loss diagnostics, PNG and JPEG
  boundary, shared lexical mapping, and differential oracle are unchanged by
  F-X079 (`crates/rdocx/src/html.rs:548`, `crates/rdocx/src/html.rs:1444`,
  `crates/rdocx/src/html.rs:1488`, `crates/rdocx/src/html.rs:1786`,
  `crates/rdocx/tests/integration_test.rs:102`).
- The F-X080 package, Pandoc, and Python repairs remain intact. The Pandoc
  installer retains its checksum, bounds, exact two aliases, and fail-closed
  member policy (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:30`,
  `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:88`). The Python mapper remains exhaustive
  for both new native variants (`crates/rdocx-py/src/lib.rs:66`).

## Milestone gate

The M22 gate requires one representative modern document to author and render
equations, rebuild fields and a table of contents, perform advanced merge and
comparison, inventory embedded content, and round-trip its modern package
variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold. F-X079 is still in progress, and
F-238 plus F-X078 remain pending (`docs/sprints/CURRENT_SPRINT.md:37`,
`docs/sprints/CURRENT_SPRINT.md:38`, `docs/sprints/CURRENT_SPRINT.md:39`). The
implemented prefix retains its focused technical gates, but B1 and B2 prevent
the F-X079 release approval boundary from being clean.

## Not found

- **Interaction**: zero additional interaction findings were found among
  F-X077, F-239, F-X080, and the F-X079 metadata carriers.
- **Duplication**: zero new duplicate helpers were found.
- **Layering**: zero dependency-direction findings were found. No `oxml-*`
  package gained a dependency on either document family.
- **Harness**: zero unexplained output deltas were found. The independent check
  reports all 49 entries match, consistent with the feature completion records
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`,
  `docs/sprints/AS_BUILT.md:11831`).
- **Dependencies**: zero new external dependency or feature findings were
  found. The manifest and lockfile changes are version-carrier updates only.
- **Surface**: zero unplanned public API findings were found beyond the approved
  additive `oxml-core` validator (`.claude/plans/F-X077-design.md:79`).

Focused evidence passed all 97 sprint-workflow tests with one expected
registry skip, the five selected release metadata, notes, and workflow tests,
release-note check and deterministic render, `cargo metadata --no-deps`, the
exact `oxml-layout` package inventory, the Python generic error-class test, all
five MHTML unit tests, the shared XML lexical matrix, the 49-entry hash harness,
skill-adapter drift, prose, and `git diff --check`. The implementing session
reports the full gate green at `d4b3437`, but B1 applies until that result is
present in the canonical sprint state.
