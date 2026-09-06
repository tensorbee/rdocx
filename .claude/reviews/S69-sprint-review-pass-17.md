# S69 sprint review, pass 17

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`da5009f5516876d1d17caf4822b3a12f823de486` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 114 files and 11,724 changed
lines, comprising 10,510 additions and 1,214 deletions. The 26 crate
directories with changed files are `oxml-chart`, `oxml-cli-support`,
`oxml-core`, `oxml-drawing`, `oxml-layout`, `oxml-media`, `oxml-opc`,
`oxml-pdf`, `oxml-sml`, `rdocx`, `rdocx-cli`, `rdocx-html`, `rdocx-layout`,
`rdocx-opc`, `rdocx-oxml`, `rdocx-pdf`, `rdocx-py`, `rdocx-wasm`, `rpptx`,
`rpptx-chart`, `rpptx-cli`, `rpptx-layout`, `rpptx-oxml`, `rpptx-py`,
`rpptx-render`, and `rpptx-wasm`. The `oxml-py-support` package also changes
effective version through workspace inheritance.
**Pass authority**: pass 17 extends the default three-pass bound because the
user scheduled a new dependency-prefix review at the F-X078 prepared stable
v0.13.0 release boundary. This pass reviews the newly integrated release
carriers and notes after the clean pass-16 M22 prefix rather than resampling an
unchanged candidate.
**Verdict**: 0 blocking, 1 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

### S1, preparation records understate the completed exact-HEAD verification

`.claude/plans/F-X078-design.md:139`
`.claude/scratch/F-X078-progress.md:20`
`.claude/scratch/S69-run.json:286`

The design checklist still leaves the full verification and its packaging,
asset, binding, WASM, dependency, supply-chain, notes, and hash riders
unchecked. The in-flight progress note likewise says to run the microscope and
full verification next. The canonical sprint state now records that full gate
as passed at exact prepared HEAD
`da5009f5516876d1d17caf4822b3a12f823de486`, and the clean microscope artifact
is already committed. This does not invalidate the gate, but the delivery
records disagree about whether the preparation prerequisite happened. Mark the
completed verification item and update the in-flight next step to the separate
`/release v0.13.0` approval boundary without marking publication complete.

## Nice-to-have

None.

## F-X078 prepared-release boundary

- The workspace version is 0.13.0. Its nine stable-group pins are 0.13.0,
  while the shared OOXML and PowerPoint dependency pins remain at their
  published 0.10.0 boundary (`Cargo.toml:34`, `Cargo.toml:55`,
  `Cargo.toml:78`). Metadata inspection found exactly seven publishable 0.13.0
  crates and four unpublished inherited carriers.
- The stable workflow publishes exactly `rdocx-opc`, `rdocx-oxml`,
  `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli` in
  dependency order (`.github/workflows/publish.yml:61`). The separately routed
  incubating family remains the exact 15-package 0.10.0 allowlist
  (`.github/workflows/publish.yml:78`).
- Python support, both Python binding crates, and both WASM crates remain
  outside crates.io publication. The carrier regression derives the exact
  seven publishable packages from all eleven inherited 0.13.0 manifests and
  separately proves all 16 explicit incubating manifests remain 0.10.0
  (`scripts/test_sprint_workflow.py:4964`,
  `scripts/test_sprint_workflow.py:5036`,
  `scripts/test_sprint_workflow.py:5109`).
- The published-shared proof passed in this review. It packaged
  `rdocx-layout@0.13.0`, inspected its normalized manifest for exact path-free
  `oxml-layout@0.10.0`, and resolved the packaged crate without an
  `oxml-layout` patch (`scripts/test_sprint_workflow.py:5142`,
  `scripts/test_sprint_workflow.py:5174`,
  `scripts/test_sprint_workflow.py:5182`,
  `scripts/test_sprint_workflow.py:5188`).
- The reviewed notes cover the complete M22 surface and name only the exact
  seven-package stable family. They state the shared 0.10.0 dependency and the
  binding publication exclusions (`CHANGELOG.md:11`, `CHANGELOG.md:43`,
  `CHANGELOG.md:46`, `CHANGELOG.md:48`). The notes render and validate
  deterministically.
- Git history from `v0.12.0` through the prepared HEAD has one authenticated
  author. The only new external GitHub record is Issue 68, an open question
  about a future JSON-like DSL rather than an implementation of any selected
  M22 change. Its exclusion agrees with the explicit empty selected-family
  inventory in the notes and HLD (`CHANGELOG.md:56`,
  `docs/hld/12-testing-strategy.md:1855`). No contribution notification is
  required.
- The HLD consistently distinguishes prepared 0.13.0 source from the latest
  published 0.12.0 stable family, retains shared 0.10.0 as the current source
  boundary, and reserves external mutation for `/release`
  (`docs/hld/03-architecture.md:850`, `docs/hld/10-bindings-spec.md:1156`,
  `docs/hld/15-build-and-toolchain.md:449`,
  `docs/hld/15-build-and-toolchain.md:479`). The stable tag remains absent
  locally and remotely. Separate final approval has not been given or implied.

## Prior finding closure

- Passes 1 through 7 remain closed. The full prefix retains fail-closed MHTML
  preflight for legacy background attributes, quote-aware CSS resources,
  string-form imports, escaped syntax, multiple resources, and unresolved or
  external resources (`crates/rdocx/src/html.rs:1488`,
  `crates/rdocx/src/html.rs:1885`). The unsafe and over-limit MHTML regression
  passed in this review.
- Passes 14 and 15 remain closed. The composed M22 test has independent
  predicates for rebuilt TOC bytes, sectioned merge boundaries, body
  comparison, header comparison, unsupported XML, package class, and VBA bytes
  (`crates/rdocx/tests/integration_test.rs:411`,
  `crates/rdocx/tests/integration_test.rs:445`,
  `crates/rdocx/tests/integration_test.rs:483`). It passed in this review.
  Inherited markup-compatibility value prefixes are materialized and the
  decoded part bound is rechecked (`crates/rdocx/src/flat_opc.rs:674`,
  `crates/rdocx/src/flat_opc.rs:731`, `crates/rdocx/src/flat_opc.rs:309`).
  Transitional and Strict alternative-format relationships still own opaque
  target classification on both import and export
  (`crates/rdocx/src/flat_opc.rs:416`,
  `crates/rdocx/src/flat_opc.rs:453`).
- F-X077 remains the sole strict lexical policy. The shared lexical matrix
  passed, and Flat OPC still calls the shared validator
  (`crates/oxml-core/src/xml.rs:37`, `crates/rdocx/src/flat_opc.rs:314`).
- F-X080 remains intact. Hosted CI enumerates all 24 bundled fonts and all six
  legal files (`.github/workflows/ci.yml:533`, `.github/workflows/ci.yml:567`).
  The authenticated Pandoc installer retains its exact digest, download and
  extraction bounds, two reviewed aliases, path checks, and fail-closed member
  policy (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:79`). Python still maps MHTML and package
  import failures into the existing public exception families
  (`crates/rdocx-py/src/lib.rs:66`).
- Passes 8 and 9 remain closed. The aggregate CI gate names only its filtered
  dependencies, while the release-regression and package jobs remain
  independent mandatory jobs (`.github/workflows/ci.yml:366`,
  `.github/workflows/ci.yml:518`, `.github/workflows/ci.yml:656`).
- F-X079 remains truthful across HLD and current source. The exact 15-package
  shared and PowerPoint family is published at 0.10.0, and stable 0.13.0 source
  pins that published boundary (`docs/hld/03-architecture.md:841`,
  `docs/hld/03-architecture.md:850`).

## Milestone gate

The M22 end gate requires one representative modern document to author and
render equations, rebuild fields and a table of contents, perform advanced
merge and comparison, inventory embedded content, and round-trip its modern
package variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`).

The functional gate holds. The source-built composed test exercises and
independently asserts every named operation, including deterministic equation
rendering, TOC replacement, field update, sectioned merge, embedded VBA
inventory, document and header comparison, DOTM identity, executable bytes,
and unsupported XML after Flat OPC reopen
(`crates/rdocx/tests/integration_test.rs:360`,
`crates/rdocx/tests/integration_test.rs:403`,
`crates/rdocx/tests/integration_test.rs:411`,
`crates/rdocx/tests/integration_test.rs:423`,
`crates/rdocx/tests/integration_test.rs:445`,
`crates/rdocx/tests/integration_test.rs:464`,
`crates/rdocx/tests/integration_test.rs:483`). It passed in this review and in
the exact-HEAD full gate.

The prepared-release checkpoint also holds technically. The current state
records `/verify --full` passed at the exact reviewed preparation HEAD with all
49 harness entries unchanged (`.claude/scratch/S69-run.json:286`). The full
97-test workflow module, release-note check, registry-only shared dependency
proof, hash harness, prose check, generated-skill check, and full-prefix diff
check passed independently in this review. S1 is the remaining record
reconciliation.

The sprint release definition of done is intentionally unfinished at this
boundary. F-X078 remains in progress, and actual publication, registry owners,
tag target, release-body bytes, exclusions, and any notification URLs remain
behind the separate `/release v0.13.0` approval
(`docs/sprints/CURRENT_SPRINT.md:39`,
`docs/hld/14-development-backlog.md:3783`).

## Not found

- **Interaction**: zero findings. The F-X078 carrier preparation preserves the
  completed F-X077, F-239, F-X080, F-X079, and F-238 contracts.
- **Duplication**: zero findings. The release preparation adds no runtime
  helper, and Flat OPC and MHTML continue to reuse their existing shared
  validators and package owners.
- **Layering**: zero findings. Metadata inspection found no `oxml-*` dependency
  on an `rdocx-*` or `rpptx-*` crate.
- **Harness**: zero findings. Exact-HEAD verification and the independent check
  both report all 49 entries unchanged (`.claude/scratch/S69-run.json:286`).
- **Gate**: zero technical findings. The M22 composite predicate and release
  preparation gates pass. S1 concerns the stale preparation record only.
- **Docs and ledgers**: one should-fix finding, S1. Zero additional findings.
  The live sprint, backlog, tracker, completion ledger, HLD, and sprint state
  otherwise agree that five stories are complete and F-X078 is in progress
  (`docs/sprints/CURRENT_SPRINT.md:34`, `docs/sprints/BACKLOG.md:532`,
  `docs/sprints/SPRINT_TRACKER.md:391`).
- **Dependencies**: zero findings. No unapproved external dependency or feature
  flag was added, and the shared and stable package trains remain separated.
- **Surface**: zero findings. F-X078 adds no runtime API. The prefix's native
  Flat OPC, package-class, MHTML, and lexical surfaces are all called for by
  their approved stories.
- **Limits and security**: zero findings. Resource preflight, archive limits,
  XML validation, opaque alternative-format handling, and atomic publication
  remain fail closed.
- **CI and release**: zero technical findings. Package inventory, Pandoc
  extraction, Python mapping, exact allowlists, carrier coverage, release-note
  truth, contribution exclusion, and approval separation remain coherent.
