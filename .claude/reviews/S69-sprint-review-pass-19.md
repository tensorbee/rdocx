# S69 sprint review, pass 19

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`0b74db9414a90f2294c3660a4db910b1297e40ab` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 119 files and 12,510 changed
lines, comprising 11,287 additions and 1,223 deletions. The 26 crate
directories with changed files are `oxml-chart`, `oxml-cli-support`,
`oxml-core`, `oxml-drawing`, `oxml-layout`, `oxml-media`, `oxml-opc`,
`oxml-pdf`, `oxml-sml`, `rdocx`, `rdocx-cli`, `rdocx-html`, `rdocx-layout`,
`rdocx-opc`, `rdocx-oxml`, `rdocx-pdf`, `rdocx-py`, `rdocx-wasm`, `rpptx`,
`rpptx-chart`, `rpptx-cli`, `rpptx-layout`, `rpptx-oxml`, `rpptx-py`,
`rpptx-render`, and `rpptx-wasm`. The `oxml-py-support` package also changes
effective stable and shared dependency versions through workspace inheritance.

**Pass authority**: pass 19 is the scheduled F-X081 release-recovery
dependency-prefix boundary. The user explicitly requested as many passes as
required and approved continuing the recovery after the immutable partial
v0.13.0 attempt. This records the extension beyond the default global review
bound before the separate final publication approval.

**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## F-X081 recovery preparation

- The workspace stable version remains 0.13.0 while the exact shared and
  PowerPoint dependency set is prepared at 0.11.0 (`Cargo.toml:34`,
  `Cargo.toml:55`, `Cargo.toml:65`). The carrier regression derives the exact
  15 publishable packages, the unpublished `rpptx-wasm` preparation member,
  all manifests, workspace pins, lock records, README examples, source
  assertions, and the CI literal (`scripts/test_sprint_workflow.py:5470`,
  `scripts/test_sprint_workflow.py:5488`,
  `scripts/test_sprint_workflow.py:5500`,
  `scripts/test_sprint_workflow.py:5513`,
  `scripts/test_sprint_workflow.py:5523`,
  `scripts/test_sprint_workflow.py:5545`,
  `scripts/test_sprint_workflow.py:5577`).
- The incubating publication allowlist contains those same 15 packages in
  dependency order and excludes stable crates and `rpptx-wasm`
  (`.github/workflows/publish.yml:78`, `.github/workflows/publish.yml:81`,
  `.github/workflows/publish.yml:99`, `.github/workflows/publish.yml:109`).
  The stable allowlist remains separately guarded by the `v*` namespace
  (`.github/workflows/publish.yml:61`).
- The selected release notes describe only the four additive Word main
  content-type constants since `rpptx-v0.10.0`, the 15-package compatibility
  boundary, stable exclusion, and the empty evidence-derived contribution
  inventory (`CHANGELOG.md:7`, `CHANGELOG.md:18`, `CHANGELOG.md:30`,
  `CHANGELOG.md:36`, `CHANGELOG.md:39`). The mutation-sensitive notes test
  checks these claims and rejects issue or pull-request links
  (`scripts/test_sprint_workflow.py:5639`).
- The actual selected source delta contains the four public constants and
  their existing relationship-vocabulary regression
  (`crates/oxml-opc/src/content_types.rs:22`,
  `crates/oxml-opc/src/relationship.rs:429`). It adds no dependency edge,
  trait, feature, parser, serializer, or stable runtime surface.
- The HLD and backlog consistently record current prepared source, latest
  published coherent families, the immutable partial v0.13.0 attempt, and the
  required order from shared 0.11.0 to stable 0.13.1
  (`docs/hld/03-architecture.md:841`,
  `docs/hld/10-bindings-spec.md:1156`,
  `docs/hld/12-testing-strategy.md:1855`,
  `docs/hld/14-development-backlog.md:3795`,
  `docs/hld/15-build-and-toolchain.md:474`). Issue 69 is accurately excluded
  from both recovery releases as a separate performance follow-up
  (`docs/hld/14-development-backlog.md:3784`).

The exact clean-tree full verification is recorded passed at the reviewed SHA
with all 49 deterministic hash entries unchanged
(`.claude/scratch/S69-run.json:326`). The observed gate includes workspace
formatting, all-target all-feature Clippy, changed-crate and workspace tests,
the no-default font path, both WASM graphs, rustdoc, README inventories,
the exact patched 22-package dry run, archive-size checks, and `cargo deny`.
The packaged `rdocx@0.13.0` compiled successfully against the reviewed local
shared 0.11.0 graph, directly exercising the boundary that stopped the
immutable v0.13.0 publication. Both Python binding graphs also compile.

## Milestone gate

The M22 gate requires one representative modern document to author and render
equations, rebuild fields and a table of contents, perform advanced merge and
comparison, inventory embedded content, and round-trip a modern package
variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`). The composed source-built test
asserts equation rendering, TOC cache replacement, field evaluation, sectioned
merge, VBA inventory, full-story comparison, DOTM identity, executable bytes,
and unsupported XML after Flat OPC reopen
(`crates/rdocx/tests/integration_test.rs:361`,
`crates/rdocx/tests/integration_test.rs:403`,
`crates/rdocx/tests/integration_test.rs:411`,
`crates/rdocx/tests/integration_test.rs:423`,
`crates/rdocx/tests/integration_test.rs:445`,
`crates/rdocx/tests/integration_test.rs:464`,
`crates/rdocx/tests/integration_test.rs:468`,
`crates/rdocx/tests/integration_test.rs:483`). It passed in the exact-HEAD
workspace gate. F-X081 changes version and release carriers only, so it does
not weaken that behavioral evidence.

## Prior closure and interaction audit

All prior pass-18 closures remain intact. F-X077 remains the sole shared XML
lexical validator. F-238 still uses it for Flat OPC and preserves package class,
opaque alternative-format parts, inherited namespaces, signatures, and binary
payloads. F-239 retains bounded fail-closed MHTML resource preflight and stable
loss diagnostics. F-X080 retains the repaired font inventory, bounded Pandoc
installer, and Python adapter gate. F-X079 remains the immutable published
0.10.0 shared-family boundary. F-X081 changes no executable implementation in
those features.

The immutable `v0.13.0` tag and its five published low-level stable crates are
recorded without reuse, movement, deletion, or an unauthorized rerun
(`docs/hld/14-development-backlog.md:3772`,
`docs/hld/14-development-backlog.md:3781`). F-X082 remains ordered after
successful shared publication and requires a new registry-only packaged
`rdocx` proof against shared 0.11.0
(`docs/hld/14-development-backlog.md:3816`,
`docs/hld/14-development-backlog.md:3824`). No external mutation is authorized
by this review.

## Not found

- **Interaction**: zero findings. The shared 0.11.0 recovery supplies the API
  required by F-238 without changing the stable package version or completed
  M22 behavior.
- **Duplication**: zero findings. F-X081 adds no runtime helper, second package
  inventory, or alternate release path.
- **Layering**: zero findings. No `oxml-*` crate gains an `rdocx-*` or
  `rpptx-*` dependency.
- **Harness**: zero findings. All 49 entries are recorded unchanged at the
  exact reviewed SHA, and the baseline is untouched.
- **Gate**: zero findings. The composed M22 gate and the release carrier,
  notes, packaging, binding, archive, and supply-chain gates pass.
- **Docs**: zero findings. The five planned HLD files, backlog, current sprint,
  and release notes agree on prepared and published state.
- **Dependencies**: zero findings. The complete 0.11.0 family moves in lockstep
  with no new third-party dependency or forbidden reverse edge.
- **Surface**: zero findings. The release recovery exposes only the four
  already reviewed additive `oxml-opc` constants.
- **Release safety**: zero findings. Stable publication, bindings, WASM, npm,
  PyPI, tag creation, pushing, and registry publication remain outside this
  review and behind separate final approval.

## Remaining release procedure

This clean audit does not authorize publication. After this artifact is
committed alone, the integrator must record pass 19 with the explicit review
extension and repeat `/verify --full` at that review commit. Only then may
`/release rpptx-v0.11.0` perform final registry and tag preflight and request
the separate approval immediately before its first external mutation.
