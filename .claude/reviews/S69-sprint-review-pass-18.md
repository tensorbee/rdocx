# S69 sprint review, pass 18

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`572875ab1648cf22844955740333f201b0f2799a` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 115 files and 11,926 changed
lines, comprising 10,712 additions and 1,214 deletions. The 26 crate
directories with changed files are `oxml-chart`, `oxml-cli-support`,
`oxml-core`, `oxml-drawing`, `oxml-layout`, `oxml-media`, `oxml-opc`,
`oxml-pdf`, `oxml-sml`, `rdocx`, `rdocx-cli`, `rdocx-html`, `rdocx-layout`,
`rdocx-opc`, `rdocx-oxml`, `rdocx-pdf`, `rdocx-py`, `rdocx-wasm`, `rpptx`,
`rpptx-chart`, `rpptx-cli`, `rpptx-layout`, `rpptx-oxml`, `rpptx-py`,
`rpptx-render`, and `rpptx-wasm`. The `oxml-py-support` package also changes
effective version through workspace inheritance.

**Pass authority**: pass 18 is the scheduled F-X078 prepared-release
remediation boundary. The user explicitly extended the global pass number
because the preceding dependency-prefix boundaries were clean and requested
review of the pass-17 remediation before release.

**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Pass-17 closure

Pass-17 S1 is closed. The F-X078 design now records the completed full gate,
including its packaging, asset, binding, WASM, dependency, supply-chain,
release-note, and hash riders (`.claude/plans/F-X078-design.md:139`). The
progress record identifies the exact verified preparation SHA, reports all 49
hash entries unchanged, confirms no publication has begun, and directs the
next operator to reconcile the clean review and repeat full verification at
that review commit (`.claude/scratch/F-X078-progress.md:17`,
`.claude/scratch/F-X078-progress.md:24`). The sprint state independently binds
the successful full gate to
`da5009f5516876d1d17caf4822b3a12f823de486`
(`.claude/scratch/S69-run.json:286`).

The two commits after that verified preparation SHA change only the pass-17
review artifact and the completed design checkbox. They do not change source,
package manifests, release notes, CI, or release carriers. The workflow still
requires the clean-review commit, review record, and repeated full gate at the
resulting exact HEAD before `/release` (`.claude/commands/run-sprint.md:207`).
That required post-review sequence is pending by design and is not a sprint
defect.

## F-X078 stable v0.13.0 preparation

- The workspace version is 0.13.0. Stable-group internal pins are 0.13.0,
  while shared OOXML and PowerPoint dependencies retain their published 0.10.0
  boundary (`Cargo.toml:34`, `Cargo.toml:55`, `Cargo.toml:71`). Metadata
  inspection found exactly seven publishable 0.13.0 packages:
  `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`, `rdocx-pdf`,
  `rdocx`, and `rdocx-cli`.
- The publication workflow names those same seven packages in dependency order
  and no others (`.github/workflows/publish.yml:61`). Its separately routed
  incubating allowlist remains the exact 15-package shared and PowerPoint
  family at 0.10.0 (`.github/workflows/publish.yml:78`).
- `oxml-py-support`, `rdocx-py`, `rpptx-py`, `rdocx-wasm`, and `rpptx-wasm`
  remain excluded from crates.io publication
  (`crates/oxml-py-support/Cargo.toml:6`,
  `crates/rdocx-py/Cargo.toml:6`, `crates/rpptx-py/Cargo.toml:6`,
  `crates/rdocx-wasm/Cargo.toml:14`, `crates/rpptx-wasm/Cargo.toml:14`). The
  carrier regression derives the exact publishable subset from the inherited
  0.13.0 manifests and separately proves the incubating manifests remain
  0.10.0 (`scripts/test_sprint_workflow.py:4964`,
  `scripts/test_sprint_workflow.py:5036`,
  `scripts/test_sprint_workflow.py:5109`).
- The workflow proves the shared dependency boundary through metadata,
  packaging, and registry-only resolution before publication
  (`.github/workflows/publish.yml:23`, `.github/workflows/publish.yml:35`). The
  focused published-shared regression passed and verified a path-free
  `oxml-layout@0.10.0` dependency in packaged `rdocx-layout@0.13.0`
  (`scripts/test_sprint_workflow.py:5142`,
  `scripts/test_sprint_workflow.py:5174`,
  `scripts/test_sprint_workflow.py:5188`).
- The v0.13.0 notes cover the M22 Word depth, field and TOC, merge and
  comparison, embedded-content, package identity, Flat OPC, MHTML, and shared
  lexical work (`CHANGELOG.md:11`, `CHANGELOG.md:19`). They name the exact
  seven-package stable family, the shared 0.10.0 dependency boundary, and the
  binding and WASM exclusions (`CHANGELOG.md:43`, `CHANGELOG.md:46`,
  `CHANGELOG.md:48`). The notes check and deterministic render passed.
- The selected-family contribution inventory is empty and explicitly recorded
  (`CHANGELOG.md:56`, `docs/hld/12-testing-strategy.md:1855`). Git history has
  one authenticated author after `v0.12.0`. The only newer external issue is a
  request about a future JSON-like DSL rather than implementation of a selected
  v0.13 change. No pull request or other selected-family contribution was
  omitted.
- The local and remote `v0.13.0` tags are absent. Registry probes found none of
  the seven selected packages at 0.13.0. Publication, ownership checks, tag,
  release-body verification, and any applicable notifications therefore
  remain correctly unchecked behind separate final approval
  (`.claude/plans/F-X078-design.md:141`,
  `docs/hld/14-development-backlog.md:3783`).

## Integrated prefix and milestone gate

The M22 gate requires a representative modern document to author and render an
equation, rebuild fields and a table of contents, perform advanced merge and
comparison, inventory embedded content, and round-trip the modern package
without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`). The composed test independently
asserts equation rendering, TOC replacement, field update, sectioned merge,
embedded VBA inventory, document and header comparison, DOTM identity,
executable bytes, and unsupported XML after Flat OPC reopen
(`crates/rdocx/tests/integration_test.rs:360`,
`crates/rdocx/tests/integration_test.rs:403`,
`crates/rdocx/tests/integration_test.rs:411`,
`crates/rdocx/tests/integration_test.rs:423`,
`crates/rdocx/tests/integration_test.rs:445`,
`crates/rdocx/tests/integration_test.rs:464`,
`crates/rdocx/tests/integration_test.rs:483`). It passed at the recorded exact
preparation SHA. No executable code changed between that gate and this review.

Prior closures remain intact:

- F-X077 remains the single strict lexical policy, and Flat OPC continues to
  call the shared validator (`crates/oxml-core/src/xml.rs:37`,
  `crates/rdocx/src/flat_opc.rs:314`).
- F-239 retains fail-closed MHTML preflight for legacy background attributes,
  quote-aware CSS resources, string-form imports, escapes, multiple resources,
  and unresolved or external resources (`crates/rdocx/src/html.rs:1488`,
  `crates/rdocx/src/html.rs:1885`).
- F-238 still materializes inherited markup-compatibility value prefixes,
  rechecks the decoded part bound, and preserves opaque transitional and Strict
  alternative-format relationships (`crates/rdocx/src/flat_opc.rs:309`,
  `crates/rdocx/src/flat_opc.rs:416`,
  `crates/rdocx/src/flat_opc.rs:674`,
  `crates/rdocx/src/flat_opc.rs:731`).
- F-X080 hosted CI enumerates all 24 bundled fonts and all six legal files
  (`.github/workflows/ci.yml:533`, `.github/workflows/ci.yml:567`). Its release
  regressions and independent package job remain mandatory inputs to the
  aggregate gate (`.github/workflows/ci.yml:366`,
  `.github/workflows/ci.yml:518`, `.github/workflows/ci.yml:656`). The pinned
  Pandoc installer retains its exact digest, download and extraction bounds,
  reviewed aliases, and fail-closed archive-member policy
  (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:79`).
- F-X079 remains truthful across current source and HLD. The shared and
  PowerPoint family is published at 0.10.0, stable 0.13.0 source pins that
  boundary, and the latest published stable family remains 0.12.0
  (`docs/hld/03-architecture.md:841`,
  `docs/hld/03-architecture.md:850`,
  `docs/hld/10-bindings-spec.md:1156`).

The full 97-test sprint workflow module passed in this review, including the
release-carrier, note-truth, selected-family, F-X080, and mutation regressions.
The registry-only shared dependency proof also passed. The release-note check,
metadata inventory, prose gate, generated-skill gate, and full-prefix diff
check passed. The recorded full verification reports all 49 hash entries
unchanged (`.claude/scratch/S69-run.json:287`).

## HLD, ledgers, and state

The HLD consistently distinguishes prepared stable 0.13.0 source from the
latest published stable 0.12.0 family, retains the published shared 0.10.0
boundary, records the exact package sets and exclusions, and reserves all
external mutation for `/release` (`docs/hld/03-architecture.md:850`,
`docs/hld/10-bindings-spec.md:1156`,
`docs/hld/15-build-and-toolchain.md:449`,
`docs/hld/15-build-and-toolchain.md:486`). The current sprint, backlog, and
state correctly keep F-X078 in progress while the five non-release stories are
complete (`docs/sprints/CURRENT_SPRINT.md:34`,
`docs/sprints/CURRENT_SPRINT.md:39`, `docs/sprints/BACKLOG.md:532`,
`docs/sprints/SPRINT_TRACKER.md:391`). The unchecked publication evidence is
consistent with the absent tag and registry versions.

## Not found

- **Interaction**: zero findings. F-X078 preparation and pass-17 remediation do
  not alter the completed F-X077, F-238, F-239, F-X079, or F-X080 behavior.
- **Duplication**: zero findings. The release preparation adds no runtime
  helper, policy copy, or second package-family inventory.
- **Layering**: zero findings. Metadata inspection found no forbidden
  `oxml-*` dependency on an `rdocx-*` or `rpptx-*` crate.
- **Harness**: zero findings. The exact preparation gate records all 49 entries
  unchanged, and no executable or baseline file changed afterward
  (`.claude/scratch/S69-run.json:286`).
- **Gate**: zero findings. The composed M22 predicate, package inventories,
  release-note check, shared dependency proof, and recorded full gate pass.
- **Docs and ledgers**: zero findings. The design, progress record, HLD, sprint
  trackers, completion ledger, and state agree with the prepared but
  unpublished boundary.
- **Dependencies**: zero findings. The stable 0.13.0 and shared 0.10.0 trains
  remain separated, with exact pins and no unapproved dependency or feature
  flag.
- **Surface**: zero findings. F-X078 adds no runtime API. The prefix's Flat OPC,
  package-class, MHTML, and lexical surfaces remain within their approved
  stories.
- **Limits and security**: zero findings. Resource preflight, archive limits,
  XML validation, opaque alternative-format preservation, and atomic package
  writing remain fail closed.
- **CI and release**: zero findings. Hosted CI fixes, exact allowlists,
  publication exclusions, carrier coverage, release-note truth, contribution
  inventory, tag and registry absence, and approval separation are coherent.

## Remaining release procedure

This clean audit does not authorize publication. After this artifact is
committed alone, the integrator must record the clean review at that exact
commit and rerun `/verify --full` there. Only `/release v0.13.0` may then seek
the separate final approval immediately before its first external mutation
(`.claude/commands/run-sprint.md:207`, `.claude/commands/run-sprint.md:210`).
