# F-X078, Tag v0.13.0

**Status**: approved
**Sprint**: S69
**Size**: S
**Depends on**: F-238, F-239, F-X077, F-X079

## Problem

The exact seven-package stable Word family is published at 0.12.0, while the
workspace carriers and release regressions still name that version in
`Cargo.toml:34`, `.github/workflows/publish.yml:24`, and
`scripts/test_sprint_workflow.py:4780`. S65 through S69 add the complete M22
Word-depth boundary, but `CHANGELOG.md:3` still says that no post-0.12.0
changes have been recorded.

F-X078 must prepare one coherent 0.13.0 stable family at the exact reviewed
M22 SHA, without publishing the shared or PowerPoint family and without
granting crates.io authority to Python or WASM packages. F-X077 changes a
published shared API, so F-X079 first publishes the required incubating family
and stable dependency pins rather than folding two families into this tag.

## Spec reference

- `docs/hld/03-architecture.md`, "Versioning".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability", "Packaging",
  and "WASM".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy", release regressions,
  registry proofs, and package checks.
- `docs/hld/14-development-backlog.md`, "F-X078, Tag v0.13.0".
- `docs/hld/15-build-and-toolchain.md`, "Publishing" and "Release process".
- `.claude/commands/release.md`, stable-family selection, approval,
  publication, verification, and notification requirements.

## Approach

After F-238, F-239, F-X077, and F-X079 are completed, confirm the reviewed
dependency graph resolves against the published shared 0.10.0 family. Move
every stable workspace version, stable workspace pin, lock record, README requirement,
source assertion, CI literal, Python project carrier, WASM carrier, workflow
preflight, and release regression from 0.12.0 to 0.13.0 in lockstep. Keep the
shared and PowerPoint family at its separately published 0.10.0 boundary and
keep every binding publication flag unchanged.

Prepare the `CHANGELOG.md` section headed `v0.13.0` from the reviewed stable
public API and behavior diff since `v0.12.0`. Cover the complete M22 Word-depth
surface, compatibility guidance, and every included authenticated external
issue or pull request. Classify each included contribution as direct or a
hardened equivalent, link it in the rendered notes, and prepare one
release-bound thank-you comment per record. An empty contribution inventory is
valid only when the reviewed Git history and delivery records prove it.

Publish exactly `rdocx-opc`, `rdocx-oxml`, `rdocx-layout`, `rdocx-html`,
`rdocx-pdf`, `rdocx`, and `rdocx-cli` in dependency order. Shared,
PowerPoint, Python, WASM, npm, and PyPI packages remain outside this story's
publication authority. After one clean full gate and sprint review at the
exact prepared SHA, stop for the separate final approval required by
`/release v0.13.0`. Complete the F-ID only after registry entries, owners, tag
target, release-body bytes, selected-family exclusions, and every applicable
notification URL verify.

No new runtime API, dependency, crate, module, file, feature, trait, generic,
wrapper, or builder is introduced. The public source changes from F-238 and
F-239 are additive pre-1.0 Rust surfaces whose compatibility notes are
recorded in the release body.

## Rejected alternatives

- Publish an incubating-family tag under the same story. The family namespaces
  and publication approvals are intentionally separate.
- Prepare 0.13.0 before the dependency diff is final. A later shared pin move
  would invalidate both the package graph and the release evidence.
- Publish Python, WASM, npm, or PyPI artifacts. This story grants no such
  authority.
- Reuse the 0.12.0 changelog body or omit contribution reconciliation. The tag
  must describe and credit the reviewed post-0.12.0 stable-family delta.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `test_stable_release_family_is_prepared_at_0_13_0` | Stable carriers, pins, lock entries, READMEs, binding metadata, publication flags, CI literals, and the allowlist agree at 0.13.0. |
| regression | published shared-family proof | Packaged `rdocx-layout@0.13.0` resolves the exact reviewed registry `oxml-layout` version without a local shared patch. |
| workflow | `python3 -m unittest scripts.test_sprint_workflow` | Family selection, order, notes, contribution inventory, approval, notification, and mutation authority remain pinned. |
| release notes | `release-notes v0.13.0 --check` and `--render` | One deterministic stable body describes only the selected family and directly links and credits every included contribution. |
| metadata | `cargo metadata --no-deps` | Exactly seven stable crates publish at 0.13.0 with the reviewed shared dependency pins. |
| packaging | exact patched 22-package workspace dry run | Every archive verifies below 10 MiB with complete required asset inventories. |
| integration | Python metadata assertions and stable WASM checks | Carriers track 0.13.0 without gaining publication authority. |
| release | `/release v0.13.0` | Seven registry entries, owners, annotated tag SHA, release body, exclusions, and applicable notification URLs verify. |

The **test gate is release**. Preparation and every local gate pass at one
reviewed SHA. Completion additionally requires separately approved real
publication, independent registry verification, byte-identical release notes,
and all required contribution notifications.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Release scripting and version strings**. Inspect every stable carrier,
  workflow literal, regression expectation, changelog section, and
  notification. Require `/verify --full` and separate immediate approval.
- **Public API of a published crate**. Record the additive pre-1.0 F-238 and
  F-239 surfaces, run the exact patched workspace dry run, inspect archives,
  and enforce the 10 MiB limits.
- **Crate dependency graph**. Inspect the complete dependency diff first.
  Prove the exact stable allowlist and registry-only resolution against the
  reviewed published shared-family boundary.
- **Bundled fonts and assets**. Verify the complete `oxml-layout` font and
  legal inventory, no duplicated fonts in `rdocx-layout`, and the PowerPoint
  template in the 22-package dry run.
- **WASM or PyO3 bindings**. Update metadata only, keep every binding
  unpublished, exclude both Python crates from workspace tests, and run the
  stable WASM target checks.

## Hash harness

Expected unchanged across the complete deterministic set. This story changes
release metadata only. Any output delta blocks preparation.

## Implementation checklist

- [x] Confirm F-238, F-239, F-X077, and F-X079 are completed at the reviewed
      M22 gate.
- [x] Confirm stable package resolution uses the published shared 0.10.0
      boundary from F-X079.
- [x] Move every stable version carrier and internal stable dependency pin to
      0.13.0.
- [x] Update stable carrier, workflow, registry, package, and notification
      regressions.
- [x] Build and verify the selected-family contribution inventory.
- [x] Prepare exact `v0.13.0` release notes and compatibility guidance.
- [ ] Run `/verify --full`, packaging, asset, binding, WASM, dependency,
      supply-chain, notes, and hash gates at one exact SHA.
- [ ] Stop at `/release v0.13.0` for separate final approval.
- [ ] Verify every publication, owner, tag, body, exclusion, and applicable
      notification before completion.

## Open questions

None. The story fixes the stable family, target version, publication
exclusions, dependency inspection, exact-SHA gate, and separate approval
boundary. The contribution inventory is derived from reviewed repository and
GitHub evidence during implementation rather than guessed during design.
