# F-X079, Tag rpptx-v0.10.0

**Status**: approved
**Sprint**: S69
**Size**: S
**Depends on**: F-X077, F-X080

## Problem

F-X077 adds the shared strict XML 1.0 validator to the published `oxml-core`
API. The current exact 15-package incubating family remains at 0.9.0, so the
stable 0.13.0 packages cannot depend on that new registry API until one
coherent shared-family release exists.

The release must publish the exact incubating family at 0.10.0, update stable
source pins to that published boundary, and exclude stable Word, binding,
WASM, Python, npm, and PyPI packages. It must derive contribution credit from
the selected-family diff rather than borrowing stable-only outcomes.

## Spec reference

- `docs/hld/03-architecture.md`, "Versioning" and package-family boundaries.
- `docs/hld/10-bindings-spec.md`, "Packaging" and "WASM".
- `docs/hld/12-testing-strategy.md`, release regressions and registry-family
  proofs.
- `docs/hld/14-development-backlog.md`, "F-X079, Tag rpptx-v0.10.0".
- `docs/hld/15-build-and-toolchain.md`, "Publishing" and "Release process".
- `.claude/commands/release.md`, incubating selection, approval, publication,
  verification, and notification requirements.

## Approach

After F-X077 and F-X080 complete, move the exact 15 publishable shared and PowerPoint
manifests, workspace pins, lock records, README requirements, CI literals,
source assertions, release regression, and unpublished `rpptx-wasm`
preparation carrier from 0.9.0 to 0.10.0. Update stable source dependency pins
to the published shared 0.10.0 boundary without changing the stable workspace
version or granting stable publication authority.

Prepare a reviewed `rpptx-v0.10.0` changelog section covering the shared
lexical-validator ownership and any other selected-family changes since
`rpptx-v0.9.0`. Build the exact selected-family contribution inventory from
repository and GitHub evidence. Link and credit each included external record,
classify direct versus hardened-equivalent outcomes, and prepare the required
release-bound comments. An empty inventory is valid only if the reviewed diff
proves it.

Publish exactly `oxml-core`, `oxml-opc`, `oxml-media`, `oxml-layout`,
`oxml-drawing`, `oxml-pdf`, `oxml-sml`, `oxml-cli-support`, `oxml-chart`,
`rpptx-oxml`, `rpptx-chart`, `rpptx-layout`, `rpptx-render`, `rpptx`, and
`rpptx-cli` in dependency order. After a clean exact-HEAD full verification and
sprint review, stop for the separate final approval required by
`/release rpptx-v0.10.0`. Complete only after all 15 registry entries and
owners, the tag SHA, release body, stable exclusion, absent `rpptx-wasm`, and
applicable notifications verify.

## Rejected alternatives

- Publish only `oxml-core`. The incubating family is a lockstep 15-package
  contract.
- Let stable 0.13.0 depend on unpublished workspace code. Registry consumers
  must resolve the new shared API from crates.io.
- Publish the stable family under this tag. It has its own F-X078 gate and
  approval.
- Publish bindings, WASM, npm, Python, or PyPI packages. This story grants no
  such authority.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `test_incubating_release_family_is_prepared_at_0_10_0` | Exact manifests, pins, lock entries, READMEs, source and CI literals, publication flags, and allowlist agree at 0.10.0. |
| regression | stable carrier isolation regression | Stable packages remain at 0.12.0 while shared dependency pins move to 0.10.0 and bindings remain unpublished. |
| workflow | `python3 -m unittest scripts.test_sprint_workflow` | Family selection, order, notes, contribution inventory, approval, and mutation authority remain pinned. |
| release notes | `release-notes rpptx-v0.10.0 --check` and `--render` | One deterministic incubating-family body describes only the selected diff and exact contribution inventory. |
| packaging | exact patched 22-package workspace dry run | Every archive verifies below 10 MiB with complete font, legal, ICC, and template inventories. |
| release | `/release rpptx-v0.10.0` | Fifteen registry entries, owners, tag SHA, release body, stable exclusion, absent `rpptx-wasm`, and applicable notifications verify. |

The **test gate is release**. Preparation and every local gate pass at one
reviewed SHA. Completion additionally requires separately approved real
publication and independent registry, owner, tag, release-body, exclusion, and
notification verification.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Release scripting and version strings**. Inspect every incubating carrier,
  workflow literal, regression expectation, changelog section, contribution,
  notification, and lock entry. Require `/verify --full` and separate approval
  before tagging.
- **Public API of a published crate**. Record the additive `oxml-core` API,
  run the exact patched workspace dry run, inspect archives, and enforce the
  10 MiB limits.
- **Crate dependency graph**. Prove the exact 15-package publication order and
  the stable source dependency pins while excluding stable publication.
- **Bundled fonts and assets**. Verify the complete `oxml-layout` font and
  legal inventory, no duplicated fonts in `rdocx-layout`, and
  `rpptx/assets/default.pptx`.
- **WASM or PyO3 bindings**. Keep every binding unpublished, exclude both
  Python crates from workspace tests, and run both WASM target graphs.

## Hash harness

Expected unchanged across the complete deterministic set. This story changes
release metadata only. Any output delta blocks preparation.

## Implementation checklist

- [ ] Complete and verify F-X077.
- [ ] Complete and verify F-X080 so hosted CI release gates are current.
- [ ] Move every incubating version carrier and stable shared dependency pin
      to 0.10.0.
- [ ] Update exact carrier, isolation, workflow, package, and release-note
      regressions.
- [ ] Build and verify the selected-family contribution inventory.
- [ ] Prepare exact `rpptx-v0.10.0` release notes and compatibility guidance.
- [ ] Run `/verify --full`, packaging, assets, bindings, WASM, dependency,
      supply-chain, notes, and hash gates at one exact SHA.
- [ ] Stop at `/release rpptx-v0.10.0` for separate final approval.
- [ ] Verify every publication, owner, tag, body, exclusion, and applicable
      notification before completion.

## Open questions

None. S69 already requires a separate incubating-family release when the
shared validator changes a published shared API. Real publication retains its
own final go or no-go immediately before external mutation.
