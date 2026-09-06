# F-X081, Tag rpptx-v0.11.0

**Status**: approved
**Sprint**: S69
**Size**: S
**Depends on**: F-238, F-X079, F-X080

## Problem

The immutable v0.13.0 workflow published five stable packages, then failed
while verifying `rdocx`. The packaged facade resolved `oxml-opc@0.10.0` from
crates.io, but F-238 added the four `WORD_*` main content-type constants after
that shared release. The locally patched workspace dry run hid this registry
gap.

## Spec reference

- `docs/hld/03-architecture.md`, "Versioning".
- `docs/hld/10-bindings-spec.md`, "Packaging" and "WASM".
- `docs/hld/12-testing-strategy.md`, release regressions and registry proofs.
- `docs/hld/14-development-backlog.md`, "F-X081, Tag rpptx-v0.11.0".
- `docs/hld/15-build-and-toolchain.md`, "Publishing" and "Release process".
- `.claude/commands/release.md`, incubating publication and verification.

## Approach

Move the exact 15 publishable shared and PowerPoint packages, their workspace
pins, 16 lock records, README requirements, source assertions, CI literals,
workflow preflight, and unpublished `rpptx-wasm` carrier from 0.10.0 to 0.11.0.
Keep the stable workspace at 0.13.0 during this release. Prepare selected notes
covering only the additive `oxml-opc` Word package-class constants since
`rpptx-v0.10.0` and an evidence-derived contribution inventory.

After full verification and a clean sprint review at one exact SHA, stop for
the separate approval required by `/release rpptx-v0.11.0`. Complete only after
all 15 registry entries, owners, tag target, release body, stable exclusion,
absent `rpptx-wasm`, and applicable notifications verify.

## Rejected alternatives

- Publish only `oxml-opc@0.10.1`. The incubating family is a lockstep
  15-package contract, and the additive pre-1.0 API requires a minor boundary.
- Rerun v0.13.0. Duplicate versions would stop before the missing registry API
  is repaired.
- Move or delete v0.13.0. Its tag and published package bytes are immutable.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `test_incubating_release_family_is_prepared_at_0_11_0` | Every incubating carrier, pin, lock record, README, source assertion, CI literal, publication flag, and allowlist agrees at 0.11.0. |
| release notes | `test_release_notes_rpptx_v0_11_0_cover_word_package_constants` | The selected-family body covers only the additive `oxml-opc` constants and exact contribution inventory. |
| packaging | exact patched 22-package workspace dry run | Every archive verifies below 10 MiB with the complete asset inventories. |
| release | `/release rpptx-v0.11.0` | Fifteen registry entries, owners, tag SHA, release body, exclusions, and notifications verify. |

The **test gate is release**. Local preparation and the complete publication
evidence must both pass.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Release scripting and version strings**. Inspect every manifest, pin,
  lock record, README, CI literal, workflow preflight, test, and changelog
  carrier. Require `/verify --full` and separate approval before tagging.
- **Public API of a published crate**. Record the additive pre-1.0 minor
  boundary, run package dry runs, and enforce archive inventories and limits.
- **Crate dependency graph**. Prove the exact 15-package order and stable
  exclusion without adding a reverse family edge.
- **Bundled fonts and assets**. Verify the complete `oxml-layout` font and
  legal inventory, no stable duplication, and the `rpptx` template.
- **WASM or PyO3 bindings**. Keep bindings unpublished and run both WASM target
  graphs plus Python metadata checks.

## Hash harness

Expected unchanged across all 49 entries. Any delta blocks preparation.

## Implementation checklist

- [ ] Preserve and record the exact partial v0.13.0 external state.
- [ ] Move every incubating version carrier and stable shared dependency pin
      to 0.11.0.
- [ ] Update carrier, workflow, notes, and publication regressions.
- [ ] Build and verify the selected-family contribution inventory.
- [ ] Run `/verify --full` and every routed release check.
- [ ] Reach a clean microscope and sprint review at the exact prepared SHA.
- [ ] Stop at `/release rpptx-v0.11.0` for separate final approval.
- [ ] Verify all 15 registry entries and release evidence before completion.

## Open questions

None. The immutable partial-release precedent fixes the full shared-family
recovery and preserves a separate final publication approval.
