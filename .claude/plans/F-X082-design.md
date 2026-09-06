# F-X082, Tag v0.13.1

**Status**: approved
**Sprint**: S69
**Size**: S
**Depends on**: F-238, F-239, F-X077, F-X081

## Problem

Five stable packages already exist immutably at 0.13.0, while `rdocx` and
`rdocx-cli` do not. After F-X081 publishes the missing shared contract, one
new stable patch family must publish all seven packages from one reviewed SHA.

## Spec reference

- `docs/hld/03-architecture.md`, "Versioning".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability",
  "Packaging", and "WASM".
- `docs/hld/12-testing-strategy.md`, release regressions and registry proofs.
- `docs/hld/14-development-backlog.md`, "F-X082, Tag v0.13.1".
- `docs/hld/15-build-and-toolchain.md`, "Publishing" and "Release process".
- `.claude/commands/release.md`, stable publication and verification.

## Approach

After F-X081 verifies every shared 0.11.0 registry entry, move the workspace
and all stable carriers from 0.13.0 to 0.13.1. Pin every shared dependency to
0.11.0. Update lock records, READMEs, source and CI literals, Python and WASM
metadata, workflow preflights, release regressions, and selected stable notes.

Add a registry-only proof that packages `rdocx` with local stable dependencies
but no shared patches, then compiles the normalized package against published
shared 0.11.0. This directly covers the v0.13.0 failure. Stop for separate
approval at `/release v0.13.1` only after full verification and clean review.

## Rejected alternatives

- Publish only `rdocx` and `rdocx-cli` at 0.13.0. That would mix source and
  dependency boundaries inside one declared stable family.
- Reuse or move v0.13.0. Published versions and release tags are immutable.
- Fold the shared release into this tag. Each family has its own package set
  and approval boundary.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `test_stable_release_family_is_prepared_at_0_13_1` | Stable carriers, pins, lock records, READMEs, binding metadata, CI literals, and allowlist agree at 0.13.1 while shared is 0.11.0. |
| regression | `test_prepared_rdocx_0_13_1_requires_published_shared_0_11_0` | Packaged `rdocx` compiles against registry-only shared crates and exercises the constants missing from `oxml-opc@0.10.0`. |
| release notes | `test_release_notes_v0_13_1_cover_recovery_and_word_outcomes` | Notes cover the complete M22 result, immutable partial v0.13.0 attempt, compatibility, and exact inventory. |
| packaging | exact patched 22-package workspace dry run | Every archive verifies below 10 MiB with required assets. |
| release | `/release v0.13.1` | Seven registry entries, owners, tag SHA, body, exclusions, and notifications verify. |

The **test gate is release**. Local preparation, registry-only compilation,
and complete publication evidence must all pass.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Release scripting and version strings**. Inspect all stable and shared
  carriers, release notes, tests, and workflow literals. Require `/verify
  --full` and separate approval before tagging.
- **Public API of a published crate**. Verify the recovery graph against the
  actual registry, run package dry runs, and enforce archive limits.
- **Crate dependency graph**. Package and compile `rdocx` without shared local
  patches, and prove the exact seven-package allowlist.
- **Bundled fonts and assets**. Verify shared font and legal inventories, no
  stable font duplication, and the `rpptx` template.
- **WASM or PyO3 bindings**. Move metadata only, retain publication exclusions,
  and run both WASM and Python binding checks.

## Hash harness

Expected unchanged across all 49 entries. Any delta blocks preparation.

## Implementation checklist

- [x] Verify F-X081 publication and the immutable v0.13.0 partial state.
- [x] Move every stable carrier to 0.13.1 and shared pin to 0.11.0.
- [x] Add the registry-only `rdocx` package verification gate.
- [x] Prepare exact recovery notes and contribution inventory.
- [x] Run `/verify --full` and every routed release check.
- [ ] Reach a clean microscope and sprint review at the exact prepared SHA.
- [ ] Stop at `/release v0.13.1` for separate final approval.
- [ ] Verify all seven registry entries and release evidence before completion.

## Open questions

None. The immutable partial release requires a complete stable patch family
after the separately published shared recovery.
