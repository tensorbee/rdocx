# F-010, Reserve crate names

**Status**: completed
**Sprint**: S02
**Size**: S
**Depends on**: none

## Problem

The future extraction depends on crates.io names that are not yet controlled by
the project. Publishing a minimal 0.0.0 placeholder permanently reserves a
name, so the exact list, owner identity, package contents, and stop conditions
must be agreed before any external mutation.

## Spec reference

- `docs/hld/03-architecture.md`, "The target crate graph".
- `docs/hld/15-build-and-toolchain.md`, "Publishing".
- `docs/hld/13-risks-and-open-questions.md`, "Package-name squatting".

## Approach

Create minimal no-dependency 0.0.0 placeholder crates in a temporary directory,
outside the workspace and without adding permanent repository crates. Each
contains only its manifest, licence, README, and a documented empty library or
binary target. Inspect and dry-run every archive, obtain a final explicit
go/no-go for the complete list, then publish one at a time and stop at the first
failure. Confirm `cargo info <name>@0.0.0` and the intended owner after each
publication.

## Rejected alternatives

- Add placeholders to the workspace. That would pollute the real dependency
  graph with crates that have no implementation yet.
- Test unversioned `cargo info`. A squatted or later version could make that
  pass without proving this project published 0.0.0.
- Publish the whole list in parallel. Partial failure would make ownership and
  recovery harder to reason about.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| packaging | `cargo package --list` and `cargo publish --dry-run` per placeholder | Every archive is minimal, licensed, dependency-free, and publishable |
| external gate | `cargo info <name>@0.0.0` | Every approved exact name resolves at the reserved version and has the intended owner |

The backlog test gate is exact-version `cargo info` success for every approved
name.

## HLD impact

- `docs/hld/13-risks-and-open-questions.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- New crate and file. Explicit approval is required for the temporary
  placeholder packages and their exact names before creation.
- Public API of a published crate. The placeholders expose no usable API. Run
  package inspection, publish dry-runs, and archive-size checks.
- Release scripting and version strings. Inspect every 0.0.0 manifest and
  publish command before the irreversible crates.io action.

## Hash harness

Expected to remain unchanged. Temporary placeholder packages do not touch the
workspace or sample outputs.

## Implementation checklist

- [x] Resolve the exact crates.io name list and owner identity.
- [x] Confirm credentials can create every approved name without exposing them.
- [x] Create minimal temporary placeholder packages outside the workspace.
- [x] Inspect package lists, licences, manifests, and dry-run archives.
- [x] Obtain final approval listing every irreversible publication.
- [x] Publish sequentially, verify exact version and ownership, and stop on the
      first failure.
- [x] Update the HLD to distinguish reserved crates.io names from unchecked
      PyPI or npm names.

## Open questions

None. Reserve the approved 14 names in the HLD 15 publishing graph for the
existing crates.io owner `mantissaman` (Atul Sharma): `oxml-core`, `oxml-opc`,
`oxml-media`, `oxml-drawing`, `oxml-layout`, `oxml-pdf`, `oxml-sml`,
`oxml-cli-support`, `rpptx-oxml`, `rpptx-layout`, `rpptx-render`, `rpptx-chart`,
`rpptx`, and `rpptx-cli`. Exclude `oxml-py-support`, `rpptx-wasm`, and
`rpptx-py` because their documented distribution channels are not crates.io.
Require a separate final go/no-go immediately before publication.
