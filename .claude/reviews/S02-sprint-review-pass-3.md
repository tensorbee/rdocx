# S02 sprint review, pass 3

**Reviewed**: `sprint/s02` at `5f9dd716170b` against `main` at
`5f25df0bde1e`, 67 files, 2,827 changed lines, crates: `rdocx-opc`,
`rdocx-oxml`, and `rdocx`

**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

This pass was required after the published `v0.4.0` mainline was merged into
S02 and F-012 was retargeted to `v0.4.1`. It supersedes pass 2 for the release
candidate without rewriting that historical review.

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

The gate requires the workspace tests, a deterministic hash baseline that
reproduces on another machine, and the release tag and publication.

The full gate passed at the reviewed SHA. The workspace test suite, Clippy,
no-default-font path, WASM check, rustdoc, prose check, adapter drift check,
package dry-run, archive inspection, and supply-chain check passed. The hash
harness matched all 28 entries. All eight workspace packages inherit 0.4.1,
with the six internal dependency pins matching it at `Cargo.toml:15` and
`Cargo.toml:27`. The WASM package remains unpublished at
`crates/rdocx-wasm/Cargo.toml:13`.

The dry-run produced exactly seven archives. All were below 10 MiB, and the
`rdocx-layout` archive contained 20 TTF files plus the required licence and
notice files. Real publication is restricted to the seven rdocx crates by
`.github/workflows/publish.yml:23` and is covered by the regression assertion
at `scripts/test_sprint_workflow.py:59`. The reserved `oxml-*` and `rpptx*`
packages cannot be selected by that workflow.

The final external part of the gate remains intentionally pending. F-012 stays
in progress at `docs/sprints/CURRENT_SPRINT.md:35` until `/release v0.4.1`
receives its separate final approval and verifies all seven registry entries
and the GitHub release.

## Not found

- **Interaction**: the `v0.4.0` contract and rendering changes compile and pass
  together with the S02 relationship, setter, cache, and unit work.
- **Duplication**: no competing helper or second release path was introduced.
- **Layering**: no `oxml-*` or `rpptx*` crate exists in this workspace, and no
  dependency crossed the prohibited family boundary.
- **Harness**: S02 has no delta relative to the merged `main` baseline, and all
  28 entries match.
- **Docs**: the sprint contract names `v0.4.1`, records the post-close forward
  merge at `docs/sprints/CURRENT_SPRINT.md:45`, and records the publication
  embargo at `docs/hld/15-build-and-toolchain.md:126`.
- **Dependencies**: only workspace version pins changed. No external dependency
  was added.
- **Surface**: the public surface remains limited to the approved F-007 and
  F-008 additions reviewed in the feature passes.
