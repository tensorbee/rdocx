# F-012, Tag v0.4.1

**Status**: completed
**Sprint**: S02
**Size**: S
**Depends on**: F-003 through F-011

## Problem

M1 needs a known-good published state immediately before extraction begins.
The existing `scripts/release.sh` requires `main`, commits, tags, and pushes
directly. More importantly,
`/run-sprint` forbids tags and pushes other than `sprint/s02`, while
`/close-sprint` is the only tagging workflow and explicitly refuses release
tags. F-012 cannot satisfy its test gate under the current command authority.

## Spec reference

- `docs/hld/15-build-and-toolchain.md`, "Release process".
- `docs/hld/15-build-and-toolchain.md`, "Publishing".
- `docs/hld/14-development-backlog.md`, "Milestone 1, Preparation and safety net".

## Approach

Add a dedicated `/release VERSION` command and generated Codex adapter. It is
the only command allowed to create `v*` release tags or publish packages.
`/close-sprint` remains the only command allowed to update `main` or create an
`sNN` sprint tag. Update the shared workflow and `/run-sprint` so a release
story pauses after a clean full verification and sprint review, reports the
exact release command, and requires a separate final go/no-go immediately
before external mutation.

Prepare an exact 0.4.1 commit before those gates. `/release v0.4.1` verifies a
clean reviewed SHA, creates and pushes the tag, verifies all seven current
publishable crates plus the GitHub release, then permits F-012 to complete its
ledger-only finalization. It must not use the current error-swallowing
`--no-verify` publisher as proof of success.

F-012 is last after F-003 through F-011. Adding F-010 and F-011 to the formal
plan matches the approved S02 contract even though the backlog currently lists
only F-003 through F-009.

## Rejected alternatives

- Run `scripts/release.sh` from the sprint branch. It refuses non-main branches
  and its global version replacement is a documented risk.
- Treat sprint tag `s02` as the release tag. Sprint and release tags have
  different meanings and trigger different workflows.
- Mark the story complete after a dry-run only. The backlog gate explicitly
  requires the release tag to build and publish.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| clean-clone integration | `/verify --full` at the exact release SHA | Workspace, harness, docs, packaging, wasm, and supply chain are green before tagging |
| packaging | seven `cargo publish --dry-run` archives plus size and bundled-font contents | Every current publishable crate is releasable and licensed |
| release | exact `cargo info <crate>@0.4.1` and GitHub release inspection | The tag workflow published all seven crates and created the release without a swallowed failure |
| publication boundary | inspect the workflow allowlist | No `oxml-*` or `rpptx*` development package can be published by an rdocx release tag |

The backlog test gate is a clean-clone build and successful publication from
the exact v0.4.1 release tag.

## Release-line reconciliation

`v0.4.0` was published from `main` after S02 branched, with contract changes
and rendering fixes. S02 merges that exact release before publishing so the
new known-good boundary includes both lines of work. Publishing `v0.3.1` from
the older branch would omit the `v0.4.0` changes and would leave sprint closure
trying to reconcile a lower workspace version into `main`.

## HLD impact

- `docs/hld/11-migration-plan.md`
- `docs/hld/13-risks-and-open-questions.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- Release scripting and version strings. Inspect every Cargo.toml, Cargo.lock,
  and README change before tagging.
- Public API of published crates. Confirm additive F-008 semver impact and run
  all publish dry-runs plus the archive-size assertion.
- Bundled fonts. Inspect the rdocx-layout archive for every TTF, licence, and
  notice before publishing.

## Hash harness

Expected to remain unchanged at the exact release SHA and on the Linux CI
runner triggered for that commit.

## Implementation checklist

- [x] Implement the approved release authority without weakening `/run-sprint`
      or `/close-sprint` boundaries.
- [x] Add and validate the dedicated release command and generated adapter.
- [x] Confirm F-003 through F-011 are completed and integrated.
- [x] Merge the published v0.4.0 source and prepare the exact 0.4.1 version and
      internal dependency-pin diff.
- [x] Restrict the publication workflow to the seven released rdocx crates.
- [x] Run the full gate and release packaging checks from a clean clone.
- [x] Obtain explicit go/no-go immediately before the irreversible tag push.
- [x] Push v0.4.1 through the authorized mechanism and verify all seven crates
      plus the GitHub release.

## Open questions

None. Add the approved dedicated release command, and make F-012 depend on
F-003 through F-011. A separate final go/no-go remains mandatory before the
release tag or publication is attempted.
