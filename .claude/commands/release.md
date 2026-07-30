---
description: Release an already prepared and reviewed workspace version. The only command that creates and pushes v* release tags or starts crates.io publication.
---

# /release vX.Y.Z

Release the exact reviewed sprint SHA. This is the only command allowed to
create or push a `v*` release tag, or to start publication of the lockstep Rust
crate family. It never merges to `main` and never creates an `sNN` sprint tag.

The version bump is prepared and committed through its F-ID before this command
runs. This command does not edit versions, create a release commit, or repair a
red gate.

## Preconditions

Refuse before any tag or push if one check fails:

1. The argument is exactly `vX.Y.Z`, and `[workspace.package].version` is
   exactly `X.Y.Z`.
2. The current branch is the active `sprint/sNN` branch and the tree is clean.
3. The release F-ID named `Tag vX.Y.Z` is `reviewed` in the sprint run state,
   remains `in-progress` in both delivery trackers, and every dependency is
   completed.
4. The latest recorded `/verify --full` passed at the current HEAD with the
   declared hash-harness result.
5. The latest recorded `/sprint-review SNN` is clean at the current HEAD, and
   its review file reports zero blocking findings.
6. `cargo publish --workspace --dry-run` passes from the clean tree and produces
   exactly the seven packages listed below. A dry-run uploads nothing. Every
   archive is below 10 MiB. The `rdocx-layout` archive contains all bundled TTF
   files, `LICENSE-Caladea`, `NOTICE-Caladea`, and the OFL licence.
7. The seven publishable packages are exactly `rdocx-opc`, `rdocx-oxml`,
   `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and `rdocx-cli`, all at
   `X.Y.Z`. `rdocx-wasm` may inherit the workspace version, but remains
   `publish = false` and is not a crates.io package. The workflow contains an
   explicit allowlist for those seven packages. It must not publish an
   `oxml-*` or `rpptx*` package while PowerPoint development is incomplete.
8. The tag is absent locally and from `origin`. Fetch the remote tag namespace
   before deciding. Refuse a conflicting or already-published version rather
   than treating it as success.

## Final approval

Report the exact HEAD SHA, tag, seven crates, version, remote, and workflow that
will run. Ask for a separate explicit go or no-go immediately before the first
external mutation. Approval given earlier in the feature or sprint does not
count at this boundary.

## Release

After approval, preserve this order:

1. Push the active `sprint/sNN` branch at the reviewed HEAD.
2. Create annotated tag `vX.Y.Z` at that exact HEAD with message
   `Release vX.Y.Z`.
3. Push only that tag. The tag starts `.github/workflows/publish.yml`, which
   publishes the seven crates with verification and creates the GitHub release.
4. Watch the workflow through completion. A failed job is a failed release.
   Do not rerun blindly and do not convert an error into an "already published"
   success.
5. Verify `cargo info <crate>@X.Y.Z` for all seven packages, verify the owner,
   and inspect the GitHub release tag and target SHA.

If branch push succeeds but tag push fails, report that exact state. If the tag
push succeeds but publication fails, retain the tag and report the failed crate
and workflow. Do not delete or move a published release tag.

## Finalise the release F-ID

Only after all seven registry versions and the GitHub release are verified:

1. Create the F-ID's `AS_BUILT.md` entry with the release evidence.
2. Complete its sprint tracker and backlog records, clear its owner, and set
   its design plan to completed.
3. Record the release F-ID completed in sprint state.
4. Re-run the sprint's ledger checks and continue `/run-sprint` to its final
   review and `/close-sprint` handoff.

## Refused situations

- A version bump or uncommitted change is still required.
- Verification or sprint review covers a different SHA.
- A local dry-run is offered as a substitute for successful publication.
- Any command would merge to `main` or create an `sNN` tag.
- The user has not given the separate final approval.
