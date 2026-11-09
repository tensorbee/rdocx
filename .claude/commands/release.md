---
description: Release an already prepared and reviewed Rust and Python package family from a closed sprint main merge. The only command that creates and pushes v* or rpptx-v* release tags or starts registry publication.
---

# /release {vX.Y.Z | rpptx-vX.Y.Z}

Release one complete Rust, CLI and Python family from the exact `main` merge
commit produced by `/close-sprint`. This command alone may create or push a
release tag or start crates.io and PyPI publication. It never edits versions,
merges branches, or creates a sprint tag. Historical `py-*` tags remain
immutable and readable, but cannot start a new publication.

## Family contract

For `vX.Y.Z`, the selected crates.io set is exactly `rdocx-opc`,
`rdocx-oxml`, `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and
`rdocx-cli`. The selected Python project is `rdocx` from `rdocx-py`.
For `rpptx-vX.Y.Z`, the selected crates.io set is exactly `oxml-core`,
`oxml-opc`, `oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`,
`oxml-sml`, `oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`,
`rpptx-layout`, `rpptx-render`, `rpptx`, and `rpptx-cli`. The selected
Python project is `rpptx` from `rpptx-py`. The exact selected Rust manifests,
workspace dependency pins when present, binding crate and `pyproject.toml`
version must match `X.Y.Z`. Binding and WASM crates are not published to
crates.io. Each release contains six selected CLI archives, six `cp39-abi3`
wheels, one source distribution and one `SHA256SUMS` covering all thirteen
payloads.

## Preconditions

Refuse before any tag or push if one check fails:

1. The argument is exactly `vX.Y.Z` or `rpptx-vX.Y.Z`, with no suffix or
   extra prefix. Run `python3 scripts/sprint_workflow.py
   unified-release-family <requested-tag>` and inspect all selected version
   carriers and exact package allowlists.
2. The current branch is clean `main`, at the same SHA as `origin/main`. The
   sprint that prepared this release has already closed through
   `/close-sprint`, and its `sNN` tag points to this exact merge commit. The
   selected release preparation F-ID is `done` in the sprint delivery records
   and its plan is completed. All its dependencies are done.
3. The committed sprint review is clean at the reviewed sprint SHA and the
   `/close-sprint` report records the merge. Compare the merge commit tree with
   that reviewed sprint tree. Only the reviewed close-sprint tracker summary
   may differ. Refuse a material source, release-workflow, command, changelog,
   manifest, lockfile, or HLD difference. Do not run `/sprint-review` on
   `main`, because that command writes tracked review files.
4. Consume the closed sprint's authenticated full verification at producer H
   and its validated completion/review metadata at R when this clean main
   merge M has exactly R's Git tree and all relevant tools, oracle inputs,
   configuration and retained artifacts remain unchanged. Use the existing
   `closure_evidence_problems(load(sprint), reviewed_R)` check and step3's
   actual main/review tree comparison. Current clean main, origin/main and
   sprint tag identities remain required. Record H, R and M truthfully. Do
   not claim the commands ran at M or repeat them merely because merging
   changed the commit SHA. No new cache service or command bridge is needed.
   A latest failed full attempt blocks reuse of an older green one. Missing
   evidence, a material tree delta or changed inputs requires the missing or
   invalidated gates. Both requested families reuse this common applicable
   full evidence. Fresh registry, publisher and approval checks remain due.
5. Run `python3 scripts/sprint_workflow.py release-notes <requested-tag>
   --check`, render the exact `CHANGELOG.md` body, and inspect the family-only
   claims, compatibility guidance, included issue and pull-request links, and
   authenticated contributor credit. Rebuild the selected contribution
   inventory from reviewed evidence.
6. Consume the actual locally patched publication dry run included in step4's
   applicable full evidence, with its retained verified archives and exact
   22-package union below10MiB. Do not execute it a second time. Confirm the
   selected family's crates.io versions and PyPI version are freshly absent.
   A duplicate or partial prior publication is a failed release.
7. Inspect `.github/workflows/wheels.yml`. Its tag path must select exactly
   one Rust allowlist, one CLI binary, and one Python distribution, validate
   six CLI archives and seven Python artifacts, verify their provenance,
   publish through the `pypi` environment and create one GitHub release only
   after both registries succeed. The PyPI trusted publisher for the selected
   project must identify this repository, `wheels.yml`, and environment
   `pypi`, with no long-lived token.
8. Fetch the remote tag namespaces. The requested tag must be absent locally
   and on `origin`. Confirm the selected crates.io and PyPI owner or
   maintainer roles that will be checked after publication.

The reviewed release preparation includes a manually dispatched `wheels.yml`
build-only run on the prepared source SHA. It builds both Python projects,
uploads artifacts, and creates no tag, registry file or GitHub release. Its
selected six wheels and source distribution were inspected for exact name,
version, Markdown description, summary, author, keywords, classifiers,
project links and `cp39-abi3` target tags. They were installed and tested in
clean Python 3.9 and 3.12 environments, with strict typing and stub checks
under Python 3.12. Recheck that evidence against the main tree comparison.

## Final approval

Report the exact clean `main` SHA, sprint tag and reviewed sprint SHA,
requested release tag, selected crate and Python distribution sets, fourteen
GitHub assets, selected version, remote, `wheels.yml` workflow, trusted
publisher identity, absent registry versions, verified owner roles and the
rendered notes from `CHANGELOG.md`. Report every included issue and pull
request URL, authenticated contributor and planned release-bound comment.
Ask for a separate explicit go or no-go immediately before the first external
mutation. Earlier feature or sprint approval does not count.

## Release

After approval, preserve this order:

1. Create one annotated tag for the requested argument at the exact reviewed
   `main` SHA with message `Release <requested-tag>`.
2. Push only that tag. It starts `.github/workflows/wheels.yml`. The tag
   selects the matching Rust, CLI and Python family. The workflow builds,
   validates and attests the thirteen payloads, verifies every downloaded
   subject, writes and checks complete `SHA256SUMS`, then publishes the selected
   crates and Python distribution. Manual dispatch cannot reach either
   registry or the GitHub release job.
3. Watch the workflow through completion. A failed job is a failed release.
   Do not rerun blindly or convert authentication, network, compilation,
   duplicate-version, attestation, asset or registry failures into success.
   If one registry succeeds and the other fails, retain the immutable tag,
   report the exact partial state, and do not create a GitHub release.
4. Verify `cargo info <package>@X.Y.Z` for every selected crate, every owner,
   the selected exact PyPI version and all seven Python files. Reinstall the
   PyPI version in clean Python 3.9 and 3.12 environments, rerun the priority
   runtime suites, typing and stub checks, and verify PyPI roles.
5. Download all fourteen GitHub release assets. Check the exact inventory,
   `sha256sum --check SHA256SUMS`, and `gh attestation verify FILE -R
   tensorbee/rdocx` for each CLI archive, wheel and source distribution.
   Inspect the release target SHA. Fetch its body and compare it byte for byte
   with a fresh rendered reviewed changelog body from the tagged SHA.
6. Notify each included GitHub issue and pull request only after publication
   and release-body verification succeed. Each comment links the tag and
   release, names the included outcome, states direct or hardened-equivalent
   provenance, and thanks the authenticated reporter or contributor. Record
   every resulting comment URL in the release report. Do not change a record's
   state unless separately authorized.

Do not commit a post-release ledger edit to `main` under this command. The
completed sprint delivery records describe release preparation. The tag,
registries, GitHub release, attestation checks and notification URLs are the
publication evidence. If the tag push succeeds but publication fails, retain
the tag and report the failed job. Never delete or move a published release
tag. If a notification fails, retain the release and report that exact record.

## Refused situations

- A dirty or unreviewed `main` tree, missing close-sprint provenance, or a
  material merge-tree difference.
- An absent full gate, unapproved requested tag, duplicate version, missing
  artifact, incomplete attestation, or mismatched publisher identity.
- A request to merge, move an existing tag, publish another family, or edit
  the prepared release source during this command.

## Strict SemVer publication preflight

Before the separate final SHA approval or any tag, obtain fresh completed CI
and signed native authority through the same reviewed private collector used by
`/verify --full`, selecting the requested family. Require actual final-source
CI gate and run success with no later relevant failed attempt. A successful
wheels rehearsal cannot replace CI. Preserve truthful H, R and M identities and
require current main proof when the publication boundary requires it.

```bash
python3 scripts/sprint_workflow.py release-semver TAG --check-evidence DIR [DIR ...] --release-candidate
```

This consumes the authenticated default/all native suite without rerunning it,
refreshes the selected-family complete crates.io and PyPI inventories, requires
all selected candidates absent and strictly above every stable same-line
version including yanked versions. An occupied version in the other family is
not a duplicate selected candidate. Equal published versions remain mandatory
API baselines during ordinary CI. Retain the independent owner/publisher,
remote-tag, exact package, archive-size, installed artifact and final approval
checks. No SemVer receipt authorizes a yank, publication, message or tag.
