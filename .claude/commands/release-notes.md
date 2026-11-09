---
description: Prepare and validate meaningful reviewed release notes in CHANGELOG.md for one unified Rust and Python family tag.
---

# /release-notes {vX.Y.Z | rpptx-vX.Y.Z}

Prepare the human-written release record that `/release` publishes unchanged.
This ceremony edits only `CHANGELOG.md`, including the matching section and
any necessary `Unreleased` cleanup. It does not create a tag, push, publish, or
create a GitHub release.

## Inputs

Choose exactly one complete family from the requested tag:

- `vX.Y.Z` is the stable rdocx crates, CLI and Python distribution.
- `rpptx-vX.Y.Z` is the incubating OOXML and PowerPoint crates, CLI and
  Python distribution.

Historical `py-*` changelog sections remain readable, but no new Python-only
release tag is prepared.

Find the current sprint story assigned to the exact requested tag. Read its
design plan and dependencies before collecting claims. Refuse a tag that has no release preparation story in the active sprint or
whose plan is not approved.

## Evidence

Build the notes from reviewed repository evidence, not memory:

1. Find the previous tag in the selected family. Use the repository root when
   the family has no previous tag.
2. Read the release story and design plan, completed `AS_BUILT.md` entries in
   the selected commit range, and the reviewed commits in that range.
3. Inspect merged pull requests and every GitHub issue or pull request named by
   the release story, design plan, completed delivery records, or reviewed
   commits in the range. Include an unmerged pull request when its reported
   behavior or reference implementation landed through a hardened equivalent.
   Use the GitHub record for its authenticated author and user-facing purpose.
   Do not infer contributor identity from an unauthenticated display name or
   commit trailer.
4. Read the current compatibility and migration contract in `CHANGELOG.md`,
   the release plan, and the relevant HLD sections.
5. Separate changes for the selected family from changes for the other
   families. Family notes cover the selected Rust, CLI and Python surfaces together,
   excluding the other family and npm packages. Exclude internal workflow work unless it changes a user-visible
   release or compatibility promise.
6. Build one contribution inventory for the selected family. For each included
   GitHub record, capture its URL, authenticated author handle, user-visible
   outcome, whether it landed directly or through a hardened equivalent, and
   current open or closed state. Record why any candidate belongs only to the
   other family. This inventory is part of the release report, not a new
   tracked file.

Every highlight, addition, fix, compatibility statement, and contributor must
be supported by that evidence. Refuse unverifiable claims or credit instead of
inventing them.

The contribution inventory and the notes must reconcile exactly. Every
included issue and pull request appears as a direct Markdown link in the
release section, normally beside the addition or fix it supports. Every
authenticated external reporter or contributor appears in `Contributors` with
the specific outcome they helped shape. Do not collapse verified contributions
into a generic community thank-you or credit an issue number without its
author.

## Write the reviewed section

Add exactly one second-level heading whose text is the requested tag. Keep the
sections below in this order and give every section meaningful text:

```markdown
## v0.8.0

### Highlights

Describe the most important user outcome.

### Added

Describe user-visible capabilities.

### Fixed

Describe user-visible fixes, or state that there are no user-facing fixes.

### Compatibility

State required migration action, supported compatibility, or that no action is
required.

### Contributors

Credit verified contributors and maintainers whose work is in this release,
with direct links to the relevant issues or pull requests.
```

Do not copy the example prose. Do not use placeholders such as `TBD`, `TODO`,
`FIXME`, `CHANGEME`, `PLACEHOLDER`, or `???`. Keep the existing `Unreleased`
material accurate after moving claims into the versioned section.

Write for users rather than reproducing a commit log. Highlights state the
release outcome. Added and Fixed name observable behavior. Compatibility says
what a caller must change or explicitly says that no action is required.
Contributor credit explains why each linked external record matters.

## Validate and review

Run both deterministic modes:

```bash
python3 scripts/sprint_workflow.py release-notes <requested-tag> --check
python3 scripts/sprint_workflow.py release-notes <requested-tag> --render
```

The rendered output is the exact GitHub release body. Inspect it alongside the
release plan, the commit range, the pull requests, the contributor evidence,
and the `CHANGELOG.md` diff. Run the prose gate and generated-skill drift gate
before handing the release story to review.

Compare the rendered body against the contribution inventory. Require every
included record link and every verified external handle to survive rendering.
Prepare the record-specific notification list that `/release` will use after a
successful publication. Each planned comment names the final tag, links the
GitHub release, states whether the work landed directly or through a hardened
equivalent, and thanks the authenticated contributor. Do not post these
comments during release-note preparation.

Report the selected family, previous tag or repository root, evidence range,
included issues and pull requests, authenticated contributors, direct versus
hardened-equivalent classification, notification list, compatibility
conclusion, and the successful check and render commands.

## Refused situations

- The tag is not exactly `vX.Y.Z`, `rpptx-vX.Y.Z`,
  `py-rdocx-vX.Y.Z`, or `py-rpptx-vX.Y.Z`.
- The active sprint has no approved release story for the exact tag.
- The versioned heading is missing, duplicated, incomplete, empty, or contains
  a placeholder.
- A note belongs only to the other release family.
- An included issue or pull request is absent from the rendered notes, lacks a
  direct link, or has no verified author classification.
- A verified external reporter or contributor is missing from `Contributors`
  or is credited without the specific included outcome.
- A product claim or contributor identity cannot be verified from the reviewed
  record.
- The requested action would tag, push, publish, or create a release. Continue
  through `/release` after review and full sprint verification.

## Compatibility evidence claims

Inspect the selected family's authenticated strict SemVer evidence and actual
compiler witnesses before describing compatibility. An initial unpublished
minor's empty compatible line and an all-skipped migration call do not prove
old-line compatibility. Preserve published notes verbatim. Issue429 recovery
and any optional old-line yank remain separate, explicitly qualified actions.
Do not claim issue430 fixed or included until its actual acceptance succeeds.
Use actual final-source CI and installed artifact outcomes, with pending native
or hosted checks stated honestly.
