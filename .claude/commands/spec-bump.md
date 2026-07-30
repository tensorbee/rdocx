---
description: Version the spec set. Two stages, audit then apply, so a tagged spec version is never a snapshot of known drift.
---

# /spec-bump vX.Y [--apply]

Give `docs/hld/` a version number and a tag, so a spec can be cited by version
rather than by "whatever main said that day".

**Two stages, and the first is not optional.** A spec version that was tagged
over known drift is worse than no version, because it lends authority to a
document that does not describe the code.

## Stage 1, audit

```text
/spec-bump vX.Y
```

1. Refuse a dirty working tree.
2. Run `/audit-spec` in full.
3. Report the drift by bucket. **Any `contradiction`, or any `shape
   divergence`, stops the bump.** Those must be resolved by fixing the spec or
   fixing the code.
4. `missing` findings owned by a `pending` story do not block. The spec is
   allowed to describe work that is scheduled and not yet built. That is what
   the set is for. Record them in the bump commit message so a reader of the
   tag knows what was intentionally ahead of the code.
5. `undocumented` findings block, because they are the case where the code
   moved and the spec did not notice.

Stage 1 changes nothing. If it is clean enough, it prints the exact `--apply`
command.

## Stage 2, apply

```text
/spec-bump vX.Y --apply
```

Refuse unless stage 1 has just run clean against the current HEAD.

1. **Set the version** in `docs/hld/README.md`, under "Living status", as a
   `**Spec version**: vX.Y, YYYY-MM-DD` line. The first bump introduces that
   line, later bumps replace it.
2. **Update the set index** in the same file if documents were added, removed
   or renumbered since the last bump.
3. **Record what is intentionally ahead of the code.** The `missing` findings
   stage 1 allowed through, as a short list with their owning story IDs. A
   reader of the tag needs to know the difference between "not built yet" and
   "the spec lied".
4. **Commit once:**

   ```
   spec vX.Y, short summary of what changed in the set

   One paragraph: what moved in the set since the previous version, and what
   is intentionally ahead of the code.

   Tests, not applicable
   Harness, unchanged
   ```

5. **Tag** `spec-vX.Y`, annotated, with the same summary.
6. **Do not push.** Report the tag and the exact push command. Pushing a tag is
   a human decision here, exactly as it is in `/close-sprint`.

## Numbering

- `v0.x` while the set is pre-v1. A major bump signals a scope change, a minor
  bump signals revisions.
- The spec version is **not** the crate version. The Rust crates move on their
  own schedule through `/release`, after a version F-ID prepares and verifies
  the exact release commit.
- After the workspace reaches v1, follow semver on the set: patch for
  clarifications, minor for additive scope, major for a breaking change to a
  documented contract.

## When to bump

- At a milestone boundary, if the set changed meaningfully during it.
- After resolving an open question in `13-risks-and-open-questions.md` that
  moved scope.
- Before showing the set to anyone outside the project. A tag is citable, a
  moving branch is not.

## Refused situations

- **A dirty working tree.** Commit first.
- **Stage 2 without a clean stage 1 against the current HEAD.**
- **Unresolved `contradiction` or `shape divergence` findings.**
- **A version that already exists as a tag**, or one that is not strictly
  greater than the current. Versions go forward.
- **Pushing the tag.** Report the command. The human runs it.
- **Bumping the crate version.** Different thing, different script.
