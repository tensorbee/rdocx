# S02 sprint review, pass 1

**Reviewed**: `main...sprint/s02` at `ab1c4cd`, 64 files, 2,710 changed
lines, crates: `rdocx`, `rdocx-opc`, `rdocx-oxml`
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, the release tag does not reproduce the hash baseline on CI

`.github/workflows/publish.yml:19`

The `v*` workflow proceeds directly from checkout to crates.io publication.
The normal hash-harness job does not cover tag or sprint-branch pushes because
`.github/workflows/ci.yml:3` limits push CI to `main`. The M1 gate requires the
baseline to reproduce on a second machine before the release is accepted, and
the F-012 plan expects the Linux runner at the release SHA to prove it. Add the
hash harness to the tag workflow before `cargo publish --workspace`, so a
runner mismatch blocks publication rather than being discovered after the
known-good release is live.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

The workspace tests and local hash harness pass, and the clean-clone package
gate verified all seven 0.3.1 archives. The gate does not yet hold because the
second-machine harness is absent from the release workflow and v0.3.1 has not
been tagged or published.

## Not found

No additional interaction, duplication, layering, documentation, dependency or
public-surface findings. The six completed stories compose without an
unexplained output delta, and all fourteen placeholder reservations remain
separate from the seven 0.3.1 release packages.
