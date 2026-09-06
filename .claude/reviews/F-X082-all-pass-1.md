# F-X082, all aspects, pass 1

**Reviewed**: the 26-file working diff, 305 insertions and 88 deletions
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Not found

- Correctness: all eleven inherited stable carriers and nine internal stable
  pins move together to 0.13.1, while the shared and PowerPoint family remains
  at 0.11.0 in `Cargo.toml:34-78`.
- Contract: the exact seven-package publication set, unpublished binding and
  WASM carriers, README requirements, CI literal, and release notes are pinned
  by `scripts/test_sprint_workflow.py:4991-5163`.
- Registry isolation: the opt-in gate packages normalized local stable crates,
  checks that their shared requirements contain no path, compiles a fresh
  consumer, and proves each shared package source is the registry at 0.11.0 in
  `scripts/test_sprint_workflow.py:5173-5320`.
- Workflow safety: the stable preflight and registry proof run before package
  publication, retain the stable-only condition, and preserve the exact
  allowlist in `.github/workflows/publish.yml:20-86`.
- Release truth: the notes identify the immutable partial v0.13.0 attempt,
  complete M22 surface, exact stable inventory, shared 0.11.0 boundary, and
  empty contribution inventory in `CHANGELOG.md:7-49`.
- Panics: the new regression code handles subprocess failure through explicit
  assertions and operates only on locally generated normalized archives in
  `scripts/test_sprint_workflow.py:5182-5319`.
- Structure: no trait, generic parameter, feature, crate, module, or production
  runtime surface is introduced.
- HLD discipline: all five files named by the approved plan describe the
  prepared 0.13.1 state and retain immutable v0.13.0 evidence.
- Tests: the three named gates were observed failing against the 0.13.0
  carriers and missing notes, then passing after preparation. The registry-only
  proof also passed with a fresh Cargo home against crates.io shared 0.11.0.
- Output stability: the hash harness reports all 49 entries unchanged.
