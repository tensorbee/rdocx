# F-007, correctness, pass 1

**Reviewed**: working-tree implementation diff, 3 files, 91 insertions and 7 deletions
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Not found

- Correctness: no wrong relationship resolution, target selection, duplicate
  relationship, or metadata update behavior found.
- Contract: the diff implements the approved OPC routing change and no XML
  parser or serializer expansion.
- Panics: no new panic path on package input found. Test-only unwraps assert
  fixture and result invariants.
- OOXML: the package-level relationship uses the required package namespace,
  writes an internal relative target for new documents, and preserves an
  existing custom target.
- Tests: the custom-target gate fails against the pre-feature behavior and
  proves load, mutation, in-place save, relationship retention, and absence of
  an orphaned conventional part.
- Structure: the diff adds no trait, generic parameter, wrapper, module, file,
  feature flag, or dependency edge.
