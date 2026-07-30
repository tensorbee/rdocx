# F-010, correctness, pass 1

**Reviewed**: working diff, 3 files, 22 insertions and 11 deletions, plus the
14 exact-version and owner results recorded by the external gate
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Not found

- Correctness: all fourteen approved names resolve at version 0.0.0 and report
  `mantissaman (Atul Sharma)` as owner.
- Contract: the reserved set exactly matches the approved HLD 15 list, and the
  three non-crates.io names remain excluded.
- Panics: no runtime code changed.
- OOXML: no parser, serializer, namespace, child-order, or preservation code
  changed.
- Tests: every temporary package passed package inspection and publish dry-run.
  The exact-version gate failed before publication and passes after publication.
- Structure: no permanent crate, module, dependency, trait, generic, wrapper,
  feature flag, or source file was added to the repository.
