# F-X080, all aspects, pass 1

**Reviewed**: working-tree diff, 7 files and 245 changed lines, comprising 217
additions and 28 deletions
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, the non-file rejection claim includes accepted directories

`.claude/plans/F-X080-design.md:47`
`scripts/install_pinned_pandoc.py:85`

The amended contract says every other non-file entry is rejected, but the
extractor intentionally accepts directories before rejecting unsupported
member types. The same overbroad wording appears in all three HLD updates. A
valid archive directory is therefore a counterexample to the documented test
predicate. Narrow the prose to unsupported non-file member types while keeping
the exact alias exception and directory handling unchanged.

## Smells

None.

## Nitpicks

None.

## Not found

No other correctness, contract, panic, archive-safety, test-strength, or
structure findings. The inventory remains explicit, the 160 MiB arithmetic is
correct, exact alias name and target matching is fail closed, hardlinks and
special entries still reject, no version or dependency carrier changed, and
the Python adapter remains on its established generic public error surface.
