# F-012, all, pass 1

**Reviewed**: `c7cea6c..0582a2e`, 30 files, 508 changed lines
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, run-sprint completes the release story before the release gate

`.claude/commands/run-sprint.md:160`

Section 7 still directs the orchestrator to finalise every integrated F-ID,
mark every plan completed and mark every feature completed before sprint review.
That includes the release F-ID, even though the new exception requires it to
stay reviewed until `/release` verifies publication. The exception at line 198
therefore becomes unreachable under the command's own earlier steps. Exclude a
real-publication release F-ID from section 7, retain its in-progress ledgers and
approved plan through the first sprint review, then create its AS_BUILT and
tracker records only after `/release` succeeds.

## Smells

None.

## Nitpicks

None.

## Not found

No additional correctness, contract, panic, OOXML, test or structure findings.
The manifest pins are lockstep, `rdocx-wasm` remains unpublished, publication
verification fails closed, and release, sprint and spec tag authorities remain
separate.
