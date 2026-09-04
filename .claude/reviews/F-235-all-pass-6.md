# F-235, all, pass 6

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 3,770 insertions and 450 deletions
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, inline raw-boundary identity includes preceding text values

`crates/rdocx/src/comparison.rs:3580`
`crates/rdocx/src/comparison.rs:3716`
`crates/rdocx/tests/regression_test.rs:13636`

The remediation admits direct raw children to full-slice alignment only when
each raw key has the same complete sequence of preceding attributed-unit
signatures. That sequence includes text values, not only the raw child's
logical boundary. Replacing `alpha` with same-length `omega` before an
unchanged end-boundary raw child therefore makes the keys unequal even though
the raw bytes and Word or Character boundary position are unchanged. If the
edited control also splits `omega beta` across two runs, comparison falls back
to per-child alignment and revises the unchanged `beta` unit in addition to
the real replacement. The new regression covers physical segmentation with
identical visible text, but not the supported combination of a granular text
edit and segmentation change at an unchanged significant raw boundary. Raw
identity should constrain its bytes and logical position without making all
preceding visible values part of the owner shell.

## Smells

None.

## Nitpicks

None.

## Not found

The pass 5 unchanged-text case is remediated. Direct significant raw content
with matching keys stays on one attributed sequence, and the focused Word and
Character cases preserve it exactly once in tracked, accepted, and rejected
views. Raw-byte differences make the keys unequal and remain significant.

No additional correctness, contract, panic, OOXML, public API, default
compatibility, ignore precedence, story filtering, source preservation,
postcondition atomicity, exact-record gate, dependency, or structure finding
was established. The shared TextBox markers, hyperlink logical boundaries,
left-biased ignore policies, fixed policy records, and independent seven-story
suppression remain intact.

The complete regression binary passed with 285 tests passed and 3 ignored. The
three comparison unit tests and focused inline-control and policy-matrix tests
passed. `cargo clippy -p rdocx --all-targets -- -D warnings`,
`cargo fmt --all --check`, and `git diff --check` passed.
