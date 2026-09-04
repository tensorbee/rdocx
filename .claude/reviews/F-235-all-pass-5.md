# F-235, all, pass 5

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 3,708 insertions and 451 deletions
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, a significant inline-control raw boundary disables cross-run granularity

`crates/rdocx/src/comparison.rs:3574`
`crates/rdocx/src/comparison.rs:3616`
`crates/rdocx/src/comparison.rs:4700`

The pass 4 remediation aligns all direct inline-control runs as one attributed
sequence only when every modeled `w:sdtContent` child is a run. A significant
unchanged `SdtContent::RawXml` child makes that condition false. The fallback
then gives each direct run one complete-run signature before aligning the raw
child. For example, an original sequence containing one `alpha beta` run
followed by a foreign raw child and an edited sequence containing `alpha ` and
`beta` runs followed by the same raw child emits a deletion and insertion for
`beta` under both Word and Character policies. The visible text and the raw
child's logical boundary are unchanged. This leaves physical run segmentation
observable whenever an inline control also contains significant unmodelled
content, contrary to the attributed-unit and atomic non-text contract.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 4 D2 is remediated. `ignore_formatting` now selects the attributed path,
and the Run-granularity regression proves edited text accepts with the original
run properties. Pass 4 D3 is remediated by selecting one collision-free marker
from both staged inputs for the main story and each related part, then reusing
those names for accepted, edited, rejected, and original projections. The
inside-host collision regression passes. The run-only one-run and two-run
inline-control cases pass under Word and Character granularity.

No additional correctness, contract, panic, OOXML, public API, default
compatibility, ignore precedence, story filtering, raw preservation,
postcondition atomicity, exact-record gate, dependency, or structure finding
was established. The fixed policy matrix remains exact and independently
covers all seven story categories.

The complete regression binary passed with 285 tests passed and 3 ignored. The
three comparison unit tests and all four focused policy regressions passed.
`cargo clippy -p rdocx --all-targets -- -D warnings`, `cargo fmt --all --check`,
and `git diff --check` passed.
