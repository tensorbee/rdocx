# F-229, all aspects, pass 1

**Reviewed**: uncommitted worker diff, 13 files, 2,268 additions and 10 deletions
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, math settings do not invalidate reusable layout context
`crates/rdocx-layout/src/engine.rs:543`

`ReusableEngineContext` does not retain `LayoutInput.math_properties`, and its
comparison at line 678 does not compare the new field. A document whose
OfficeMath settings change can therefore reuse paragraph layout produced with
the old font, spacing, margins, justification, or limit placement.

### D2, the golden does not connect recorded Word geometry to Rust geometry
`crates/rdocx-layout/src/math.rs:1602`

The test compares Rust output with constants duplicated from the manifest, but
only checks that the manifest contains the raw Word page and equation bounds.
Neither the test nor the harness derives the per-expression normalized values
from the Word PDF. Replacing the normalized manifest geometry with arbitrary
Rust output leaves both the Word validation and Rust golden green, so the named
test does not prove agreement with the external oracle.

### D3, the raster mutation test does not exercise the rendered page
`crates/rdocx-layout/src/math.rs:1666`

The SSIM assertions compare two synthetic uniform byte arrays. They do not
rasterize the OfficeMath page or perturb its baseline, delimiter, or operator.
The test can remain green if the actual PDF or PNG output changes arbitrarily,
so it does not satisfy the mutation-sensitive deterministic page golden in the
approved plan.

### D4, Word token validation accepts non-exact output
`scripts/officemath_oracle_harness.py:107`

The manifest declares exact Word PDF tokens, but `is_subsequence` accepts any
number of inserted or duplicated tokens. A Word export that gains visible
unexpected equation text therefore passes the oracle gate.

## Smells

None.

## Nitpicks

None.

## Not found

No additional findings in correctness, contract, panics, OOXML preservation,
tests, or structure.
