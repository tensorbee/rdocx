# F-235, all, pass 3

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 2,850 insertions and 202 deletions
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, policy is not applied to run content owned by an inline content control
`crates/rdocx/src/comparison.rs:3545`
`crates/rdocx/src/comparison.rs:3552`

The `SdtContent::Run` branch compares the legacy complete-run signature and
always writes a complete deletion plus insertion when text differs. It does not
use attributed units, the selected granularity, or the whitespace, field, and
comment ignore flags. For example, changing `alpha beta` to `alpha brave beta`
inside an unchanged inline `w:sdt` under Word granularity revises both complete
runs instead of only `brave `. A whitespace-only change in the same owner still
creates revisions even when whitespace is ignored. The policy is therefore not
threaded through all supported modeled content-control content.

### D2, the private text-box marker collides with preserved producer XML
`crates/rdocx/src/comparison.rs:968`
`crates/rdocx/src/comparison.rs:974`
`crates/rdocx/src/comparison.rs:1089`

Restoration identifies every element whose expanded name is
`{urn:rdocx:comparison:private}textBoxHost` as an internal marker. It accepts
more markers than were inserted, then replaces extra matches with the empty
string. A supported document can already contain that expanded name as an
unmodelled raw run child beside a real text box. Comparing with TextBox ignored
then replaces or deletes the producer child during restoration and fails the
staged projection instead of preserving the raw child exactly. Removing the old
text sentinel moved the collision into arbitrary producer XML rather than
eliminating it.

### D3, hyperlink raw-child boundaries still depend on physical run segmentation
`crates/rdocx/src/comparison.rs:2169`
`crates/rdocx/src/comparison.rs:2171`

The hyperlink shell projects `run_start` and `run_end` through attributed policy
boundaries, but compares `link.extra_xml` unchanged. Each entry in that field
contains a relative physical run boundary. If an unchanged hyperlink has a raw
child after the same logical text on both sides, splitting the preceding text
from one run into two changes only that stored boundary. Word or Character
comparison rejects the pair as a changed owner shell before granular alignment.
The segmentation remediation therefore covers plain hyperlinks but not the raw
content that the contract requires to remain significant and source-positioned.

### D4, the named matrix still does not freeze exact records for every policy
`crates/rdocx/tests/regression_test.rs:13770`
`crates/rdocx/tests/regression_test.rs:13792`
`crates/rdocx/tests/regression_test.rs:13815`

The whitespace, field, and comment cases assert only two kinds, order values,
story values, a shared owner, and an empty ignored result. They never fix the
record content or exact owner. The six ordinary story expectations assert only
the total count, then derive each ignored expectation by filtering records
produced by the implementation under test. A mutant that corrupts content or
ownership consistently in the ordinary and ignored calls remains green. The
gate is deterministic and independently suppresses all seven story categories,
but it is not the required fixed exact normalized record matrix for every
policy.

## Smells

None.

## Nitpicks

None.

## Not found

The TextBox mask now retains significant same-run host siblings and no longer
uses user text as its sentinel. Raw-significant attributed units remain in the
policy projections. Plain hyperlink run segmentation is normalized. Supported
unchanged content-control shells compare successfully, and the atomic failure
trigger is narrowly scoped to test builds after candidate staging. Duplicate
ignored-story options are validated before staging, and the story suppression
checks select Main, Header, Footer, Comment, TextBox, Footnote, and Endnote
independently.

The complete regression binary passed with 284 tests passed and 3 ignored. No
separate arithmetic or indexing panic was established in the reviewed paths.
No structure-rule or public-surface defect was found. The implementation adds
no dependency, trait, generic, feature flag, module, test binary, Python API,
WASM API, or CLI API.
