# F-235, all, pass 4

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 3,306 insertions and 289 deletions
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, granular alignment stops at each direct inline-control run

`crates/rdocx/src/comparison.rs:3502`
`crates/rdocx/src/comparison.rs:3653`
`crates/rdocx/src/comparison.rs:4547`

The pass 3 remediation applies attributed splitting only after the outer
content-control alignment has paired one `SdtContent::Run` with one other run.
That outer alignment still assigns one complete-run signature to each direct
run child, and the helper then builds attributed units from only that single
pair. An unchanged inline control whose text is represented by one run on the
left and two runs on the right therefore cannot align the common Word or
Character units across the physical run boundary. It emits a replacement and
an insertion even when the visible text and significant raw content are
unchanged. This violates policy-wide granularity, deterministic attributed-unit
alignment, and minimal revision fragments for supported direct inline runs.

### D2, Run granularity leaks edited formatting inside a direct inline control

`crates/rdocx/src/comparison.rs:2039`
`crates/rdocx/src/comparison.rs:3620`
`crates/rdocx/src/comparison.rs:4619`

`ignore_formatting` alone does not select the attributed path. When text and
formatting both change in a direct `SdtContent::Run` under Run granularity, the
legacy branch inserts `run_xml(right)` instead of the policy-aware inserted run
that copies the original run properties. Acceptance therefore contains edited
formatting that the policy declared ignored, rather than retaining the original
formatting bytes left-biased. The focused inline-control test covers Word
granularity and ignored whitespace, but not this Run plus ignored-formatting
combination.

### D3, ignored TextBox normalization can choose different private placeholders

`crates/rdocx/src/comparison.rs:335`
`crates/rdocx/src/comparison.rs:946`
`crates/rdocx/src/comparison.rs:1435`
`crates/rdocx/src/comparison.rs:1458`

The comparison mask correctly chooses one collision-free expanded name across
both inputs and restoration now enforces exact cardinality. The acceptance
projection does not reuse that shared choice. It normalizes each package or
story independently by calling `text_box_marker_local(source, source, ...)`.
If an ignored original text box contains a producer element named
`{urn:rdocx:comparison:private}textBoxHost0` and the edited ignored box does
not, the accepted document selects `textBoxHost1` while the edited projection
selects `textBoxHost0`. Those marker bytes remain significant after the host is
masked, so the staged acceptance proof rejects an otherwise supported ignored
TextBox change. The collision regression places its producer element outside
the ignored host and does not exercise this projection divergence.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 3 D3 is remediated by projecting hyperlink raw-child positions through
logical policy boundaries. Pass 3 D4 is remediated with fixed literal records
covering kind, content, story, owner, and order, and the seven story categories
are suppressed independently. Namespace-aware TextBox collision detection and
exact restoration cardinality are present. No additional correctness,
contract, panic, OOXML, public-surface, dependency, or structure finding was
established.

The complete regression binary passed with 285 tests passed and 3 ignored. The
three comparison unit tests passed. The four named focused regressions passed,
and `git diff --check` was clean.
