# F-235, all, pass 2

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 2,145 insertions and 129 deletions
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, ignoring a text box also ignores significant siblings in its containing run
`crates/rdocx/src/comparison.rs:744`
`crates/rdocx/src/comparison.rs:941`

The ignored TextBox path identifies the complete `w:r` that contains each text
box. It replaces the entire original run with a placeholder and removes the
entire edited run before comparing the host story. The same complete-run
removal is used for revision counting, id seeding, and postcondition
normalization. If one run contains ordinary text or another significant child
beside a drawing-backed text box, changes to that host content are silently
preserved from the original. Its ids and existing revisions are also excluded.
Only the text-box subtree is the ignored story. Significant siblings in its run
must remain in the host-story comparison.

### D2, the text-box placeholder can collide with ordinary document text
`crates/rdocx/src/comparison.rs:937`
`crates/rdocx/src/comparison.rs:2329`

The masking token is a normal `w:t` value, and the attributed comparison treats
any text with that prefix and suffix as an ignored text-box unit whenever the
TextBox policy is selected. A document containing legitimate text such as
`__rdocx_f234_text_box_0__` can therefore be mistaken for an internal marker.
Two such user values with different ordinals compare as the same ignored unit,
and a real text box plus matching user text can make placeholder restoration
fail its occurrence check. Internal ownership must not be inferred from a
user-visible text value.

### D3, granular hyperlink comparison still rejects harmless run segmentation changes
`crates/rdocx/src/comparison.rs:1837`

The hyperlink equality guard compares complete `HyperlinkSpan` values,
including `run_start` and `run_end`. Word or Character comparison of one
unchanged hyperlink represented by one run on the left and two equivalent runs
on the right is rejected before attributed-unit alignment. The relationship,
anchor, attributes, and XML owner shell can all be unchanged. The varying run
bounds describe the content segmentation that granular comparison is meant to
align, not a changed hyperlink shell.

### D4, policy postconditions still remove significant raw content with ignored text
`crates/rdocx/src/comparison.rs:3932`

The attributed signature now includes raw children when deciding granular
actions, but policy normalization filters every ignored unit before computing
the package projection. An ignored whitespace, field, or comment unit that
owns a raw child is removed together with that child. Acceptance and
cross-story move checks therefore do not verify raw content that the contract
declares significant. The focused raw test can catch the current writer, but
the staged package proof itself remains blind to loss, substitution, or
cross-story movement of this raw XML.

### D5, the staged atomicity test canonizes failure for a supported comparison
`crates/rdocx/tests/regression_test.rs:13996`
`crates/rdocx/tests/regression_test.rs:14017`

The second atomicity case constructs identical existing run-level
content-control shells and changes only the neighboring paragraph text under
Word granularity. Modeled content inside existing content-control shells and
ordinary adjacent text are part of the supported F-234 boundary. The test then
requires `compare_with_options` to fail its acceptance postcondition. This
turns a valid feature regression into the mechanism used to test atomic
failure. The postcondition failure test needs a deliberate invalid or injected
failure that does not make supported input erroneous.

### D6, the regression gate is still not an exact matrix over every policy
`crates/rdocx/tests/regression_test.rs:13520`
`crates/rdocx/tests/regression_test.rs:13747`
`crates/rdocx/tests/regression_test.rs:13783`

The named policy-matrix gate asserts only revision-container counts for
granularity and three ignore flags. It does not compare exact normalized
records, so wrong kind, content, order, story, or ownership can preserve the
same counts. It also omits `ignored_stories`. The separate story test combines
Header and Footnote in one option, while the independent loop contains neither
kind. A mutant where either selection suppresses both categories can remain
green. The approved gate requires exact declared record deltas and independent
coverage of every story category.

## Smells

None.

## Nitpicks

None.

## Not found

The leading and empty-paragraph insertion prefix is now placed inside the
paragraph owner. Paragraph numbering is removed from the ignore-formatting
projection. Same-bound hyperlink runs are source-correlated. Adjacent changed
units with the same owners are coalesced. Run granularity routes whitespace,
field, and comment ignores through attributed units. Duplicate ignored-story
options are validated before staging, and commit remains after both package
postconditions.

No separate arithmetic or panic defect was found in the new grouping code.
Group ranges and optional sides are derived from alignment action invariants.
No structure-rule or public-surface defect was found. The implementation adds
no dependency, trait, generic, feature flag, module, test binary, Python API,
WASM API, or CLI API.
