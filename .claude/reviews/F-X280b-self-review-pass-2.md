# F-X280b, self-review, pass 2

**Reviewed**: Successor working diff against claimed Base2e3c0780, thirteen
source, test, plan and HLD paths,2257additions and101deletions. The production
remediation since Root pass1 is limited to the existing Word validator.
**Verdict**: 0 additional defects, 0 smells, 0 nitpicks.

This follows Root's independent pass1, which found D1 through D4 after the
first self-review. The earlier review record is retained unchanged. This pass
does not replace the required independent Root pass2. No file outside the
review directory was modified during this pass.

## Defects

No additional defects found. The four Root findings have these local remedies.

- D1: `crates/rdocx/src/document.rs:82854` recognizes the actual picture,
  WPS and WPG drawing container namespaces already admitted by the existing
  story reader. Qualified picture and actual authored text-box a:ext values
  now reach the existing extent check. Arbitrary foreign payload keeps its
  semantic stop. Genuine picture and textbox controls pass before mutation.
  Negative a:ext values reproduce actual RED and become required part-named
  Error findings after the fix. The actual picture wp:extent stays unchanged.
- D2: `crates/rdocx/src/document.rs:83519` restricts the style-reference reader
  to actual story roles after lexical validation. Main and story bytes still
  reach the same semantic reader and final actual Document-open predicate.
  All four related UTF-16 auxiliary variants reproduce actual RED, while
  identical unconnected copies are valid. Connected and unconnected variants
  now validate and retain exact auxiliary bytes through save and reopen.
- D3: `crates/rdocx/src/document.rs:83486` keys targets by actual normalized
  part identity and unions the story role before scanning. The source's
  relationship-local lookup remains unchanged. Actual custom-first and alias
  cases reproduce RED. Both relationship orders, exact, case and dot aliases,
  clean headers and header-only schema faults now satisfy the discriminator.
- D4: `crates/rdocx/tests/integration_test.rs:40116` calls the real
  `clone_content` publisher on an imported repeated-id paragraph. The existing
  fragment ownership check refuses before remapping and publication. Source
  and destination paragraph bytes, held and fresh locations, package bytes
  and save/reopen remain intact. An unrelated clone then succeeds with the
  imported repeats preserved. This was a missing acceptance control and is
  GREEN on the first frozen implementation, not a newly claimed product RED.

## Smells

None found.

## Nitpicks

None found.

## Not found

Correctness, contract, panics, OOXML, tests and structure were checked with no
additional finding. Known drawing paths were traced to actual existing picture
and textbox publishers and namespace classification. Non-story encoding was
challenged through actual related and unconnected paths before the style
reader. Role evidence is accumulated before deduplication. Imported-repeat
multiplicity is still distinct from fragment ownership and reserved custody.
No helper framework, public API, module, crate, generic parameter, trait or
test binary was added. HLD intent and final gate debt remain unchanged.

## Evidence and exclusions

The immutable initial freeze and Root review lineage remain applicable.
`/private/tmp/rdocx-s96-f280b-local-evidence/review-pass1-remediation`
retains the actual initial three-RED/one-GREEN source restore patch, source
copies and binary, sealed supplementary WPS RED and its replacement, the
small production-only remediation diff, successor full patch, hashes and
copied GREEN binary. The first unsealed WPS observation remains separately
logged. It was repeated to bind exact post-format source, not to select a
favorable outcome.

The affected Word cohort now passes16tests. Strict scoped Word Clippy with
all targets and features passes. Formatting, prose and diff whitespace pass.
Unchanged OPC, PPT, bookmark/private and CLI observations are reused with their
original source applicability. They were not rerun as confirmation or renamed
newer installed/native evidence. Root's independent pass2, prepared handoff,
final integrated workspace, native, schema, rendering, installed Python,
packaging and hosted CI gates remain pending. No commit, push, completion or
release is claimed.
