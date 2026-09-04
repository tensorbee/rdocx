# F-235, all, pass 1

**Reviewed**: Working diff from claim base `82a7d5a`, 3 files, 1,263 insertions and 108 deletions
**Verdict**: 9 defects, 0 smells, 0 nitpicks

## Defects

### D1, leading granular insertions are serialized outside the paragraph
`crates/rdocx/src/comparison.rs:1871`

The interleaver copies the paragraph prefix only when it first consumes an
original unit. An insertion aligned before the first original unit is appended
while `output` is still empty. For example, Word or Character comparison of
`world` against `hello world` writes the insertion before `<w:p>`. An insertion
into an empty paragraph has the same failure because no original owner is ever
consumed. This produces an invalid owner hierarchy instead of a revision inside
the paragraph.

### D2, selected stories are still validated and influence ids before they are skipped
`crates/rdocx/src/comparison.rs:245`
`crates/rdocx/src/comparison.rs:257`
`crates/rdocx/src/comparison.rs:685`

The declared precedence requires a selected story to be skipped before shell
comparison, existing-revision checks, or revision allocation. Main-story
revisions are still rejected unconditionally, and main XML always seeds the id
allocator even when `Main` is ignored. Text-box host and run-owner counts are
also compared before the `TextBox` ignore check at line 709. A valid edited host
that adds or removes an ignored text box therefore errors instead of preserving
the original text-box bytes and comparing the rest of the host story.

### D3, paragraph numbering remains significant when formatting is ignored
`crates/rdocx/src/comparison.rs:3622`
`crates/rdocx/src/comparison.rs:304`

The policy-normalized paragraph signature always includes `numId` and `ilvl`.
The formatting writer correctly retains the original paragraph properties when
`ignore_formatting` is set, so a numbering-only edit leaves the accepted
candidate with original numbering while the edited policy projection still has
edited numbering. The acceptance proof then rejects a comparison that the
policy says must be ignored.

### D4, granular comparison cannot correlate runs owned by hyperlinks
`crates/rdocx/src/comparison.rs:1560`
`crates/rdocx/src/comparison.rs:1862`
`crates/rdocx/src/comparison.rs:4028`

Every Word or Character paragraph is routed to the granular comparator, but its
source correlation scans only `w:r` elements at paragraph depth one. Runs owned
by an unchanged `w:hyperlink` are present in `CT_P::runs` but are nested one
level deeper in the source XML. Editing visible hyperlink text with unchanged
hyperlink bounds therefore fails the run-count guard instead of producing a
granular revision. This regresses a supported visible-text owner from the F-234
comparison boundary.

### D5, raw run children attached to ignored units become insignificant
`crates/rdocx/src/comparison.rs:1928`
`crates/rdocx/src/comparison.rs:1958`
`crates/rdocx/src/comparison.rs:2050`

Whether a unit is ignored is fixed before raw children are attributed. The
signature for an ignored unit then returns only `ignored:whitespace`,
`ignored:field`, or `ignored:comment` and omits the attached raw bytes. With
`ignore_whitespace`, for example, different foreign children immediately after
an all-whitespace text node compare equal and the original is retained
silently. Raw content is required to remain significant and to be emitted
exactly once, independent of text ignore policy.

### D6, adjacent granular edits are not coalesced into minimal revisions
`crates/rdocx/src/comparison.rs:1748`
`crates/rdocx/src/comparison.rs:1763`

The implementation allocates and serializes one deletion and insertion wrapper
for every aligned unit. There is no coalescing step between alignment and
serialization. A Character comparison of `abc` against `xyz` therefore emits
three adjacent deletion and insertion pairs even though all units have the same
source owner and properties. This violates the approved requirement to
coalesce identical ownership before writing the smallest fixed-prefix
fragments.

### D7, Run granularity does not preserve ignored sub-run content left-biased
`crates/rdocx/src/comparison.rs:1585`
`crates/rdocx/src/comparison.rs:1618`
`crates/rdocx/src/comparison.rs:3855`

Run mode projects ignored fields, comments, or textual whitespace only for
alignment. When significant content in the same run also changes, the entire
original run is deleted and the entire edited run is inserted. The insertion
helper preserves only formatting and copies the edited ignored content. For
example, a text change beside a whitespace change in one run accepts the edited
whitespace rather than retaining the original whitespace bytes. The ignore
policies are therefore not left-biased under the default granularity.

### D8, the regression gate does not prove its declared policy matrix
`crates/rdocx/tests/regression_test.rs:13278`
`crates/rdocx/tests/regression_test.rs:13298`
`crates/rdocx/tests/regression_test.rs:13466`

The named gate covers only default, Word, and Character plus whitespace
options. It never exercises formatting, fields, comments, or selected story
policies, and it compares only two executions of each option with each other.
It does not compare one policy with another or assert the declared record
deltas. Most selected-story assertions also require only that old text occurs,
which remains true inside an ordinary deletion if the ignore is removed. These
tests can stay green when individual policy branches are bypassed, so the
regression gate is not mutation-sensitive to the promised contract.

### D9, the atomic policy test exercises metadata validation only
`crates/rdocx/tests/regression_test.rs:13542`

The test named `invalid_comparison_policy_leaves_package_and_caches_unchanged`
passes default options and an invalid timestamp. It repeats the pre-policy
metadata guard and does not exercise an option-driven validation failure or an
acceptance or rejection postcondition failure after staged work has begun. It
therefore does not establish the test-plan requirement that policy validation
and postcondition failures leave package state and caches unchanged.

## Smells

None.

## Nitpicks

None.

## Not found

No separate panic or arithmetic defect was found. Indexed source spans are
guarded by correlation counts, and unit owners originate from the enumerated
run slice. No structure-rule violation was found. The public enum, options
value, re-export, and additive method follow the approved native-only surface,
and the legacy method delegates to default options. No additional dependency,
trait, generic, module, file, test binary, Python API, WASM API, or CLI API was
introduced.
