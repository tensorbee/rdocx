# F-232, correctness, pass 6

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 12 files and 3,319 changed lines, with 3,238 insertions and 81 deletions. All 239 `rdocx-layout` unit tests, its doc test, and all 24 `toc_` regression tests pass.
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, namespace reinjection does not find the actual end of the run start tag
`crates/rdocx/src/field.rs:1525`

The isolated-run builder treats the first `>` byte as the end of the start tag.
A literal `>` is legal inside a quoted XML attribute value, so a valid
instruction run such as `<q:r producer:value=">">` makes this helper insert
namespace declarations into the attribute value. The projected paragraph then
fails to parse and a supported wrapped TOC cannot be rebuilt. The helper also
injects every inherited declaration without excluding a declaration repeated
locally on the run, which can create a duplicate `xmlns` attribute. The new
inherited-alias regression covers neither valid start-tag shape.

### D2, a terminal hyperlink revision is still sorted after a following control
`crates/rdocx/src/field.rs:1977`

Only a revision-only hyperlink has `preserved_raw_before`, so a hyperlink that
contains an ordinary run followed by a revision falls back to its high-bit
encoded hyperlink slot. If a content control follows that hyperlink at the
same flattened run boundary, its small raw position sorts before the terminal
revision even though the revision is serialized inside the earlier hyperlink.
For example, a terminal accepted insertion containing `SEQ Chapter` followed
by a control containing a selected `TC` field evaluates the TC first and loses
its sequence prefix. Layout already represents this terminal position as
`BeforeRaw`, but facade source discovery does not share that rule. The pass-5
test covers only a hyperlink with no ordinary runs.

### D3, bookmark-selected TOC scope discards the marker positions
`crates/rdocx/src/field.rs:1635`

Validated bookmark starts and ends carry paragraph, run, and raw positions, but
the stored named range keeps only its two paragraph indexes. Source discovery
therefore treats every accepted field in either boundary paragraph as inside
the `\b` scope. A single paragraph can contain a TC field before the bookmark,
another TC field inside it, and a third after it. This implementation includes
all three instead of only the bookmarked entry. Reversed-marker validation now
uses the exact positions, but valid range selection does not.

### D4, old result exclusion drops unowned sources in its boundary paragraphs
`crates/rdocx/src/field.rs:1877`

The recorded TOC ownership is a byte range from after the separator run through
the matching end run, but discovery skips every paragraph from the begin
paragraph through the end paragraph. A TC or SEQ field serialized before the
TOC begin marker in the instruction paragraph, or after the end marker in the
end paragraph, lies outside the owned cache and must remain eligible. It is
silently omitted because its paragraph index falls in this coarse range. This
contradicts the contract to exclude only the old owned result range.

### D5, a content control inside an accepted insertion is not a typed TOC owner
`crates/rdocx/src/field.rs:1232`
`crates/rdocx-oxml/src/revision.rs:117`

The insertion model reparses its content as a paragraph, so a valid inline
content control inside `w:ins` is present in the typed accepted projection.
The raw ownership predicate nevertheless permits `w:sdt` only below a
paragraph or another control content container, not below the insertion. A
complex TOC whose begin, instruction, and separator runs are in that nested
control is ignored instead of rebuilt. This is the inverse composition of the
new control-owned-revision case and is not covered by the separate wrapper
tests.

### D6, layout drops content controls nested inside accepted insertions
`crates/rdocx-layout/src/engine.rs:479`
`crates/rdocx-oxml/src/revision.rs:117`

Facade discovery walks the insertion's paragraph projection, including its
content controls. Layout instead projects only `RevisionContent::Runs`, whose
lower-level parser does not include those controls. A Heading 1 paragraph with
long text inside `w:ins/w:sdt/w:sdtContent` contributes that text to the TOC
title and bookmark, while deterministic pagination measures none of it. Later
target pages can therefore be wrong. Direct inline controls and revisions
directly inside controls now lay out correctly, but their opposite nesting
order still splits the accepted projection.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-5 D1 direct trigger: only the first typed `w:sdtContent` in one control
  owns block or inline descendants. A second content container stays opaque.
- Pass-5 D2 direct trigger: property-change revisions now contribute to the
  32-element depth limit, and an overdeep owner is invalidated before field
  discovery.
- Pass-5 D3 direct trigger: namespace aliases inherited from a wrapper are
  copied onto isolated instruction runs. D1 above concerns legal start-tag
  spellings that the byte insertion corrupts.
- Pass-5 D4 direct trigger: accepted insertions and move-to revisions directly
  inside a content control contribute headings, SEQ, TC, and simple TOC
  diagnostics. D5 and D6 concern the inverse nesting order.
- Pass-5 D5 direct trigger: a revision-only hyperlink uses its retained owner
  position relative to a same-boundary control. D2 above concerns a terminal
  revision in a hyperlink that also has an ordinary run.
- Pass-5 D6 direct trigger: runs in a direct inline content control now
  participate in layout and can shift later TOC page targets.
- Panics and bounds: no new reachable indexing, slicing, splice, conversion,
  recursion, or arithmetic panic was found.
- OOXML generation and preservation: no additional generated-child order,
  whitespace, fixed-prefix write, or unowned-byte loss was found beyond D1.
- Diagnostics: direct and accepted-revision simple TOCs are counted, and no
  additional diagnostic omission was found beyond the ownership composition in
  D5.
- Test gate: the focused layout and TOC suites pass. No additional
  mutation-sensitivity gap was found beyond the missing trigger shapes for D1
  through D6.
- Collision-safe substitution, nested TOC rejection, lazy maximum bookmark id
  allocation, checked outline conversion, case-insensitive sequence identity,
  same-boundary bookmark order, and distinct-page target association remain
  correct.
- Atomicity: observed parse, layout, reopen, and substitution errors remain
  outside the live document commit path.
- Structure and surface: the content-control revision accessor has current
  facade and layout consumers. No unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, Python, WASM, or CLI surface
  was introduced.
