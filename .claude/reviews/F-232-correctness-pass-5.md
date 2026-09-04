# F-232, correctness, pass 5

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 11 files and 2,651 changed lines, with 2,607 insertions and 44 deletions. All 238 `rdocx-layout` unit tests, its doc test, and all 18 `toc_` regression tests pass.
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, a second content container crosses the typed block boundary
`crates/rdocx/src/field.rs:1004`

The block classifier accepts every `w:sdtContent` whose parent was classified as
a content control. The typed `CT_Sdt` parser accepts only its first
`w:sdtContent` and preserves a later one as raw XML. A valid content control can
therefore contain its typed content container followed by a second container
holding a complex TOC. The raw scanner marks the second container and its
paragraph as typed, so rebuild can mutate or reject a field that the object
model deliberately preserved as opaque. The direct-body regression closes the
pass-4 spelling but does not enforce the content control's single-container
grammar.

### D2, property revisions can make scanner depth disagree with the typed parser
`crates/rdocx/src/field.rs:1061`

The ownership scanner increments revision depth only for accepted `w:ins` and
`w:moveTo` inline owners. The OXML parser's limit counts all revision elements,
including property-change revisions, anywhere inside a captured revision. For
example, 32 nested valid insertions can contain a run property change and then
a complex TOC marker. OXML rejects the outer revision because the property
change reaches depth 33, while this predicate still accepts the field marker at
depth 32. The result is field discovery inside raw XML even though the direct
33-insertion regression now passes.

### D3, wrapped instruction projection drops inherited namespace aliases
`crates/rdocx/src/field.rs:1310`

Rebuild copies the accepted instruction runs into a new paragraph that declares
only the `w` prefix. A valid wrapper can declare a different WordprocessingML
alias, such as `q`, and its descendant `q:r` and `q:fldChar` elements can inherit
that declaration. The original balanced span parses, but the copied run slices
no longer have the wrapper's `xmlns:q` declaration. Parsing the projected
paragraph then fails instead of rebuilding this prefix-tolerant wrapped TOC.

### D4, accepted revisions directly inside content controls disappear from source projection
`crates/rdocx/src/field.rs:1777`

`CT_Sdt` recognizes revisions inside `w:sdtContent`, records them separately,
and retains their serialized element as raw content. This traversal visits runs,
nested controls, paragraphs, tables, rows, and cells, but skips raw content and
has no access to the control's typed revision collection. The raw ownership
scanner also rejects an insertion whose immediate parent is `w:sdtContent`.
Consequently accepted heading text, SEQ or TC fields, and simple TOC diagnostics
inside a valid control-level insertion or move-to revision are all omitted.

### D5, revision-only hyperlinks still break cross-type source order
`crates/rdocx/src/field.rs:1749`

This merge treats the second value of every paragraph revision tuple as a raw
XML position. Revisions parsed from hyperlinks instead store a high-bit encoded
hyperlink index there. A revision-only hyperlink retains its real raw position
on the hyperlink while its encoded tuple value sorts after ordinary raw slots.
If such a hyperlink contains an accepted `SEQ` field before a same-boundary
content control containing `TC`, accepted source evaluation processes the TC
first and loses its sequence prefix. The pass-4 direct insertion and control
tests do not exercise this hyperlink representation.

### D6, layout omits inline content-control runs used by TOC discovery
`crates/rdocx-layout/src/engine.rs:220`

Facade source discovery includes accepted runs from a paragraph's direct
content controls, but the layout projection emits only paragraph revisions and
ordinary runs. A heading whose text is inside an inline content control is
therefore used as a TOC title and bookmark target while pagination measures the
paragraph without that text. With enough controlled heading text to reflow, the
omission can shift later targets and produce page numbers that disagree with
the accepted Word view. The shared accepted-projection contract is still split
between facade discovery and layout.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-4 D1 direct trigger: a direct `w:sdtContent` below `w:body` is ignored.
  D1 above is the adjacent duplicate-container grammar case inside a typed
  content control.
- Pass-4 D2 direct trigger: a field inside 33 nested insertions is ignored. D2
  above is the property-revision contribution to the same parser depth limit.
- Pass-4 D3 direct triggers: insertion, hyperlink, and content-control wrappers
  using the declared `w` prefix are balanced and rebuild successfully. D3 above
  concerns an alias inherited from a removed wrapper.
- Pass-4 D4 direct trigger: direct revisions and content controls are merged by
  their raw positions in both orders. D5 above concerns the distinct encoded
  slot representation of hyperlink revisions.
- Pass-4 D5: whole-bookmark qualification now compares the raw position at a
  shared run boundary and rejects the reversed-marker example.
- Pass-4 D6: immediate simple TOCs in accepted direct insertions and move-to
  revisions are counted. D4 above concerns revisions directly owned by a
  content control.
- Panics and bounds: no additional reachable indexing, slicing, splice,
  conversion, recursion, or arithmetic panic was found.
- OOXML generation and preservation: no additional generated-child order,
  whitespace, or unowned-byte loss was found beyond D1 through D4.
- Exact raw ordering: same-boundary direct controls, revisions, ordinary runs,
  bookmarks, and both SEQ/TC direct orders are covered. No additional ordering
  fault was found beyond D5.
- Diagnostics: unsupported direct and accepted-revision simple TOCs are counted,
  and no additional diagnostic omission was found beyond D4.
- Test gate: the focused layout and TOC suites pass. No additional
  mutation-sensitivity gap was found beyond the missing trigger shapes for D1
  through D6.
- Collision-safe substitution, nested TOC rejection, lazy maximum bookmark id
  allocation, checked outline conversion, case-insensitive sequence identity,
  same-boundary bookmark order, and distinct-page target association remain
  correct.
- Atomicity: observed parse, layout, reopen, and substitution errors remain
  outside the live document commit path.
- Structure and surface: no unjustified trait, generic, forwarding wrapper,
  module, feature flag, crate, dependency, Python, WASM, or CLI surface was
  introduced.
