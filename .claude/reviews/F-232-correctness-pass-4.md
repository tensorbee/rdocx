# F-232, correctness, pass 4

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 11 files and 2,314 changed lines, with 2,270 insertions and 44 deletions. All 238 `rdocx-layout` unit tests, its doc test, and all 12 `toc_` regression tests pass.
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, block-level Word wrapper names still impersonate typed ownership
`crates/rdocx/src/field.rs:908`

Paragraph ownership accepts any descendant chain made from `tbl`, `tr`, `tc`,
`sdt`, and `sdtContent`, without checking the parent-child grammar that the
typed body parser accepts. For example, a direct `w:sdtContent` below `w:body`
can contain a paragraph and complex TOC. The body parser retains that entire
element as raw XML, but this predicate marks its paragraph as typed and the
scanner can rebuild or reject a field inside the opaque subtree. The new inline
checks close the pass-3 examples below a real typed paragraph, but malformed
same-namespace block chains still cross the typed ownership boundary instead of
being ignored.

### D2, revision ownership does not enforce the parser's nesting bound
`crates/rdocx/src/field.rs:957`

The raw scanner accepts arbitrarily deep `w:ins` and `w:moveTo` chains whenever
each wrapper has a parseable id and an author. The OXML revision parser rejects
depth greater than 32 and retains that wrapper as raw XML. A field marker inside
a 33-level chain is therefore treated as owned by this scanner although it has
no typed revision projection. With a deeply wrapped begin marker followed by
direct instruction and separator runs, rebuild discovers the field and then
fails to reparse it instead of ignoring the opaque revision content.

### D3, a TOC instruction inside a valid inline wrapper is reconstructed as malformed XML
`crates/rdocx/src/field.rs:1154`

The scanner now intentionally accepts field runs in typed hyperlinks, content
controls, insertions, and move-to revisions. `result_start` is recorded at the
end of the separator run, but this slice ends before the closing tags of any
wrapper around that run. For a valid insertion containing the TOC begin,
instruction, and separator runs, the synthetic paragraph appends an end run and
paragraph close while `w:ins` is still open. XML parsing then rejects the
rebuild. Supported wrapped TOCs must be reparsed through their complete typed
owner structure rather than turned into an unbalanced fragment.

### D4, accepted source projection reorders controls and revisions at one boundary
`crates/rdocx/src/field.rs:1589`

The paragraph model records raw slots for both content controls and revisions,
but this traversal emits every control at a run boundary before every revision
at that boundary. Their actual serialized order is not consulted. If a valid
insertion containing `SEQ Chapter` precedes a content control containing a `TC`
field at boundary zero, source order is SEQ then TC, while this helper evaluates
TC then SEQ. The generated TC entry consequently loses its sequence prefix.
The reverse arrangement is also forced into the same order. This violates the
document-order and shared accepted-projection contract even though the current
revision test interleaves revisions only with direct runs.

### D5, reversed bookmarks at one run boundary pass validation
`crates/rdocx/src/field.rs:1266`

Bookmark ordering is compared only as `(paragraph, run)`, even though typed
markers carry their raw position within that run boundary. A paragraph can
serialize `bookmarkEnd`, then an inline content control, then the matching
`bookmarkStart` while every item has direct-run index zero. The comparison sees
equal positions and accepts the malformed range. A TOC scoped with `\\b` to
that name can then discover sources from the reversed range instead of rejecting
the staged rebuild atomically. The raw whole-paragraph scanner fixes partial
reuse, but it does not supply the ordered positions used for range validation.

### D6, simple TOCs in accepted revisions are not reported
`crates/rdocx/src/field.rs:837`

Simple TOC counting requires `w:fldSimple` to be a direct child of the typed
paragraph. A valid insertion or move-to wrapper containing a simple TOC is part
of the accepted revision projection, but its immediate parent is the revision,
so it contributes no diagnostic. Rebuild silently reports the default count,
or rebuilds another supported TOC without counting this unsupported valid one,
instead of preserving it with the required diagnostic.

## Smells

None.

## Nitpicks

None.

## Not found

- Exact pass-3 D1 examples: a direct `w:sdtContent` below a typed paragraph and
  an insertion without required id and author metadata no longer qualify as
  inline owners. D1 and D2 above concern adjacent block and depth mismatches.
- Pass-3 D2: direct outline conversion now uses `checked_add` before converting
  to the supported `u8` level, and the `u32::MAX` regression remains inside the
  `Result` boundary.
- Pass-3 D3: the basic accepted insertion cases for heading text, SEQ, and TC
  are now visible. D4 concerns their cross-type order at a shared boundary.
- Pass-3 D4: direct-child tokenization now prevents reuse when a content control
  prefix lies outside a bookmark. D5 concerns reversed marker order at an equal
  direct-run boundary.
- Panics and bounds: no new reachable indexing, slicing, splice, conversion, or
  arithmetic panic was found.
- OOXML generation and preservation: no additional generated-child order,
  namespace, prefix-write, whitespace, or unowned-byte loss was found beyond
  D1 through D3.
- Test gate: the focused layout and TOC suites pass, and no additional
  mutation-sensitivity gap was found beyond the missing inputs for D1 through
  D6.
- Collision-safe substitution, nested TOC rejection, lazy maximum bookmark id
  allocation, case-insensitive sequence identity, and distinct-page target
  association remain correct.
- Atomicity: observed parse, layout, reopen, and substitution errors remain
  outside the live document commit path.
- Structure and surface: no unjustified trait, generic, forwarding wrapper,
  module, feature flag, crate, dependency, Python, WASM, or CLI surface was
  introduced.
