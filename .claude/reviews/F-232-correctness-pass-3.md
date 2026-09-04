# F-232, correctness, pass 3

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 11 files and 1,927 changed lines, with 1,884 insertions and 43 deletions. All 238 `rdocx-layout` unit tests, its doc test, and all nine `toc_` regression tests pass.
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, Word-shaped wrapper names are mistaken for a typed owner path
`crates/rdocx/src/field.rs:920`

The new owner check accepts a run when every intervening element has one of
seven Word local names. It does not validate their parent-child grammar or
prove that the OXML parser projected each wrapper. For example, a direct
`w:sdtContent` below `w:p`, or a malformed `w:ins` without the required id and
author, can contain a Word run whose TOC markers pass this whitelist even
though the paragraph parser retains the wrapper as raw XML. The scanner then
discovers a field that `parse_dynamic_toc_field` cannot project and rejects the
rebuild instead of ignoring the opaque subtree. The foreign-wrapper regression
closes the exact pass-2 examples, but pass-2 D1 remains open for same-namespace
opaque wrapper chains.

### D2, the largest parsed outline level can panic during source discovery
`crates/rdocx/src/field.rs:1451`

An outline-selected TOC applied to a paragraph containing
`<w:outlineLvl w:val="4294967295"/>` reaches `level + 1` while `level` is
`u32::MAX`. This overflows and panics in checked builds before the `u8`
conversion can reject the out-of-domain value. In release builds it wraps to
zero, so behavior also differs by build profile. An invalid direct outline
level must be rejected or ignored through checked arithmetic without escaping
the atomic `Result` boundary.

### D3, accepted revision content is absent from TOC source discovery
`crates/rdocx/src/field.rs:1341`

Field discovery iterates `paragraph.runs()`, and heading title extraction later
uses `paragraph.text()` at `crates/rdocx/src/field.rs:1387`. Those APIs collect
direct and content-control runs but omit the paragraph's typed revision runs at
`crates/rdocx-oxml/src/text.rs:1835`. A Heading 1 paragraph whose text is in a
valid insertion is therefore omitted, and mixed direct plus inserted text
produces a truncated entry. Accepted inserted SEQ or TC content is likewise
invisible. Deterministic pagination uses the accepted revision projection, so
source discovery and the final layout disagree about the document being
rebuilt.

### D4, whole-paragraph bookmark reuse ignores positioned inline content
`crates/rdocx/src/field.rs:1235`

The whole-paragraph test compares only direct-run indexes. Bookmark markers
also carry raw ordering, while content controls and revisions can occupy the
same direct-run boundary. A valid paragraph can place a run-bearing inline
content control before a bookmark, then put its direct heading run inside the
bookmark. The marker indexes are still zero and `paragraph.runs.len()`, so this
code classifies the bookmark as covering the whole paragraph even though the
source title includes the content-control prefix. Rebuild then reuses that
partial bookmark for the generated hyperlink and reports no allocation. Stable
first-marker ordering is fixed, but only a complete ordered paragraph position
can establish that a candidate is actually whole.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-2 D2 ordering: bookmark starts are consumed in marker encounter order,
  and the first qualifying candidate wins repeatably. D4 concerns candidate
  qualification, not the repaired ordering.
- Pass-2 D3 sequence identity: TOC and SEQ identifiers now use the same ASCII
  case-insensitive comparison, and the mixed-case exact tuple covers it.
- Pass-2 D4 page association: one rebuild now binds two exact entry and
  bookmark tuples to distinct final pages, so an all-targets-first-page mutant
  fails.
- Collision-safe substitution and nested ownership: placeholders remain unique
  to their owned result spans, and overlapping TOC ranges are rejected before
  edits.
- Maximum bookmark id behavior: allocation remains lazy, and the final
  representable id is returned before exhaustion.
- Bounds and panics: no additional reachable indexing, slicing, splice, or
  arithmetic panic was found beyond D2.
- OOXML schema order and verbatim preservation: no additional generated-child
  order, namespace, whitespace, prefix-write, or unowned-byte loss was found
  beyond D1 and D4.
- Test gate: no additional mutation-sensitivity gap was found beyond the
  missing edge coverage corresponding to D1 through D4.
- Structure and surface: no unjustified trait, generic, forwarding wrapper,
  module, feature flag, crate, dependency, Python, WASM, or CLI surface was
  introduced.
- Atomicity: apart from the panic in D2, staged parse, reopen, layout, and
  substitution failures remain outside the live document commit path.
