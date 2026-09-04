# F-232, correctness, pass 7

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 3,845 changed lines, with 3,755 insertions and 90 deletions. All 30 focused `toc_` regression tests, all 240 `rdocx-layout` unit tests, and its doc test pass.
**Verdict**: 2 defects, 0 smells, 0 nitpicks

## Defects

### D1, wrapper-internal runs do not retain exact old-result positions
`crates/rdocx/src/field.rs:1240`
`crates/rdocx/src/field.rs:2200`

The raw ownership scan assigns one position to an accepted revision or content
control, then every nested run inherits that same position. Source discovery
does the same when it overwrites the nested paragraph positions with the outer
owner position. The new boundary comparison therefore cannot distinguish two
runs on opposite sides of a TOC marker when they share one wrapper. For
example, a valid accepted insertion can contain a selected `TC` or matching
`SEQ` field followed by the TOC begin, instruction, and separator runs. The
source and separator receive the same coordinate, so the source is classified
as owned result content and omitted. Conversely, a selected field in the end
paragraph before an end-marker run in the same wrapper compares equal to the
end position and is incorrectly retained. The direct-run pass-6 regression
does not compose its boundary fields with an accepted revision or content
control.

### D2, tracked change bars ignore revisions in the newly projected control compositions
`crates/rdocx-layout/src/engine.rs:622`
`crates/rdocx-layout/src/engine.rs:651`

Tracked layout now projects revisions directly inside inline content controls
and content controls inside insertions or move destinations, but the paragraph
visibility predicate still inspects only direct paragraph runs and direct
paragraph revisions. Its revision helper also examines `RevisionContent::Runs`
instead of the accepted revision's typed paragraph projection. A paragraph
whose only changed text is `w:sdt/w:sdtContent/w:ins/w:r`, or
`w:ins/w:sdt/w:sdtContent/w:r`, renders the tracked text and decoration but is
reported as having no visible revision. It consequently omits the required
paragraph change bar. The new projection tests assert text only and cannot
detect this mismatch.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-6 D1 direct trigger: isolated instruction runs now use the XML reader's
  parsed start-tag boundary and skip namespace declarations already local to
  the run. Quoted greater-than signs and repeated declarations are covered.
- Pass-6 D2 direct trigger: facade ordering now gives a terminal hyperlink
  revision `BeforeRaw`, matching layout and keeping it before a following
  same-boundary control.
- Pass-6 D3 direct trigger: bookmark source ranges retain paragraph, run, and
  raw-child coordinates, and the direct before, inside, and after TC fields are
  filtered correctly.
- Pass-6 D4 direct trigger: direct sources before the TOC begin marker and after
  the end marker remain eligible, while direct sources in both result-boundary
  interiors are excluded. D1 above concerns positions collapsed inside one
  wrapper.
- Pass-6 D5 direct trigger: a content control directly inside an accepted
  insertion or move destination is recognized as a typed complex-TOC owner.
- Pass-6 D6 direct trigger: accepted insertion and move-destination paragraph
  projections now carry nested content-control text into layout, and the
  long-heading pagination regression is exact.
- Contract and public surface: the native additive report and operation match
  the approved plan, and Python, WASM, and CLI surfaces remain unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML generation and preservation: no fresh expanded-name, schema-order,
  wrapper-balance, fixed-prefix, whitespace, unmodelled-subtree, or package-byte
  preservation defect was found.
- Diagnostics: direct and accepted-revision simple TOCs retain stable counts.
  No additional diagnostic omission was found.
- Test gate: the pinned differential metadata and the existing direct trigger
  matrix remain mutation-sensitive. The two missing composed shapes are
  identified in D1 and D2.
- Collision-safe substitution, nested TOC rejection, lazy maximum bookmark-id
  allocation, checked outline conversion, ASCII case-insensitive sequence
  identity, same-boundary bookmark ordering, distinct page-target association,
  and atomic staged commit remain correct.
- Structure: no unjustified trait, generic, forwarding wrapper, module, feature
  flag, crate, dependency, or published binding surface was introduced.
