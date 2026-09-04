# F-232, correctness, pass 10

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 4,657 changed lines, with 4,547 insertions and 110 deletions. All 39 focused `toc_` regression tests, all 242 `rdocx-layout` unit tests, its doc test, all 368 `rdocx-oxml` unit tests, its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `git diff --check`, and the changed-file prose check pass.
**Verdict**: 2 defects, 0 smells, 0 nitpicks

## Defects

### D1, replacement retains cached result runs from the end paragraph
`crates/rdocx/src/field.rs:3105`

The replacement closes the begin paragraph and emits the generated entries,
then copies the complete end-paragraph prefix from its opening tag through the
byte before the end-marker run. That prefix contains any old cached-result
runs in the end paragraph, so those owned runs are reinserted instead of being
removed. The new boundary-heading regression places `old boundary` before the
end marker, but it checks only the filtered entry title and bookmark text. The
stale run therefore survives outside the generated bookmark while the test
passes. Rebuild is required to replace the whole owned result range and retain
only the structural paragraph prefix needed to contain the end marker.

### D2, whole-paragraph bookmark reuse is unsafe for a partial boundary source
`crates/rdocx/src/field.rs:274`
`crates/rdocx/src/field.rs:2868`

A heading on a TOC boundary paragraph now contributes only its unowned runs,
but bookmark selection can still reuse a pre-existing bookmark that covers the
whole paragraph. Reused bookmarks skip insertion and are not narrowed to that
unowned range. For a heading before the begin marker, a valid whole-paragraph
bookmark has its end marker after the separator. Result replacement deletes
that marker, leaving the generated hyperlink and PAGEREF with an unmatched
target and causing an otherwise valid rebuild to fail or produce a dangling
link when no page field is requested. On the end boundary, the same reuse can
target cached-result content in addition to the post-end heading. Boundary
sources need a bookmark whose range matches the surviving unowned fragment.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-9 D1: direct comments and processing instructions now consume the same
  raw-child slots as the typed paragraph parser, with both forms covered.
- Pass-9 D2: direct simple fields advance the modeled run boundary only when
  their parsed instruction name is nonempty. Missing and empty instructions
  remain raw and are covered.
- Pass-9 D3: direct hyperlink runs now consume the shared nested ordinal, so a
  selected revision field immediately before the end-marker run stays inside
  the old result.
- Pass-9 D4: heading titles on a boundary paragraph now filter each accepted
  run through exact old-result ownership, and newly allocated bookmarks cover
  the same post-end text. D1 and D2 cover adjacent replacement and reuse paths.
- Contract and public surface: the additive native rebuild operation and report
  remain within the approved plan. Python, WASM, and CLI surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML ownership and ordering: no fresh expanded-name, typed-owner,
  raw-coordinate, wrapper-balance, generated-child order, fixed-prefix,
  whitespace, or namespace defect was found beyond D1 and D2.
- Verbatim preservation: unowned field scaffolding, raw XML, relationships, and
  untouched package parts remain preserved outside the boundary cases in D1
  and D2.
- Facade and layout parity: accepted nested controls, insertion and
  move-destination text, hyperlink ordering, positioned tables, final page
  targets, and tracked visibility remain aligned.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No new diagnostic omission was
  found.
- Test gate: the pinned differential metadata and exact entry, link, level, and
  distinct-page assertions remain intact. The missing mutation-sensitive
  boundary shapes are identified in D1 and D2.
- Bookmark validation, stable whole-paragraph reuse away from TOC boundaries,
  collision-safe owned substitution, nested TOC rejection, maximum bookmark-id
  allocation, checked outline conversion, case-insensitive sequence identity,
  and atomic staged commit remain correct.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
