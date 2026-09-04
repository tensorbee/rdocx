# F-232, correctness, pass 12

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 5,049 changed lines, with 4,939 insertions and 110 deletions. All 44 focused `toc_` regression tests, all 242 `rdocx-layout` unit tests, its doc test, all 368 `rdocx-oxml` unit tests, its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `git diff --check`, and the changed-file prose check pass.
**Verdict**: 2 defects, 0 smells, 0 nitpicks

## Defects

### D1, an end-fragment target skips surviving text inside the end wrapper
`crates/rdocx/src/field.rs:1092`
`crates/rdocx/src/field.rs:2969`

When the end-marker run is inside an accepted hyperlink, revision, or content
control, the scan changes `end_run_end` from the end of that run to the end of
the outer wrapper. Bookmark insertion then uses that extended offset as the
start of the surviving end fragment. Source discovery still treats a later run
inside the same wrapper as unowned because it compares the run against the end
marker at `crates/rdocx/src/field.rs:2618`. Put the real heading text after the
end marker but before the wrapper closes. The generated TOC entry includes that
heading, while its hyperlink and PAGEREF bookmark starts after the wrapper and
does not cover the heading. If the wrapper contains all remaining paragraph
text, the generated target is empty. The new control-prefix test avoids the
trigger by placing `Real heading` after `</w:sdt>` at
`crates/rdocx/tests/regression_test.rs:2384`.

### D2, bookmark repair excludes valid non-whole crossing ranges
`crates/rdocx/src/field.rs:2151`
`crates/rdocx/src/field.rs:2234`
`crates/rdocx/src/field.rs:2955`

Repair candidates are populated only from bookmarks classified as covering
every content token in one paragraph. A valid partial or cross-paragraph
bookmark with one marker outside the TOC result and the other inside it is not
in `all_whole_paragraphs`, so result replacement still deletes one marker and
leaves the original reference target unmatched. For example, a begin-boundary
paragraph can have an ordinary run, then a bookmark start, then a heading run,
the TOC instruction and separator, and the bookmark end. The leading run keeps
that bookmark out of the whole-paragraph set, while the end marker is removed
with the result. This happens whether the entry uses hyperlink plus PAGEREF,
hyperlink alone, or neither, because original repair and generated target
allocation are separate paths. The new regressions at
`crates/rdocx/tests/regression_test.rs:2425` and
`crates/rdocx/tests/regression_test.rs:2505` cover only whole-paragraph
bookmarks.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-11 D1: direct self-closing paragraph properties now advance the exact
  structural-prefix boundary. The aliased form with producer attributes is
  covered through save and reopen.
- Pass-11 D2: an accepted end-marker content control now retains its complete
  prefix through `w:sdtContent`, including modeled properties, end properties,
  identity, binding, type, and ordered raw slots.
- Pass-11 D3, whole-paragraph trigger: all nested whole-paragraph bookmarks on
  a begin boundary and the whole-paragraph bookmark on an end boundary are
  repaired around the surviving fragment. The target-free policy is covered.
  D2 identifies the remaining crossing-range class.
- Contract and public surface: the additive native rebuild operation and report
  remain within the approved plan. Python, WASM, and CLI surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML ownership and ordering: no fresh expanded-name, direct-owner,
  raw-coordinate, wrapper-balance, generated-child-order, fixed-prefix, or
  namespace defect was found beyond D1 and D2.
- Verbatim preservation: paragraph and content-control structural prefixes,
  unowned field scaffolding, raw XML, relationships, and untouched package
  parts remain preserved outside D2's crossing-marker case.
- Exact result removal: cached result text in middle and boundary paragraphs is
  removed without retaining stale runs.
- Facade and layout parity: accepted nested controls, insertion and
  move-destination text, hyperlink ordering, positioned tables, final page
  targets, and tracked visibility remain aligned outside D1's target scope.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No new diagnostic omission was
  found.
- Test gate: the pinned differential metadata and exact entry, link, level, and
  distinct-page assertions remain intact. The missing mutation-sensitive
  wrapper-local target and non-whole crossing bookmark shapes are D1 and D2.
- Stable whole-paragraph reuse away from TOC boundaries, collision-safe owned
  substitution, nested TOC rejection, maximum bookmark-id allocation, checked
  outline conversion, case-insensitive sequence identity, and atomic staged
  commit remain correct.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
