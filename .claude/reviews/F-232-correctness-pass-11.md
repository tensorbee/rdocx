# F-232, correctness, pass 11

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 4,818 changed lines, with 4,708 insertions and 110 deletions. All 41 focused `toc_` regression tests, all 242 `rdocx-layout` unit tests, its doc test, all 368 `rdocx-oxml` unit tests, its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `git diff --check`, and the changed-file prose check pass.
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, a self-closing end-paragraph property element is deleted
`crates/rdocx/src/field.rs:2990`
`crates/rdocx/src/field.rs:3159`

The paragraph-boundary scan advances `content_start` when a direct `w:pPr`
is closed through an `End` event, but its `Empty` branch never advances that
boundary. A valid end paragraph using `<w:pPr/>`, including an aliased spelling
or one carrying producer attributes, therefore leaves `content_start`
immediately after the paragraph opening. Replacement copies only through that
offset, so the property element is consumed with the old result instead of
being retained. The new end-boundary regression uses a nonempty property
element at `crates/rdocx/tests/regression_test.rs:2308`, so it does not exercise
the self-closing form.

### D2, reconstructing an end-marker content control drops its properties
`crates/rdocx/src/field.rs:1881`
`crates/rdocx/src/field.rs:3160`

The scanner records only each accepted wrapper's lexical start tag, and the
replacement copies only those ranges. If the end-marker run is inside
`w:sdt/w:sdtContent`, the `w:sdtPr` and optional `w:sdtEndPr` nodes occur
between those two start tags and are not copied. Rebuild silently removes the
control's alias, tag, id, data binding, type, and preserved raw property slots,
which are modeled as control identity and producer state at
`crates/rdocx-oxml/src/content_control.rs:143`. Retaining the bare wrapper keeps
the XML balanced, but it does not preserve the unrelated control metadata.

### D3, boundary-fragment rebuild leaves the original bookmark unmatched
`crates/rdocx/src/field.rs:274`
`crates/rdocx/src/field.rs:347`

The pass-10 remediation correctly allocates a new target instead of reusing a
whole-paragraph bookmark on a TOC boundary, but it does not preserve or repair
the old bookmark. On a begin fragment, the old start marker survives before
the heading while its end marker after the separator is removed. On an end
fragment, the old start marker before the end field is removed while its end
marker after the heading survives. The operation accepts unmatched markers, so
the rebuild succeeds after corrupting an existing bookmark and any reference
to it. This also occurs when both hyperlinking and PAGEREF are disabled because
the allocation loop skips sources that need no generated bookmark at
`crates/rdocx/src/field.rs:268`. The two new tests only locate the generated
`_Toc1` target at `crates/rdocx/tests/regression_test.rs:2380` and
`crates/rdocx/tests/regression_test.rs:2411`. They never assert that
`wholeBegin` or `wholeEnd` remains a valid matched range.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-10 D1, ordinary cached runs before the end marker are now removed. A
  nonempty direct `w:pPr` and the accepted wrapper starts needed for balance
  are retained. D1 and D2 cover the remaining structural-prefix forms.
- Pass-10 D2, a boundary-fragment source no longer reuses a whole-paragraph
  bookmark as its generated TOC target. D3 covers preservation of the original
  bookmark itself.
- Contract and public surface: the additive native rebuild operation and report
  remain within the approved plan. Python, WASM, and CLI surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML ownership and ordering: no fresh expanded-name, direct-owner,
  raw-coordinate, wrapper-balance, generated-child-order, fixed-prefix, or
  namespace defect was found beyond D1 through D3.
- Verbatim preservation: unowned field scaffolding, raw XML, relationships, and
  untouched package parts remain preserved outside the structural and bookmark
  boundary cases in D1 through D3.
- Exact result removal: cached result text in ordinary middle and end-boundary
  runs is removed, and the post-end source title and generated bookmark cover
  the same surviving text outside D1.
- Facade and layout parity: accepted nested controls, insertion and
  move-destination text, hyperlink ordering, positioned tables, final page
  targets, and tracked visibility remain aligned.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No new diagnostic omission was
  found.
- Test gate: the pinned differential metadata and exact entry, link, level, and
  distinct-page assertions remain intact. The missing mutation-sensitive
  structural and bookmark shapes are identified in D1 through D3.
- Stable whole-paragraph reuse away from TOC boundaries, collision-safe owned
  substitution, nested TOC rejection, maximum bookmark-id allocation, checked
  outline conversion, case-insensitive sequence identity, and atomic staged
  commit remain correct.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
