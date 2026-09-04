# F-232, correctness, pass 13

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 5,496 changed lines, with 5,386 insertions and 110 deletions. All 48 focused `toc_` regression tests, all 242 `rdocx-layout` unit tests, its doc test, all 368 `rdocx-oxml` unit tests, its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `git diff --check`, and the changed-file prose check pass.
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, final wrapper-local bookmark markers disappear from the typed bookmark model
`crates/rdocx/src/field.rs:3179`
`crates/rdocx/src/field.rs:3519`
`crates/rdocx-oxml/src/text.rs:2266`
`crates/rdocx-oxml/src/text.rs:2281`
`crates/rdocx-oxml/src/text.rs:2355`
`crates/rdocx/src/comments.rs:133`
`crates/rdocx-layout/src/engine.rs:5867`
`crates/rdocx/tests/regression_test.rs:2461`

The final relocation moves every generated or repaired end-boundary bookmark
start after the end run. When that run is inside an accepted hyperlink,
revision, or content control, the bookmark start becomes a child of that
wrapper. Begin-boundary repair can likewise insert a replacement bookmark end
at a wrapped begin run. The paragraph parser records direct bookmark elements
only. It consumes each of these wrappers through its wrapper-specific branch,
without projecting nested bookmark markers into `paragraph.bookmark_markers`.
The public bookmark facade and layout target insertion then inspect only that
direct marker vector.

Consequently, the post-end wrapper regression leaves `_Toc1` with a raw nested
start and a typed direct end after the final reopen. `Document::bookmarks()` no
longer exposes a matched target, a later layout cannot resolve the PAGEREF
target, and a second `rebuild_toc()` rejects the document as having an unmatched
bookmark at `crates/rdocx/src/field.rs:2232`. The regression verifies raw byte
order and range text only. It does not query the reopened bookmark facade,
layout the completed document, or rebuild it again. The same loss applies to an
exactly-one-consumed original bookmark whose replacement marker is inserted at
a supported wrapped boundary, regardless of whether the TOC entry requests
hyperlink plus PAGEREF, hyperlink only, or no generated target.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-12 D1, raw XML scope: the exact end-run boundary is retained and final
  relocation puts the generated start before post-end text inside hyperlink,
  accepted insertion, and content-control wrappers. D1 identifies the typed
  projection failure after that raw placement.
- Pass-12 D2, direct-boundary cases: every valid direct partial and
  cross-paragraph range with exactly one consumed marker is now discovered and
  repaired. The hyperlink plus PAGEREF, hyperlink-only, and target-free direct
  cases pass. D1 identifies the remaining accepted-wrapper boundary failure.
- Contract and public surface: the additive native rebuild operation and report
  remain within the approved plan. Python, WASM, and CLI surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML ownership and ordering: no fresh expanded-name, direct-owner,
  raw-coordinate, wrapper-balance, generated-child-order, fixed-prefix,
  structural-prefix, or namespace defect was found beyond D1.
- Verbatim preservation: accepted wrapper metadata, paragraph properties,
  unowned field scaffolding, raw XML, relationships, and untouched package
  parts remain preserved.
- Deterministic page resolution: the temporary direct target used for initial
  pagination and its final wrapper-local position differ only by zero-width
  field and wrapper scaffolding. No incorrect first-build page value was found.
  D1 covers later layout after the final reopen.
- Exact result removal: cached result text in middle and boundary paragraphs is
  removed without retaining stale runs.
- Facade and layout parity: accepted nested controls, insertion and
  move-destination text, hyperlink ordering, positioned tables, tracked
  visibility, and first-build page substitution remain aligned outside D1's
  bookmark projection failure.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No new diagnostic omission was
  found.
- Test gate: the pinned differential metadata and exact entry, link, level,
  raw target range, distinct-page, partial-range, cross-paragraph, and
  target-policy assertions remain intact. The missing reopened facade, later
  layout, and repeat-rebuild assertions are the mutation-sensitivity gap in D1.
- Stable whole-paragraph reuse away from TOC boundaries, collision-safe owned
  substitution, nested TOC rejection, maximum bookmark-id allocation, checked
  outline conversion, case-insensitive sequence identity, and atomic staged
  commit remain correct.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
