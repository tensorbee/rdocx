# F-232, correctness, pass 14

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 15 files and 6,211 changed lines, with 6,044 insertions and 167 deletions. All 49 focused `toc_` regression tests, all 242 `rdocx-layout` unit tests and its doc test, all 369 `rdocx-oxml` unit tests and its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `git diff --check`, and the prose check pass.
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, wrapper-local public bookmark ranges use a different coordinate system from their text
`crates/rdocx/src/comments.rs:194`
`crates/rdocx/src/comments.rs:198`
`crates/rdocx/src/comments.rs:213`
`crates/rdocx-oxml/src/text.rs:4436`
`crates/rdocx-oxml/src/text.rs:4454`
`crates/rdocx/tests/regression_test.rs:2492`

`Document::bookmarks()` still constructs its public `RunRange` from each
marker's direct `run_index`, but it validates direction and extracts text from
the new accepted-view `projected_run_index`. Those coordinates diverge for
markers projected out of revisions and content controls. A marker after one
run inside a revision can have direct run index one at an outer boundary whose
direct end marker has run index zero. A control-local start and end both have
direct run index zero even when they surround visible control runs. The facade
therefore reports a matched bookmark with no issue and the right text, while
returning a reversed or empty public range. The repaired pass-13 regression
checks `text()` after reopen but never checks `range()`. Callers cannot use the
reported range as the documented half-open bookmark boundary.

### D2, wrapper-local marker raw coordinates collapse tracked layout and TOC bookmark scope
`crates/rdocx-oxml/src/text.rs:4508`
`crates/rdocx-layout/src/engine.rs:6059`
`crates/rdocx-layout/src/engine.rs:6072`
`crates/rdocx/src/field.rs:2222`
`crates/rdocx/src/field.rs:2272`
`crates/rdocx/src/field.rs:2717`

The control projection assigns every nested marker direct coordinates
`run_index = 0` and `raw_before = 0`, then its parent only replaces those with
the outer control slot. Tracked bookmark resolution ignores the accepted
coordinate and consumes those direct coordinates, so a complete bookmark
around control text becomes empty. A revision-local marker can instead exceed
the outer paragraph direct-run count and make tracked resolution fail. The TOC
`\b` scope has the same loss. Bookmark state records only direct run and raw
positions, and `TocDocumentPosition` has no nested ordinal, while accepted TC,
SEQ, and heading runs do. A scope around a source inside an inline control or a
nested accepted owner can therefore exclude the selected source or include a
neighbour at the same outer boundary.

### D3, a nested content control loses its locally inherited namespace aliases during projection
`crates/rdocx-oxml/src/text.rs:4398`
`crates/rdocx-oxml/src/text.rs:4419`
`crates/rdocx-oxml/src/text.rs:4495`
`crates/rdocx-oxml/src/text.rs:4517`
`crates/rdocx-oxml/src/text.rs:7340`

Projection recurses into a nested `CT_Sdt` with the unchanged prefix scope from
the outer control. It does not carry namespace declarations from the nested
control root. A valid nested control such as `q:sdt xmlns:q="...wordprocessingml..."`
whose `q:bookmarkStart` and `q:bookmarkEnd` children inherit that declaration
parses as typed content, but the later marker projection cannot resolve `q` and
drops both markers. The new parser test uses only the inherited `w` spelling
and one-level wrappers, so this accepted typed path is not covered. The raw TOC
scanner remains namespace-aware, leaving raw ownership, facade discovery, and
layout target discovery inconsistent.

### D4, complex-field collapse leaves projected bookmark boundaries at removed run indexes
`crates/rdocx-oxml/src/text.rs:1582`
`crates/rdocx-oxml/src/text.rs:1610`
`crates/rdocx-oxml/src/text.rs:1721`
`crates/rdocx-oxml/src/text.rs:1751`
`crates/rdocx-oxml/src/text.rs:2395`
`crates/rdocx-oxml/src/text.rs:2501`
`crates/rdocx-layout/src/engine.rs:6007`

Parsing increments the accepted projection for every raw run in a complex
field. The later complex-field pass replaces the complete run sequence with
one typed field run. Its boundary remapper updates `BookmarkMarker.run_index`
only and leaves `projected_run_index` unchanged. A bookmark that starts or ends
after a valid same-paragraph complex field therefore retains an accepted index
based on runs that no longer exist. Accepted layout rejects an index beyond the
post-collapse run count, while public bookmark text clamps it to the end and
can silently return an empty or truncated range. No regression combines a
bookmark boundary after a collapsed complex field with REF or PAGEREF layout.

### D5, bookmark and comment mutation do not maintain accepted marker coordinates
`crates/rdocx-oxml/src/text.rs:385`
`crates/rdocx-oxml/src/text.rs:1968`
`crates/rdocx-oxml/src/text.rs:1985`
`crates/rdocx-oxml/src/text.rs:2301`
`crates/rdocx-oxml/src/text.rs:2333`
`crates/rdocx/src/comments.rs:831`
`crates/rdocx/tests/regression_test.rs:6840`

Inserting a direct run shifts only each bookmark's direct index. It never
shifts `projected_run_index`, even though the inserted run joins the accepted
run stream. Comment creation uses this helper for its reference run. Adding a
comment before a bookmark end that follows an inline control consequently
makes in-memory bookmark text and later target placement stop one run early
until save and reopen recalculates the projection. New bookmark markers have a
second form of the same error. Their accepted coordinate is initialized from
the direct run index, which omits accepted revision and control runs preceding
that direct boundary. The existing comment mutation regression checks direct
ranges and serialized ordering, not bookmark text or layout before reopen.

### D6, the public bookmark facade omits table and block-control paragraphs
`crates/rdocx/src/comments.rs:129`
`crates/rdocx/src/comments.rs:131`
`crates/rdocx/src/comments.rs:799`
`crates/rdocx/src/field.rs:6252`
`crates/rdocx-layout/src/engine.rs:5931`
`docs/hld/03-architecture.md:430`
`docs/hld/08-rendering-spec.md:1055`

The public facade still enumerates only body entries that are direct
paragraphs, and its position helper rejects tables and content controls. TOC
source discovery and layout now recursively traverse table and block-control
paragraphs. A rebuilt table heading can therefore own a valid generated
bookmark that resolves its PAGEREF in layout, but `Document::bookmarks()` does
not expose that same bookmark. This contradicts the revised facade correlation
contract and the rendering contract's shared typed main-story correlation.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-13 D1's direct trigger is closed. Generated starts inside hyperlink,
  accepted insertion, and inline-control owners are projected once, survive
  reopen, resolve a later accepted PAGEREF layout, and permit a second rebuild.
  D1 through D5 identify adjacent coordinate contracts that the focused case
  does not exercise.
- Direct-marker deduplication and standard-prefix wrapper order are correct.
  Direct, hyperlink, accepted revision, and inline-control markers retain
  accepted document order without duplicating serialized raw XML.
- Opaque ownership is correct for the tested cases. Foreign wrappers,
  malformed same-namespace wrappers, deleted revisions, and move-from
  revisions do not enter the accepted marker projection.
- Generated and repaired end-boundary bookmarks retain their tested hyperlink,
  revision, and control owner chains and post-end text. The tested target
  policies retain valid references and permit a repeat rebuild.
- Correctness outside D1 through D6: no additional wrong-result, atomicity,
  deterministic page-substitution, stale-result, source-order, or repeat-build
  defect was found.
- Contract and public surface outside D1 and D6: the additive native rebuild
  operation and report remain within the approved plan. Python, WASM, and CLI
  surfaces are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion,
  arithmetic, or recursion panic was found.
- OOXML ownership and ordering outside D2 and D3: no additional expanded-name,
  direct-owner, wrapper-balance, schema-order, fixed-prefix, raw-slot, or
  structural-prefix defect was found.
- Verbatim preservation: accepted wrapper metadata, paragraph properties,
  unowned field scaffolding, raw XML, relationships, and untouched package
  parts remain preserved.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No additional diagnostic omission
  was found.
- Test gate: the pinned differential metadata and exact entry, link, level,
  raw target range, distinct-page, boundary-fragment, repair-policy, and repeat
  rebuild assertions remain intact. D1 through D5 are the remaining mutation
  sensitivity gaps.
- Stable whole-paragraph reuse away from TOC boundaries, collision-safe owned
  substitution, nested TOC rejection, maximum bookmark-id allocation, checked
  outline conversion, case-insensitive sequence identity, exact old-result
  exclusion, and atomic staged commit remain correct.
- Structure and dependencies: no unjustified trait, generic, forwarding
  wrapper, module, feature flag, crate, dependency, or published binding
  surface was introduced.
