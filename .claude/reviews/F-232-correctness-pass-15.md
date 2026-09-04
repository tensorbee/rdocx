# F-232, correctness, pass 15

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 15 files and 7,219 changed lines, with 6,893 insertions and 326 deletions. All 51 focused `toc_` regression tests, the full `rdocx` regression binary with 237 passes and 2 ignored tests, all 371 `rdocx-oxml` unit tests and its doc test, all 244 `rdocx-layout` unit tests and its doc test, `cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, `git diff --check`, and the prose check pass.
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, mutation refresh loses Word aliases inherited above the paragraph
`crates/rdocx-oxml/src/text.rs:2248`
`crates/rdocx-oxml/src/text.rs:2531`
`crates/rdocx-oxml/src/text.rs:2555`
`crates/rdocx-oxml/src/text.rs:4628`
`crates/rdocx-oxml/src/document.rs:870`
`crates/rdocx-oxml/src/document.rs:896`
`crates/rdocx-oxml/src/document.rs:962`
`crates/rdocx-oxml/src/text.rs:7671`

The new mutation repair serializes only the changed paragraph and reparses it
through `CT_P::from_xml`, whose initial namespace scope contains only the
canonical `w`, `r`, and `mc` bindings. The original document parse instead
inherits aliases declared on `w:document`, and document serialization retains
those declarations at that ancestor. A paragraph can therefore initially
project raw `q:bookmarkStart` and `q:bookmarkEnd` elements whose `q` alias is
declared only on the document root. Inserting a comment run, direct run, or
bookmark invokes the standalone refresh. The fragment still contains the raw
`q:` markers but no declaration for `q`, so `word_prefixes_at` no longer
recognizes them and both markers disappear from the in-memory projection. A
later full save and reopen restores them from the retained root declaration,
which makes bookmark discovery and REF or PAGEREF behavior depend on whether a
mutation has occurred and whether the package has been reopened. The namespace
regression declares `q` locally on the nested control, so it does not exercise
the ancestor-only alias that the document parser supports.

### D2, reversed markers at the same accepted boundary are reported as a valid bookmark
`crates/rdocx/src/comments.rs:168`
`crates/rdocx/src/comments.rs:195`
`crates/rdocx/src/comments.rs:199`
`crates/rdocx-layout/src/engine.rs:6003`
`crates/rdocx-layout/src/engine.rs:6020`
`crates/rdocx-layout/src/engine.rs:6042`
`crates/rdocx/tests/regression_test.rs:2571`
`docs/hld/03-architecture.md:973`

The facade groups markers by id and detects reversal only by comparing their
accepted `(body_index, run_index)` positions. It discards encounter order when
the positions are equal. Layout target discovery makes the same position-only
comparison. Thus `<w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="1"
w:name="bad"/>` at one run boundary is exposed as a valid empty range rather
than the documented reversed-marker issue, and layout accepts the same invalid
pair as a zero-width target. The repaired reversed-marker regression inserts a
visible run between its end and start, so its positions differ and the test
does not cover this shared-boundary case. Marker order is already available in
the paragraph projection and must participate in direction qualification when
the accepted coordinates tie.

### D3, an inline content control projects block children that violate its owner grammar
`crates/rdocx-oxml/src/text.rs:2607`
`crates/rdocx-oxml/src/content_control.rs:578`
`crates/rdocx-oxml/src/content_control.rs:582`
`crates/rdocx-oxml/src/text.rs:1873`
`crates/rdocx-oxml/src/text.rs:4748`
`crates/rdocx-layout/src/engine.rs:346`
`crates/rdocx/src/field.rs:1310`
`crates/rdocx/src/field.rs:1698`
`crates/rdocx/src/field.rs:2971`
`docs/hld/03-architecture.md:510`

The paragraph parser accepts any direct Word `w:sdt` as an inline typed
control, but the context-free content-control parser also types block-level
paragraph, table, row, and cell children. Accepted and tracked run projection,
bookmark projection, layout, and the typed TOC source path then flatten those
block children into the containing paragraph. The raw ownership scanner
correctly applies distinct block and inline parent-child grammars and rejects
the same path. An invalid inline control such as
`w:p/w:sdt/w:sdtContent/w:p` can consequently contribute heading text, TC or
SEQ fields, bookmark targets, and rendered runs through the typed paths while
remaining excluded by raw TOC ownership. This violates the documented rule
that same-namespace wrappers qualify only when their owner grammar matches and
that Word-shaped content below unsupported paths remains opaque. The control
parser needs its block or inline owner context before classifying children, or
the inline projections need to exclude block variants.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-14 D1 through D6 are closed for their direct triggers. Public bookmark
  ranges and text now share recursive accepted coordinates, accepted and
  tracked marker indexes survive nested owners and complex-field collapse,
  nested controls retain local aliases, mutation reparses marker projection,
  and the facade traverses table and block-control paragraphs. D1 through D3
  above are adjacent cases outside those regressions.
- Correctness outside D1 through D3: no additional wrong-result, atomicity,
  deterministic page-substitution, stale-result, source-order, or repeat-build
  defect was found.
- Contract and public surface outside D2: the additive native rebuild operation
  and report remain within the approved plan. Python, WASM, and CLI surfaces
  are unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion,
  arithmetic, recursion-depth, or allocation panic was found.
- OOXML namespace, ownership, and ordering outside D1 and D3: no additional
  expanded-name, direct-owner, wrapper-balance, schema-order, fixed-prefix,
  raw-slot, structural-prefix, or revision-depth defect was found.
- Verbatim preservation: control properties and end properties, paragraph
  properties, unowned field scaffolding, comments and processing instructions,
  raw XML, relationships, and untouched package parts remain preserved.
- Accepted and tracked projection outside D2 and D3: hyperlink, insertion,
  move-to, nested-control, and direct marker coordinates remain aligned across
  facade and layout. Deleted and move-from content stays excluded from the
  accepted view.
- Bookmark repair: generated and repaired boundary-fragment markers retain
  their tested hyperlink, revision, and content-control owner chains. Crossing,
  partial, and exactly-one-consumed original bookmarks retain valid tested
  references for every target policy.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No additional diagnostic omission
  was found.
- Test gate: the pinned differential metadata and exact entry, link, level,
  raw target range, distinct-page, boundary-fragment, repair-policy, and repeat
  rebuild assertions remain intact. D1 through D3 identify the remaining
  mutation-sensitive and malformed-input gaps.
- Stable whole-paragraph reuse away from TOC boundaries, collision-safe owned
  substitution, nested TOC rejection, maximum bookmark-id allocation, checked
  outline conversion, case-insensitive sequence identity, exact old-result
  exclusion, and atomic staged commit remain correct.
- Structure and dependencies: no additional unjustified trait, generic,
  forwarding wrapper, module, feature flag, crate, dependency, or published
  binding surface was introduced. The cross-layer inconsistency caused by the
  context-free control parser is reported as D3 rather than duplicated as a
  smell.
