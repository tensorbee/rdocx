# F-232, correctness, pass 8

**Reviewed**: Full uncommitted working diff from worker HEAD `9109a9a`, 13 files and 4,128 changed lines, with 4,022 insertions and 106 deletions. All 32 focused `toc_` regression tests, all 241 `rdocx-layout` unit tests, its doc test, `cargo fmt --all --check`, and `git diff --check` pass.
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, wrapper positions still use a synthetic raw slot that disagrees with the typed parser
`crates/rdocx/src/field.rs:1290`
`crates/rdocx/src/field.rs:2313`

The raw ownership scan assigns every otherwise uninherited content control,
insertion, and move destination `Raw(0)`. The accepted typed projection instead
uses the owner's actual paragraph raw slot and gives a revision in a hyperlink
`BeforeRaw`, `AfterRaw`, or the hyperlink's preserved raw slot. Adding a nested
leaf ordinal cannot reconcile those different outer coordinates. For example,
a terminal hyperlink insertion can contain the TOC separator followed by a
selected `TC` field. The scan records the separator at `Raw(0)`, while source
discovery places every run in that insertion at `BeforeRaw`. The post-separator
TC therefore compares before the result start and is incorrectly included.
Likewise, a control or revision after an ordered raw paragraph child has a
nonzero typed raw slot, so a source before the separator inside that same owner
compares after the synthetic `Raw(0)` result start and is incorrectly omitted.
The new total-order tests place their TOC owners in raw slot zero and do not
exercise either mismatch.

### D2, simple fields retained as hyperlink raw XML still advance ownership positions
`crates/rdocx/src/field.rs:1270`
`crates/rdocx-oxml/src/text.rs:4158`

The ownership scanner treats `w:fldSimple` under a hyperlink as a modeled run
and advances the paragraph run boundary. The typed hyperlink parser models only
direct runs and content revisions, retaining a simple field as raw XML instead.
The two projections then assign different positions to every following run.
For example, a hyperlink containing a simple field before a direct complex TOC
makes the scanner record the separator one boundary later than the accepted
paragraph model. A selected TC field immediately after the separator is then
treated as outside the owned old result and becomes a spurious rebuilt entry.
The same drift occurs for any direct simple field shape that the typed parser
rejects and preserves as raw.

### D3, direct control runs hide property-only revisions from tracked change bars
`crates/rdocx-layout/src/engine.rs:668`

The recursive tracked-visibility fix descends into control-owned paragraphs,
tables, cells, nested controls, and content revision wrappers, but explicitly
returns false for every direct `SdtContent::Run`. A direct run in an inline
content control can carry `w:rPrChange` or another modeled run-property
revision. That property-only change is projected and rendered, yet the outer
paragraph is still reported as having no visible revision and receives no
change bar. The new two-order test covers visible insertion wrappers only, not
a property revision on a direct control-owned run.

## Smells

None.

## Nitpicks

None.

## Not found

- Pass-7 D1 direct trigger: runs before and after the separator or end marker
  inside one raw-slot-zero accepted insertion or inline control now receive a
  total nested order and are classified correctly. D1 and D2 concern cases
  where the raw and typed outer coordinate systems do not match.
- Pass-7 D2 direct trigger: visible content revisions inside inline controls and
  controls inside insertions now mark the tracked paragraph, and both nesting
  orders render exactly one change bar. D3 concerns direct run-property
  revisions rather than content revision wrappers.
- Contract and public surface: the native additive report and rebuild operation
  remain within the approved plan, and Python, WASM, and CLI surfaces remain
  unchanged.
- Panics and bounds: no new reachable indexing, slicing, conversion, recursion,
  or arithmetic panic was found.
- OOXML generation and preservation: no fresh expanded-name, generated-child
  order, wrapper-balance, fixed-prefix, whitespace, unmodelled-subtree, or
  package-byte preservation defect was found beyond the position disagreement
  in D1 and D2.
- Accepted projection and layout: nested content-control text, accepted
  insertion and move-destination text, positioned table traversal, terminal
  hyperlink ordering, and final page-target association remain aligned. No
  additional parity defect was found.
- Diagnostics: supported and unsupported complex TOCs plus direct and accepted
  revision simple TOCs retain stable counts. No additional diagnostic omission
  was found.
- Test gate: the pinned differential metadata and existing mutation-sensitive
  cases remain intact. Missing trigger shapes are identified in D1 through D3.
- Bookmark scope, whole-paragraph bookmark reuse, collision-safe owned
  substitution, nested TOC rejection, lazy maximum bookmark-id allocation,
  checked outline conversion, ASCII case-insensitive sequence identity,
  distinct page-target association, and atomic staged commit remain correct.
- Structure: no unjustified trait, generic, forwarding wrapper, module, feature
  flag, crate, dependency, or published binding surface was introduced.
