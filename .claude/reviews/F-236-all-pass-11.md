# F-236, all, pass 11

**Reviewed**: Pass-11 uncommitted implementation diff against `dbb5ab1`, excluding the ten earlier review artifacts, 7 files and 5,806 changed lines, comprising 5,800 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all ten prior reviews and their closure evidence
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, same-kind control owners can still claim overlapping removal ranges
`crates/rdocx/src/embedded.rs:828`
`crates/rdocx/src/embedded.rs:1528`
`crates/rdocx/src/embedded.rs:1065`

The new collision guard compares only OLE ranges against ActiveX ranges. It
does not compare two ranges returned by the control scan. A run-owned
`w:pict` can contain its own `w:control` plus a VML text box whose nested run
contains another `w:pict/w:control`. Both controls are inventoried, and the
outer picture range strictly contains the inner picture range. Removing the
outer ActiveX identity deletes both owner ranges while removing only the outer
story relationship and properties graph. Re-inventory then omits the now
ownerless inner relationship and commits. The mixed-kind reproducer is fixed,
but same-kind nesting still violates selected-owner isolation and the
valid-package postcondition.

### D2, undeclared general entities are accepted in relationship-owning XML
`crates/rdocx/src/embedded.rs:1246`
`crates/rdocx/src/embedded.rs:1632`

Both XML readers accept every general-reference event encountered inside the
document element. The ActiveX reader now rejects every document type, so a
named reference such as `&producer;` inside `ax:ocx` cannot have a declaration
and is not well-formed XML, yet the catch-all returns the root relationship id.
The story reader has the same acceptance path for an undeclared reference
inside a Word root while retaining owner references elsewhere in that root.
Inventory and mutation therefore trust malformed relationship-owning XML
instead of distinguishing legal character and predefined references from
undeclared general entities.

### D3, source identities are not required to be normalized Pack part names
`.claude/plans/F-236-design.md:105`
`crates/rdocx/src/embedded.rs:334`
`crates/rdocx/src/embedded.rs:1038`

The approved identity is a normalized source part, but inventory scans package
map keys without validating them and the public identity guard checks only a
small subset of Pack URI syntax. It accepts raw spaces, fragments, queries,
non-ASCII bytes, non-canonical percent encodings, and segments ending in a dot.
For example, a header stored and related as `/word/header 1.xml` can produce an
inventory identity, and extract, replace, and remove all accept that same
schema-invalid source string. Target normalization is now strict, but the
other half of the public identity and relationship source scope is not.

### D4, WordprocessingCanvas text boxes are absent from owner discovery
`crates/rdocx/src/embedded.rs:26`
`crates/rdocx/src/embedded.rs:1285`
`crates/rdocx/src/embedded.rs:2102`

The DrawingML path state supports direct WordprocessingShape graphic data and
WordprocessingGroup graphic data only. Office 2010 also defines
`a:graphicData` with the WordprocessingCanvas URI, whose `wpc:wpc` content can
contain `wps:wsp`. A valid path such as
`a:graphicData/wpc:wpc/wps:wsp/wps:txbx/w:txbxContent` resets to `Other` at
the graphic-data URI. Any OLE object or ActiveX control in that text-box story
is therefore omitted from inventory and cannot be extracted, replaced, or
removed.

### D5, VML text boxes beneath a run-owned Word object are not recognized
`crates/rdocx/src/embedded.rs:2043`
`crates/rdocx/src/embedded.rs:2075`

Legacy text-box discovery can begin only at a run-owned `w:pict`. VML shapes
are also valid children of `w:object`, and those shapes can contain
`v:textbox/w:txbxContent`. A schema-positioned nested owner at
`w:r/w:object/v:shape/v:textbox/w:txbxContent/w:p/w:r/w:object` leaves the
first VML shape in the `Other` state because its parent is not classified as a
legacy picture. The nested executable owner is invisible, and an outer OLE
owner can consequently be removed together with that undisclosed nested
content.

### D6, an empty valid ProcessContent value makes owners opaque
`crates/rdocx/src/embedded.rs:1770`
`crates/rdocx/src/embedded.rs:1870`

Markup Compatibility permits `mc:ProcessContent` to contain a list of zero or
more qualified names, including an empty value. The shared QName-list checker
instead requires at least one token. A harmless
`mc:ProcessContent=""` on a story ancestor therefore marks the entire ancestry
invalid and hides otherwise schema-positioned embedded owners. The prior MC
grammar and QName regressions cover malformed and non-empty values, but not
this valid empty-list case.

## Smells

None.

## Nitpicks

None.

## Not found

All 41 findings from passes 1 through 10 are closed for their cited
reproductions. In particular, pass 10's mixed OLE and ActiveX owner ranges now
fail closed, all VML shape primitives are recognized, validated MC wrappers
are transparent to text-box paths, ActiveX document types are rejected, and
relationship targets require normalized Pack URI syntax. The earlier graph,
signature MIME and incoming-edge, relationship singleton, root anchoring,
owner cardinality, story MIME, MC vocabulary, raw preservation, and grouped
DrawingML cases also remain closed.

No additional findings were found in signature invalidation or removal,
failure atomicity, hashing and exact extraction, panic safety, dependency
direction, public API shape, or repository structure. All 39 focused
`word_embedded_` regressions pass with default features and with all features.
`cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, and
`git diff --check dbb5ab1` pass.
