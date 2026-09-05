# F-236, all, pass 10

**Reviewed**: Pass-10 uncommitted implementation diff against `dbb5ab1`, excluding the nine earlier review artifacts, 7 files and 5,395 changed lines, comprising 5,389 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all nine prior reviews and their closure evidence
**Verdict**: 5 defects, 0 smells, 0 nitpicks

## Defects

### D1, OLE and ActiveX identities can claim the same removable owner range
`crates/rdocx/src/embedded.rs:339`
`crates/rdocx/src/embedded.rs:364`
`crates/rdocx/src/embedded.rs:950`

OLE and control references are collected in two independent scans, so no check
can detect that one `w:object` contains both an `o:OLEObject` and a `w:control`.
Each scan accepts its own single child and assigns the complete `w:object` range
to that identity. Removal then rescans only the selected kind and deletes that
complete range. Removing the OLE identity therefore also deletes the ActiveX
owner XML while leaving its story relationship and properties graph in place,
and removing the ActiveX identity has the symmetric effect on the OLE graph.
The candidate re-inventory ignores the now-unreferenced relationship and commits
the unintended second removal. This violates both selected-owner isolation and
the valid-package postcondition.

### D2, valid VML text boxes are recognized only beneath `v:shape`
`crates/rdocx/src/embedded.rs:1950`
`crates/rdocx/src/embedded.rs:1956`

The VML path state recognizes only a literal `v:shape` as the parent that can
lead to `v:textbox`. VML also permits text boxes on shape primitives such as
`v:rect`, `v:oval`, and `v:roundrect`. A valid
`w:pict/v:rect/v:textbox/w:txbxContent` story therefore resets to `Other` at the
rectangle, and every OLE object or control in that text box is omitted from the
security inventory and cannot be extracted, replaced, or removed.

### D3, compatibility branches interrupt otherwise valid text-box paths
`crates/rdocx/src/embedded.rs:1205`
`crates/rdocx/src/embedded.rs:1929`

Text-box state always reads the immediate open node rather than skipping a
validated `mc:AlternateContent` and its branch. Every MC node has an `Other`
text-box state, so a valid path such as
`v:textbox/mc:AlternateContent/mc:Choice/w:txbxContent` fails the story-root
test even though the compatibility grammar itself is accepted. The scanner
already treats MC wrappers as transparent for run ownership, but not for the
VML and DrawingML paths that establish a nested Word story. Executable content
in these conforming compatibility-wrapped text boxes is consequently absent
from inventory.

### D4, ActiveX properties accept a document type inside the root element
`crates/rdocx/src/embedded.rs:1133`

The ActiveX reader explicitly rejects misplaced XML declarations, but every
other event encountered while `depth > 0` is accepted by the catch-all arm.
That includes `Event::DocType`, so a properties part such as
`<ax:ocx ...><!DOCTYPE x></ax:ocx>` reaches EOF with a valid relationship id
even though a document type is not legal inside the document element. Inventory
and mutation then trust malformed relationship-owning XML instead of failing
closed. The new declaration tests do not exercise this remaining document
grammar event.

### D5, relationship targets are not required to be normalized Pack URIs
`crates/rdocx/src/embedded.rs:880`
`crates/rdocx/src/embedded.rs:911`
`crates/rdocx/src/embedded.rs:936`

The internal-target guard rejects empty values, control bytes, and traversal
expressed with slash-separated `..` segments. It does not reject a backslash,
while the public source-identity validator explicitly does. A relationship
target such as `..\payload.bin` is treated as one ordinary segment, and a ZIP
part with the corresponding backslash-containing name can satisfy the existence
and content-type checks. The resulting non-Pack-URI target is inventoried and
can be mutated even though the approved contract requires a normalized safe
target.

## Smells

None.

## Nitpicks

None.

## Not found

All 36 findings from passes 1 through 9 are closed for their cited
reproductions. In particular, preservation QNames now require effective
`Ignorable` semantics, `MustUnderstand` is limited to scanner-supported
vocabularies, ActiveX declarations and properties MIME are checked, and the
main document permits at most one VBA project relationship. The earlier graph,
signature MIME and incoming-edge, root anchoring, MC grammar, raw preservation,
owner cardinality, nested group, and story MIME cases also remain closed.

No additional findings were found in signature invalidation or removal,
failure atomicity, hashing and exact extraction, panic safety, dependency
direction, public API shape, or repository structure. All 34 focused
`word_embedded_` regressions pass with default features and with all features.
`cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, and
`git diff --check dbb5ab1` pass.
