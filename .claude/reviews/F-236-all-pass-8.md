# F-236, all, pass 8

**Reviewed**: Pass-8 uncommitted implementation diff against `dbb5ab1`, excluding the seven earlier review artifacts, 7 files and 4,501 changed lines, comprising 4,495 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all seven prior reviews and their closure evidence
**Verdict**: 8 defects, 0 smells, 0 nitpicks

## Defects

### D1, owner cardinality still ignores relationship-less OLE and control children
`crates/rdocx/src/embedded.rs:1161`
`crates/rdocx/src/embedded.rs:1165`
`crates/rdocx/src/embedded.rs:1223`
`crates/rdocx/src/embedded.rs:1252`

The pass-7 fix counts collected relationship ids, not schema children. An
`o:OLEObject` contributes to that count only when it has a nonempty relationship
attribute. A `w:object` containing one relationship-less `o:OLEObject` followed
by one valid relationship-owned child therefore has an id count of one and is
accepted. Control discovery has the same conditional behavior. Removing the
inventoried identity then deletes the complete `w:object` or `w:pict`, including
the otherwise opaque extra child. Duplicate same-id children now fail, but
mixed missing-id and valid children still bypass ambiguous-owner rejection and
raw preservation.

### D2, ActiveX removal deletes targets of unrelated producer relationships
`crates/rdocx/src/embedded.rs:561`
`crates/rdocx/src/embedded.rs:581`
`crates/rdocx/src/embedded.rs:590`
`crates/rdocx/src/embedded.rs:591`

Before deleting an unreachable ActiveX properties part, removal records the
target of every internal relationship in that scope, not only the validated
ActiveX binary. After deleting the properties part and its relationship set,
it deletes each recorded target that is now unreachable. A producer-defined
relationship from the properties part to an otherwise unreferenced custom part
therefore causes that custom part and its content-type override to be deleted
with the selected control. The contract allows only validated owned candidates
to be cleaned up and explicitly requires unrelated orphans to survive.

### D3, compatibility rules on ordinary owner ancestors are never validated
`crates/rdocx/src/embedded.rs:1132`
`crates/rdocx/src/embedded.rs:1180`
`crates/rdocx/src/embedded.rs:1410`

Compatibility-rule attribute validation is invoked only while constructing an
`mc:AlternateContent` state and while checking its `mc:Choice` or
`mc:Fallback`. Ordinary Word and drawing ancestors do not carry any validation
result. A story such as `w:hdr mc:Ignorable="unbound"` with a normal embedded
owner below it is consequently treated as valid schema ancestry and remains
actionable. The pass-7 regression puts each invalid rule value on
`mc:AlternateContent`, so it does not close the cited general case of an
otherwise ordinary wrapper carrying malformed MC rules.

### D4, ProcessContent does not require a matching Ignorable declaration
`crates/rdocx/src/embedded.rs:1439`
`crates/rdocx/src/embedded.rs:1443`
`crates/rdocx/src/embedded.rs:1457`

The QName-list validator proves only syntax, an in-scope binding, and that the
binding is not the MC namespace. A `ProcessContent` token must also name a
namespace declared ignorable on the same element or an ancestor. An alternate
content owner with `xmlns:x="urn:producer" mc:ProcessContent="x:item"` but no
matching `mc:Ignorable` therefore passes this implementation. Embedded content
below that nonconformant MC ancestry can be inventoried and removed instead of
remaining opaque.

### D5, valid preservation wildcards make embedded owners disappear
`crates/rdocx/src/embedded.rs:1444`
`crates/rdocx/src/embedded.rs:1471`
`crates/rdocx/src/embedded.rs:1478`

The shared MC QName-list grammar permits a namespace-qualified wildcard, but
the validator enables `prefix:*` only for `ProcessContent`. It explicitly
disables the wildcard for `PreserveElements` and `PreserveAttributes`, causing
the `*` local part to fail the NCName check. A valid wrapper using
`mc:PreserveElements="w14:*"` is marked invalid as a whole, so a supported OLE
or control owner below it is omitted and cannot be extracted, replaced, or
removed.

### D6, DrawingML text-box ownership ignores the required graphicData URI
`crates/rdocx/src/embedded.rs:1668`
`crates/rdocx/src/embedded.rs:1674`
`crates/rdocx/src/embedded.rs:1686`
`crates/rdocx/tests/regression_test.rs:16304`

The DrawingML state machine enters `GraphicData` from the expanded element name
alone and accepts either word-processing shape vocabulary beneath it. It never
requires `a:graphicData/@uri`, or proves that the URI selects the direct shape
or group vocabulary actually encountered. Missing and mismatched URI owners
therefore become actionable. The grouped DrawingML regression itself omits the
required URI, so its green result proves acceptance of a schema-invalid
fixture rather than closure for a valid grouped text box.

### D7, a story content type is not matched to its document root
`crates/rdocx/src/embedded.rs:311`
`crates/rdocx/src/embedded.rs:449`
`crates/rdocx/src/embedded.rs:1547`
`crates/rdocx/src/embedded.rs:1576`

Relationship-less parts enter the scan from one of six supported content
types, but the scanner independently accepts any of the six recognized Word
roots. It never carries the expected part kind into root validation. For
example, a part declared as footnotes but rooted at `w:hdr`, or a header part
rooted at `w:comments`, can expose an embedded owner to inventory and removal.
The owner path is locally plausible but is not schema-positioned for the OPC
part that contains it, so the malformed graph should fail closed.

### D8, declarations and document types outside the story root are ignored
`crates/rdocx/src/embedded.rs:1281`
`crates/rdocx/src/embedded.rs:1288`
`crates/rdocx/src/embedded.rs:1299`
`crates/rdocx/src/embedded.rs:1335`

The root validity check records element count plus outside text, CDATA, and
general references, while every declaration and document-type event falls
through the catch-all arm. A second XML declaration or a document type after
the closed story root therefore leaves the count at one and the outside flag
clear. References collected from the root survive EOF and remain removable
even though the source is not a well-formed single XML document. The pass-7
outside-root regression covers text only and does not exercise these remaining
document-grammar events.

## Smells

None.

## Nitpicks

None.

## Not found

The pass-7 cases for non-whitespace and character content outside the root,
duplicate same-id OLE children, supported story parts without relationship
sets, and the named unbound or malformed MC values have concrete closure. All
earlier findings remain closed for their cited reproductions, including exact
target-mode handling, complete ordinary story ancestry, nested text boxes,
package and VBA signature MIME checks, shared VBA targets, and nested group
discovery. D1, D3, D4, and D8 above identify adjacent cases in the same broader
cardinality, MC, and root-grammar requirements.

No additional findings were found in payload hashing and extraction,
replacement identity, package or VBA signature removal and invalidation,
target normalization, failure atomicity, panic safety, additive public API
shape, dependency direction, or repository structure. All 21 focused
`word_embedded_` regressions pass. `cargo check -p rdocx --all-targets` and
`git diff --check dbb5ab1` pass.
