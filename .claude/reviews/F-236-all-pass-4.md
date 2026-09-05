# F-236, all, pass 4

**Reviewed**: Pass-4 remediated uncommitted implementation diff against `dbb5ab1`, 7 files and 2,830 changed lines, comprising 2,824 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all three prior reviews and their closure evidence
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, malformed compatibility ancestry is accepted as schema-positioned
`crates/rdocx/src/embedded.rs:1136`

The story validator removes every `mc:AlternateContent`, `mc:Choice`, and
`mc:Fallback` node before validating the owner path. It never proves that a
choice or fallback is an immediate child of one alternate-content container,
or that the container has a valid branch structure. A path such as
`w:hdr/mc:Choice/w:p/w:r/w:object` therefore reduces to the accepted
header-to-paragraph path. Inventory and removal can act inside malformed
compatibility XML that should remain opaque under the fail-closed
schema-position contract.

### D2, valid text-box story content is invisible to the inventory
`crates/rdocx/src/embedded.rs:1111`

The story-path classifier has no `w:txbxContent` case, and the validator always
starts from the package part's outer story root. A valid paragraph and run
inside `w:txbxContent`, including a nested relationship-owned `w:object` or
`w:control`, consequently fails ancestry validation because the drawing or VML
path between the outer story and the text-box paragraph is classified as
unsupported. Office can activate executable content in that schema-positioned
text box, but `embedded_content` omits it and callers cannot identify or remove
it through the promised inventory.

### D3, VBA removal strips signatures before shared-target reachability is known
`crates/rdocx/src/embedded.rs:419`, `crates/rdocx/src/embedded.rs:491`

Both mutation policies remove the selected VBA project's signature
relationships before the later reachability check for the project part. If an
unrelated internal relationship also targets that VBA project, removal of the
main-document owner retains the shared project but has already detached its
VBA signature. The operation therefore mutates and invalidates a shared target
instead of preserving its graph. The existing shared-target regression covers
OLE and ActiveX, but not this VBA path.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 1 defects D1 through D5 and both pass 2 defects remain closed. ActiveX
binary multiplicity, shared ActiveX properties, package-signature incoming
edges, synchronized signature invalidation, compatibility-wrapper byte
preservation, invalid target modes, and invalid paragraph-to-run ancestry have
implementation and regression coverage.

Pass 3's exact `w:hdr/w:pPr/w:p/w:r/w:object` reproducer is now rejected by the
story-path state machine. D1 and D2 above show that complete schema-position
coverage is not yet closed for compatibility grammar and nested Word text-box
stories.

No additional findings were found in target normalization, package-signature
cleanup, mutation atomicity, panic safety, namespace-expanded relationship
attributes, raw byte-range removal, public API shape, dependency direction, or
repository structure. The focused `word_embedded` regression selection passes
all 6 tests.
