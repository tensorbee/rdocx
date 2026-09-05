# F-237, all, pass 7

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the six earlier review artifacts, 17 files and 5,222 changed lines, comprising 5,160 additions and 62 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, entity references outside a story root bypass document-element validation
`crates/rdocx-oxml/src/glossary.rs:386`
`crates/rdocx-oxml/src/glossary.rs:399`
`crates/rdocx/src/field.rs:7841`
`crates/rdocx/src/field.rs:7867`

Both root scanners reject non-whitespace `Text` and `CData` while their stacks
are empty, but numeric and general character references arrive from quick XML
as `Event::GeneralRef` and fall through the wildcard arms. A source such as
`&#65;<w:glossaryDocument ...>` or the equivalent before or after a package
story root is therefore accepted even though it contains character data
outside the XML document element. This reopens the malformed-root hole for both
building-block loading and legacy-form inventory.

### D2, glossary entries accept the body before the optional properties element
`crates/rdocx-oxml/src/glossary.rs:451`
`crates/rdocx-oxml/src/glossary.rs:459`
`crates/rdocx-oxml/src/glossary.rs:475`
`crates/rdocx-oxml/src/glossary.rs:483`

The direct-entry parser now rejects duplicate `w:docPartPr` and
`w:docPartBody` children, but it does not record which child was encountered
first. It consequently accepts `w:docPartBody` followed by `w:docPartPr`, even
though the entry sequence permits the optional properties only before the
required body. Inventory exposes this invalid entry, and a replacement edits
the existing structural spans in place, retaining the invalid order after a
successful reopen.

### D3, duplicate modeled glossary properties survive a successful typed replacement
`crates/rdocx-oxml/src/glossary.rs:560`
`crates/rdocx-oxml/src/glossary.rs:565`
`crates/rdocx-oxml/src/glossary.rs:586`
`crates/rdocx-oxml/src/glossary.rs:591`
`crates/rdocx-oxml/src/glossary.rs:132`

Within `w:docPartPr`, a second direct modeled singleton such as `w:name` or
`w:description` is downgraded to extra raw XML instead of making the owner
malformed. On mutation, the first typed occurrence is rewritten and the second
occurrence is replayed from `extra_xml`. Reopen projects the first occurrence,
so the requested facade value compares equal and the duplicate schema-invalid
child is committed. This differs from the fail-closed handling already added
for duplicate category children and direct entry properties or bodies.

### D4, the legacy-form projection accepts malformed singleton values and cardinalities
`crates/rdocx-oxml/src/text.rs:1687`
`crates/rdocx-oxml/src/text.rs:1693`
`crates/rdocx-oxml/src/text.rs:1701`
`crates/rdocx-oxml/src/text.rs:1725`
`crates/rdocx-oxml/src/text.rs:1731`
`crates/rdocx-oxml/src/text.rs:1741`

The `w:ffData` projector has no seen-state for singleton metadata or
kind-specific singleton children. Repeated `w:name`, `w:enabled`,
`w:calcOnExit`, `w:default`, `w:checked`, or `w:result` elements silently
overwrite the prior projection. Invalid boolean tokens are also converted to
`true`, while invalid numeric tokens are treated as absent and fall through to
defaults. These malformed owners are published as typed forms and remain
selectable for mutation instead of setting `legacy_form_parse_error` and
failing closed.

### D5, note ownership uses attribute local names without namespace resolution
`crates/rdocx/src/field.rs:8287`
`crates/rdocx/src/field.rs:8290`
`crates/rdocx/src/field.rs:8291`
`crates/rdocx/src/field.rs:8296`

The package-story scanner decides whether a footnote or endnote is a normal
typed owner by matching `id` and `type` attribute local names only. A foreign
`x:id="2"` can therefore make a note with no valid `w:id` eligible, and a
foreign attribute occurring after the Word attribute can override its value.
The same problem applies to `type`. Forms under an invalid note, or under a
separator disguised by foreign attributes, can consequently acquire public
part-and-ordinal identities.

### D6, part-only deduplication hides a relationship-inappropriate story role
`crates/rdocx/src/field.rs:151`
`crates/rdocx/src/field.rs:157`
`crates/rdocx/src/field.rs:180`
`crates/rdocx/src/field.rs:183`

Story relationships are collected with their header, footer, footnote, or
endnote role, then deduplicated only by resolved part name before root
validation. If a target is reached by both a header and footer relationship,
one role is silently discarded. A `w:hdr` target retained as `Header` therefore
passes even though the same package graph also uses it as a footer target. The
new relationship-appropriate root check must validate every relationship role
or reject conflicting roles before part identity deduplication.

## Smells

None.

## Nitpicks

None.

## Not found

All thirty-five exact findings from passes 1 through 6 are closed. In
particular, optional glossary metadata removal now removes complete elements,
duplicate direct `w:docPartPr` and `w:docPartBody` children reject, each
retained package-story role requires its exact root, and glossary mutation uses
parser-recorded structural spans rather than lexical byte search. No additional
findings were found in selected-entry isolation, source-order identity for
nested fields and controls, cached-display mutation, unrelated raw-subtree
preservation, staged failure atomicity, panic safety, dependency direction,
public API shape, HLD file scope, or repository structure. All 19 focused
glossary unit tests and all 22 focused F-237 integration tests pass.
`git diff --check` passes for the implementation and HLD diff.
