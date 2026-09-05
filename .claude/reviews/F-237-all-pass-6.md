# F-237, all, pass 6

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the five earlier review artifacts, 17 files and 4,984 changed lines, comprising 4,937 additions and 47 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, removing typed glossary metadata leaves schema-invalid valueless elements
`crates/rdocx-oxml/src/glossary.rs:722`
`crates/rdocx-oxml/src/glossary.rs:788`
`crates/rdocx-oxml/src/glossary.rs:827`
`crates/rdocx/src/building_block.rs:152`

When a replacement changes an existing optional scalar to `None`, the property
patcher removes only its `w:val` attribute and retains the element. The same
path is used for category children, and surplus repeated `w:type` or
`w:behavior` children are passed a missing desired value instead of being
removed. Removing a description, category, gallery, GUID, or behavior, or
changing an AutoText entry to a building block, can therefore emit elements
such as `<w:description/>` or `<w:type/>` even though these property elements
require their value attribute. Reopen projects the valueless element as absent,
so the facade equality check accepts and commits the schema-invalid glossary.

### D2, duplicate direct entry properties and bodies are accepted ambiguously
`crates/rdocx-oxml/src/glossary.rs:439`
`crates/rdocx-oxml/src/glossary.rs:449`
`crates/rdocx-oxml/src/glossary.rs:451`
`crates/rdocx-oxml/src/glossary.rs:478`

The entry parser stores direct `w:docPartPr` and `w:docPartBody` children in
single options but never counts them. Each later occurrence silently replaces
the earlier captured source. A `w:docPart` with two direct property elements or
two direct bodies consequently opens, exposes only the last occurrence, and can
be mutated successfully while the duplicate remains in the saved entry. These
children have maximum cardinality one, so the malformed and ambiguous entry
must fail before inventory or mutation rather than acquire a selectable
identity.

### D3, package form stories do not enforce one relationship-appropriate root
`crates/rdocx/src/field.rs:55`
`crates/rdocx/src/field.rs:7818`
`crates/rdocx/src/field.rs:7839`
`crates/rdocx/src/field.rs:7841`

Header and footer relationships are collapsed to one scanner kind that accepts
either `w:hdr` or `w:ftr`, and the scanner records neither root presence nor
root closure. A header relationship targeting a `w:ftr` document is therefore
inventoried and mutated as a supported header. Quick XML also permits two
top-level roots, and after the first closes the empty stack makes a second root
eligible for the same scan. EOF validates only that the stack is empty, so the
forms from both roots receive ordinals and can be edited in a part that is not
a well-formed, relationship-appropriate Word story.

### D4, lexical glossary replacement can target an identical byte sequence in a comment
`crates/rdocx-oxml/src/glossary.rs:269`
`crates/rdocx-oxml/src/glossary.rs:274`
`crates/rdocx-oxml/src/glossary.rs:278`
`crates/rdocx-oxml/src/glossary.rs:413`

Glossary serialization relocates retained entries, properties, and bodies by
searching for their raw byte strings rather than retaining structural spans. A
valid comment before an entry can contain the exact bytes of that entry, or a
comment inside an entry can contain the exact bytes of its later property or
body child. The first lexical match is then rewritten inside the comment while
the typed element remains unchanged. Staged reopen rejects the requested value,
so failure stays atomic, but a valid supported entry becomes uneditable and the
selected-entry replacement contract is not met.

## Smells

None.

## Nitpicks

None.

## Not found

All thirty-one exact findings from passes 1 through 5 are closed, including the
four pass-5 cases for empty `w:docPartPr`, content outside the glossary root,
exactly one direct `w:docParts`, and block forms in package stories. No
additional findings were found in internal relationship target safety,
explicit glossary content-type policy, legacy-form kind and bounds validation,
source-order identity for nested fields and inline controls, form raw-subtree
preservation, failure atomicity, panic safety, dependency direction, public API
shape, HLD file scope, or repository structure. All 16 focused glossary unit
tests and all 21 focused F-237 integration tests pass. `git diff --check` passes
for the implementation and HLD diff.
