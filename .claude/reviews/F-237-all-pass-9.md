# F-237, all, pass 9

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the eight earlier review artifacts, 17 files and 5,917 changed lines, comprising 5,837 additions and 80 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 8 defects, 0 smells, 0 nitpicks

## Defects

### D1, nested category cardinality and order are not validated on parse
`crates/rdocx-oxml/src/glossary.rs:1138`
`crates/rdocx-oxml/src/glossary.rs:1170`

The pass-8 property-order gate validates the direct `w:docPartPr` sequence,
but `w:category` is projected by collecting all direct `w:name` and
`w:gallery` values independently and taking only the first name and gallery.
It therefore accepts duplicate names or galleries and accepts `w:gallery`
before `w:name`. Changing an unrelated property such as description reuses the
unchanged category bytes, staged reopen produces the same first-value
projection, and replacement commits the schema-invalid nested property. The
existing duplicate-category test catches only a change to the category slot
itself.

### D2, an empty form-kind element leaves descendant projection active
`crates/rdocx-oxml/src/text.rs:1591`
`crates/rdocx-oxml/src/text.rs:1616`
`crates/rdocx-oxml/src/text.rs:1938`

The empty-element path calls `set_legacy_form_kind`, which records a
`kind_depth`, but no end event exists to clear that depth. A valid empty
`w:checkBox` followed by an unmodelled wrapper containing `w:checked` therefore
projects that nested descendant as though it were a direct checkbox child.
Inventory reports a checked value sourced from the unsupported subtree instead
of preserving the subtree without typed meaning, and subsequent mutation can
act on an identity whose value was established at the wrong schema position.

### D3, known unprojected form singletons can be repeated or combined
`crates/rdocx-oxml/src/text.rs:1846`
`crates/rdocx-oxml/src/text.rs:1860`
`crates/rdocx-oxml/src/text.rs:1884`

The order tables recognize common children such as `w:entryMacro` and
kind-specific children such as `w:type`, `w:format`, `w:size`, and
`w:sizeAuto`, but the validator rejects only a decreasing slot. It accepts a
repeated singleton at the same slot, and it accepts both mutually exclusive
checkbox size choices because they share slot zero. These malformed owners
remain selectable. A value change can then commit while retaining the invalid
known children because only the bounded value children have explicit seen
state.

### D4, glossary and package-story scanners accept misplaced XML declarations
`crates/rdocx-oxml/src/glossary.rs:386`
`crates/rdocx-oxml/src/glossary.rs:405`
`crates/rdocx/src/field.rs:7864`
`crates/rdocx/src/field.rs:7895`

Both document scanners validate text, CDATA, character references, roots, and
closure, but every declaration and document-type event falls through the
wildcard arm. A second XML declaration after a closed glossary or header root,
or a declaration nested inside that root, therefore reaches a successful EOF.
The malformed glossary can be replaced and the malformed story can expose and
mutate forms even though neither input is one valid XML document.

### D5, package-owned form stories do not validate their content types
`crates/rdocx/src/field.rs:170`
`crates/rdocx/src/field.rs:193`

Story discovery validates relationship type, exact target mode, target safety,
part existence, and root kind, but it never requires the target part's content
type to match header, footer, footnotes, or endnotes. A header relationship can
therefore target a part overridden as `application/octet-stream`, or one
resolved only through the generic XML default, and a `w:hdr` byte stream still
becomes an editable supported story. Relationship role, MIME, and root must
agree before a part-scoped form identity is published.

### D6, typed form and glossary values accept ambiguous or malformed attributes
`crates/rdocx-oxml/src/glossary.rs:1211`
`crates/rdocx-oxml/src/text.rs:6545`

Both value readers return as soon as they find the first matching Word
attribute. They do not exhaust the iterator to detect a duplicate `w:val`, or
two differently prefixed attributes with the same expanded Word name. They
also turn an entity-unescape failure into the original raw text. A name,
description, text default, or list entry with duplicate expanded values or an
undefined entity is consequently published as a supported typed value rather
than making its owner malformed, and an unrelated mutation can retain and
commit the ambiguous source.

### D7, encoded note attributes are not decoded before ownership checks
`crates/rdocx/src/field.rs:8327`
`crates/rdocx/src/field.rs:8337`

The repaired note classifier resolves the attribute namespace correctly, but
it parses the raw attribute bytes without XML entity decoding. Legal values
such as `w:id="&#49;"` and `w:type="norm&#97;l"` are therefore rejected even
though their decoded values are `1` and `normal`. Forms under those supported
normal notes disappear from inventory and cannot be mutated, so signed-ID
support remains dependent on one lexical spelling of the value.

### D8, duplicate package-story relationship identities are not rejected
`crates/rdocx/src/field.rs:157`
`crates/rdocx/src/field.rs:196`
`crates/rdocx/src/field.rs:199`

Unlike glossary loading, form-story discovery never validates relationship ID
uniqueness. Two header relationships with the same ID and distinct valid
targets are both collected, while duplicate IDs to one target are silently
hidden by part-name deduplication. Either malformed OPC graph can therefore
publish editable form identities. The collector must reject duplicate story
relationship identities before role collection or deduplication.

## Smells

None.

## Nitpicks

None.

## Not found

All six exact pass-8 findings have concrete closure for their cited cases.
Direct modeled glossary properties now enforce their declared order and
required values. `w:ffData` and the selected form-kind children reject the
tested out-of-order sequences, and missing, contradictory, and out-of-range
form kinds fail closed. Explicitly normal footnotes and endnotes accept zero
and negative IDs. Multiple footnotes or endnotes relationships reject before
inventory or mutation.

No additional findings were found in source-order identity for the covered
nested fields and controls, cached-display updates, selected-entry structural
span replacement, unrelated raw-subtree preservation, staged failure
atomicity, panic safety, public API shape, dependency direction, HLD file
scope, or repository structure. All 24 focused glossary unit tests, all 6
focused legacy-form unit tests, and all 27 focused F-237 integration tests
pass. `git diff --check 4ba8b6b` passes.
