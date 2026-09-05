# F-237, all, pass 1

**Reviewed**: Uncommitted working tree against `4ba8b6b`, 16 files and 2,309 changed lines, comprising 2,306 additions and 3 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 7 defects, 0 smells, 0 nitpicks

## Defects

### D1, an explicitly internal glossary relationship is rejected
`crates/rdocx/src/building_block.rs:71`

The glossary loader rejects every relationship with a present `TargetMode`,
including the valid OPC spelling `TargetMode="Internal"`. The contract requires
exactly one internal relationship, and explicit `Internal` has the same
semantics as an absent target mode. A conforming producer that writes the
explicit value therefore makes the whole document fail to open.

### D2, form story relationships are not proven safe and internal
`crates/rdocx/src/field.rs:7600`

The story discovery helper excludes only the exact value `External`, then
normalizes the target without validating its mode or checking for package-root
escape. A header, footer, footnote, or endnote relationship with an unknown
mode such as `ProducerDefined`, or a target such as `../../outside.xml`, is
treated as an internal Word story when the resolved package part exists. Form
inventory and mutation can consequently act through a relationship that does
not satisfy the supported-internal-story contract instead of failing closed.

### D3, changing a glossary property drops unmodelled root attributes
`crates/rdocx-oxml/src/glossary.rs:101`

Once any supported property changes, `CT_DocPartPr` constructs a new root with
only `xmlns:w`. It does not retain attributes from the parsed `w:docPartPr`
root. For example, replacing the description of an entry whose properties root
has `x:producer="keep"` silently deletes that attribute. This violates the
explicit requirement to preserve every unmodelled attribute while rewriting
only changed typed slots.

### D4, a metadata-only replacement rewrites the unchanged building-block body
`crates/rdocx-oxml/src/glossary.rs:216`

Whenever either properties or body differs, `CT_DocPart::to_xml` serializes
both slots. A replacement that changes only `description` therefore
canonicalizes the unchanged `w:docPartBody`, changing its prefix and whitespace
and dropping any root attributes or namespace declarations that are not part
of `CT_Body`. The contract says fixed-prefix output applies only to changed
typed slots and requires unrelated body XML and whitespace boundaries to
remain exact.

### D5, inventoried forms in table-owned content controls cannot be selected reliably
`crates/rdocx/src/field.rs:328`

Inventory traverses table-level and row-level content controls in source order
through `collect_table_paragraphs`, but the mutation walker iterates only
`table.rows` and `row.cells`. A form inside one of those content controls is
assigned an ordinal by `legacy_form_fields`, then
`set_legacy_form_field_value` skips it and either reports the valid identity as
stale or applies the ordinal to a later form. The staged value check prevents a
bad commit in some cases, but the advertised inventoried identity is not
editable as required.

### D6, namespace declarations leak from one sibling into the next
`crates/rdocx-oxml/src/glossary.rs:235`
`crates/rdocx-oxml/src/text.rs:1455`

Both new parsers replace one shared prefix vector on every start or empty event
and never restore the previous bindings when that element ends. A valid
unmodelled sibling such as `<x:raw xmlns:q="urn:producer"/>` therefore changes
how a later `q:docPart` or `q:ddList` is interpreted even when `q` is bound to
WordprocessingML in the enclosing scope. Glossary entries or form kinds can be
silently missed, contrary to namespace-aware, prefix-tolerant parsing and
unmodelled-subtree preservation.

### D7, form value patching can overwrite an unrelated attribute
`crates/rdocx-oxml/src/text.rs:3690`

The value writer locates the serialized `w:val` attribute by searching the
whole start tag for the attribute key bytes, rather than using the parsed
attribute's position. If an earlier unmodelled attribute value contains that
text and another attribute follows it, for example
`x:a="w:val" x:b="keep" w:val="0"`, the subsequent equals-sign search selects
`x:b` and replaces its value instead. The requested form value remains stale
while unrelated producer metadata is corrupted, violating both typed mutation
and exact attribute preservation.

## Smells

None.

## Nitpicks

None.

## Not found

No additional findings in public API shape, dependency direction, schema child
ordering of newly emitted supported glossary properties, mutation commit
atomicity, panic safety, HLD scope, or repository structure. The focused
`rdocx-oxml` glossary unit test passes, and all 7 focused F-237 integration tests
pass.
