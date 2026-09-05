# F-237, all, pass 2

**Reviewed**: Full uncommitted working tree against `4ba8b6b`, 16 implementation and HLD files and 2,717 changed lines, comprising 2,714 additions and 3 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, changing a glossary property drops unmodelled content owned by that property
`crates/rdocx-oxml/src/glossary.rs:131`
`crates/rdocx-oxml/src/glossary.rs:163`

An unchanged property slot reuses its retained XML, but a changed slot is
rebuilt from only the supported values. For example, changing the category of
`<w:category x:producer="keep"><w:name w:val="old"/><x:child/></w:category>`
drops both the producer attribute and foreign child. The same loss applies to
unmodelled attributes on changed scalar properties. This violates the
requirement to preserve every unmodelled attribute and child while changing
only the selected typed value.

### D2, an existing foreign `w` binding makes supported glossary edits fail
`crates/rdocx-oxml/src/glossary.rs:498`

The serializer treats the lexical presence of `xmlns:w` as proof that `w` is
bound to WordprocessingML. A valid prefix-tolerant entry can use `q` for Word
and bind `w` to a producer namespace on `q:docPartPr`. Changing a property then
emits a `w:` property under the foreign binding. The staged reopen cannot see
the requested value and rejects the replacement, so a valid inventoried entry
cannot be edited.

### D3, an empty glossary body invents default section properties
`crates/rdocx-oxml/src/glossary.rs:694`

The empty-element form `<w:docPartBody/>` is parsed through `CT_Body::new()`,
which supplies a default letter `sectPr`. The equivalent expanded form
`<w:docPartBody></w:docPartBody>` parses with no section properties. Inventory
therefore reports content that was never present, and a later body edit can
serialize the invented section subtree into the glossary entry.

### D4, a content-type default is accepted where the contract requires an override
`crates/rdocx/src/building_block.rs:85`

Glossary validation uses the resolved content type, which falls back to a
`Default` mapping by extension. A package with no glossary `Override` but a
default `.xml` mapping to the glossary MIME type is therefore accepted. The
affected packaging HLD explicitly requires the target's override to use the
Word glossary content type, so this malformed package should fail before
mutation.

### D5, legacy forms inside inline content controls are omitted
`crates/rdocx/src/field.rs:220`
`crates/rdocx/src/field.rs:305`

Both inventory and mutation inspect only `paragraph.runs`. A paragraph's typed
`content_controls` can own `SdtContent::Run` values containing begin field
characters and `w:ffData`, but those runs are never visited. Such forms are in
a supported Word story and are represented by the existing object model, yet
they are absent from inventory and cannot be mutated.

### D6, inserting a missing form value can emit an unbound `w` prefix
`crates/rdocx-oxml/src/text.rs:3585`
`crates/rdocx-oxml/src/text.rs:3684`

When the form-kind element uses the default Word namespace, its QName has no
prefix and the mutation code substitutes the literal prefix `w`. It does not
ensure that `w` is declared in scope. A valid default-namespace checkbox or
drop-down with no existing value and no `xmlns:w` binding therefore receives
an undeclared `w:checked` or `w:result` element. The staged reopen cannot
observe the requested value, so mutation of that valid prefix-tolerant form
fails.

## Smells

None.

## Nitpicks

None.

## Not found

All seven pass-1 defects have concrete code and regression-test closure. No
additional findings were found in panic safety, dependency direction, public
API structure, schema child ordering, staged commit atomicity, or HLD file
scope. The focused `rdocx-oxml` glossary and legacy-form unit tests pass, and
all 10 focused F-237 integration tests pass.
