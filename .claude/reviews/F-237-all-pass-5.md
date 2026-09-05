# F-237, all, pass 5

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the four earlier review artifacts, 17 files and 4,543 changed lines, comprising 4,496 additions and 47 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, a valid empty docPartPr entry cannot be replaced
`crates/rdocx-oxml/src/glossary.rs:510`
`crates/rdocx-oxml/src/glossary.rs:120`

The optional-properties fix handles an absent `w:docPartPr`, but the equally
valid empty-element form `<w:docPartPr/>` retains neither a root start nor a
root end. Once a caller supplies the required nonempty facade name, changed
property output contains the new child without its `w:docPartPr` wrapper. A
staged reopen therefore cannot recover the replacement and rejects an
otherwise supported entry instead of editing it.

### D2, non-whitespace text outside the glossary root is accepted
`crates/rdocx-oxml/src/glossary.rs:358`
`crates/rdocx-oxml/src/glossary.rs:363`

The pass-4 fix rejects a second start or empty element after the root, but all
text events are still ignored. A part such as
`<w:glossaryDocument ...><w:docParts/></w:glossaryDocument>garbage` reaches EOF
with both root flags set and opens successfully. Non-whitespace character data
outside the document element is not well-formed XML, so this remains a
malformed-root glossary graph that must fail closed.

### D3, the required single docParts container is not validated
`crates/rdocx-oxml/src/glossary.rs:316`
`crates/rdocx-oxml/src/glossary.rs:363`

Root parsing recognizes entries when the current ancestry happens to be
`w:glossaryDocument/w:docParts`, but it never records the cardinality of that
container. An expanded glossary root with no `w:docParts`, or one with two
direct `w:docParts` children, therefore opens and exposes zero or concatenated
entries. The glossary grammar requires exactly one such container, and the HLD
requires malformed-root graphs to fail before mutation.

### D4, forms in block content of package-owned stories are omitted
`crates/rdocx/src/field.rs:74`
`crates/rdocx/src/field.rs:79`

Header, footer, footnote, and endnote inventory is limited to the top-level
paragraph vectors returned by the existing story models. Legal block content
such as a direct `w:tbl` or block `w:sdt` in a header, or a table inside a
normal footnote, is retained outside those vectors. A form in one of those
paragraphs is consequently absent from `legacy_form_fields` and cannot be
selected for mutation, despite the public contract covering supported
internal Word stories rather than only their direct paragraphs.

## Smells

None.

## Nitpicks

None.

## Not found

The exact findings from passes 1 through 4 are otherwise closed, including
absent optional `w:docPartPr`, a second top-level element, duplicate category
children, instruction source-order ordinals, interleaved nested controls, and
duplicate direct `w:ffData` rejection. No additional findings were found in
relationship target-mode and content-type policy, selected-entry and
form-value failure atomicity, namespace-aware matching, changed-slot schema
order, unrelated raw-subtree preservation, public API structure, dependency
direction, HLD file scope, or repository structure. All 14 glossary unit tests
and all 19 focused F-237 integration tests pass. `git diff --check` passes for
the implementation and HLD diff.
