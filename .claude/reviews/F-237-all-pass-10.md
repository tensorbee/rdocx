# F-237, all, pass 10

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the nine earlier review artifacts, 17 files and 6,620 changed lines, comprising 6,535 additions and 85 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, XML declaration contents are never validated
`crates/rdocx-oxml/src/glossary.rs:466`
`crates/rdocx/src/field.rs:7965`

The pass-9 preflights now enforce declaration placement and cardinality, but
both accept every `Event::Decl` without examining its required version or its
remaining pseudo-attributes. Quick XML classifies `<?xml?>`, a declaration
whose first pseudo-attribute is `encoding`, duplicate pseudo-attributes, and
invalid `standalone` values as declaration events. Those malformed
declarations therefore pass both the glossary and package-story preflights,
after which a supported glossary entry or form can be inventoried and mutated
inside input that is not one valid XML document.

### D2, glossary category children remain optional when the schema requires both
`crates/rdocx-oxml/src/glossary.rs:1335`
`crates/rdocx-oxml/src/glossary.rs:197`
`crates/rdocx/src/building_block.rs:152`

`category_values` returns two independent options without requiring the direct
`w:name` followed by direct `w:gallery` pair. The writer likewise emits a
`w:category` when either value is present and silently omits the other child.
An input category containing only one required child is accepted, and callers
can manufacture the same invalid structure by replacing a block with only
`category` or only `gallery`. Staged reopen repeats the permissive parse, so the
facade commits schema-invalid glossary XML instead of rejecting atomically.

### D3, known legacy-form leaf properties skip required value and leaf-shape validation
`crates/rdocx-oxml/src/text.rs:1754`
`crates/rdocx-oxml/src/text.rs:1794`
`crates/rdocx-oxml/src/text.rs:1825`

The common `w:entryMacro` and `w:exitMacro` children and the kind-specific
`w:type`, `w:format`, `w:size`, and `w:sizeAuto` children are checked only for
singleton cardinality. Required `w:val` values, type and size tokens, and the
fact that these are leaf elements are not validated. For example,
`<w:textInput><w:type/></w:textInput>` or a `w:size` with child content remains
a selectable typed form. A later value change can commit that malformed owner
because the skipped property bytes are retained and the reopened projection
accepts them again.

### D4, bounded legacy-form ranges and list cardinality are not enforced
`crates/rdocx-oxml/src/text.rs:1813`
`crates/rdocx-oxml/src/text.rs:1872`
`crates/rdocx-oxml/src/text.rs:2003`

Numeric form values are accepted as any platform `usize`, and each direct
`w:listEntry` is pushed without a cardinality cap. This admits a zero or
out-of-schema `w:maxLength`, a drop-down with more than 25 entries, and strings
beyond the schema limits for modeled names, defaults, formats, macros, and
entries. Bounds are part of the planned malformed-owner and mutation gate, but
these owners are inventoried and can survive staged mutation as supported
forms.

## Smells

None.

## Nitpicks

None.

## Not found

All eight exact pass-9 findings have concrete closure for their cited cases.
Nested categories now reject duplicate and reversed modeled children. Empty
form-kind elements clear their scope, and known unprojected singleton children
reject duplicate occurrences and the checkbox size choice conflict. Glossary
and package-story declaration and document-type placement is checked before
the structural scan. Package stories require the role-specific content-type
override. Typed projected values reject duplicate expanded attributes and
unescape failures. Note attributes are decoded before signed normal-note
classification, and duplicate main-document relationship IDs reject story
discovery.

No additional findings were found in complete package-story root matching,
source-part and source-order identities, nested field or content-control
traversal, selected-value structural replacement, cached-display updates,
unrelated raw-subtree preservation, staged failure atomicity, panic safety,
public API shape, dependency direction, HLD file scope, or repository
structure. All 27 focused glossary unit tests, all 8 focused legacy-form unit
tests, and all 31 focused F-237 integration tests pass. `git diff --check
4ba8b6b` passes.
