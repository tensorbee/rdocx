# F-237, all, pass 4

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the three earlier review artifacts, 17 files and 4,221 changed lines, comprising 4,175 additions and 46 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, a schema-valid entry without docPartPr prevents the document from opening
`crates/rdocx-oxml/src/glossary.rs:408`
`crates/rdocx/src/building_block.rs:130`

`w:docPartPr` is optional on `w:docPart`, but `parse_doc_part` requires a
captured properties element. A glossary containing a valid entry with only its
required `w:docPartBody` therefore makes `Document::from_bytes` fail. The facade
already maps an absent typed name to the empty string, so the entry should stay
preserved with absent bounded metadata instead of making an otherwise readable
document reject its relationship-owned glossary.

### D2, a second top-level element is accepted as part of a glossary document
`crates/rdocx-oxml/src/glossary.rs:290`
`crates/rdocx-oxml/src/glossary.rs:329`

After the closing `w:glossaryDocument`, `root_seen` remains true and the stack
is empty. A following start element therefore takes the generic stack branch,
and EOF returns success because only `root_seen` is checked. Bytes such as
`<w:glossaryDocument ...><w:docParts/></w:glossaryDocument><x:other/>` are not a
single-root XML document, yet the package opens instead of failing the
malformed-root graph as required.

### D3, changing a category also changes an unmodelled duplicate child
`crates/rdocx-oxml/src/glossary.rs:692`
`crates/rdocx-oxml/src/glossary.rs:1003`

Category projection takes only the first direct `w:name`, so a later duplicate
does not contribute to the typed value. The patcher nevertheless treats every
direct child with that local name as the modeled child and rewrites all of
them. Changing the supported category in a producer entry with two `w:name`
children therefore changes the retained duplicate too, instead of preserving
that unmodelled subtree byte-exactly.

### D4, nested form ordinals follow storage grouping rather than instruction source order
`crates/rdocx/src/field.rs:293`
`crates/rdocx/src/field.rs:405`

Both inventory and mutation traverse all nested positional arguments before
all nested switch arguments. The field parser retains a separate source-order
map precisely because those categories can be interleaved. If a nested form
owned by a switch precedes a nested positional form in the instruction, the
facade reports and mutates them in the opposite order. The advertised
source-part plus source-order ordinal identity is therefore wrong even though
the artificial order remains stable after reopen.

### D5, interleaved nested inline controls are assigned out-of-source-order ordinals
`crates/rdocx/src/field.rs:265`
`crates/rdocx/src/field.rs:268`
`crates/rdocx/src/field.rs:488`

An inline control first projects every field from all of its direct runs, then
walks nested controls in a second pass. For content ordered as direct form A,
nested inline control containing form B, direct form C, inventory returns A,
C, B. Mutation uses the same direct-field count before recursing, so ordinal 1
changes C rather than the source-order field B. This affects main and
package-owned stories because they share this control walk.

### D6, duplicate ffData owners are exposed and mutated instead of rejected
`crates/rdocx-oxml/src/text.rs:1516`
`crates/rdocx-oxml/src/text.rs:3579`
`crates/rdocx-oxml/src/text.rs:3620`

Legacy-form projection stops as soon as the first direct `w:ffData` closes, so
it cannot detect a second `w:ffData` on the same begin `w:fldChar`. The field is
still inventoried from the first subtree. On mutation, the raw-source rewriter
enters every `w:ffData` and patches matching direct value elements even after
the first value was written. A schema-invalid owner can consequently commit a
mutation that also changes the otherwise retained duplicate subtree. This
violates the planned malformed-owner rejection and exact raw preservation.

## Smells

None.

## Nitpicks

None.

## Not found

All twenty-one exact reproduction cases from passes 1 through 3 have concrete
code and named regression closure. No additional findings were found in panic
safety, relationship target-mode and content-type policy, dependency direction,
public API shape, HLD file scope, or repository structure. All 391
`rdocx-oxml` unit tests and its doc test pass. All 16 focused F-237 integration
tests pass, and `git diff --check` passes. The full `rdocx` integration binary
reported 147 passes, one ignored pinned Word oracle, and the existing
environment-dependent LibreOffice ODT oracle failure.
