# F-237, all, pass 11

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the ten earlier review artifacts, 17 files and 7,168 changed lines, comprising 7,054 additions and 114 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, XML 1.1 declarations are accepted for OOXML parts
`crates/rdocx-oxml/src/glossary.rs:532`
`crates/rdocx/src/field.rs:8031`

Both pass-10 declaration validators explicitly accept `version="1.1"` as well
as `version="1.0"`. OOXML package XML uses XML 1.0, so an XML 1.1 glossary or
package story must not enter the supported model. A glossary or header starting
with `<?xml version="1.1"?>` currently passes preflight, can expose supported
entries or forms, and can survive a staged mutation. Pass-10 D1 is therefore
only partially closed even though pseudo-attribute cardinality and order are
now checked.

### D2, help and status form leaves still accept invalid schema tokens
`crates/rdocx-oxml/src/text.rs:1818`
`crates/rdocx-oxml/src/text.rs:1821`

The known `w:helpText` and `w:statusText` leaves are marked only as singletons.
Their optional `w:type` is restricted to `text` or `autoText`, but no attribute
is read or validated. For example, a form containing
`<w:helpText w:type="bogus"/>` remains selectable, and a later typed value
change retains that malformed leaf and passes the same permissive reopen.
Pass-10 D3 and D4 added checks for the other known leaves and their facets, but
did not complete token validation for these two leaves.

### D3, required glossary collections are allowed to be empty
`crates/rdocx-oxml/src/glossary.rs:439`
`crates/rdocx-oxml/src/glossary.rs:1295`

The glossary root check requires exactly one direct `w:docParts` container but
never requires that container to contain a `w:docPart`. The shared nested-value
reader likewise returns an empty vector for empty `w:types` and `w:behaviors`
containers even though each container requires at least one respective child.
Consequently `<w:docParts/>`, `<w:types/>`, and `<w:behaviors/>` are accepted as
supported glossary grammar rather than rejected as malformed input. The first
case opens as an empty inventory, while the latter cases can survive an
otherwise valid selected-entry replacement.

### D4, building-block replacement can author invalid glossary enum values
`crates/rdocx-oxml/src/glossary.rs:1306`
`crates/rdocx-oxml/src/glossary.rs:1420`
`crates/rdocx/src/building_block.rs:157`

Typed glossary values are decoded as arbitrary strings. Neither the parser nor
`apply_block` validates the schema enumerations for `w:gallery` or
`w:behavior`. A caller can therefore replace an entry with a gallery such as
`not-a-gallery` or a behavior other than `content`, `p`, or `pg`. Serialization
writes that value, and staged reopen accepts it through the same unrestricted
string path, so the mutation commits schema-invalid OOXML despite the promised
validation and atomicity gate.

## Smells

None.

## Nitpicks

None.

## Not found

All pass-1 through pass-9 findings have concrete closure for their cited
cases. Pass-10 D2 is closed by requiring the complete ordered category pair.
The declaration validators now enforce pseudo-attribute presence, order,
cardinality, encoding-name grammar, and standalone tokens apart from D1 above.
The legacy-form reader now validates required values, leaf shape, the projected
type and boolean tokens, string bounds, numeric bounds, selection bounds, and
drop-down cardinality apart from D2 above.

No additional findings were found in relationship ownership and content types,
normalized part identity, story root and note ownership, source-order form
identity, nested field and content-control traversal, selected-value insertion
order, cached-display updates, raw subtree and namespace preservation,
structural spans, staged failure atomicity, panic safety, public API shape,
dependency direction, HLD file scope, tests, or repository structure. All 29
focused glossary unit tests, all 10 focused legacy-form unit tests, and all 33
focused F-237 integration tests pass. `cargo check -p rdocx --all-targets`,
`cargo fmt --all --check`, and `git diff --check 4ba8b6b` pass.
