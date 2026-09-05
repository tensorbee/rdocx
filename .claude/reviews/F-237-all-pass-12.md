# F-237, all, pass 12

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the eleven earlier review artifacts, 17 files and 7,425 changed lines, comprising 7,325 additions and 100 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, legacy form name and format limits use the wrong schema maxima
`crates/rdocx-oxml/src/text.rs:1791`
`crates/rdocx-oxml/src/text.rs:1902`

The bounded form projection accepts 65 characters for `w:name` and 255 for
`w:textInput/w:format`. These schema types allow at most 20 and 64 characters,
respectively. A 21-character form name or 65-character format is therefore
inventoried as supported, survives a typed value edit as retained owner XML,
and passes the same permissive staged reopen. Pass-10 D4 remains incomplete,
and its boundary regression now asserts the incorrect larger limits.

### D2, help and status text values still have no schema bounds
`crates/rdocx-oxml/src/text.rs:1820`
`crates/rdocx-oxml/src/text.rs:1824`
`crates/rdocx-oxml/src/text.rs:2126`

The pass-11 repair calls a helper for both known leaves, but that helper reads
only `w:type`. It never reads or bounds the optional `w:val`, whose maximum is
255 characters for `w:helpText` and 140 for modern `w:statusText`. An owner
with a 256-character help value or a 141-character status value remains a
supported typed form, and mutation of its actual form value retains the
malformed leaf and commits after reopen. Pass-11 D2 is closed for its token
example but not for the complete known-leaf grammar.

### D3, a present result hides an out-of-schema drop-down default index
`crates/rdocx-oxml/src/text.rs:1720`
`crates/rdocx-oxml/src/text.rs:1954`

The parser stores `w:ddList/w:default` as an unrestricted `usize`, then checks
only the index selected by `result.or(default)`. The default schema type is
restricted to 0 through 24 independently of the current result. With a valid
`w:result="0"`, a `w:default="25"`, and 25 entries, the valid result masks the
invalid default and the owner is published. A later selected-value change
retains that invalid default and staged reopen accepts it again. This is a
second unclosed part of the pass-10 D4 facet requirement.

### D4, glossary entry types still bypass their schema enumeration
`crates/rdocx-oxml/src/glossary.rs:1267`
`crates/rdocx-oxml/src/glossary.rs:1339`
`crates/rdocx/src/building_block.rs:131`

`nested_values` validates enumerations only when reading `w:behavior`, so each
direct `w:type` is still accepted as an arbitrary string. The supported type
set is `none`, `normal`, `autoExp`, `toolbar`, `speller`, `formFld`, and
`bbPlcHdr`. A value such as `not-a-type` enters the typed properties, affects
the facade classification path by omission, and survives an unrelated block
replacement because staged reopen repeats the unrestricted parse. The
pass-11 gallery and behavior cases are closed, but the adjacent typed
enumeration remains open.

### D5, building-block GUID values are accepted and authored without validation
`crates/rdocx-oxml/src/glossary.rs:1270`
`crates/rdocx/src/building_block.rs:177`

The glossary parser projects `w:guid/@w:val` through the generic string reader,
and the public facade assigns any replacement string directly. The modeled
value requires the braced uppercase GUID lexical form, but values such as
`not-a-guid` pass both paths. Serialization writes the invalid value, staged
reopen accepts it through the same generic reader, and facade equality allows
the invalid replacement to commit instead of failing atomically.

### D6, checkbox forms do not require a size choice
`crates/rdocx-oxml/src/text.rs:1713`
`crates/rdocx-oxml/src/text.rs:1905`

The checkbox grammar requires one direct `w:size` or `w:sizeAuto` choice.
Projection validates either child only when one occurs and checks only that
the two alternatives are not combined. A `w:checkBox` containing only
`w:default`, or an empty checkbox, is therefore inventoried with a defaulted
Boolean value. Setting its checked value inserts or replaces only `w:checked`,
so staged reopen again accepts the still-missing size choice and commits
schema-invalid form XML.

## Smells

None.

## Nitpicks

None.

## Not found

Pass-11 D1 is closed in both glossary and package-story declaration paths by
requiring XML version 1.0. Pass-11 D3 is closed by nonempty direct glossary
entry, type, and behavior collections. Pass-11 D4 is closed for gallery and
behavior values at both parser and facade boundaries. The exact pass-1 through
pass-9 cases and pass-10 D1 through D3 remain concretely closed. Pass-10 D4
and pass-11 D2 remain partially open as D1 through D3 above.

No additional findings were found in relationship ownership and content
types, normalized package identities, story root and note scope, source-order
form identity, nested field and content-control traversal, cached-display
updates, selected-entry structural replacement, namespace and raw-subtree
preservation, staged failure atomicity beyond the defects above, panic safety,
public API structure, dependency direction, HLD file scope, or repository
structure. All 421 `rdocx-oxml` unit tests and its doc test pass, including all
32 glossary tests and 11 focused legacy-form tests. All 35 focused F-237
integration tests pass. `cargo check -p rdocx --all-targets`, `cargo fmt --all
--check`, and `git diff --check 4ba8b6b` pass.
