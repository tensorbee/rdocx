# F-237, all, pass 14

**Reviewed**: Full pass-14 uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the thirteen earlier review artifacts, 17 files and 8,073 changed lines, comprising 7,865 additions and 208 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`, plus all thirteen prior reviews and their closure evidence
**Verdict**: 8 defects, 0 smells, 0 nitpicks

## Defects

### D1, known legacy form elements at the wrong container level remain supported
`crates/rdocx-oxml/src/text.rs:1811`
`crates/rdocx-oxml/src/text.rs:1882`

The pass-13 repair rejects a known kind child when it belongs to a different
kind, but the direct `w:ffData` path ignores every known kind child that has no
common-property slot. The kind path likewise rejects only names in the
kind-child set, so it ignores common properties and nested kind containers.
For example, `<w:ffData><w:checked/><w:textInput/></w:ffData>` and
`<w:textInput><w:name w:val="nested"/></w:textInput>` both produce a valid
text-form projection. A typed value edit retains the misplaced WordprocessingML
element and commits after the same permissive reopen. Known form vocabulary at
the wrong schema level must fail closed just as cross-kind vocabulary now does.

### D2, a valid empty text-input form cannot be edited
`crates/rdocx-oxml/src/text.rs:4220`
`crates/rdocx-oxml/src/text.rs:4296`

An empty `w:textInput` is valid because all of its bounded children are
optional, and inventory correctly derives its current text from the cached
field result. The rewriter establishes kind scope only for a start event. It
replays `<w:textInput/>` unchanged on the empty-event path, then reaches EOF
without writing `w:default` and returns a missing-value error. The public
mutation therefore rejects a valid inventoried form instead of expanding the
empty kind container and persisting its new typed value atomically.

### D3, nested glossary element-only containers still accept character data
`crates/rdocx-oxml/src/glossary.rs:659`
`crates/rdocx-oxml/src/glossary.rs:777`
`crates/rdocx-oxml/src/glossary.rs:1357`

The pass-13 guard covers only `w:glossaryDocument` and its direct
`w:docParts`. Non-whitespace text, CDATA, and character references are still
ignored or retained inside the element-only `w:docPart`, `w:docPartPr`,
`w:category`, `w:types`, and `w:behaviors` containers. For example, a valid
entry with `text` directly inside `w:docPartPr` or `w:types` opens and exposes a
building block. Replacing another supported property retains that character
data, and staged reopen accepts it again, so schema-invalid glossary grammar
can be committed through the typed facade.

### D4, glossary and story relationships do not require normalized targets
`crates/rdocx/src/building_block.rs:99`
`crates/rdocx/src/building_block.rs:114`

The shared target validator checks only emptiness, a few characters, and
package-root escape depth. It explicitly accepts empty and `.` path components,
so targets such as `stories//header.xml` and `glossary/./glossary.xml` pass and
are normalized later by `resolve_rel_target` to an existing part. The glossary
and package-story collectors can consequently publish building-block or form
identities from a non-normalized relationship graph, contrary to the approved
normalized-internal-relationship contract.

### D5, forbidden comment bodies pass every trusted XML preflight
`crates/rdocx-oxml/src/glossary.rs:483`
`crates/rdocx/src/field.rs:7952`

Both complete-document preflights use the default reader configuration, whose
comment well-formedness check is disabled. A comment containing the forbidden
`--` sequence is therefore emitted normally in a glossary or package-owned
form story. Inventory still succeeds, and selected building-block or form
mutation preserves the malformed comment and can commit. The document grammar
checks added for declarations, document types, roots, and character placement
do not cover comment lexical grammar.

### D6, reserved and invalid processing-instruction targets are accepted
`crates/rdocx-oxml/src/glossary.rs:529`
`crates/rdocx/src/field.rs:8002`

Processing instructions fall through both document preflights without target
validation. Quick XML emits a case variant such as `<?XML version="1.0"?>` as
an instruction rather than a declaration, even though every case-insensitive
spelling of `xml` is reserved. It also emits `<?1producer value?>`, whose
target is not an XML Name. Either construct can precede or occur inside an
otherwise valid glossary or package story, after which its supported entry or
form remains mutable despite the source not being well-formed XML.

### D7, invalid XML names and unbound prefixes survive glossary replacement
`crates/rdocx-oxml/src/glossary.rs:336`
`crates/rdocx-oxml/src/glossary.rs:368`

The glossary reader uses a lexical prefix table only to decide which names are
WordprocessingML. It neither validates XML Names nor rejects a prefix with no
namespace binding. An unmodelled `<producer:item/>` with no corresponding
`xmlns:producer`, an element named `<1producer/>`, or an invalid attribute name
is treated as retained unsupported content around a valid entry. A selected
replacement then preserves and commits XML that is not namespace-well-formed
instead of rejecting the malformed glossary graph.

### D8, forbidden nested character references and literal XML characters are preserved
`crates/rdocx-oxml/src/glossary.rs:720`
`crates/rdocx-oxml/src/glossary.rs:735`

The document preflight does not validate character references or the XML 1.0
character repertoire within captured entry subtrees. The property parser then
retains an unmodelled child containing an undeclared general entity, a
forbidden reference such as `&#xFFFE;`, or a literal U+0001 or U+FFFE without
examining its content. Changing a different supported property replays those
bytes, and the same parser accepts them on staged reopen. This lets typed
replacement commit a glossary part that is not well-formed XML 1.0.

## Smells

None.

## Nitpicks

None.

## Not found

All 73 findings from passes 1 through 13 are closed for their cited
reproductions. In particular, known cross-kind children now reject within each
form-kind container, direct character data rejects within `w:ffData` and the
selected kind, form rewriting targets only the direct `w:ffData` child of the
begin field character, and character content rejects at the glossary root and
direct `w:docParts` level.

No additional findings were found in glossary content-type enforcement,
source-part and ordinal identity ordering, note ownership, nested field and
content-control traversal, cached-display updates, structural source-span
replacement, namespace context for supported changed slots, staged failure
atomicity, panic safety, public API structure, dependency direction, HLD file
scope, or repository structure. All 35 focused glossary unit tests, all 14
focused legacy-form unit tests, and all 36 focused facade integration tests
pass. All-target checks pass for `oxml-opc`, `rdocx-oxml`, and `rdocx`, and
`git diff --check 4ba8b6b` passes.
