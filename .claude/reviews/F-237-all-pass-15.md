# F-237, all, pass 15

**Reviewed**: Full pass-15 uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the fourteen earlier review artifacts, 17 files and 8,884 changed lines, comprising 8,680 additions and 204 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`, plus all fourteen prior reviews and their closure evidence
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, glossary and package-story document types bypass XML grammar validation
`crates/rdocx-oxml/src/glossary.rs:512`
`crates/rdocx/src/field.rs:7979`
`crates/rdocx-oxml/src/glossary.rs:2299`
`crates/rdocx/tests/integration_test.rs:7170`

Both trusted complete-document preflights validate only the position and
cardinality of a `DocType` event. They never validate the declaration's
lexical keyword, root Name, external identifier, or internal-subset grammar.
Quick XML emits a document-type event after a permissive balanced scan, so a
lowercase `<!doctype ...>` or a declaration with an invalid root name such as
`<!DOCTYPE 1producer>` can precede an otherwise recognized glossary or header
root. The glossary entry or package-story form remains inventoried and can be
rewritten, after which the equally permissive staged reopen accepts the
malformed XML again. The existing tests cover only misplaced or duplicate
document types, not malformed declarations in the accepted prolog position.
Rejecting document types in these trusted OOXML parts, or fully validating the
XML 1.0 document-type grammar, is required before their typed owners can be
actionable.

## Smells

None.

## Nitpicks

None.

## Not found

All eighty-one findings from passes 1 through 14 are closed for their cited
reproductions. In particular, known form vocabulary now rejects at every
wrong schema level, empty text-input kinds expand with a locally bound value
prefix, and nested modeled glossary containers reject non-whitespace
character data. Glossary and package-story relationship targets now require
normalized internal pack-URI spelling. Their complete-source preflights reject
forbidden comments, invalid or reserved processing-instruction targets,
invalid QNames, unbound and reserved prefixes, duplicate expanded-name
attributes, undeclared or forbidden references, invalid UTF-8, and literal
characters outside the XML 1.0 repertoire, apart from D1.

No additional findings were found in relationship role or content-type
validation, source-part and source-order ordinal identity, note ownership,
nested field and content-control traversal, form-kind facets, selected-value
insertion order, cached-display updates, glossary child order and cardinality,
selected-entry structural replacement, namespace context, raw-subtree or
whitespace preservation, staged failure atomicity, panic safety, public API
shape, dependency direction, HLD scope, test-binary structure, or repository
structure. All 40 focused glossary unit tests, all 15 focused legacy-form unit
tests, and all 37 focused facade integration tests pass. All-target checks pass
for `oxml-opc`, `rdocx-oxml`, and `rdocx`. `cargo fmt --all --check` and
`git diff --check 4ba8b6b` pass.
