# F-236, all, pass 13

**Reviewed**: Pass-13 uncommitted implementation diff against `dbb5ab1`, excluding the twelve earlier review artifacts, 7 files and 6,535 changed lines, comprising 6,529 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all twelve prior reviews and their closure evidence
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, malformed XML declaration contents are accepted as well formed
`crates/rdocx/src/embedded.rs:1268`
`crates/rdocx/src/embedded.rs:1687`

Both relationship-owning XML scanners validate where an `Event::Decl` appears,
but discard the declaration without validating its contents. `quick_xml` emits
that event for declarations such as `<?xml?>` and
`<?xml encoding="UTF-8"?>`, even though they lack the required leading version,
and likewise does not make the scanner check duplicate pseudo-attributes or the
`standalone` token. A story part or ActiveX properties part with one of these
malformed declarations and otherwise valid ownership therefore remains
actionable. Inventory succeeds and replacement or removal can mutate a package
whose trusted ownership XML is not well formed, rather than failing closed
before mutation.

## Smells

None.

## Nitpicks

None.

## Not found

All 54 findings from passes 1 through 12 are closed for their cited
reproductions. In particular, empty MC rule lists remain actionable, effectively
ignorable extension children no longer invalidate `mc:AlternateContent`,
qualified attributes on MC elements are constrained to understood or ignorable
namespaces, general and character references obey XML 1.0 and AlternateContent
grammar, grouped WordprocessingCanvas paths use `wpg:wgp`, and DrawingML text
boxes beneath `w:object` are discovered and mutated through the complete owner
path. The earlier graph, signature MIME and incoming-edge, relationship
singleton, normalized identity, root anchoring, owner cardinality, story MIME,
MC vocabulary, raw preservation, and nested text-box cases also remain closed.

No additional findings were found in graph reachability, signature state or
removal, failure atomicity, schema ownership and child order, raw XML range
preservation, hashing and exact extraction, panic safety, dependency direction,
public API shape, tests, or repository structure. All 52 focused
`word_embedded_` regressions pass with default features and with all features.
The complete regression binary passes with 338 tests and 3 ignored tests.
`cargo check -p rdocx --all-targets`, `cargo fmt --all --check`, and
`git diff --check dbb5ab1` pass.
