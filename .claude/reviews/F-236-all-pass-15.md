# F-236, all, pass 15

**Reviewed**: Pass-15 uncommitted implementation diff against `dbb5ab1`, excluding the fourteen earlier review artifacts, 7 files and 6,863 changed lines, comprising 6,857 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all fourteen prior reviews and their closure evidence
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, a lone relationship-less embedded owner is silently ignored
`crates/rdocx/src/embedded.rs:1647`
`crates/rdocx/src/embedded.rs:1681`

The owner-finalization paths reject multiple children and mixed child kinds, but
they emit a reference only when an `r:id` was collected. A schema-positioned
`w:object` containing exactly one relationship-less `o:OLEObject`, or a
`w:object` or `w:pict` containing exactly one relationship-less `w:control`,
therefore produces neither an inventory item nor an error. A package can include
such a malformed owner in one supported story and a valid embedded item in
another. Inventory omits the malformed owner, and replacement or removal of the
valid item can commit despite the contract requiring missing relationship
metadata to fail closed. The pass-8 same-kind and pass-14 cross-kind repairs
cover a relationship-less child only when another owner child exposes the
cardinality or collision.

### D2, processing instructions with invalid XML names remain actionable
`crates/rdocx/src/embedded.rs:1330`
`crates/rdocx/src/embedded.rs:1754`

The scanners reject only the case-insensitive reserved target `xml`. The parser
also emits an instruction such as `<?1producer value?>` as `Event::PI`, even
though a processing-instruction target must be an XML Name. That event passes
the catch-all handling before or inside a valid story root and inside the
ActiveX root. The relationship-owning XML is therefore still inventoried, and a
mutation of its item or another item can commit against XML that is not
well-formed. The pass-14 regression proves the reserved-name rule but not the
general target-name grammar.

### D3, invalid element and attribute names and unbound prefixes are accepted
`crates/rdocx/src/embedded.rs:1284`
`crates/rdocx/src/embedded.rs:1425`

Both trusted scanners accept arbitrary nested start and empty events without
validating XML Names or requiring each prefix to resolve. The namespace reader
reports an undeclared prefix as unknown rather than an error, and it tokenizes
names such as `1producer` even though they are not XML Names. A valid story or
ActiveX root can therefore contain `<producer:item/>` without an `xmlns`
binding, `<1producer/>`, or an invalid attribute name alongside an otherwise
valid owner. Inventory and mutation still succeed instead of failing closed on
XML that is not namespace-well-formed.

### D4, forbidden literal XML 1.0 characters bypass character validation
`crates/rdocx/src/embedded.rs:1358`
`crates/rdocx/src/embedded.rs:1792`

Character legality is checked for `Event::GeneralRef`, but ordinary text and
CDATA inside the trusted roots fall through without validation. The parser
emits literal forbidden characters such as U+0001 or U+FFFE as content events.
An ActiveX properties root or supported Word story containing one of those
literal characters plus a valid relationship owner therefore remains
actionable. The pass-12 regression rejects the equivalent forbidden numeric
character references, but the literal form can still permit an embedded
mutation to commit against XML that is not well-formed XML 1.0.

## Smells

None.

## Nitpicks

None.

## Not found

All 58 findings from passes 1 through 14 are closed for their cited
reproductions. In particular, relationship-less cross-kind children now
participate in the owner collision check, case variants of the reserved `xml`
processing-instruction target fail closed in both scanners, and forbidden
double-hyphen comment bodies are rejected. Signature parts now also require
their exact package, legacy VBA, or Agile VBA content type before mutation.

No additional findings were found in relationship target normalization,
payload hashing and extraction, shared-target reachability, package or VBA
signature state and removal, staged failure atomicity, public API signatures,
panic safety, dependency direction, test structure, or repository structure.
All 57 focused `word_embedded_` regressions pass with default features and with
all features, and `cargo check -p rdocx --all-targets` passes.
