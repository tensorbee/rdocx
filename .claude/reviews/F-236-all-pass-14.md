# F-236, all, pass 14

**Reviewed**: Pass-14 uncommitted implementation diff against `dbb5ab1`, excluding the thirteen earlier review artifacts, 7 files and 6,719 changed lines, comprising 6,713 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all thirteen prior reviews and their closure evidence
**Verdict**: 3 defects, 0 smells, 0 nitpicks

## Defects

### D1, relationship-less cross-kind children bypass owner collision checks
`crates/rdocx/src/embedded.rs:346`
`crates/rdocx/src/embedded.rs:348`
`crates/rdocx/src/embedded.rs:831`

OLE and control children are collected by two independent scans, and the
cross-kind collision check can compare only the references those scans emit.
A `w:object` containing one valid relationship-owned `o:OLEObject` plus one
relationship-less `w:control` therefore emits only the OLE range and passes
collision validation. Removing that OLE identity deletes the complete
`w:object`, including the extra control child. The symmetric case with a valid
control and relationship-less OLE child also passes. The same-kind
relationship-less cardinality regression and the two relationship-owned
cross-kind regression do not cover this combination, which still violates
ambiguous-owner rejection and raw subtree preservation.

### D2, reserved XML processing-instruction targets bypass declaration validation
`crates/rdocx/src/embedded.rs:1329`
`crates/rdocx/src/embedded.rs:1749`

Both scanners accept processing instructions before the document element.
`quick_xml` emits only a lowercase `<?xml ...?>` as `Event::Decl`, while a
case variant such as `<?XML version="1.0"?>` is emitted as `Event::PI`. XML
reserves every case-insensitive spelling of the `xml` processing-instruction
target, so that input is not a well-formed declaration or a legal processing
instruction. It nevertheless bypasses `validate_xml_declaration`, after which
the Word owner or ActiveX binary relationship remains actionable. Replacement
or removal can therefore commit against malformed relationship-owning XML.

### D3, malformed comment bodies are accepted in both trusted XML scopes
`crates/rdocx/src/embedded.rs:1258`
`crates/rdocx/src/embedded.rs:1368`

Both relationship-owning XML readers use the default `quick_xml` reader
configuration, whose comment well-formedness check is disabled. A comment body
containing the forbidden `--` sequence can therefore be emitted as a normal
comment event while an otherwise valid OLE owner or ActiveX relationship is
still inventoried and mutable. This leaves another path for malformed trusted
XML to pass the staged preflight. The declaration, document-type, entity, and
character-reference regressions do not exercise malformed comment grammar.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 13 D1 is closed for declaration events in both scanners. Missing,
duplicated, out-of-order, unsupported-version, invalid-encoding, and invalid
standalone declaration forms fail closed, while the supported XML 1.0 form
remains actionable. All prior findings are also closed for their cited
reproductions. D1 and D2 above are adjacent bypasses of the earlier owner and
document-grammar requirements rather than failures of those exact repairs.

No additional findings were found in relationship target normalization,
payload hashing and extraction, shared-target reachability, package or VBA
signature state and removal, staged failure atomicity, public API signatures,
panic safety, dependency direction, test structure, or repository structure.
All 54 focused `word_embedded_` regressions pass with default features and with
all features, and `cargo check -p rdocx --all-targets` passes.
