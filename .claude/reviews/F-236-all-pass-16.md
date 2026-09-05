# F-236, all, pass 16

**Reviewed**: Pass-16 uncommitted implementation diff against `dbb5ab1`, excluding the fifteen earlier review artifacts, 7 files and 7,303 changed lines, comprising 7,297 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all fifteen prior reviews, the approved plan, progress record, affected HLD, and current focused test evidence
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, story document-type declarations are accepted without XML grammar validation
`crates/rdocx/src/embedded.rs:1957`
`crates/rdocx/src/embedded.rs:1962`
`crates/rdocx/tests/regression_test.rs:16717`

The story scanner validates only where a `DocType` event occurs and whether it
is repeated. It never validates the declaration's lexical keyword, root Name,
external identifier, or internal-subset grammar. `quick-xml` deliberately emits
a `DocType` event after a permissive balanced scan, including for a lowercase
`<!doctype w:hdr>` or an invalid root name such as `<!DOCTYPE 1producer>`.
Either form can precede an otherwise recognized `w:hdr` containing a valid OLE
owner, and the owner remains inventoried and removable even though the source
is not a well-formed XML 1.0 document. The existing regression covers only a
doctype after the root, so it does not exercise a malformed declaration in the
accepted prolog position. Rejecting document types in trusted Word story XML,
as the ActiveX scanner already does, or fully validating their XML grammar is
required before these owners can be actionable.

## Smells

None.

## Nitpicks

None.

## Not found

All four pass-15 findings are closed for their cited reproductions. Lone
relationship-less OLE and ActiveX owner children now fail owner finalization.
Processing-instruction targets are checked as XML Names and case variants of
the reserved `xml` target remain rejected. Start, empty, and end element QNames
are validated, used prefixes must resolve, namespace declarations obey the
reserved binding rules, and attributes are deduplicated by expanded name.
Invalid UTF-8 and forbidden literal XML 1.0 characters are rejected before
either trusted scanner acts.

No additional findings were found in relationship role and target validation,
owner source-range selection, text-box and markup-compatibility ancestry,
payload hashing and exact extraction, shared-target reachability, package or
VBA signature graph handling, signature policy semantics, staged mutation
atomicity, public API shape, panic safety, dependency direction, test-binary
structure, or repository structure. The current focused command
`cargo test -p rdocx --test regression_test word_embedded_` passes all 61 tests.
