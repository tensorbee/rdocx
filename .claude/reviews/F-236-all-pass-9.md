# F-236, all, pass 9

**Reviewed**: Pass-9 uncommitted implementation diff against `dbb5ab1`, excluding the eight earlier review artifacts, 7 files and 5,142 changed lines, comprising 5,136 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus all eight prior reviews and their closure evidence
**Verdict**: 5 defects, 0 smells, 0 nitpicks

## Defects

### D1, PreserveElements and PreserveAttributes do not require ignorable namespaces
`crates/rdocx/src/embedded.rs:1618`
`crates/rdocx/src/embedded.rs:1620`

The pass-8 repair correctly makes `ProcessContent` require a namespace named by
an in-scope `mc:Ignorable`, but the two preservation attributes call the same
QName-list validator with no ignorable-set constraint. MCE requires each
preserved element or attribute name to use a namespace identified as ignorable.
An ordinary ancestor with `xmlns:x="urn:producer" mc:PreserveElements="x:*"`
and no matching `mc:Ignorable="x"` is therefore accepted, and an embedded owner
below that nonconformant compatibility ancestry remains actionable. The new
wildcard regression proves only that a wildcard with a matching Ignorable
declaration is accepted.

### D2, MustUnderstand proves binding but never proves consumer understanding
`crates/rdocx/src/embedded.rs:1613`
`crates/rdocx/src/embedded.rs:1616`
`crates/rdocx/src/embedded.rs:1670`
`crates/rdocx/src/embedded.rs:1687`

`mc:MustUnderstand` is accepted whenever every token is a syntactically valid,
bound, non-MC namespace prefix. It does not require the namespace to belong to
the vocabulary understood by this scanner. For example, a story root can bind
`x` to an arbitrary producer URI, declare `mc:MustUnderstand="x"`, and expose a
normal OLE owner below it. MCE requires a consumer that does not understand a
named namespace to fail rather than process descendants, so inventory and
removal currently act through ancestry the implementation cannot claim to
understand.

### D3, ActiveX properties accept declarations in illegal document positions
`crates/rdocx/src/embedded.rs:1033`
`crates/rdocx/src/embedded.rs:1081`
`crates/rdocx/src/embedded.rs:1087`

The story scanner now rejects duplicate or trailing XML declarations, but the
separate ActiveX properties parser accepts every declaration whenever its
element depth is zero. It tracks neither whether a declaration was already seen
nor whether the root has closed or prolog content preceded it. A properties part
such as `<?xml version="1.0"?><ax:ocx .../><?xml version="1.0"?>` therefore
returns the binary relationship id and remains replaceable or removable even
though it is not one XML document. The package must fail closed before mutation
for malformed relationship-owning XML, not only for malformed Word story XML.

### D4, ActiveX properties root is not paired with its required content type
`crates/rdocx/src/embedded.rs:365`
`crates/rdocx/src/embedded.rs:373`
`crates/rdocx/src/embedded.rs:382`
`crates/rdocx/src/embedded.rs:383`

A `CONTROL` relationship target is checked for existence and parsed as an
`ax:ocx` document, but its content type is never resolved or required to be the
ActiveX XML properties type. A missing override or an arbitrary producer MIME
therefore still authorizes traversal to the binary relationship and later
deletion of the properties part. The pass-8 story repair established that MIME
and XML root are a joint package role. The same package-role check is missing
for the intermediate ActiveX owner part.

### D5, multiple main-document VBA project relationships are accepted
`crates/rdocx/src/embedded.rs:440`
`crates/rdocx/src/embedded.rs:442`
`crates/rdocx/src/embedded.rs:461`

Inventory loops over every `VBA_PROJECT` relationship in the main document
scope and emits each distinct relationship id. It never enforces the singleton
VBA project role. A malformed macro-enabled document with two VBA project
relationships is consequently reported as two valid executable identities, and
callers can replace either one, although the document relationship grammar
permits only one VBA project part. The implementation already enforces
singleton package signature origin, VBA signature, and ActiveX binary roles.
The same fail-closed cardinality is required here before inventory or mutation.

## Smells

None.

## Nitpicks

None.

## Not found

All eight pass-8 findings have concrete closure for their cited reproductions.
Relationship-less OLE and control children now count toward owner cardinality.
ActiveX deletion follows only its validated binary edge. MC rules are checked
on ordinary ancestors, `ProcessContent` requires an ignorable namespace, and
valid preservation wildcards remain actionable. DrawingML text-box discovery
requires the matching `a:graphicData/@uri`. Story MIME and root kind are paired,
and the Word story scanner rejects misplaced declarations and document types.

All earlier findings also remain closed for their named cases, including exact
target modes, owner and story ancestry, recursive VML and DrawingML text boxes,
shared VBA preflight, signature incoming edges and content types, complete
alternate-content child grammar, prohibited MC attributes, safe staged
publication, and raw byte preservation outside removed owner spans.

No additional findings were found in public signatures, payload hashing and
extraction, normalized target safety, package or VBA signature invalidation and
removal, failure atomicity, panic safety, dependency direction, or repository
structure. All 29 focused `word_embedded_` regressions pass.
`cargo check -p rdocx --all-targets` and `git diff --check dbb5ab1` pass.
