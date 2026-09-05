# F-238, correctness, pass 1

**Reviewed**: working diff, 13 files, 1,396 additions and 8 deletions
**Verdict**: 3 defects, 1 smell, 0 nitpicks

## Defects

### D1, malformed relationship part names become ordinary parts

`crates/rdocx/src/flat_opc.rs:358`

A part with the relationship content type whose name is not the package
relationship name and does not map to a valid owner falls through to the
ordinary-part branch. The contract requires malformed relationship owners to
fail closed rather than enter `parts`.

### D2, XML part limits are checked after allocating the event

`crates/rdocx/src/flat_opc.rs:247`

The writer copies an untrusted XML event into its output buffer before the size
check at line 298. One oversized text or CDATA event can therefore allocate
beyond the caller's part and cumulative limits before rejection.

### D3, `text/xml` parts are emitted as binary

`crates/rdocx/src/flat_opc.rs:721`

The XML classifier accepts `application/xml`, relationship XML, and `+xml`
suffixes, but omits the registered `text/xml` media type. A valid XML part with
that content type is emitted as `pkg:binaryData`, contrary to the Flat OPC data
kind contract.

## Smells

### S1, MIME validation accepts extra slash separators

`crates/rdocx/src/flat_opc.rs:729`

Checking only `split_once('/')` accepts values such as `application/xml/extra`.
That weakens the canonical one-content-type boundary and makes classification
dependent on malformed MIME input.

## Nitpicks

None.

## Not found

No contract drift, unsafe public error expansion, untrusted arithmetic panic,
schema child-order fault, unmodelled binary loss, unnecessary trait, generic,
dependency, feature, or public package model was found.
