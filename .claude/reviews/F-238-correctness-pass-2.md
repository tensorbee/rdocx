# F-238, correctness, pass 2

**Reviewed**: remediated working diff, 14 files, 1,493 additions and 8 deletions
**Verdict**: 2 defects, 1 smell, 0 nitpicks

## Defects

### D1, an empty binary part has no accepted representation

`crates/rdocx/src/flat_opc.rs:184`

The part reader accepts only a start event for `pkg:binaryData`. The equivalent
self-closing form is rejected even though it is the natural representation of
a valid empty opaque part and still satisfies the exactly-one-data-element
contract.

### D2, explicit empty relationship elements are rejected

`crates/rdocx/src/flat_opc.rs:643`

Relationship parsing accepts only self-closing `Relationship` events. XML also
permits the schema-empty element to use separate start and end tags. A strict
reader must validate that form's empty content rather than reject a valid
relationship document because of lexical style.

## Smells

### S1, class conversion does not directly prove signature invalidation

`crates/rdocx/tests/integration_test.rs:184`

The conversion comparison proves package preservation and content-type
isolation, but it does not include a retained package-signature graph or assert
the deterministic invalidation marker. Removing the invalidation call could
leave this named test green.

## Nitpicks

None.

## Not found

The pass-1 relationship routing, allocation-bound, `text/xml`, and MIME-shape
findings are remediated. No new contract drift, path publication fault,
namespace-prefix coupling, opaque payload loss, or unnecessary abstraction was
found.
