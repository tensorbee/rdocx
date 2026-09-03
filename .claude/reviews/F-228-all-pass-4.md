# F-228, all aspects, pass 4

**Reviewed**: uncommitted worker diff after pass 3 remediation and schema-slot
hardening, 11 source and test files plus the untracked grammar module
**Verdict**: 4 defects, 0 smells, 0 nitpicks

## Defects

### D1, collapsed run boundaries do not rebase equation raw slots

`crates/rdocx-oxml/src/text.rs:2106`

Run replacement remaps each equation's run boundary but leaves its raw-child
ordinal unchanged. When two old boundaries collapse onto one new boundary,
both equations can retain ordinal zero even though the corresponding raw XML
now occupies ordinal zero and ordinal one. Serialization and `items()` then
associate only the first equation with its raw node, so a mutation to the
second equation is lost. Equation ordinals need the same old-boundary raw
prefix rebasing already applied to revisions and markers.

### D2, a conflicting inherited `m` binding changes preserved raw semantics

`crates/rdocx-oxml/src/math.rs:2246`

Canonical roots deliberately replace `m` with the Transitional OfficeMath
namespace and skip any inherited `m` binding. An aliased equation can legally
arrive under an ancestor that binds `m` to a producer namespace. If an opaque
attribute or child uses that inherited prefix, the unchanged raw bytes acquire
OfficeMath meaning after serialization. The projection must fail closed when
fixed-prefix canonicalization would rebind preserved producer XML.

### D3, default namespaces are incorrectly applied to attributes

`crates/rdocx-oxml/src/math.rs:2738`

`root_attribute` resolves attributes through `expanded_name`, which maps an
unprefixed name through the default namespace. XML namespace defaults do not
apply to attributes. With a default OfficeMath namespace, an unqualified
`val` or `alnAt` is therefore accepted as a modeled math attribute and can be
removed or rewritten during mutation. Attribute expansion must require an
explicit prefix for namespaced math attributes.

### D4, the mandatory corpus gate does not exercise its preservation claim

`crates/rdocx-oxml/src/math.rs:2922`

The named test gate builds only authored typed nodes, contains no opaque
sibling, and verifies only the expression count plus one run string. It does
not prove that the complete source-built corpus preserves raw siblings through
mutation and reopen as required by the approved test plan. The public
integration test also authors only a fraction at
`crates/rdocx/tests/integration_test.rs:59`, despite the plan assigning the
complete supported corpus to that path. The exact gate must contain source raw
content and assert every expression variant, while the facade test must author
the full supported set.

## Smells

None.

## Nitpicks

None.

## Not found

All pass 1 through pass 3 findings are otherwise resolved. No new settings
relationship, legacy-boundary, dependency, or structural defect was found
beyond the four items above.
