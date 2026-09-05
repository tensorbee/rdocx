# F-236, all, pass 1

**Reviewed**: Uncommitted working tree against `dbb5ab1`, 7 files and 2,290 changed lines, comprising 2,284 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`
**Verdict**: 5 defects, 0 smells, 0 nitpicks

## Defects

### D1, ActiveX inventory accepts an ambiguous binary relationship graph
`crates/rdocx/src/embedded.rs:251`

Inventory resolves the one relationship id written on the `ax:ocx` root and
checks that relationship's type, but it never verifies that the properties
part has exactly one `ACTIVEX_CONTROL_BINARY` relationship. A properties part
with the referenced binary relationship plus a second binary relationship is
therefore accepted. This violates the approved exactly-one graph contract and
lets removal silently treat the second binary as an unrelated owned candidate
instead of rejecting the malformed graph.

### D2, shared ActiveX properties are retained after their required relationship is removed
`crates/rdocx/src/embedded.rs:407`

ActiveX removal deletes the properties part's binary relationship before
checking whether the properties part is still reachable. If an unrelated
internal relationship also targets that properties part, the later
`delete_if_unreachable` call retains the shared XML part but its root still
contains the removed relationship id. The mutation commits a broken retained
part instead of preserving the shared target or failing closed.

### D3, removing package signatures can leave unrelated relationships dangling
`crates/rdocx/src/embedded.rs:1398`

The signature graph validator checks the outgoing signature relationships but
does not reject unrelated incoming relationships to the origin or signature
parts. `remove_package_signatures` then deletes those parts unconditionally.
A package in which another source also targets `sig1.xml` therefore retains
that unrelated relationship with a missing target, contrary to the valid
package and remove-only-validated-infrastructure contracts.

### D4, synchronized facade mutations can leave stale signatures reported as present
`crates/rdocx/src/embedded.rs:497`

Without the optional digital-signatures feature, reopening a staged mutation
recognizes invalidation only from the new marker or a missing manifest target.
Existing atomic facade operations such as redaction, revision resolution, and
watermark mutation can change signed part bytes and commit a fully synchronized
candidate without first writing that marker. Reopen then resets
`package_signatures_invalidated` to false, the later staged comparison sees no
pending delta, and `embedded_content` reports `Present` for signature evidence
whose digest no longer matches the package.

### D5, the required compatibility-wrapper preservation case is not tested
`crates/rdocx/tests/regression_test.rs:15216`

The round-trip test adds one raw sibling before the selected `w:object`, but it
does not construct an `mc:AlternateContent`, `mc:Choice`, or `mc:Fallback`
owner path. Removing compatibility-container handling from `xml_references`
would still leave this test green. The approved test plan and parser risk rider
both require compatibility-wrapper coverage alongside prefix aliases and
byte-exact retained XML.

## Smells

None.

## Nitpicks

None.

## Not found

No additional findings in panic safety or repository structure. The additive
public types, private implementation module, dependency edge, error variant,
and existing integration binary follow the approved structural contract.
