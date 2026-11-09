# F-X272 Issue430 design, Codex final 8

REPAIR REQUIRED. Two defects, one smell, zero nitpicks. Root read the complete
555-line V11, 125144 bytes, SHA-256
eb111cbbe39c32e61a13244416f52bc2815d340fdaa2f55197ed56ac700fcf12.
The matching handoff is 2215 bytes. Author terminal23224 returned its natural
exit0 before review. Both artifacts are frozen at
/private/tmp/rdocx-cross-review/issue430-v11-final-frozen.
This is the direct Root final, not a fourth design peer or a counter reset.
No product source, design adoption or implementation is accepted.

## Defects

### N15, reopened, rejecting an addition must restore the separator-only baseline

/private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V11.md:302
sets A430-18 up as a special-only original part plus an added note, then says
Reject restores the definition and leaves N greater than zero. The original
contains no normal definition to restore. Reject must remove the added normal
definition and reference, preserve the original specials and part, and retire
the library markers at settlement. In this case N is zero, R is zero,
E is false and base is one, so the predicate retains the part without retaining
the addition. The current expectation violates the reject-to-original
postcondition at V11:235 and contradicts its own predicate at V11:132.

N15's destination-state invariant is reopened by this regression. Repair
this acceptance row and any repeated special-only addition expectation, retaining
all 45 rows. Include the corresponding accept control. This needs no new user
policy, token representation or native gate.

### N21, the planned CLI regression command names no existing test carrier

/private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V11.md:337
requires cargo test -p rdocx-cli --test integration_test issue430.
The existing CLI carrier is crates/rdocx-cli/tests/integration.rs:1, and
crates/rdocx-cli/Cargo.toml:32 declares the binary without an alternate
integration_test target. The native rdocx crate has an integration_test.rs,
but that is a different package. The proposed CLI gate cannot execute the new
regressions as written.

Use the actual existing CLI integration target and preserve the separate native
integration_test and regression_test targets. Do not create another test file or
binary to make the incorrect command work.

## Smells

### S1, the decision summary contradicts the supported author and date selectors

/private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V11.md:551
says a single public id, author or date selector cannot resolve a unit whose
elements have different ids or metadata, and suggests a token API would be
required. Different ids alone do not prevent an author or date selector from
covering every element. V11:266 already gives the compatible control: ids1,2,3
with the same author and matching date are fully selected by those existing
methods. The actual predicates at crates/rdocx/src/revision.rs:804 and :807
do not compare ids for those selectors.

Keep the operative per-element and partial-unit rules. Rewrite the summary per
selector: a scope is full when its own predicate matches every element, partial
when it matches some, and empty when it matches none. No new public API or user
decision is needed to state that contract.

## Repairs retained and review limits

N19 and N20 are closed at the mechanism level by V11:114, :130 and :531.
Every open unit votes, including retained tombstones. Reopening does not require
a current-versus-prior class or token chronology. Before new custody is emitted
into a settled part, the producer retires completed imported markers. The
specified producer convention prevents completed imported provenance from
poisoning the next open lifetime. The mixed 0/1 plus 1/1 all-rejected creation
therefore returns to absent, while the three fully completed transactions
leave no old bits to vote. A new group representation is not required by this
review. Runtime correctness remains unproved.

N13 and N14 remain closed at the design level. The deletion-direction scalar
facts at V11:41 and :539 retain actual ids1,2,4 and the note marker outside the
deletion. N16, N17 and N18 retain their principal repairs: unowned invalid and
opaque content prevents removal, completed library custody is retired, and
numeric token coincidence follows actual selector semantics.

The review checked all five issue rows, all45 acceptance rows, parser/type
policy, marker grammar, owner association, lifecycle, part predicate,
listing/selection, package cleanup, HLD impact, test plan and evidence timing.
No additional correctness, namespace/schema-order, atomicity, public-surface
or structure finding is asserted. No speculative foreign-marker provenance
finding is added beyond the explicitly stated producer convention.

Native evidence is bounded fixture evidence. Multi-reference behavior,
library-selected resolution, mixed lifetimes, absent-part cleanup, pinned SDK,
installed bindings and the final native/schema/hash/CI union remain pending.
Attribution correction after checking the prior artifact:
/private/tmp/rdocx-deepseek-harness/issue430-contract/DEEPSEEK_PROPOSAL_V10.md:284
correctly removes the added normal definition and reference in A430-18.
V11 introduced this expectation regression. The incorrect CLI target already
appears at V10:317 and was missed by the earlier review chain, including Root.
Root's initial wording wrongly grouped both as retained prior errors. The
findings and requested repairs are unchanged by this correction.

Preserve V11 and its handoff. A targeted author repair should change the
contradictory expectation, command and summary without regenerating the entire
contract or relaxing any row or gate. Root reviews the repair directly afterward.
All three design peers remain consumed. STOP and return to orchestration.
