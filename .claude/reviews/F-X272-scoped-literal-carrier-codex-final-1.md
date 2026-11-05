# F-X272, scoped literal carrier, Root final pass 1

Root read the complete1177line increment in three lower product files and
existing progress, plus the copy-boundary trace. Patch SHA256
`5ccc1dfd1bec62993dc7626fc7982262b2b3894be3e80c873b40a7606bd7c8d7`.
Canonical contract is4cf5e902. This bounded Root final does not reset peer caps.

Verdict: one defect, zero smells and zero nitpicks. Lower carrier acceptance
and Native event forwarding wait for repair. No product source changed during
review. The finding is source-derived, not an invented historical test RED.

## D1, generated paragraphs in self-closing cells have no companion

At crates/rdocx-oxml/src/table.rs:3073, the actual Empty tc branch constructs
CT_Tc::new and pushes it without a paragraph companion. CT_Tc::new at
crates/rdocx-oxml/src/table.rs:2461creates one actual generated paragraph. The
new owner walk at crates/rdocx-oxml/src/placeholder.rs:1523requires a path entry
for every visited paragraph and returns an error when this one is absent.

Thus a previously admitted self-closing cell inside a row produces a paragraph
that the parser companion does not describe. A no-match scoped table call can
now fail with missing constructed paragraph source companion. An empty cell
before a later selected Field can likewise prevent the real edit.

Record this actual generated paragraph at its constructor push path with its
true Generated/None extra-slot maps. Do not infer Original authority from the
empty cell or add a generic missing-map fallback. Preserve row/control ordering
and transactional companion publication. Cover a no-match table, empty cell
with retained root attributes, and an empty cell alongside a selected Field,
including nested/control row ownership. Produce actual baseline/successor
observations rather than labelling this source trace itself as executed RED.

## Inspected implementation and evidence

The substantive new hidden entry owns original prepare and the same work/
registry through all four root branches. Actual factory component paths are
sorted once and binary-searched during the prior traversal order. Original
selected facts and a genuine literal event use checked arithmetic before the
ledger commit. Matching shares the original run/stretch predicate, preserves
crossing skips and uses the original batch mutation/writeback engine.

The finite final block serializer binds its live concrete writer under the
existing restoring receipt. Nested fresh writers retain generation demand.
Actual root frames and final generated w closure retain foreign w and opaque
shadows. Source correspondence remains separate from genuine incoming credit.
Earlier model String copies are explicitly not proved by the XML guard and
retain their actual local/current/fact-copy obligations, without a new quota.

Mechanical support reconstructs all four final snapshots and authenticates
90stage source copies. Final binary
`58e5cb6b2fd16755b4426ba658f05b6139f82f15ec32868ad8e00e079926b162`
passes85affected tests. Five scoped gates and three metadata checks pass. The
20root/frame cases are five contexts times four roots in one test, not20tests
or universal acceptance. Initial setup, real absent-w failure and mistaken
inherited-only-q fixture expectation remain separately classified. Native is
unchanged and unqualified by this increment.

Contract, panic, XML preservation, tests, structure and evidence yielded no
additional established finding in this bounded read. D1 is correctness and
compatibility. Existing unsupported selectors are not promoted by maps. Whole
current-role, constructor, deferred/capture and Native grammar obligations,
final integrated/native/performance riders and provider acceptance remain open.
