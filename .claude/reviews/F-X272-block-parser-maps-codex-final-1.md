# F-X272, block parser map producers, Root final pass 1

Root read the complete1084line increment in four product files and existing
progress. Patch SHA256
`219f231a6050d0e0fab97c2853565666be12711851ebb8929c0f0b179edc2864`.
This bounded final review does not reset exhausted implementation peer passes.

Verdict: zero defects, one structural smell and zero nitpicks. Acceptance waits
for the narrow structural repair. No product file was changed during review.

## S1, remove the new forwarding-only helper

`crates/rdocx-oxml/src/text.rs:27154` adds
raw_with_external_bindings_and_map solely to forward its arguments plus None
to the substantive helper at27162. This adds an unnecessary place to inspect
and violates the repository rule against new forwarding-only wrappers. Expose
the substantive concrete helper directly to its table callers, with explicit
None emission where appropriate, while retaining existing compatibility
entrypoints. No behavior or map policy needs to change.

## Reviewed behavior and evidence

Actual successful constructor pushes supply finite table, row, cell and SDT
component paths. Root companions stage all child outputs until successful
terminal return. Option fallback and later Result failure discard pending
entries. Empty paragraph None slots match actual retained attributes. No
equality, pointer lookup or traversal zip discovers paragraph identity.

Actual SDT capture intervals pass through the same namespace insertion emitter.
Inserted declarations are Generated. Unchanged prefix, opening and suffix
intervals retain only their actual correspondence. Canonicalized openings
remain Generated, and absent capture proof stays absent. Existing namespace
selection, emitted bytes and error order remain in the substantive producer.
Binding a capture clears direct selector facts. Legacy signatures remain
compatibility entries and do not request paragraph companion collection.

Root read all three new tests. They discriminate five nested nonempty owners
and an empty owner, controls at table/row/cell levels, actual part7/offset100
wrapper provenance, ordinary model parity, late Option/Result rollback and
Generated-only closure insertion. Direct simple Fields coexist with wrappers,
but this is not complete deferred Field current-role qualification.

Mechanical authentication reconstructs all five final snapshots exactly and
pins54stage source copies, three saved binaries and the actual raw processes.
The final binary
`f33d449ee0ba4b000547bae54bbeb7a1a646d73b51a636b4fb2b6c59e5d96450`
passes86affected tests, including three new and83existing checks. Five scoped
gates and three metadata checks pass. Two prior Clippy style diagnostics are
retained as lint failures, not historical product RED.

Correctness, contract, panic, OOXML preservation, tests and evidence yielded
zero new findings within this producer scope. Structure has S1 above. Public
shared-work replacement, the approved literal event, actual final block sink,
foreign-w Publication root repair, generic constructor/current-work closure
and Native grammar consumers remain open. No provider, integrated rider,
performance, native-Office or sprint completion follows from this review.
