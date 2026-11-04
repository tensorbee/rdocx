# F-X272, partial affine model increment, final pass 1

**Reviewed**: Root's frozen borrowed-model affine increment in `root-affine-model-stage`, three existing paths, 338 added lines and one removed line. Production change is confined to `crates/rdocx-oxml/src/text.rs`. The other paths are the existing F272 plan and progress record. All canonical and frozen before/after SHA-256 values match the supplied manifest. Parent supplied sprint/s96 revision101c34c4a6d5fa62d9ba0e0ff745fd44c9aeaf13. No Git command, source edit, build, test, network operation or remediation was performed during this audit.

**Verdict**: 0 new defects, 0 new smells, 0 nitpicks in this partial increment. This is not full C1 acceptance, a provider checkpoint, a peer-cap reset or the final integrated F272 review.

## Defects

None found in the reviewed increment.

## Smells

None found in the reviewed increment.

## Nitpicks

None.

## Correctness and contract examined

`crates/rdocx-oxml/src/text.rs:2010` evaluates the root at ancestry1 as A+B. The recurrence at `text.rs:2036` preserves local work and sums every child occurrence's coefficient. Its offset adds the child's coefficient times the actual ancestry delta, plus the child's offset. For nonnegative local work and deltas, this is exactly the former occurrence walk's sum. Memoization avoids recomputing a child but does not remove any edge charge.

The per-call key at `text.rs:2021` is the actual borrowed Field allocation. No immutable FieldSource or shared result Arc identity is substituted for current mutable Field identity. Distinct current clones retain distinct local accounting, while genuinely shared descendant allocations are summarized once. Keys and borrowed child references stay stable during the immutable call. Safe ownership of boxed operands and immutable result owners supplies the acyclic graph assumed by this postorder.

The local expression at `text.rs:2068` retains raw source and original scope multiplicity, current instruction/cache bytes, legacy names/values/choices/defaults/author text, both checked payload terms and ordered normalized/pending source. The complete prior production walk is retained at `text.rs:2180`. A read-only comparison confirmed that its body is byte-identical to the frozen old method after removal of the added diagnostic counter.

The actual selected child graph at `text.rs:2137` preserves the old graph. Ordered results traverse accepted run paths and count actual Revision path segments, then add one cache ancestor. Other results use the existing selected-cache helper at `text.rs:2644`, including typed display consistency, simple cache precedence, parsed-child eligibility and the established source fallback. Current nested instruction arguments and switch arguments add zero ancestry. Raw/current instruction divergence does not silently select original operands.

Caller checks include initial child capture at `text.rs:1333`, cached-run replacement at `text.rs:1655`, atomic ordered normalization at `text.rs:1870`, attachment validation at `text.rs:2408`, source writing at `text.rs:14841` and retained-root publication at `text.rs:19736`. They retain their existing admission and publication ordering around the changed method. No public signature, model field, namespace selection, OOXML child ordering or serialization representation changed in this increment.

## Overflow, scheduling and evidence limits

The field postorder at `text.rs:2023` is iterative. A memo hit skips computation, while the parent's retained edge list still charges each occurrence. Missing child summaries produce a checked error. Coefficient sums, ancestry products, offset sums and root evaluation use checked arithmetic. For finite acyclic graphs, each nonnegative partial coefficient and offset is bounded by the corresponding complete root work, so the new factorization does not introduce a stricter numerical success boundary.

Local validation now occurs at `text.rs:2036`, after descendant scheduling at `text.rs:2028`, rather than before scheduling in the retained reference at `text.rs:2186`. Models with multiple independent invalid/overflowing components can therefore expose a different first error message. No exact error-message precedence contract was found for this private calculation, so this is not classified as a defect. The frozen controls do not contain a dedicated near-usize-boundary overflow fixture. The numerical refusal assessment here follows the checked expressions and nonnegative recurrence, not a claimed executed overflow boundary test.

The work counters measure Field summary computations, not all helper work. Existing accepted-run path production and cache-display selection remain separate consumers. This review does not turn those helpers, repeated calls, parser-side construction or all derived operations into an iterative or whole-operation scheduling proof.

## Acceptance evidence examined

All seven supplied receipt/log hashes match, including the initial failed test expectation and prose1. Child and wrapper outcomes remain separate. Recorded final observations are two affine controls, 52 field-filter controls, 49 capture controls and changed-crate all-target/all-feature Clippy0. Counts overlap and are not summed. Source formatting success is author-reported in the handoff rather than represented by a dedicated receipt in the supplied set. No independent formatting or other gate was run in this read-only audit.

The fixture at `text.rs:32057` uses a parsed ordered result containing an actual insertion revision and complex child field. The parser establishes the ordered paragraph at `text.rs:6090`, and accepted paths resolve actual revision ownership through `text.rs:11278`. The test compares the unchanged reference, ancestry rebasing, selected versus stale caches and legacy metadata. Divergent current clones explicitly share source and ordered-result Arcs. The asserted unique allocation count and second-occurrence delta reject source-identity memoization and lost edge multiplicity.

The shared DAG fixture at `text.rs:32151` compares exact numerical work and actual reference versus summary visits. Depths2/4/8 preserve reference work while visits7/31/511 become5/9/17. This supports the stated Field-computation observation, not an end-to-end speedup claim. The initial expected delta omitted the existing factor124 on the64-byte argument term. Its correction is a fixture repair, not a historical production RED.

Prose1 retains the same nine preexisting progress-file semicolon findings. It is not a full prose pass. The plan retains the partial boundary at `.claude/plans/F-X272-design.md:3812`.

## Unfinished full C1 and final gates

Parser-populated summaries before construction, original aggregate initial-field child admission, complete weighted owner/table maxima, all constructor actions and positive/Option/raw outcomes, projected namespace contexts, standalone Result error consumption and original/derived source correspondence remain unfinished. Cross-facade selected/counterpart reservations and checked mutator error propagation remain open. Q-C1-01 exact raw fallback and Q-C1-02 direct legacy infallible reads only remain binding.

Issue430, HLD reconciliation, native G02, schema/native qualification, exact performance, final integrated and installed consumers, hash, CI, packaging and release approval remain due. None is discharged by numerical equivalence of this borrowed-model method.

## Not found

No new defect was found in numerical recurrence, selected-cache policy, ordered revision ancestry, current allocation identity, edge multiplicity, checked arithmetic or the reviewed public caller transitions. No new panic, schema-order change, raw-XML loss, unsupported abstraction or public API change was found in this increment. Review ends here. No remediation was performed.
