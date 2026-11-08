# S96, unused Word scope successors, design support pass 1

## Scope and verdict

Independent read-only investigation of the exact proposed replacement preparation mode on sprint/s96 at cec87f9501eb13d0d5d60572472b227483262620 with the existing dirty source preserved. This is the same design investigation after interruption, not a new cycle. Root remains the sole product writer.

Verdict: needs-design for the proposed API shape. There are zero correctness defects in the stated allocation-elision invariant, one structural design smell and zero nitpicks. This is not implementation approval or a performance finding.

## Structural design smell

### S1, a new forwarding-only entrypoint conflicts with the structural rule

`AGENTS.md:103`, `crates/oxml-py-support/src/lib.rs:412`, `crates/oxml-py-support/src/lib.rs:432`

The proposed `prepare_replacement_scopes` entrypoint, when implemented only by forwarding to the existing private allocator with another policy argument, introduces exactly the new forwarding-only wrapper the repository prohibits. The existing `prepare_scopes` and `prepare_local_scopes` wrappers are precedent in the source, not an exemption for a new construct.

A substantive one-file arrangement is available without duplicating the allocator: promote the existing preparation implementation as the policy-owning callable, with a closed concrete policy for its three actual cases, global successors, local successors and replacement births. Existing public entrypoints may retain their existing contracts. The two Word Conservative consumers select the replacement case directly. This avoids adding another forwarding-only layer and avoids exposing a fourth unsupported combination of two independent booleans. The exact callable signature and policy type still need contract adoption. This report does not approve an unwritten alternate API.

## Semantics examined

At `crates/oxml-py-support/src/lib.rs:321`, each `state` call creates one actual shared ScopeGeneration. The current preparation loop at 480 allocates a document successor and one successor per supplied retired scope, then allocates births at 492. Skipping only the unused owner successor creations can preserve the same retirement and birth behavior.

The document successor must remain fully allocated at vector index zero. `PreparedRevision::document_scope` reads it at 259, `capture_prepared` captures it at 406, and `commit_scopes` installs it at 546. An empty retired-owner slice or empty birth slice does not remove that requirement. An omitted owner successor must return None through the new mode's `successor` contract, without changing the existing methods' ordered successor behavior at 251.

Generation reservation must remain exact. With prior next generation g, N supplied retired scopes and B births, the replacement document is g+1, birth j is g+N+2+j for zero-based j, and the next generation is g+N+1+B. Omitting N allocations must not omit those N generation positions. Preserve both overflow checks at 439 and 445, including refusal when unused virtual positions would overflow. This also preserves subsequent operation generations, not merely the immediate returned births.

All supplied retired states must remain in the prepared retirement inventory. Ownership, already-retired and duplicate checks at 463 must retain their current order. Dropping a preparation must leave the counter and existing states unchanged. Commit must retain foreign/stale/retired checks at 511 before the first retirement, then the same serial, operation label and scope written to every actual old state at 528. No lazy owner registry, source allowance, parser or publication shortcut is justified by this change.

## Consumer coverage

The actual Word splice consumer at `crates/rdocx-py/src/document.rs:6816` distinguishes AppendBody from Conservative. Only Conservative is in the proposal. It prepares the current retired inventory and new births at 6836, then consumes births for owner states and the snapshot at 6898 and 6914. Its returned handle also uses `capture_prepared` and `document_scope` at 6941 and 6971. Commit consumes the complete prepared revision at 7004. It does not consume an owner successor.

The second actual Word consumer is `prepare_word_revision` at `crates/rdocx-py/src/document.rs:11199`. Conservative deliberately has no old-owner matches at 11282 and creates new owner states from births at 11483, plus its snapshot birth at 11495. Its preparation branch at 11465 must select the proposed mode only for Conservative. ElementXml, RunText, Signing, ModernFormValues, AppendBody and other existing publication kinds keep their current global or local policy. Publication at 11509 still commits before installing the prepared owners and indexes. No Word caller of `PreparedRevision::successor` was found in the complete crate search.

PowerPoint is an explicit unchanged consumer. `crates/rpptx-py/src/presentation.rs:4133` prepares normal successors, which its owner update reads at 4150. Slide, layout and master inventory states read successors at 4202, 4206 and 4210. Element XML replacement reads them at 4569, 4573, 4574 and 4577. These are real required successors and must not be optimized away. The existing global and local support tests also exercise the public successor contract, for example `crates/oxml-py-support/src/lib.rs:1054` and 1266.

## Contract adoption and required discriminators

The current F268 lease list at `.claude/plans/F-X268-design.md:563` does not include the shared support source, and the new callable and mode contract are not yet adopted there. Root must reconcile the canonical existing plan and its HLD work list before implementation. This does not reset historical delivery state or a review cap. The shared crate is unpublished at `crates/oxml-py-support/Cargo.toml:6`, but its internal API still requires an explicit consumer contract.

Required local evidence before any timing inference:

- Observe actual ScopeGeneration creations at the allocation owner, not a constant marker near a caller. Reset after fixture construction. A genuine predecessor must complete its state and retirement semantics before failing the reduced-creation assertion. For N retired scopes and B births, the new global replacement preparation creates exactly 1+B states, while the unchanged default creates 1+N+B. This is not a count of Vec, HashSet or all heap allocations.
- Compare default and replacement preparations for zero, one and many mixed scopes, empty and multiple births, dropped preparations, prepared future capture, committed serial, document generation, every birth generation and scope, first retirement causes, and one subsequent mutation. Check `successor(0)` and `successor(usize::MAX)` explicitly in both policies.
- Preserve foreign, duplicate, already-retired, stale preparation, foreign commit/capture, serial overflow and generation overflow refusals before publication. Include overflow caused solely by the skipped virtual successor positions. Existing refusal controls at `crates/oxml-py-support/src/lib.rs:1098` and 1469 provide the relevant baseline behavior.
- Exercise both actual Word Conservative consumers through focused fresh installed calls. Cover cold and warm splices, returned inserted handles, held stale handles with the same first operation label, save/reopen values and byte-identical refusals. A helper-only allocation test does not prove the binding selected the new policy. Keep an AppendBody retention control and a local Word edit beside the changed cases.
- Preserve the default PowerPoint successor and prepared-capture controls. Run unique scoped checks, then an independent implementation microscope. Final installed/typing, memory, compatibility, artifact, native and integrated gates remain due where affected.

## Limits and handback

The source proves that the two specified Word consumers do not read owner successors. It does not establish what fraction of reply or cold-splice wall time those allocations consume. The separate exact-original cold diagnostic establishes prepared-splice use and first-activation work, not scope-allocation causality. The clean monitored timing evidence supplied by Root remains separate from this design investigation. Any implementation needs a fresh qualified artifact and actual measurements under the unchanged protocols and thresholds.

No correctness defect, panic-path defect, OOXML defect or nitpick was found in the proposed invariant. S1 leaves the exact forwarding-entrypoint proposal at needs-design. This pass performed no compilation, tests, product imports, benchmarks, source changes, plan changes, scratch changes, commits, pushes or status updates. It wrote only this review and returns control to Root.
