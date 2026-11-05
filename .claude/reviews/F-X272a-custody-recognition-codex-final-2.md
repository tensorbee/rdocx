# F-X272a custody recognition, Root bounded final pass 2

Reviewed repair-v2's285line two-path source repair and v3's107line successor
delta, the existing first-pass invariant and the affected source consumers.
Verdict: D1 remains open on an equivalent publisher and custody consumer.
One defect, zero smells, zero nitpicks. This pass is not whole-feature acceptance.

V3 manifest SHA256 is
`332e0e815beacc8caee76b1763180657993809cea65561ae2ae4358a6c264839`.
Delta SHA256 is
`96b6b58bc3e8cf84a0632719b9f57b3c5fb5b67061826aa240ffc32c0c1e7428`.
Frozen comparison.rs SHA256 is
`67e38ebc8fd4bc27b5d7591b75a42b814826169ee75a97419f2ef8107d3df100`.
Citations refer to the v3 repository-relative snapshot.

## Defects

### D1, opaque-owner exclusion does not yet reach revision rewriting and unit membership

`crates/rdocx/src/comparison.rs:5409`
`crates/rdocx/src/comparison.rs:4557`
`crates/rdocx/src/comparison.rs:4679`
`crates/rdocx/src/revision.rs:3284`

The repaired bookmark decoder and first-paragraph emitter now use actual
admitted paragraph membership. The added-note revision unifier still rewrites
every XmlTree node with revision metadata. XmlTree intentionally restarts its
modeled flag at a lexical Word txbxContent, including below an otherwise opaque
ancestor. Thus revision metadata is not proof that this new producer may edit
the node. A foreign opaque note subtree containing Word txbxContent, paragraph
and insertion markup reaches the unifier and has its stored revision id and
declarations rewritten. The new actual paragraph inventory excludes that
subtree, but the unifier never consults its admission boundary.

The same weak predicate also reaches custody membership, which scans every
metadata-bearing descendant of the real definition, and the definition-removal
revision count. Fixing only the unifier leaves opaque revision-shaped bytes
eligible for partial-unit refusal or incorrect grouping/counting. This is the
same physical-owner invariant as pass1 D1 on additional publishers/consumers,
not a new fourth peer or a marker-format decision.

The author supplied this source lead after freezing v3. Root checked the
unifier, XmlTree reset and both downstream consumers. No runtime counterexample
for this variant was executed in this review, and no successful forged output
is claimed. Existing source guarantees still require preserving the opaque
bytes and refusing no otherwise supported edit merely because they resemble
active revisions. Use actual producer and admitted-owner membership throughout.
Do not globally change the inherited XmlTree textbox policy as a shortcut.
Retain genuine admitted textbox, table and control revision positives beside
the opaque negative, including selectors and actual resolved-element counts.

## Smells

None found in the bounded repair.

## Nitpicks

None.

## Closed portions and evidence limits

The original opaque N/B association, first admitted paragraph selection and
normal signed-id story ownership gaps have concrete old-source discriminators.
The finite Native bridge reuses actual story inventory and conveys source
ranges without edit authority. It leaves public signatures unchanged and
preserves empty/unknown type legacy classification. Raw bookmark leaves remain
distinct from their admitted containing path. The temporary namespace envelope
keeps opaque child bytes untouched, while v3 restores the existing modeled-root
publication decision across whitespace, comment and PI prefixes.

The introduced raw-bookmark-leaf predicate failure and duplicate81 fixture
setup error remain separate. The corrected181 old-source control really fails.
V3's leading-gap failure is an actual test-only v2-production RED, while its
opaque-B atomic-refusal control already passes. None is relabelled as native
acceptance or another model's failure.

Mechanical support authenticated v2's83custody files and exact repair from v1
plus actual Git Base NativeDoc, then v3's45custody files,12source snapshots and
exact two-path delta. Current49integration and5unit named passes, zero ignored,
and Clippy exit0 are source-stable. Q-A430-29 is explicitly excluded, not passed.
V2 CLI, regression and prose remain at their earlier source identities after
v3 production changes. Unchanged fixture files alone are not fresh v3 proof.
Python AST proves syntax only. The full45A/14G map retains final obligations.

No additional public API, panic, ordering or structure finding is recorded in
the examined repair. The Native file lease is returned and its pinned32line
change remains isolated for later semantic integration. Q-A430-29, exact Q28
native exception, installed bindings, schema/native, final full gates and
provider readiness remain unchanged. Root modified only this review record
and returns to orchestration for separate invariant-based remediation.
