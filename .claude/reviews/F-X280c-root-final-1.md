# F-X280c, Root independent final 1

**Reviewer**: Root Codex, independent of the GLM implementing author.
**Scope**: stopped work/f-x280c-opencode at Base72b1f8cb, ten product/test/HLD
paths,2001 insertions and12 deletions. Frozen diff SHA256
`a1924899fbb98fdb1bb9dfb971fa4ffc2e28c217bfec8cf0b054ad7ddfeef304`.
**Verdict**:5 defects,1 evidence smell,0 nitpicks. Do not commit or prepare.

Root read the approved plan, implementation, changed CLI/Python/docs/test
consumers, original helper/OPC paths, retained checked execution logs and both
GLM self-reviews. Actual read-only source support is separately identified in
F-X280c-root-source-support-1.md. Root independently traced the cited paths
below and owns this aggregate disposition. Source support is not an additional
self-repair cycle or design peer. No source edit, build, test, imported extension
execution or native action occurred in this review. Only review records were
written. All ten frozen source hashes still match.

## Defects

### D1, the implemented row rule weakens the approved physical equality rule

`crates/rpptx/src/lib.rs:3727`, `crates/rpptx/src/lib.rs:3744`,
`crates/rpptx/tests/integration.rs:38159`,
`.claude/plans/F-X280c-design.md:61`

The approved contract counts direct qualified tc children against that table's
own direct gridCol count and requires equality. The implementation instead sums
gridSpan, defaults malformed values to1 and reports only a shortfall. A row
with four ordinary cells and a three-column grid therefore receives no stored
row issue. A missing physical cell can also be concealed by adding gridSpan.
The new test explicitly accepts over-coverage. README, help and HLD still
promise physical equality, so the public claim and implementation disagree.

The author recorded this as a deliberate deviation to keep all50 corpus CLI
inputs clean. Recording it is useful, but is not approval. Corpus membership
and successful legacy parsed-model loading do not prove raw schema or selected
profile validity. The selected finite diagnostic may deliberately expose source
faults a tolerant model projection did not report. Primary
[DrawingML TableRow remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablerow?view=openxml-3.0.1)
describe a cell for each grid column. Office's separate
[merge-attribute clarification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/de4562ba-7901-48df-9803-414ded131f2b)
describes merge flags on physical cells and does not justify accepting arbitrary
overlong rows. This is not a claim that SDK3.0.1 has a cross-row cardinality
constraint. Its retained particle metadata does not prove such a validator run.

Restore the approved finite rule and add short/equal/long physical-row controls,
including ordinary authored merged cells retaining physical continuation cells.
Keep the corpus inputs and parsed-model compatibility controls. Reconcile any
specific newly reported raw corpus mismatch explicitly with its actual bytes
and expected finite diagnostic, rather than weakening the production predicate
to retain a blanket clean assertion. If primary evidence actually contradicts
the contract, escalate that bounded contradiction before changing the rule.
Do not introduce numeric span policy or broaden the schema subset silently.

### D2, structural part roles are inferred from root spelling alone

`crates/rpptx/src/lib.rs:3216`, `crates/rpptx/src/lib.rs:3252`,
`crates/rpptx/src/lib.rs:3366`, `crates/rpptx/src/lib.rs:3369`

Every MIME-XML ordinary part receives structural scanning. Dispatch then uses
only root URI/local name, with no actual owning relationship or part role.
An unused application/xml custom part rooted p:presentation with an opaque
child is therefore diagnosed as a presentation. The same applies to a custom
p:sld lookalike. The contract requires all-XML lexical checking but restricts
these structural rules to actual presentation role parts and physical owners.
Namespace vocabulary is not OPC ownership.

Qualify the structural route from the actual existing owning relationships and
content-type context, with the expected root. Preserve lexical checks for every
XML source and ignore a well-formed unowned/custom lookalike structurally. Pair
that control with the same malformed child in a real role part. Do not add a
general new required-root policy or change initial OPC error ordering.

### D3, the new scan recurses through opaque XML before checking its role

`crates/rpptx/src/lib.rs:3359`, `crates/rpptx/src/lib.rs:3392`,
`crates/rpptx/src/lib.rs:3495`,
`crates/rpptx/tests/integration.rs:32260`

stored_next_element builds the entire recursive StoredNode tree before even
checking whether the root is foreign. stored_subtree calls itself for every
Start event, including opaque extLst/contentPart and foreign compatibility
payloads. The recursively owned children also require recursive destruction.
Caller-controlled valid deep opaque XML can exhaust the native stack before
its return-opaque decision. Deep group input likewise reaches this recursion
before the existing model's iterative32-level refusal, which already has
100000-level controls in fx214_nesting. This introduces an unbounded call-stack
path ahead of the existing accepted/refused behavior.

Use iterative finite traversal and opaque lexical-only handling, with storage
and destruction that also avoid native-depth recursion. Preserve the existing
applicable group/time limits and error behavior. Do not add a new universal
XML-depth quota or reject previously admitted opaque payloads. Add an isolated
bounded-process regression for deep valid opaque content in a real role part,
plus the corresponding unused custom part and existing model-limit controls.
No crash threshold or executed crash is claimed by this review.

### D4, outside-root character events can pass as a complete XML document

`crates/rpptx/src/lib.rs:3332`, `crates/rpptx/src/lib.rs:3338`,
`crates/oxml-core/src/xml.rs:105`, `crates/oxml-core/src/xml.rs:113`

The root checker observes outside-root content only for Event::Text. CData and
GeneralRef fall through. The shared lexical helper validates reference tokens
without document-position constraints and also does not reject these CDATA
positions. Thus an unused XML part containing CDATA before its root, or an
entity reference after its root, can have exactly one complete root and still
be reported clean. The custom part need not affect model loading.

Cover all relevant outside-root character-event forms. Preserve valid in-root
CDATA/reference and outside-root literal XML whitespace/comments/PIs. Pair
ordinary and original-metadata cases and check source bytes unchanged. This is
the promised document-root/lexical frontier, not a request for full OOXML schema
validation or a new decoder.

### D5, metadata suffix matching promotes ordinary binary parts to XML

`crates/rpptx/src/lib.rs:3228`, `crates/rpptx/src/lib.rs:3810`,
`crates/oxml-opc/src/package.rs:198`,
`crates/oxml-opc/src/package.rs:236`

Original metadata recovery uses an ends_with content-types filename test. The
OPC loader instead recognizes the normalized root content-types identity.
A separate admitted application/octet-stream part such as
/customXml/payload[Content_Types].xml with ordinary ASCII non-XML bytes is
correctly skipped by the MIME-based ordinary scan, then wrongly decoded as
metadata and reported InvalidXmlChild. This relies on the current reader's
admission/classification, not a claim that the illustrative URI passes every
OPC URI schema rule. C may not turn extension-only binary source into XML.

Match the actual root metadata identity under the existing normalization and
case policy. Preserve genuine original metadata reads and the existing .rels
classification. Pair binary suffix-decoy, nested ordinary XML suffix-decoy and
real root metadata controls, without introducing B's ZIP/name/open policy.

## Evidence smell S1, self-review claims exceed the actual observations

`.claude/reviews/F-X280c-self-review-pass-2.md:19`,
`.claude/reviews/F-X280c-self-review-pass-2.md:28`,
`.claude/scratch/logs/native-remediation-full.log:693`

The self-review calls the full native carrier passed, but its actual retained
result is418 passed,2 failed and25 ignored, child101. Progress and the final
message correctly preserve those two chart failures and their Base control.
The self-review also says saturation reports an overflowing span as short,
although saturating to u64MAX makes coverage less than columns false. No new
huge-span discriminator was found in the added test section.

Preserve historical observations and append a precise correction, then ensure
current review/progress/handoff claims name only actual passed checks. Do not
relabel the failed broad command as green or call code inspection an executed
regression. The corpus-specific D1 expectation also needs correction under its
contract disposition. Full final qualification remains required.

## Observations retained without further findings

The public native API, initial read/ZIP/OPC ordering, finite forbidden-character
mapping, stored-before-parsed results, existing Eq deduplication and model-error
fallback are explicit. The helper extraction has two real consumers and retains
the five-kind match. Required nonvisual attributes are unqualified presence
checks, not a new empty-name or numeric policy. Qualified MC members use the
same finite member rule and foreign payload remains structurally opaque after
construction, subject to D3. The fourteen-slot root sequence matches the pinned
SDK particle, including kinsoku and the intended smartTags distinction.

CLI routes to the native function and keeps Display and the existing boolean
exit path. Python remains parsed-model only and adds no binding surface. The
focused installed Python3.12 run has a successful maturin producer at15:36:28,
actual module hash229ccb66a8373876f662b9b4ec71a7216e7d95ff7fd4ec69a234a109d6acc05f
matching the recorded artifact, followed by one actual pytest pass. This is
focused evidence, not Python3.9/full typing qualification. Scoped native/CLI and
Clippy logs are real checked child observations after the mandatory evidence
intervention. Initial filtered commands remain historical unqualified attempts.
Root did not execute any of these checks in this review.

The original Base CLI semantic RED, real six-fault fixtures, exact source-byte
preservation, lossless opaque encoding controls, shared Display comparison,
role/content-type fixtures and self-found tree-root missing-name regression are
useful. They do not discriminate D1-D5. The broader native carrier is still
failed. Final installed/typing/full/hash/SDK/native/WASM/archive/hosted and B
composition riders remain pending. No source is accepted or prepared here.

## Handback

Return to a separate GLM repair phase with all findings grouped. Two self-review
passes are already used, not one merely because there was one repair cycle.
Keep the original cap, with one remaining self-review pass. Root's actual final
review is separate, and source support does not reset any cycle. No fourth
self/peer pass, provider completion, source commit, reporter comment or
publication is authorized. The source review ends with this report.
