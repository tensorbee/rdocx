# F-X279d, Author crossing comment replies with initials

**Status**: approved
**Sprint**: S96
**Size**: M
**Depends on**: none
**Test gate**: regression

## Operative Root adoption

The final integrated Python gate at a78eba3c exposed an existing root-move
consumer rejecting the crossing replies this story authors. The bounded
correction reuses qualified comment ownership and selected marker projection,
retaining actual source ordinals and generic paired-range nesting checks.
The existing reference-restoration helper in Native document.rs may call that
crate-private proof. This narrow lease extension preserves root-only moves,
thread metadata, unrelated reply anchors and atomic refusal. The existing
Python failure and the one/three-reply Native regression decide the repair.
No review counter, public signature or acceptance requirement changes.

Root adopts the exact authoring split after independent Root final1, which
reviewed draft2e06c7a2e620fc9653c33f35e69522677179ace10d53b9f8b30c43c262527d13.
This approval changes status and ownership only. All retained draft/proposed
wording below is historical preparation and cannot override this disposition.
D owns A279-1 authoring/initials and reply-specific lifecycle. B retains P1/P2,
C6 and actual F272 dependency. No criterion, native rider or exhausted review
cap is removed or reset. The actual baseline callable proof is accepted.
Normal clean claim and source revalidation are required before implementation.

## Split disposition and retained lineage

This draft extracts the complete A279-1 AUTHORING scope, its C1 authoring
matrix and reply-specific C10 held-state controls from approved F-X279b.
P1/P2 comparison and selective-resolution COMPOSITION remain entirely B,
including C6 and the actual F-X272 implementation prerequisite. A279-3/5/6
remain B, A279-4 remains A and complete A279-2/C2 remains C. No original
acceptance row disappears. Root must adopt and approve this exact split before
claim, implementation or any dependency change becomes operative.

The original three exhausted B design peers, Root final2 lineage and original
three-pass implementation cap remain applicable without reset. This draft
creates no new peer allowance. Retain the prior review reports and their exact
historical applicability. No child or provider completion is recorded here.

## Spec reference

Existing replies author only definition/thread metadata. They omit initials
and their own range/reference anchors. The approved A279-1 requires crossing
markers around the selected parent's actual range, not a nested marker shortcut
or an invented range for a point/orphan parent.

Read the retained A279-1 section, C1/C10 rows and Root P1/P2 ownership
reconciliation in `.claude/plans/F-X279b-design.md`. The original Issue434 body
pin is `9e4dcae4a526819d7a89517ef936a53a26c606df7d9b0bcb1de30cc5ca527815`
and historical PR444 head is `fac363eb9701918ad0d9d6672e386921506b7c02`.
Those retained captures are contract lineage, not a new GitHub fetch or oracle.

- `docs/hld/03-architecture.md`, Shared Word document operations, checked comment authoring, paragraph
  identifiers, thread-root resolution and staged publication.
- `docs/hld/04-opc-and-packaging.md`, Part naming and Package integrity, comment ranges/references, related
  comments parts, untouched producer preservation and checked publication.
- `docs/hld/10-bindings-spec.md`, Native Word facade stability and Python Word document operations,
  comment/reply APIs, held handles, revision publication and CLI operation records.
- `docs/hld/14-development-backlog.md`, F-X279 and F-X279d Issue434 acceptance registration.
- Existing CLI README comment examples and the approved B public signatures.

## Independence and actual consumer proof

Canonical preparation Base is `cb896e72bad1b6693b01ad3dcc242fbc9770a969`.
The eventual claim must use Root's then-clean reviewed canonical descendant,
not a fabricated new Base or the active dirty F272 worker.

The authoring seam predates C2 and unfinished C1. At clean C2 claim Base
`b03b206be3d6028ac979323304a3419c7f4b1374`, Native comments.rs already exposes
comment_anchor2840, comment_anchor_text2851, reply_to_with_date1687,
reply_to_staged1701, remove_comment1857, resolve_comment1786,
comment_source_markers2657 and the existing private comment_ownership_at3437.
These are source observations from git objects, not a new runtime success.
Existing staging, story/source locations, identifier allocation and source-anchor
writers are baseline APIs. They do not require ScopedWordOperation, new source
carriers, template lineage, comparison or import service.

Current canonical comments.rs has the same authoring seam at1774-1876, actual
source marker spans at2767, and dedicated comment_anchor/anchor_text projections
at2950/2961. Use qualified source identities and the existing exact endpoint
policy. Generic paired story-marker nesting is not authority to replace the
required crossing topology with nested markers. Preserve unrelated marker kinds
and their existing contracts. No Native document.rs edit is leased. Existing
crate-private story/source helpers can be called unchanged from comments.rs.
If a genuinely missing helper blocks authoring, report its exact seam before
expanding the lease or inventing a provider-none assertion.

Python's existing document.rs16657 reply_to stages the dated Native API and
prepares/commits one Word revision before installing the candidate. CLI main.rs564
and commands.rs1715 already parse and publish comment replies. Extend these
existing consumers, rather than introducing a replacement authoring stack.

C2 markup/styles is integrated neighbor code at canonical94775888, preserved
by the current Base. D reuses its existing authoring helper without changing
A279-2, lower comments admission or style policy. D needs no new C2 contract or
unfinished C2 callable. This is not a claim that C's reviewed local source
satisfies its formal final/native provider gate. B and the umbrella keep their
actual F272 barriers. Proposed D independence is limited to this authoring
path and still requires Root review and clean claim.

## Callable and compatibility contract

```rust
impl Document {
    pub fn reply_to(&mut self, parent_id: i32, author: &str, text: &str) -> Result<i32>;
    pub fn reply_to_with_date(&mut self, parent_id: i32, author: &str, text: &str, date: Option<&str>) -> Result<i32>;
    pub fn reply_to_with_initials(&mut self, parent_id: i32, author: &str, initials: Option<&str>, text: &str, date: Option<&str>) -> Result<i32>;
}
```

The first two signatures stay unchanged and are compatibility shims to the
new substantive method. reply_to supplies initials=None and date=None.
reply_to_with_date supplies initials=None. Python becomes
`reply_to(parent_id, *, author, text, date=None, initials=None) -> int`, with the
same stub signature. CLI comment reply gains optional --initials. Existing
omitted calls remain valid. This is an additive Rust0.17.0 API and Python/CLI
keyword/flag addition, not a new public provenance or namespace API.

Omitted initials emit no attribute, Some("B") emits w:initials="B", and
Some("") emits an empty attribute exactly as add_comment does. Use existing
reject_non_xml_comment and validate_comment_date. An unknown parent retains
`comment id {id} does not exist`. A known parent without the required
projectable range/reference pair refuses with
`comment id {id} has no range to reply to`. Imported point comments and
reference-only replies retain their original shape. Do not author ranges on
untouched imports or fabricate a parent range.

## Acceptance criteria

1. All three Rust entrypoints, Python and CLI author one reply definition and
   one thread row with the actual requested id/author/date/initials semantics.
   Save/reopen preserves those values and the parent paraId link. Multi-line
   text remains one paragraph per line, with C2's existing first-only annotation
   behavior preserved, not reimplemented.
2. For a parent with exactly one story containing its unique start, end and
   reference, insert the reply start immediately after the parent start. Keep
   parent covered content unchanged. Emit the parent end and parent reference,
   then reply end and its own new reference run. The reply covers the parent's
   text, and its reference follows the parent's. A nested ordering that places
   the reply reference before the parent reference fails the discriminator.
3. Repeated replies and a reply to a reply preserve the same crossing pattern
   around the selected parent, with separate ids and references. A second reply
   uses the approved pattern around the first reply. Multi-paragraph parents
   receive the new start in the parent's start paragraph and the new end plus
   reference in its end paragraph, after its own end/reference. Multi-line
   comment text creates only one body reply range.
4. Missing or ambiguous markers, duplicated starts/ends/references, split-story
   endpoints, a nonprojectable parent or an existing anchor-policy refusal
   publish nothing. Preserve document, package, styles, comment parts, identifier
   availability, layout caches and Python revision on failure. Do not widen the
   existing supported story/endpoint grammar or validate opaque subtrees as a
   new typed input. Valid qualified aliases and unrelated producer bytes remain
   supported under existing source policy.
5. Successful authoring uses a staged candidate and commit_staged_mutation.
   Rust invalidates layout once. Python advances revision once and stales held
   views on success. Held and fresh handles, comment projections and reopened
   package values agree. Failure keeps held views usable and unchanged.
6. Removing the authored reply removes its own start/end/reference, definition
   and commentsExtended row, while preserving parent anchors/content. Existing
   descendant-thread removal semantics remain. resolve_comment keeps existing
   thread-root behavior. move_comment still refuses replies. Save/reopen and
   subsequent removal/resolution are checked on actual authored crossing XML.
7. Do not rewrite untouched imported comment definitions/parts or synthesize
   commentsIds. Preserve existing C2 markup/styles/MC namespace and atomic
   fixed-writer refusal behavior. No comparison or import implementation is
   claimed. P1/P2 and C6 stay B and require actual F272 completion.

## Consumer map

| Input and operation | Required observation | Wrong behavior rejected |
|---|---|---|
| All Rust reply forms, Python keyword, CLI flag | Same saved metadata and crossing anchors | Only one shim fixed, omitted arguments broken |
| Initials omitted, empty, nonempty and invalid XML, valid/invalid/omitted date | Exact attribute omission/value and atomic refusal | Empty treated as absent, validation after commit |
| One parent, sibling replies, reply to reply | Exact expanded marker ids/order and separate reference runs | Nested order, shared reference, wrong selected parent |
| One and multiple parent paragraphs, multi-line reply | Actual endpoint paragraphs and one body range | All ends in first paragraph, one range per comment line |
| Known orphan/point, missing parent, duplicate or cross-story markers | Defined error and unchanged bytes/state | Invented parent range, first-match ambiguous selection |
| Actual supported body/table/control/related-story parent and qualified aliases | Existing endpoint policy, text/opaque preservation and reopened ids | Body-only ordinal guess or changed unsupported grammar |
| Held Python paragraph/run/comment views on failure and success | Failure unchanged, success single publication/staleness | Partial commit or extra revision |
| Author, save/reopen, remove reply and resolve thread | Parent retained, reply anchors removed, root resolved | Lost parent anchors or orphaned thread metadata |
| Existing C2 namespace/style refusals and untouched imported part | Existing behavior and optional commentsIds preserved | Lower serializer overwrite or incidental markup rewrite |

Actual Native crossing add compared/resolved as a revision is NOT D evidence.
That is B's P1/P2/C6 matrix, including unrelated revised-owner preservation,
protected-body refusal and revised-comment snapshot refusal under F272.

## Source lease and implementation shape

Only existing files, after Root approval and normal claim:

- crates/rdocx/src/comments.rs, existing reply, ownership, marker and staged
  lifecycle helpers plus their existing unit test module.
  crates/rdocx/tests/regression_test.rs and integration_test.rs, only existing
  reply authoring, imported-anchor and held-state regression regions.
- crates/rdocx-py/src/document.rs, only the existing reply binding and its
  publication path. crates/rdocx-py/python/rdocx/_rdocx.pyi, its exact signature.
- crates/rdocx-cli/src/main.rs and commands.rs, existing reply parsing/dispatch.
  crates/rdocx-cli/tests/integration.rs, existing CLI regression carrier.
- crates/rdocx-py/tests/test_core.py, disjoint reply authoring/held-state nodes.
- Narrow docs/hld/03-architecture.md,04-opc-and-packaging.md,10-bindings-spec.md
  reply-authoring paragraphs, coordinated with Root's C1 prose.

No Native document.rs, template.rs, comparison.rs, field.rs or lower OXML file
lease. No new source/module/crate/test binary/trait/generic/feature flag. Reuse
one substantive reply/anchor implementation through actual callers. Do not copy
comparison custody logic or introduce another marker policy beside the existing
source authority. Shared Python binding/test files are finite serial regions,
not a concurrent whole-file authoring right. Root coordinates source ownership
and serial integration with C1 and the other children.

## Test plan

**Test gate**: regression. Default iteration is scoped cargo check. Batch related
Native/CLI cases in existing binaries. Preserve actual Base RED separately from
fixture errors, then actual GREEN with exact source/binary/log custody. Run
independent microscope and /verify --scoped on the approved changed paths.
Fresh Python extension tests must bind the actual implementation and stub.
Do not run unrelated full suites or native oracle work per local increment.

## HLD impact

- docs/hld/03-architecture.md
- docs/hld/04-opc-and-packaging.md
- docs/hld/10-bindings-spec.md

Document crossing authoring and initials only. B retains comparison, selective
resolution and import prose. HLD14 is this integrator preparation registration.

## Risk routing

- Public API and PyO3: exact Rust/Python/CLI compatibility calls, strict installed
  typing/stub checks, archive/size and WASM riders in final integrated union.
- XML creation/relocation: HLD04/06 sequence and namespace rules, actual expanded
  marker identities/order in each supported owner, opaque/alias preservation.
- Shared authoring/cached state: same input across Rust/Python/CLI, held/fresh
  handles, save/reopen/remove/resolve and explicit failure/success publication.
- Contributed ordering contract: retained Issue434/PR444 pins and every A1 row,
  not donor output or a self-authored fixture as external native acceptance.
- External Word: fresh exact-candidate open on the actual installed runtime with
  candidate SHA, version/build, custody and no repair/thread/range/resolved-state
  observations. Preserve the historical constants until that actual final run.
- Evidence ownership: original review caps, final composition and scoped versus
  integrated policy remain separate. No local status implies provider completion.

## Hash harness and final obligations

Crossing reply XML is an intentional behavior change. Record any actual affected
hash delta in a separately labelled reviewed behavior commit at integration,
never a baseline rewrite by this worker. Untouched imports must stay unchanged.
Final union runs hash/schema, installed Python3.9/3.12 and strict typing,
CLI/SDK, archive/size, WASM, hosted CI and fresh Word evidence once as applicable.
B's comparison/resolution and all original parent riders remain mandatory.
A self-hash or historical Word version cannot satisfy the final native gate.

## Implementation checklist

- [x] Independent split review, Root adoption/approval and clean normal claim.
- [x] Actual baseline authoring/anchor seam revalidated at claim, lease confirmed.
- [x] Batched RED/GREEN for complete A1 authoring/initials and refusal matrix.
- [x] Existing C2 neighbor and unrelated marker/opaque behavior preserved.
- [x] Fresh Python, CLI and lifecycle/held-state controls with source custody.
- [x] Narrow HLD changes, independent microscope and scoped handoff.
- [ ] Serial integration and final native/installed/schema/hash/union riders.

## Open questions

No new product choice remains. Root approved and clean-claimed the authoring
split. Local implementation, scoped evidence and independent pass2 are complete.
The plan stays approved and the feature stays in-progress until actual final
integrated acceptance. Historical draft/proposed text above is preparation
lineage and does not override the operative adoption. No review cap is reset.

## D1 projectable reference endpoint clarification

The supported crossing authoring endpoint is the selected parent's actual end
boundary followed by its reference run in the same location. The parent end
must physically precede that reference. The reference run may retain formatting,
root-attribute records and empty Text, but no unrelated text, Field, Drawing or
opaque payload may be absorbed into the reply range.

A reference inside the parent's covered text or a mixed reference run without
an exact supported endpoint is nonprojectable for this authoring operation.
Refuse atomically with the existing no-range error while preserving its imported
read/save shape. Do not truncate the parent, expand the reply into unrelated
text, reorder parent anchors or introduce a new splitting grammar. Validate
projectability before reserving the new id. This closes the concrete D1 finding
within the approved supported-endpoint/refusal contract, not a security check
or a new marker policy.
