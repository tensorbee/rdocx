# F-X279c, Author comment markup and styles without changing reply ranges

**Status**: approved
**Sprint**: S96
**Size**: M
**Depends on**: none
**Test gate**: regression

## Split approval and review lineage

Root approves this exact split after independent final review in
`.claude/reviews/F-X279c-split-design-codex-final-1.md`. The reviewed clean
product Base is `987027d41cf87b476f46427482f4693e823cb702`. Only the recorded
planning/citation/claim metadata may enter its descendant claim commit. The
worker path is `/private/tmp/rdocx-s96-f279c-codex`, branch
`work/f-x279c-codex`, cut only after the normal clean claim. No product code is
approved as implemented. F-X272 remains dirty in its separate recovery tree,
whose lower comments source lease is frozen for semantic integration.

This child extracts the complete original A279-2 and C2 only from the approved
F-X279b contract SHA256 `4b2f1e57b204cb8f77d4b9b748bb375cde1eadf04aa404e954562f356a7bdd17`. Its existing three design peers and Root
final2 lineage remain applicable history. This is a bounded split review by
Root, not a new design-peer allowance. No design or implementation counter is
reset. The original implementation cap remains three. F-X279b retains A279-1,
A279-3, A279-5, A279-6, all P1/P2 and C3/C4/C5/C6/C8-C14 composition work and
its actual F-X272 implementation prerequisite. F-X279a retains all A279-4.
The umbrella waits for all three children and actual F-X272 completion.

## Problem

At the clean Base, Native `comments.rs:2087` creates plain comment paragraphs.
`push_comment_definition:1639` and `reply_to_staged:1701` share this path, but
neither creates CommentText/CommentReference styles or annotation runs.
Lower `comments.rs:115` preserves root attributes, while its private
push_preserved_attributes at361 forcibly rebinds canonical w/w14. Existing
Python authoring already reaches these Native methods. No Python/CLI signature
is missing for A279-2. Initials and crossing reply ranges belong to B.

## Spec reference

- HLD03, Shared Word document operations, comment mutation and staged publication.
- HLD04, Part naming, Word comment part creation. Package integrity, comment
  deletion/threaded-comment preservation and Checked related-story publication.
- HLD10, Native Word facade stability, Python comment authoring/lifecycle.
  Python API shape, held handles and atomic revision publication.
- HLD14, F-X279c and original F-X279 regression gate.
- S96 SPRINT_PLAN original Issue434 A279-2 and C2, retained in SPRINT_PLAN.md.
- ECMA-376 Part3 Markup Compatibility, mc:Ignorable, and WordprocessingML
  CT_Comments/CT_Comment/CT_P/CT_R child order, as already adopted by F-X279b.

## Approach

Reuse the original behavior below without a new public API, parser policy,
namespace engine, input allowance, quota, trait, generic, module or test binary.
Only existing files are leased. No historical donor source is adopted.

### Complete inherited A279-2 contract

On paragraphs created by `add_comment`, `add_comment_on_text` and `reply_to`, the first paragraph serializes in this order:

```xml
<w:p w14:paraId="...">
  <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
  <w:r><w:t>text</w:t></w:r>
</w:p>
```

Later paragraphs have `CommentText` and no annotation run. `w:annotationRef` is the existing raw child after `w:rPr`, not a new enum.

When either style id is absent, authoring creates it after `w:latentStyles` and after existing `w:style` siblings, never at index 0. `CommentText` is a paragraph style and `CommentReference` is a character style. `w:basedOn` is written only when `Normal` or `DefaultParagraphFont` respectively already exists. If either base is absent, that style has no `w:basedOn`. If the style id exists with the wrong type, authoring refuses atomically and writes neither the comment nor a second style. Existing style bytes stay. Style names, priority and donor sizes are not acceptance values.

For a comments part this story creates, the root declares `xmlns:mc` and `mc:Ignorable` including `w14`, in addition to the existing `xmlns:w` and `xmlns:w14` behavior. If a preserved `mc:Ignorable` exists, add `w14` when missing and keep the other tokens and order. Do not duplicate `xmlns:mc`. On an imported part that is actually authored, identify an existing Ignorable attribute by its Markup Compatibility namespace URI and local name, not solely the literal prefix. Reuse an existing correctly bound namespace prefix when available. Otherwise use canonical mc if free, or a fresh unused prefix if mc belongs to preserved producer content. Do not rebind or delete foreign declarations/attributes. Merge w14 into the one existing MC token list without changing the other token order. New parts retain the exact canonical mc expectation from Issue434. Extend C2 with aliased MC and occupied foreign-mc controls beside untouched-save controls. This is Root technical reconciliation of imported-preservation invariants, not a new public API or a blanket refusal. An untouched reopen and save of an imported comments part does not gain `mc` or styles merely because `to_xml` runs. Imported unknown root attributes and children stay verbatim. `commentsIds.xml` is not created. An imported part of that name stays byte for byte, including relationships and content type.

Fixed-writer admission precedes authoring publication. Inspect the original comments root attributes before push_preserved_attributes can replace xmlns:w or xmlns:w14. An explicit canonical w binding to a foreign URI, or canonical w14 binding to a foreign URI, is incompatible with this bounded fixed-prefix writer and refuses atomically with the existing location-qualified XML error. Do not silently canonicalize the imported URI. Apply the same original-binding check at actual owned comment/paragraph destinations where their retained declarations would change the expanded type of a generated fixed-prefix element or attribute. The check concerns typed writer destinations, not an attempt to validate or rewrite opaque payload XML. Use the existing namespace/context machinery in the current comments/OXML files, without a new public namespace API, parser policy or quota. Validate before either primary add path or reply publishes its staged candidate, including reply_to_staged which bypasses push_comment_definition. CT_Comments serialization must not retain the current forced foreign root rebinding.

A genuine Word/W14 alias with canonical names absent or correctly bound remains supported. For accepted authoring the fixed w14 declaration is correctly bound before adding that Ignorable token. Occupied foreign mc still uses the existing correctly bound MC alias or a fresh unused alias, without blanket alias refusal. A valid imported W14 alias beside foreign canonical w14 and a preserved producer w14:opaque child is the atomic-refusal control. A corresponding foreign canonical w root is separate. Root, existing comment/style/package bytes, cached layout and held Python revision remain unchanged on refusal. Untouched imported source-save stays byte-identical and does not invoke authoring merely to add styles or declarations. Positive alias/MC controls inspect expanded generated names and retained producer identities after save/reopen. This is technical closure of imported-preservation and atomicity, not a newly promised general prefix-renaming engine.

The pin `WORD_COMMENT_CANDIDATE_SHA256` at `comments.rs:3679` is `b7e1f39a5af80d9928ed671fa45557d485c2c70c8761439a09df66c027274995`, Word `16.104`, build `16.104.25121423`. That is historical. Authoring changes the candidate bytes. Update the constant only in the labelled behavior commit with old and new self-hash. Do not present the self-hash as Word evidence. Root observed installed Word `16.114.1 (26102031)`. A fresh native open of the new candidate on the actually installed version is mandatory and unperformed. The ignored `word_opens_comment_reply_and_resolved_thread_without_repair` currently generates the candidate and checks its digest plus installed bundle version. It does not itself open Word. Keep it as the candidate-generation/pin check, update both WORD_VERSION and WORD_BUILD to the actual final bundle values alongside the new candidate digest, and preserve old/new identities in the labelled behavior commit. The inspected installed bundle is CFBundleShortVersionString16.114.1 and CFBundleVersion16.114.26102031, not a runtime/open observation. Actual native acceptance additionally requires a separate CUA-controlled open of those exact new bytes without repair and the required thread/range/resolved-state observations. Record candidate SHA, actual running Word version/build, custody and the observed result. Existing Word custody and generated-folder permissions still govern. No AppleScript fallback is implied. Q-A430-28 applies only to its named shared-footnote-owner observation and cannot waive this comment gate. Until actual execution, A279-2 native evidence remains pending.

### Operative pin and native-evidence timing

Root's split instruction supersedes only the historical pin-update timing in
that inherited paragraph. Preserve the historical WORD_COMMENT_CANDIDATE_SHA256,
WORD_VERSION and WORD_BUILD constants until actual new native execution is
available. Local regressions must record the actual candidate artifact and SHA
without presenting the historical ignored self-hash test as passing on new
bytes. Its current old-candidate assertion remains historical, not a new failure
or a waiver. After actual final native execution, the integrator reconciles
those existing constants and records old/new identities in the labelled
behavior commit. No old candidate digest or installed bundle version is an
actual native-open result. The exact new candidate must open without repair
and retain comment/thread/resolved state on the actual running Word build.
The B crossing-range candidate remains a separate composition check after B.

### Concrete source lease and implementation boundary

- `crates/rdocx/src/comments.rs`: existing authoring helpers, staged failure
  paths, comment paragraph/style construction and existing unit-test module.
  Include all add paths and reply_to_staged, which bypasses push_comment_definition.
- `crates/rdocx-oxml/src/comments.rs`: minimal original canonical w/w14 checks
  before root preservation, removal of forced rebinding, and write_comment/
  actual paragraph destination admission. Preserve the rest of the reader.
  Root has frozen its C1 hunks in this file until child integration. Integrate
  semantic hunks only, retaining C1 parser controls and emitted-root scope work.
- `crates/rdocx-py/tests/test_core.py`: disjoint new issue434 comment-markup
  tests in the existing carrier. No edits to C1 tests, Python source or stub.
- Narrow comment-markup paragraphs only in HLD03/HLD04/HLD10. Root coordinates
  these sections against C1. No C1 source, template, comparison or import lease.

Use public CT_Comments.root_attributes for authoring-side MC token/alias merge.
The clean lower serializer already preserves those attributes. Existing
namespace_prefixes_at/capture_attributes and numbering namespace helpers can
inspect borrowed original declarations in the comments module. Existing
CT_P::raw_is_root_attributes/is_root_attribute_record identify retained root
records before the paragraph writer rewrites them. Use existing quick-xml
record and namespace machinery for those actual stored records, with a private
comments-file admission helper as necessary. Do not call to_xml merely to
manufacture a validation witness. Do not require dirty C1 helpers or add an
API in text.rs. If an original destination cannot be checked within this lease,
stop and report that exact source-backed prerequisite before widening scope.
The original-binding refusal applies to actual fixed-prefix typed destinations,
not arbitrary opaque subtree content.

Use existing style builders/model operations. Validate both existing style ids
and their types before publishing the staged candidate. Preserve style ordering,
existing bytes and source namespaces. No commentsIds synthesis or relationship
rewrite belongs here. No comparison, insert_document, Field source/output work,
imported scalar ownership or Native lineage consumer is added.

## Consumer map

| Rule | Entries and output | Shared owner | Lifecycle |
|---|---|---|---|
| Comment paragraphs/styles | add_comment, add_comment_with_date, add_story_comment, add_comment_on_text, reply_to/reply_to_with_date | Existing Native authoring helpers | staged model, save, reopen |
| MC and fixed-prefix preservation | Authored CT_Comments, lower to_xml/write_comment and actual paragraph destinations | Original stored bindings, lower writer | refusal/rollback, emitted names, preserved imports |
| Python authoring | Existing add_comment/add_comment_on_text/reply_to wrappers | Existing Native policy, no new signature | live revision, held handles, bytes, reopened model |
| CLI comment authoring | Existing comment add/reply commands | Existing Native policy, no new option | final actual CLI composition |
| Native acceptance | New comment/reply/resolved candidate | Integrator's actual Word execution | exact artifact/build, no repair, preserved thread |

## Behavior contracts and evidence

| Boundary | Exact behavior | Authority | Discriminator |
|---|---|---|---|
| First comment paragraph | pPr/CommentText, annotation rPr/CommentReference plus raw annotationRef, text run | Inherited A279-2 above | child order and one annotation on multiline text |
| Styles | Missing only, correct type, basedOn only with existing base | Inherited C2 | no bases, existing correct/wrong type, atomicity |
| Imported MC | Expanded URI/local Ignorable match, token order retained, fresh alias only if needed | Approved Root namespace closure | aliased MC, foreign occupied mc, unknown attrs/children |
| Fixed w/w14 destination | Original incompatible canonical binding refuses before coercion | Approved F-X279b Root final2 D1 closure | separate foreign-w and foreign-w14 plus valid aliases |
| Untouched imports | No added styles/MC, exact retained parts/relationships | Inherited A279-2 | save bytes, commentsIds remains optional/unchanged |
| API compatibility | Existing signatures/defaults/returns unchanged | Current clean source | Python live state/revision and native rollback |

No numeric conversion or new getter/clear contract is introduced. Initials and
reply body anchoring retain their existing behavior here, with B changes pending.

## Dependency readiness

No new C1 provider is consumed. The clean Base already has the comment, style,
OPC and Python authoring APIs used here. Source inspection is not execution.
Root has independently approved the split and exact lease. The child may start
only after metadata is committed and the normal clean claim succeeds. Root C1
lower comments.rs and overlapping HLD sections stay frozen until integration.
The parser-side C1 hunks remain excluded from the child Base and untouched by
its worker. No other dirty worker material is copied. If borrowed destination
checking cannot be implemented with clean existing helpers in comments.rs,
readiness becomes blocked and Root must adjudicate the exact missing seam.

F-X279b's F-X272 dependency remains genuine. No plan approval, reviewed child
or pending native gate is a provider checkpoint. C local source readiness does
not close the umbrella, transfer P1/P2 or complete any remaining B acceptance.

## Acceptance criteria

| ID | Input | Expected result | Consumers | Local evidence | Final rider |
|---|---|---|---|---|---|
| A279-2/C2 | New add_comment_on_text and reply, with/without Normal and DefaultParagraphFont | Exact inherited markup/styles, mc+Ignorable w14, no commentsIds | Rust/Python authoring, existing CLI | issue434_comment_markup_and_optional_ids | fresh actual Word open of exact candidate |
| C2-preserve | Imported aliases, foreign mc, unknown attrs/children and untouched save | Expanded names correct, token order and unowned bytes preserved | Native/lower/Python save | extended C2 controls | union schema/installed composition |
| C2-atomic | Wrong style type, foreign canonical w/w14 or actual incompatible typed destination | Existing location-qualified error, no model/package/style/layout/revision change | Rust and Python | held/fresh/saved rollback controls | final C1 shared-consumer reconciliation |

## Test matrix

| Case | States | Observation | Wrong implementation rejected |
|---|---|---|---|
| New/multiline | add paths and reply, one/many lines | One first annotation, styles and schema order | only add fixed, all lines annotated |
| Style availability | both/one/no bases, existing correct/wrong ids | correct basedOn omission and atomic refusal | invented base, duplicate style, partial publication |
| MC aliases | canonical/alias/occupied foreign mc | one expanded Ignorable, stable tokens/producer identity | literal-prefix search or rebinding |
| Fixed writer | foreign w and foreign w14 separately, valid aliases | refuse only incompatible typed destinations | forced URI rewrite or blanket alias refusal |
| Imported lifecycle | untouched and actually authored, commentsIds absent/present | untouched bytes, optional part unchanged | global serializer adds MC or durable ids |
| State publication | failure and success with held Python views | rollback or single revision publication | partial styles/comments/revision advance |

No new rendering or font-baseline promise. Comparison/revision resolution,
source import and crossing ranges remain B tests. No second test binary.

## Test plan

**Test gate**: regression. Use existing comments.rs unit tests and disjoint
issue434 nodes in test_core.py. Authenticate actual clean-Base REDs separately
from fixture/setup failures, then batch relevant GREEN cases in existing
binaries. Default iteration is scoped cargo check, not repeated big links.
Run /microscope and /verify --scoped with applicable local evidence. Python
nodes require a freshly authenticated local extension. Final installed3.9/3.12,
strict typing/stub checks, CLI composition, schema, hash, WASM, archive, hosted
CI, fresh Word native open and all original union riders run once integrated.
No per-worker full suite, native/oracle run, baseline update or publication.

## HLD impact

- docs/hld/03-architecture.md
- docs/hld/04-opc-and-packaging.md
- docs/hld/10-bindings-spec.md

Only authored comment markup/styles and untouched imported preservation belong
to C. Crossing reply and comparison/import prose remain B. HLD14 registration
is the integrator's preparation work, not a second worker impact list.

## Risk routing

- XML creation/serializer and imported namespace preservation: HLD04 and06,
  exact CT_P/CT_R order and opaque roundtrip controls, final SDK/native union.
- Shared Native/Python/CLI behavior: differential-testing composition rules,
  real same-input authoring/save/reopen and Python rollback observations.
- Existing published behavior: HLD10 and strict SemVer/archive/size riders in
  final union, no new public API or constructor semantics.
- External Word oracle: pinned candidate and actual running version/build,
  custody and no-repair/thread observations, final only and still pending.
- Completion/capability ownership: integrator's final policy reconciliation,
  no child capability-complete row or delivery ledger before actual acceptance.

## Hash harness

Authored comment outputs intentionally change according to A279-2. Declare
and independently review any actual affected hash delta in a labelled behavior
commit. Untouched imported outputs must remain unchanged. The worker does not
record a baseline or borrow another story's exclusive baseline lease. Final
union measures and attributes actual deltas once.

## Implementation checklist

- [x] Root split review/adoption/approval and exact source lease.
- [ ] Normal clean claim at the recorded metadata-only descendant.
- [ ] Scoped actual RED/GREEN and authored/imported atomicity controls.
- [ ] Native-only helper changes plus minimal lower original-binding admission.
- [ ] Preserve C1 parser/scope hunks at integration, without whole-file adoption.
- [ ] Narrow HLD intent, independent microscope, scoped gate and handoff.
- [ ] Final fresh native candidate, installed/schema/hash/typing/CI riders.

## Open questions

No new product choice. Root approves the split and exact clean-source lease.
Normal claim and actual native evidence remain pending. If clean-helper inspection
fails during implementation, report the missing seam rather than inventing a
new API, relaxing A279-2 or consuming unfinished C1.
