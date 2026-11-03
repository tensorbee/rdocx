# F-X280a, Refuse negative authored shape sizes before publication

**Status**: approved
**Sprint**: S96
**Size**: S
**Depends on**: none
**Test gate**: regression

## Problem

Original Issue435 A280-7 requires add_textbox and add_shape to refuse a negative width or height at the call, preserving package state, while allowing zero shape extents. Tables already refuse nonpositive extents. The actual shared textbox/preset/unstyled helpers currently pass unchecked sizes to positioned_transform. This child owns the complete original A280-7 across real native and Python callers, not the validation feature in A280-1 through A280-6.

## Spec reference

- docs/hld/10-bindings-spec.md, ShapeCollection additions and owner-atomic publication at lines1022-1044, existing error families at the PowerPoint binding error contract.
- docs/hld/06-presentationml-model.md, shape-tree model and authored geometry/transform intent. No imported geometry read policy changes.
- docs/hld/01-glossary.md, EMU units. Existing signed i64 values and unit-constructor truncation remain unchanged.
- docs/hld/14-development-backlog.md, original F-X280 A280-7 and proposed child partition.

## Source and independence

Reviewed product Base is ed588356bc50982516fff571eeba21bde8d76868. Current HEAD e6fec886557f819fc28f5ff1e0196edd458b1371 differs only by the previously approved F277 split/claim metadata. SOURCE_PINS.json authenticates the complete changed-path set and the actual shape, native test and Python carriers. The worker must start at a newly reviewed metadata-only descendant, never copy unfinished canonical Word source. F276 and F281 integrated reviewed product changes already exist in the product Base. Their final qualification remains pending. They are resource predecessors, not unfinished providers of this guard. A280-7 uses no F272 or F277 product mechanism.

## Approach

Add private sized_transform(operation: &'static str, left: Emu, top: Emu, width: Emu, height: Emu) -> Result<CT_Transform2D> beside positioned_transform in existing crates/rpptx/src/lib.rs. If width.0 < 0 or height.0 < 0 return existing Error::InvalidShapeMutation with the caller's operation and message `shape width and height must not be negative, got {width} by {height} EMU`. The message formats raw i64 operands width.0 and height.0, because Emu has no Display implementation (oxml-core/src/units.rs:41). Otherwise call positioned_transform unchanged. Wire textbox_member with operation `add textbox` and preset_member/unstyled_preset_member with `add shape`, propagating with ? before constructing a shape. Do not replace positioned_transform globally.

Retain the existing build-before-append ordering in append_new_member. Preserve both native local append and candidate transaction paths and Python's existing textbox shortcut/candidate paths. Leave table_member, CT_Table::new and connector_transform unchanged. Tables still require positive extents. Connectors remain endpoint-based, including reversed and equal endpoints. Left/top are signed positions, not dimensions, and remain supported when negative.

SlideMut::add_unstyled_shape is pub(crate) and has real HTML/ODP import consumers. ShapesMut has no add_unstyled_shape method. Test the actual private helper/SlideMut route in the existing lib.rs unit carrier, rather than inventing a public method. No Python or stub source change is expected because the actual wrappers already call the shared native helpers. If actual RED/GREEN demonstrates another change is necessary, stop for a bounded design amendment before expanding the source lease.

## Consumer map

| Rule | Actual call path | Publication boundary | Required evidence |
|---|---|---|---|
| Textbox dimensions | SlideMut::add_textbox:22196 and ShapesMut::add_textbox:22349 to textbox_member:22486 | append_locally or candidate append, before append_new_member publishes at22660 | Native call-time error and unchanged package/identities |
| Styled preset dimensions | SlideMut::add_shape:22211 and ShapesMut::add_shape:22360 to preset_member:22496 | Same shared native append boundaries | Same negative/zero controls, style preserved |
| Unstyled preset dimensions | SlideMut::add_unstyled_shape:22226 to unstyled_preset_member:22517 | Actual crate-private importer authoring route | Existing lib.rs unit module tests private route and zero/positive no-style behavior |
| Python dimensions | ShapeCollection.add_textbox:2258 and add_shape:2278 in rpptx-py/src/shape.rs | add:1843 textbox fast path commits only after native add, ordinary path prepares/publishes candidate later | Actual installed Python error, package equality, held collection/shape/text handles remain usable |
| Existing table rule | Native table_member:22570 to oxml-drawing CT_Table::new:312 and Python add_table:2419 | Existing build-before-append | Negative and zero table refusal unchanged, positive table still accepted |
| Existing connector rule | connector_member:22535 to connector_transform:22879 | Existing append paths | Reversed/equal endpoints still accepted, expected geometry retained |

## Behavior and errors

No public signature, type, default, method, feature flag or error variant is added. Native textbox/preset negative dimensions return InvalidShapeMutation at the call with the existing operation labels. Actual Python rpptx_to_pyerr maps this existing variant to public rpptx.RpptxError, not ValueError. Preserve that family and its message. Unsupported preset and malformed owner precedence is not broadened into a new product guarantee. The guard is selected inside the existing builder, so existing owner-validation failures may still occur first.

Zero width or height is accepted for textboxes and ordinary/unstyled preset shapes. Both-negative dimensions refuse. Negative left/top with nonnegative dimensions remain accepted. No imported XML extent policy, upper limit, resize behavior, table zero rule, connector span rule or schema validator changes are part of this child. Successful positive/zero output remains the same bytes as the existing implementation permits. Failed calls create no shape, relationship or part and advance no published owner revision.

## Acceptance criteria

All subrows implement original A280-7. They add discriminators, not new issue scope.

| ID | Input and consumer | Required observation | Local evidence | Final evidence |
|---|---|---|---|---|
| A280-7a | Native SlideMut and ShapesMut, textbox/preset, negative width and negative height | Existing InvalidShapeMutation with correct operation and negative-size message at call, unchanged package bytes and shape IDs | issue435_new_shapes_refuse_a_negative_size in existing integration.rs | Integrated native regression |
| A280-7b | Actual unstyled helper/SlideMut route, negative/zero/positive dimensions | Negative refuses before append, zero/positive accepted, no style introduced | Existing lib.rs unit tests | Integrated regression |
| A280-7c | Same native publishers, zero width or height and negative left/top | Calls succeed, save/reopen retains exact extents/positions, ordinary valid deck has empty validate | Batched native regression with positive controls | Integrated reopen and package qualification |
| A280-7d | Installed Python root/group shapes add_textbox/add_shape, negative width/height | rpptx.RpptxError at call with must not be negative, bytes/shape count unchanged, previously held collection/shape/text handles retain authority and remain usable | test_issue435_add_textbox_and_add_shape_refuse_a_negative_size in existing test_documented_examples.py | Actual final installed Python3.9/3.12 and typing suites |
| A280-7e | Existing native/Python table and connector controls | Table negative and zero still refuse, positive table succeeds. Reversed/equal endpoint connectors retain existing behavior | Same focused batch plus existing connector/table controls | Final integrated consumers |

## Test matrix

| Cohort | Selected combinations | Rejects plausible wrong implementation |
|---|---|---|
| Native refusal | SlideMut root shortcut, ShapesMut root and populated group candidate path, textbox/preset, width<0 and height<0 distributed across callers plus both-negative helper | Python-only or one-member-only guard, a guard after append, shortcut-only coverage |
| Atomic publication | Seed existing shape, relationship-backed member and opaque package part. Compare complete serialized bytes and IDs after each failure. Python retain actual handles and reuse collection after failure | Partial child/id/part/relationship publication, hidden owner-generation advancement |
| Boundaries | width0, height0, both0 and positive dimensions, negative position | Accidentally using <=0, checking left/top, clamping or abs dimensions |
| Unstyled route | Actual private helper/SlideMut add_unstyled_shape with negative and accepted controls | Missing importer path or silently introducing default style |
| Controls | Existing table negative/zero/positive and connector reversed/equal endpoint cases | Broadly replacing positioned_transform or imposing shape policy on unrelated constructors |

Use risk-based cases, not every cross-product. Rust borrow rules do not permit an immutable shape reference to coexist with a mutable deck borrow. Assert native state/IDs after the call and use actual Python held handles to test live owner generations. Do not claim a synthetic Rust held-reference test.

## Test plan

1. Read source and actual Base, add the decisive regression module/functions in the existing carriers, and observe semantic RED against the real unmodified product. A compiler/setup failure is not RED. Preserve the exact test source, command and terminal outcome before production edits.
2. Implement the one private helper and three call-site replacements. Batch the shared Rust regression/unit controls with CARGO_BUILD_JOBS=2 when appropriate. Iterate with cargo check -p rpptx --all-targets by default. Run the new focused test names after source changes, affected Clippy with warnings denied and fmt.
3. Exercise the installed Python test against an actual child-source wheel if a suitable existing local producer is available. Otherwise retain exact source and pending installed evidence, do not invent a Python pass. Existing RpptxError and handle patterns are primary source behavior.
4. Run /verify --scoped F-X280a and a distinct read-only microscope on the exact child source. Do not run the workspace union, hash recording or external oracles per worker.
5. Final integrated owner runs /verify --full, the applicable SDK/native PowerPoint/reopen observations retained from the umbrella, installed Python versions/typing, WASM, archives/publish dry-runs, hosted CI and hash harness once over the union. All are currently unexecuted for this child. Refusal itself produces no native document to open. Accepted zero/positive outputs supply reopen/native controls.

No new source file, test binary, crate or module is authorized. Unit tests live in the existing lib.rs test module. Native and Python tests use existing entrypoints.

## HLD impact

- docs/hld/10-bindings-spec.md

Update only the ShapeCollection authoring/atomic-publication paragraph around1022-1044 to state negative textbox/preset-size refusal, accepted zero dimensions, unchanged table positivity and existing RpptxError mapping. Preserve the OfficeMath sections reserved for F277a, all Word/F272 carry and the pending validation/Display-policy sections owned by Root/F280b. HLD04 and HLD06 require reading but this constructor guard changes no package mechanism, imported model or parser contract. HLD14 split registration is integrator planning work.

## Risk routing

Public published behavior requires strict SemVer classification and final publish/archive checks. Existing public signatures stay unchanged, invalid authoring refusal is prepared for shared/PowerPoint0.15.0, no publication authorized. Binding/shared-consumer routes require native/Python same-input and held-handle atomicity checks plus final installed Python/typing and WASM. Units/boundaries require preserving EMU signed positions and truncation, with negative/zero controls. Completion/issue accounting requires final metadata consistency and every original A280 row retained. No parser, serializer, font/layout algorithm, budget, dependency edge, new framework or baseline move is introduced. Final riders add to the ordinary gate, not replace it.

## Hash harness

Expected unchanged. Refused negative authoring produces no output. Zero/positive authoring bytes are preserved. Worker does not record a baseline or edit README/archive measurement rows. Any unexpected final delta blocks acceptance pending explanation and review.

## Review lineage and coordination

The original F280 design used all three Grok peers, confirmed by F-X280-codex-preparation-final-2.md. This extraction inherits those exhausted passes and relevant dispositions. No new peer/reset is created. Root performed the affected split/contract audit and approved this child in .claude/reviews/F-X280-split-design-codex-final-1.md. The intended implementation is one complete GLM-5.3 Flash DGX high-effort story through host opencode. Keep its actual work/repair/evidence narrative separate from Root harness work. Its self-review does not count as an independent microscope. The original implementation cap remains at most three, no completed implementation review for A280-7 is claimed here. Root final read-only microscope can supply the actual independent assessment, not an author's self-review.

Canonical F272 source/build lease remains with its current owner. Proposed exception is only a reviewed metadata-only Base/claim despite preserved unrelated canonical carry, and /private/tmp/rdocx-s96-f280a-opencode instead of the normal sibling path. Root must approve that explicit exception, obtain a brief writer metadata window, and confirm disjoint worker leases. Proposed worker is host opencode, backend GLM-5.3 Flash DGX high effort, branch work/f-x280a-opencode. F277a worktree/owner and its math/native/Python/HLD sections remain intact. No worker commit, push, GitHub comment, tag or publication occurs from this draft preparation.

## Implementation checklist

- [x] Root approves partition and bounded design, actual provider independence, metadata/path exception and resource windows.
- [ ] Fresh mutable issue/PR check, reviewed metadata Base, clean worker product verification and /start-feature F-X280a --claimed.
- [ ] Actual semantic RED, private helper and three call-site implementation, focused GREEN and Python evidence or exact pending installation rider.
- [ ] Update only declared HLD section, scoped gate and independent zero-finding microscope with actual source pins.
- [ ] Prepare and validate structured handoff only after Root's acceptance, preserve branch through integration/verification.
- [ ] Integrator consumes final applicable riders and completes child without claiming parent validation/provider completion.

## Open questions

No new product choice is proposed. Root approval of the split/design, exact metadata/path exception, source/test/HLD lease and final evidence reconciliation is pending. Whole-product check and claim have not been performed. F280b remains draft with its F272/F277 barriers and Root's pending Display-policy reconciliation. This child does not approve that remainder.

## Root approved coordination and claim readiness

Root final1 approves this complete unchanged A280-7 child, with no F272/F277 product dependency. Root approves the exact metadata-only Base exception, /private/tmp/rdocx-s96-f280a-opencode and existing source/test/ShapeCollection HLD lease. The actual implementing host is opencode, backend GLM high. The user restored server availability and a fresh sentinel tool check passed, with inherited hooks preserved. A fresh Issue435 body and PR439 head are unchanged at11:44UTC. Parent/remainder remain draft and pending. Original review caps and final debt remain unchanged. Actual claim/clean worker and semantic implementation evidence are still required. Root independent microscope follows identified GLM self-review before preparation.
