# F-X273, all, pass 6

**Reviewed**: Frozen working main-remediation result on `sprint/s96`, HEAD `6506fa660372b2cfe7dd620cfaaf0959a62ba640`, under approved plan `f39cf22966952420d61a45b8d88e26179442b3b2dc2ffb418b4ba55ef77f130c`. The real-before incremental patch contains 9 files, 7,463 additions and 1,979 deletions. This includes the ignored progress record's 94 added lines. Production, tests and HLD comprise 8 tracked files, 7,369 additions and 1,979 deletions. The final successor against the previous failure handoff contains 4 files, 182 additions and 30 deletions.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the reviewed main remediation. This meets this review's exit condition. It does not complete F-X273 or qualify final native and integrated acceptance.

## Scope and identity

This independent Codex review read the complete main incremental patch during read-only preparation, then the actual final successor, affected source, test assertions, consumers and HLD against the full applicable K1, R4, producer, annotation, allocation, mixed-route and cluster contracts. The historical LOCAL provider checkpoint remains completed under its approved exception. Its tracker status does not describe this later remediation as accepted. Prior all-pass5 and provider-prefix-pass4 remain historical reviews of their respective producers, not executions on this result.

The packet is `/private/tmp/rdocx-s96-run-refresh/fx273-main-implementation-v1`. Independently recomputed identities are:

- `final-handoff.json`: `6e78c556022d1f08eba6c35a546c38a9743a2e1644790485a777e5a9a9a60247`.
- Full real-before patch: `55e605063ff9f198d9299c77b4838e7d78cfbcdee75df17de35da52900378c48`.
- Narrow final successor: `e4df171bd71f446f2241f97c6a43d46f94b5787b3d6aa0abf1b8dddedfa17b06`.
- Final seal: `5caaad46d3d65b23418c33a30088c924b473d53e446c758e49120a816c88e065`, all 724 members authenticated.
- Final child receipt: `90aff917d42e0cf05e33362b9ee3171743eab5f95e17ec4e1c36cc95eebeae34`.

All 12 final source/document/plan pins match actual files. Comparing the 3,844 acquisition source inputs to the review tree finds only the separately adopted F-X275 plan changed. That unrelated plan does not change these product artifacts. The ignored progress update is outside the acquisition snapshot and explicitly identified as such. The complete earlier 554-member failure seal and both earlier 103-member child seals also authenticate. No unrelated dirty sprint work is attributed to this feature's incremental patch. convert.rs and style_resolver.rs remain unchanged from this implementation lease's before copies.

## Defects

0 findings.

## Smells

0 findings. The preparation observation about the redundant legacy predicate is addressed at `crates/rdocx-layout/src/engine.rs:16786`. The final branch now directly appends the unchanged base and None language with the actual kern entry.

## Nitpicks

0 findings.

## Correctness and contract consumers

`crates/oxml-layout/src/font.rs:265` selects the private legacy exception during construction, rather than inheriting it after a strict constructor has refused. Public new at line230 still selects false. The existing full coverage/range/monotonic validator at line535 must succeed in one of the two actual glyph orientations for known legacy context. Invalid levels, lengths, finite positions and incompatible base data remain rejected. Compatible base replacement at line334 applies that exception only to an already legacy owner. Factory empty/shaped branches, contextual reshape and the legacy hanging slice use the same constructor. No glyph reversal, layout-level rewrite, extra stored direction or new allocation authority is introduced.

The seventh method at `crates/oxml-layout/src/font.rs:391` preserves the actual payload while assigning assembly ordinal and explicit legacy convention. `crates/rdocx-layout/src/engine.rs:16761` keeps the old eligible rich batch separate from actual excluded plain-field owners, consumes the correct queue and assigns checked global ordinals. Structural and atomic-marker routes retain their known producer context. Legacy carriers preserve former Text splitting, bidi slot movement and physical hanging behavior. Genuine rich reconstruction remains separate.

The neutral origin bridge and its consumers capture actual Group/Figure occurrences during construction and flush, then move those bindings by the actual bidi permutation. `crates/rdocx-layout/src/block.rs:229` binds annotation producers to those slots without recovering ownership from equal output values. Annotation recipes retain base-only field roles. `crates/rdocx-layout/src/engine.rs:9466`, line9921 and line10274 regenerate complete first annotation occurrences and remove continuation decorations under one shared occurrence state. Empty owners, ordinary/annotated fragment ordering and actual repeated table-header scopes remain explicit.

Field membership and subtree surgery were traced through `crates/rdocx-layout/src/block.rs:858`, line1067 and line1173, page/background/artifact wrapping, source rebinding, notes, furniture and table adapters. The checked operations move maps with actual glyph storage and validate exact field paths before publication. Paragraph, table, furniture and restart accounting retains private vectors, paths, recipes and Arc charges. `crates/rdocx-layout/src/engine.rs:8811` requires positive private legacy producer proof for cache admission. Raw and substituted maps participate in exact cache equality. The existing configured cache ceilings remain unchanged.

R4 was traced through paragraph/split/column placement, table and prepainted-row slices, anchored shape children, rotated/clipped cells, furniture composition and same-page late endnotes. `crates/rdocx-layout/src/paginator.rs:2017` scopes actual repeated headers. Line4325 carries checked page descriptors through endnote insertion, and line7242 records the actual row intervals. Child shape layers stay inside their outer anchor. Page order retains back border, watermark, body/furniture back, regular text, body/note front, furniture front and front border. Existing unsupported ordinary note/nested anchor placement is not claimed as newly implemented.

The final late-field guard at `crates/rdocx-layout/src/engine.rs:9722` refuses a contradictory legacy owner on a rich positioned leaf. Normal legacy publication remains Text through paginator line5931 onward, and annotation builders emit plain base GlyphRuns. Genuine rich late replacement remains subject to the unchanged public MultilingualGlyphRun invariant. No relaxation of output.rs validation is present. HLD08 describes these finite producer, ownership, cluster and composition rules while retaining native qualification boundaries.

## Actual evidence and discriminators

The successful CargoJSON build has a genuine successful build-finished event and matching executable compiler events for oxml-layout0.15.0 and rdocx-layout0.17.0. Both original and frozen executables match their recorded hashes: `86ea84962e1f9f93b488b6f70f25335955e4c1a81874ca6a232e2f225e033327` and `a086cf136679bef2da40d02619d1a42ac8847b4599895c14897249a002feb867`. The actual profile is opt1 with assertions and overflow checks enabled, offline dependencies and 8MiB stack. These are the default-feature affected lib artifacts, not an invented all-feature or optimized release acquisition.

Independently checked all 47 raw child outputs against exact named receipt membership: 19 shared-layout and 28 Word-layout bodies, each exit0, one started named PASS and zero ignored. The same finite set retains the original public K1 painted-advance and front-order REDs, constructor/map refusals, mixed-script/RTL/neutral/hanging and conditional-hyphen controls, origin parity, spacing/tabs, empty annotations, stored child fields, repeated headers, cache identity/bounds, notes/endnotes, columns and line-number/rotation consumers. Assertions inspect actual glyph payloads, publication variants, ownership, order, metrics and warm/fresh equality.

The new legacy control accepts a fully covered opposite-order map while strict public construction refuses that mismatch. Malformed maps remain rejected. The mixed old-Text control includes rich-initial Latin whitespace and repeated reflow. Late replacement checks positive plain output, positive publicly valid rich output and unchanged input after contradictory-owner refusal. It does not treat a fabricated rich legacy producer as a new supported Word route.

The original K1 and front-order executions genuinely exited101 at their named assertions. Their raw stderr records incorrect painted advances and a front anchor before later body text. Earlier 42 PASS/5 FAIL and 45 PASS/2 FAIL acquisitions remain immutable failed batches. The fixed99 fixture correction follows the existing PAGE placeholder producer and preserves equal-value/opposite-policy discrimination. The spacing reference retains explicit-off shaping plus the same spacing assertion. Cache tests retain both configured limits and assert the actual FIFO suffix and exact current paragraph-key identity. None of these corrections retroactively changes an earlier result.

Actual affected strict clippy9 passed in11.217seconds for both crates/all targets with warnings denied. Formatting, prose and diff checks passed at the recorded producer. Prose reports zero violations. These observations were consumed without rerunning tests or builds during review. The ignored private-face/HarfBuzz caps oracle was not executed or counted among the47.

## Not found and remaining acceptance

Correctness, contract, panics, OOXML, tests and structure: 0 additional findings in this scope. No new parser/serializer or source XML normalization, module, trait, generic, public model or public field is introduced. Public additions remain the seven approved hidden methods and the one approved neutral origin function. This main incremental scope does not re-review unrelated F271/F272/F275 edits or promote older navigation/native evidence to current acceptance.

Actual repaired-artifact native fixture/font comparisons remain pending. Native control is unavailable and this review grants no fallback or waiver. Final integrated full/risk qualification, reviewed49hash/7pixel deltas, deterministic PDF/PNG/text, installed runtime/typing/stubs, WASM, hosted gates and the ignored caps oracle remain due. The F-X272/F-X273/F-X274 lifecycle and source-qualified union still requires its applicable final evidence. Missing CLI/Python/native executions are not supplied by these Rust libtests. No baseline, tracker, completion, commit, push, tag, reporter comment or release action is authorized by this report.

End review and return the sole canonical writer lease to root. No unchanged-scope confirmation pass is requested.
