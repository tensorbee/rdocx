# F-X272 DeepSeek independent source checkpoint, Grok peer pass 2

Audit only. This is source peer 2 of at most 3. It is not Codex final, not feature acceptance, and not the pending all-pass-12 microscope. No product file was edited and no test was executed in this pass.

## Identity

| Item | Value |
|---|---|
| Reviewer | Grok, session `01a19364-efef-7bc1-97a0-a7ec4c726991` |
| Author | DeepSeek, independent source increment after peer 1 |
| Codex role | Router of this phase. No substantive final review is claimed here |
| Branch / HEAD | `sprint/s96` / `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd` |
| Plan | `.claude/plans/F-X272-design.md` `e37d2ded5174b69d322a053fe85b85afbedddd7448aacfb5033b5922454d9949` |
| Prior peer report | `.claude/reviews/F-X272-deepseek-handoff-grok-pass-1.md` `36e7c34e583bf23f2a281f2e04d9d6bb2c0f28c8dda5a07b646bddddf4c8703f` |
| Readiness text | `/private/tmp/rdocx-cross-review/fx272-independent-ready-handoff.md` `53779bfd2db79e4adb145ee2555c1a1d8f3b9eb7b1240a23b252bcd60e926662` |
| Sprint status | `CURRENT_SPRINT.md` and `BACKLOG.md` both keep F-X272 in progress. The Owner column remains codex |

The peer2 packet was authenticated before this report. All 30 manifest files match the canonical tree and `fx272-peer2-packet/current/`. The same check is repeated after the report write. `working-tree.diff` was not used as an author-attribution baseline. The three peer1 `after` sources, oxml `document.rs`, oxml `text.rs`, and facade `document.rs`, all differ from the current tree. Current oxml `document.rs` is `6d81fe0ab7a94fc723667061f96004da0c0eb711d12c5521c171014fc070d5f3`. Current oxml `text.rs` is `80b6de9c1b4fd4a0bc5d52a08da8c25be2ececb1d1049906f72b2f2a8bd97f6e`. Current facade `document.rs` is `756c7632e4e7e0b3fcbd3dd7f525851301819b2fa24b2ab4e6c0a0951f92d309`. Facade `field.rs` is now `a7efc10f7748323d4fdf9387fe14e47db844856b75186bfe005c5f7730710f6b`. HLD 03 is `e6cc502f9431dc3bda4f72c4a623ad5c78d87127fe7009343f165b96354dc5c8`. HLD 10 is unchanged from peer 1 at `c09cf197a66315737897aa597361ff2d32dfa1a109d8db67930924e985410ade`.

The live plan is `e37d2ded`, not the peer1 plan `50af8047`. The source-relevant additions are the D7 leaf, drawing, mutation, and capture sections, plus "Retained field raw declaration closure on generated paragraph roots". A later plan sentence limits the HLD impact list to `docs/hld/03-architecture.md` and `docs/hld/10-bindings-spec.md`. HLD 03 now records the shared ancestor envelope and the collapsed-hyperlink reservation. HLD 10 did not change in this checkpoint.

The readiness sentence that names progress hash `0cc66b6d` does not match the authenticated progress file `98da1adaadf536429a493e5f7bd1c1c110f8bc976850539be99f160a854b4210`. That is a provenance mismatch in the handoff text. It is not a product defect.

## Prior findings

Peer1 D1 is closed. `updated_main_field_source` at `crates/rdocx/src/field.rs:4072` no longer refuses because public `to_xml` differs from the canonical writer. It compares each direct body item from `body_item_xml` with the canonical span at `field.rs:4100`, then compares ancestor-qualified paragraph emission with the proved physical range at `field.rs:4178`. The REF projection at `field.rs:3760` reparses `main_part_with_body_spans` canonical bytes, not public `to_xml`.

Peer1 S1 is closed. `root_declares_foreign_word` at `crates/rdocx/src/document.rs:10455` inspects only the bytes before the first `>` and only a double-quoted `xmlns:w`. `main_part_with_body_spans` at `document.rs:25617` always checks canonical spans against `record.blocks`. The public Publication compare skips only a block whose own root start declares a non-Word `w`. One debug-only residual remains: `if let Ok` at `document.rs:25634` skips that public compare for every block when the public writer returns an error. The canonical check still runs, and the current writer emits double quotes, so this does not reopen S1.

## Scope checked without a new finding

The retained-field assertion `fx211_selected_run_xml_replaces_checked_existing_owner_atomically` at `crates/rdocx/src/document.rs:62167` keeps both `default` arms. It checks the expanded field name, instruction, stored result, producer namespace, and producer attribute. That matches the current plan at `.claude/plans/F-X272-design.md:3703`. The old exact-substring failure is reconciled with the approved declaration closure.

`body_item_xml_with_ancestor_declarations` at `crates/rdocx-oxml/src/document.rs:5180` admits every ancestor frame, the reconstructed root, and the item source bytes against one `retained_source_envelope` before the guard copy. The per-frame work is the same linear envelope. Field `retained_root_plan` at `crates/rdocx-oxml/src/text.rs:19316` proves projected declaration growth before `emit_retained_roots` allocates. Run Source raw growth is checked at `text.rs:5142` before the raw children are written.

Public `CT_Drawing::to_xml` and public `CT_R::to_xml` remain Publication. The new Source entries are private. D8 named-union selection, Issue 430 callable ownership, collision, import, and selective cleanup, and the final source, oracle, native, wheel, typing, WASM, archive, hosted, and release riders stay pending. They are not new defects in this checkpoint. The native design-reference script was not reviewed.

The inherited depth-0 lookup in `original_capture_binding` at `crates/rdocx-oxml/src/document.rs:408` collides numerically with a fragment-root run owner. The charge loop only considers prefixes that the event actually uses, and a fragment-root run is also the depth the run parser treats as its own frame. That collision is not, by itself, a false charge of every inherited prefix.

## Defects

### D1. Original-capture admission does not follow the drawing records the parser retains

`admit_original_capture_records` at `crates/rdocx-oxml/src/document.rs:586` is the pre-construction gate. For a typed drawing it charges a used prefix only when `depth == owner.depth` or the prefix is implicit canonical `w` (`document.rs:688`). The direct drawing parser does not retain that set.

`CT_Drawing::from_xml_with_prefixes` at `crates/rdocx-oxml/src/drawing.rs:2123` copies every drawing-scope binding that is absent from the run ancestor list into `wp:inline` or `wp:anchor` raw through `lost_bindings` (`drawing.rs:2164` and `drawing.rs:2194`). Those bytes are the serialized owner (`drawing.rs:1250` and `drawing.rs:1553`). A declaration on `w:drawing` itself is outside the run owner depth, so the preflight does not charge it, while the parser appends it to the retained raw.

After that parse, `CT_R::from_xml_with_prefixes_and_root` records every encoded run-frame binding consumed by that raw (`crates/rdocx-oxml/src/text.rs:4951`) and stores a new root-attribute record (`text.rs:5003`). Document-level `xmlns:x` reaches the run through `word_prefixes_at`. Its capture depth is the document element, not the run, so the typed rule does not charge it. Each run still allocates and later emits its own copy. The same path includes explicit ancestor canonical `w` when the raw uses it. `original_capture_typed_drawing_with_explicit_ancestor_w_is_not_charged` at `document.rs:6231` locks the opposite preflight result and never parses the document, so it does not prove that the parser skips the record. The plan's drawing discovery requires consumed full-frame bindings, including canonical `w` when consumed (`.claude/plans/F-X272-design.md:3584`).

The direct parser selects `inline` and `anchor` by local name only (`drawing.rs:2146` and `drawing.rs:2177`). The preflight also requires the wordprocessingDrawing namespace (`document.rs:648`). The AlternateContent parser does require that namespace and returns no typed drawing otherwise (`drawing.rs:566`). A modeled local-name collision is therefore retained by the direct parser and omitted by the preflight.

The opposite miss is also present. Any child of `w:drawing` that is not an inline or anchor is discarded with `read_to_end_into` (`drawing.rs:2231`). An unmodeled sibling such as `x:opaque` is still charged to the open run because it is outside the modeled namespace set (`document.rs:676`). Repeated owners can cross `retained_owner_work` for records the parser never creates.

Affected consumer: `CT_Document::from_xml` at `document.rs:4935`, which admits the whole part before the body parser builds run and drawing records. Discriminator, under-count: many direct runs whose `wp:inline` or `wp:anchor` uses a long ancestor prefix, or a long prefix declared on `w:drawing`. Admission can return `Ok` and the parse then allocates one copied declaration per run. Discriminator, false refusal: many runs whose `w:drawing` contains an unmodeled sibling outside `wp:inline` and `wp:anchor`. Admission can refuse before a parse that drops those siblings. These inputs were not executed in this pass. The three drawing tests in `fx272-scoped-04.log` cover implicit `w`, no Word use, and explicit ancestor `w` admission only.

### D2. Collapsed-hyperlink reservation omits revision bytes, and the public remover discards the error after mutation

`collapsed_hyperlink_source_total` at `crates/rdocx-oxml/src/text.rs:12173` counts the generated name, declarations, typed attributes, and `hyperlink.extra_xml`. It does not count `revision.write_xml_role`. `remove_runs_core` treats a hyperlink-scoped revision as a reason to collapse (`text.rs:12270`) and then writes that revision into the reserved buffer through `write_hyperlink_boundary` (`text.rs:19002` and `text.rs:12502`). Attribute values are also counted at their raw length (`text.rs:12225`) while `push_attribute` escapes them. The reservation is set to this short total (`text.rs:12289`) even when the allowance is larger.

The early `needed > allowance` return at `text.rs:12284` happens before the placeholder fill. The backstop at `text.rs:12523` runs only after run removal and hyperlink remapping (`text.rs:12483`). For a revision-backed collapsed hyperlink the incomplete total passes, the buffer is constructed, and the backstop then returns `collapsed hyperlink replacement exceeds its reservation`. That violates the current HLD 03 sentence at `docs/hld/03-architecture.md:2812`, which requires the total generated-root and retained-raw growth to be reserved before the first replacement is constructed.

`remove_runs` at `text.rs:12060` discards every `remove_runs_core` error. On the backstop path the public paragraph is already mutated and the empty placeholder is left unfilled, while the caller sees success. On the early path a real allowance refusal becomes a silent no-op of the whole removal. Checked callers that clone a candidate before calling the core, including `Field::set_cached_runs` at `text.rs:1623`, still drop the candidate on `Err`. The public path has no such outer candidate. This input was not executed here. `collapsed_hyperlink_source_growth_is_reserved_before_construction` covers retained raw children, not a revision-only collapsed hyperlink.

## Smells

Zero.

## Nits

Zero.

## Evidence

This pass did not run cargo, tests, lint, or format. Receipts below were read as logs.

| Receipt | Observed result |
|---|---|
| `task15/fx272-scoped-04.log` | `test result: ok. 23 passed, 0 failed, 0 ignored` |
| `task15/fx272-facade-ancestor-02.log` | `fx272_public_page_cache_update_preserves_ancestor_namespace_context` passed, 1 passed |
| `task15/fx272-red-drawing-under-01.log` | typed-selection probe returned `Ok` and failed `unwrap_err` |
| `task15/fx272-red-drawing-over-01.log` | whole-run charge probe failed with `original source capture exceeds admitted source work` |
| `task15/fx272-drawing-green-01.log` | the three drawing admission tests passed. Its dead-field warning predates the current `OriginalCaptureScope`, so `fx272-scoped-04.log` is the later 23-test receipt |
| `task15/capture-clippy-06.log` | 8 lib errors and 12 lib-test errors. No `EXIT` line. Pre-existing lint debt is not waived |
| `task15/capture-rdocx-all-03.log` | 911 lines. Last line is the scoped-replacement test `ok`. No `test result` line and no `EXIT` line. UNKNOWN, not passed |

The author corrected the facade03 claim. That credit stands. `capture-rdocx-all-02.log` remains an older tree and is not evidence for this source. No runtime result in this report was inferred from reading code. The two defects are source-path mismatches with concrete discriminators, not observed test failures.

## Outcome

REJECTED INCREMENT.

Peer1 D1 and S1 are closed. The increment is not accepted because original-capture admission disagrees with the drawing parser in both directions, and collapsed-hyperlink reservation omits revision bytes while the public remover swallows the resulting error after mutation. This does not accept or reject the whole feature. D8, Issue 430, all-pass-12, and the final riders remain pending. No dependency is unblocked here.

Remediation for a later author pass: make `admit_original_capture_records` charge the bindings the direct parser copies into serialized inline, anchor, and run metadata, and stop charging drawing children the parser discards. Include revision emission and escaped attribute growth in the collapsed total before construction. Preserve a public removal that either completes or leaves the prior paragraph installed. Update the ancestor-`w` admission test so it follows the parser. Re-authenticate the tree and return for source peer 3. Do not reset the pass count. Codex final remains after that.

The report-only write lease is returned to the controller. Source and build writes may resume only under the controller's next lease.
