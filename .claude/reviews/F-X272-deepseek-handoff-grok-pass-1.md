# F-X272, DeepSeek handoff, Grok pass 1

**Author**: DeepSeek V4.1 Flash through OpenCode, sole source owner for this increment.
**Reviewer**: Grok 4.7, actually invoked. Pass 1 of at most 3. Codex remains assigner and final reviewer. This file is not that final pass, not integration, and not feature completion.
**Effort**: Read-only audit of the frozen increment, the approved plan, the prior Codex task15 review, progress, and the task15 receipts. No product build, test, format, git mutation, or Word session was run by this review.
**Verdict**: 1 defect, 1 smell, 0 nitpicks. The increment is not accepted.

## Identity

Canonical tree `sprint/s96` at `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`, 28 dirty entries. Current files match the frozen AFTER copies and the declared hashes:

| Path | SHA-256 |
|---|---|
| `crates/rdocx-oxml/src/document.rs` | `b3f10c58113d53df265bee97510770f323e4271629a1cd75542774d7cccafde7` |
| `crates/rdocx-oxml/src/text.rs` | `257b0adc9d3c91a6859359df87e70f2d9ba5b8e216686e2c60243e71bac9d207` |
| `crates/rdocx/src/document.rs` | `b5746b38361a06beb246835e9b6685e9bfe4e758b36890737e397b0c3879839b` |
| `crates/rdocx/src/field.rs` | `bc59571e6a813642c7ee1c100fc22947bc4a22b72fb521cfdc7f2c6727e13081` |
| `.claude/plans/F-X272-design.md` | `50af80476256f68ad7c6fe27ebc0ec800d2e387aeb4d2267953ce515de067f77` |
| `docs/hld/03-architecture.md` | `832e6fd5c77ae1181df5df316d31108e98f7af00a0e9820de3c18ed3cfe9e027` |
| `docs/hld/10-bindings-spec.md` | `c09cf197a66315737897aa597361ff2d32dfa1a109d8db67930924e985410ade` |
| `.claude/scratch/F-X272-progress.md` | `97ee971f3f61526761ed07d398c808785d1aa9bf7346175d1fec4ad681e4a196` |

BEFORE copies match the declared starting hashes `8be92bb9`, `dad09b64`, and `78c0e188`. `field.rs`, the plan, and both HLD files are unchanged in this increment. Unified diffs against BEFORE are 11 lines for `text.rs`, 176 lines for oxml `document.rs`, and 222 lines for the facade `document.rs`.

## Scope

This pass reviews the incremental remediation of DEEP-FRESH-01, DEEP-FRESH-02, and DEEP-FRESH-03, including retained consumers of the new Source and Publication split. It is a partial handoff. It does not close initial aggregate capture admission, TOTAL source-growth preconstruction, the final named union, Issue430 contract or native retention, or an all-pass-12 review.

Examined product paths: `retain_foreign_word_alias_declaration`, `generated_word_start`, `word_prefixes_at`, `CT_Body::to_xml`, `write_marked`, `write_body_content`, `CT_Document::write_part`, `body_item_xml_with_ancestor_declarations`, `standalone_body_item_xml`, `CT_P::to_xml`, `CT_P::write_xml_core`, `MainPartBlocks::serialize` and `body_spans`, `updated_main_field_source`, and the two new facade tests. Equivalent body publishers were searched for `source_context.extend` and `frame.iter().cloned()`. The only remaining `root_context` hits are unrelated math layout locals.

## Defects

### D1, field-cache association treats the required writer split as a mutation

`crates/rdocx/src/field.rs:4079`

`updated_main_field_source` loads canonical bytes from `main_part_with_body_spans` and then rejects the call when `CT_Document::to_xml` differs. Any nonempty cache update reaches that check from `apply_cached_field_updates` at `crates/rdocx/src/field.rs:4004`.

This increment makes those two byte strings differ for a direct body paragraph whose root binds `w` to a non-Word URI. Public `write_body_content` now always calls `CT_P::to_xml`, which is Publication (`crates/rdocx-oxml/src/document.rs:4440` and `crates/rdocx-oxml/src/text.rs:13521`). The facade assembler still calls `body_item_xml`, which writes that same paragraph with `XmlWriteRole::Source` (`crates/rdocx-oxml/src/document.rs:4803`). The approved plan at `.claude/plans/F-X272-design.md:3396` and `:3420` requires that split: public and reentrant calls stay Publication, and Source stays an explicit private argument.

The new test records the divergence. `fx211_public_foreign_w_serialization_stays_publication` keeps `xmlns:w="urn:producer"` on the canonical part and requires public `to_xml` to rewrite it to `xmlns:rdocxProducer` (`crates/rdocx/src/document.rs:61031` and `:61046`). Receipt `fresh03e-fx211_public_foreign_w_serialization_stays_publication.log` reports that test passed. Those are the two strings the field-cache guard compares.

Before this increment, `write_body_content` used Source for that paragraph, so public `to_xml` and `body_item_xml` agreed. The guard was a real change detector. After the split it fires on an unchanged document. The error text says the writer changed during association.

Discriminator, source-proven, runtime not in the supplied receipts: the foreign-`w` PAGE paragraph from the new publication test, then `update_fields` with a resolved PAGE result. Expected result: the cache update runs against the canonical Source part, the canonical binding stays `xmlns:w="urn:producer"`, and a later public `to_xml` still rewrites to `rdocxProducer`. Actual control flow: `Err` with `current field-cache writer changed during source association` before any field byte is patched. The same refusal happens for a header field update when the main part merely contains one direct foreign-`w` paragraph, because the main-document compare runs before the story loop.

The later per-item compare at `crates/rdocx/src/field.rs:4105` is consistent, because both sides use `body_item_xml`. The early whole-document compare is the fault. Repairing it by sending the facade path back through Publication would undo DEEP-FRESH-03. The REF projection at `crates/rdocx/src/field.rs:3760` reparses `to_xml()` into the model. It is currently unreachable when D1 fires. A narrowed guard must not let that reparse replace Source bytes with Publication bytes.

The new tests never call `update_fields`, so their green receipts do not cover this consumer.

## Smells

### S1, the debug span oracle is skipped by a whole-part byte search

`crates/rdocx/src/document.rs:10456`

`declares_foreign_word_namespace` scans the canonical part for `xmlns:w="` and skips the debug compare at `crates/rdocx/src/document.rs:25624` when any value is not the Word URI, or when a quote is missing. The Source branch is narrower: only a direct body paragraph whose root attribute record binds `w` to a non-Word URI (`crates/rdocx-oxml/src/document.rs:4452` and `:4803`).

A part that contains the needle in text, raw XML, or a nested owner still disables the oracle for every sibling, including children whose Source and Publication bytes are required to match. The returned spans themselves are the head plus block lengths at `crates/rdocx/src/document.rs:9138`, so this is a lost independent check, not a proven bad offset. The new publication test never asserts a span.

The alias test at `crates/rdocx/src/document.rs:60943` likewise locks `w:flag="keep"`, text `7`, and the absence of `xmlns:=`. It does not lock the selected nonempty Word alias, the `rdocxW` fallback, or expanded names. Reverting only the alias choice would leave that test green.

## Nitpicks

None.

## Not found

Panics. The changed product lines add no `unwrap`, `expect`, or unchecked index. The `?` in the `body_item_xml_with_ancestor_declarations` match guard returns before the explicit pop. `RootBindingScope` restores on `Drop` at `crates/rdocx-oxml/src/text.rs:547`, and the `Result` is propagated only after that pop at `crates/rdocx-oxml/src/document.rs:4843`.

OOXML child order. The increment changes namespace attributes and writer role. It does not reorder schema children.

Structure. No new trait, generic, crate, module, or public signature.

Escaping. The changed writers still push attributes through quick-xml. No new hand-built element text was added.

Callback timing. `write_marked` still marks after the body start and after each child at `crates/rdocx-oxml/src/document.rs:4406`. Removing the unused source argument does not move the callback.

DEEP-FRESH-01 at the cited builder. `retain_foreign_word_alias_declaration` no longer selects the empty plain prefix that `word_prefixes_at` uses for a default Word binding (`crates/rdocx-oxml/src/text.rs:796`, `crates/rdocx-oxml/src/numbering.rs:611`). An empty prefix returns the record unchanged, so the function does not emit `xmlns:`. A later nonempty Word alias is still eligible. `generated_word_start` already skipped empty prefixes and falls back to `rdocxW` at `crates/rdocx-oxml/src/text.rs:595`. Receipt `fresh01-red-01.log` failed on `default-before-q` and `default-and-nonempty` with `source invalid xmlns:` before the repair. `fresh01-green-01.log` and `fresh03e-fx211_default_word_foreign_w_alias_declaration_is_valid.log` report the four-input test passed. This review did not re-execute it.

DEEP-FRESH-02 ancestor clone. `frame.iter().cloned()` is gone. The item path borrows ancestor and root strings for the `write_xml_core` call at `crates/rdocx-oxml/src/document.rs:4816`. A workspace search found no remaining inherited-frame clone of that shape. One owned `raw_namespace_declarations` of the part root remains inside the per-item call. That is the physical root decode the prior finding kept distinct from cloning ancestor frames. `body_frame_xml` and `standalone_body_item_xml` do not clone those frames. `let _ = work` at `crates/rdocx-oxml/src/document.rs:4785` still discards the ancestor census. That is the already pending TOTAL admission, not a new copy.

DEEP-FRESH-03 on the public body writer itself. `CT_Body::to_xml`, `write_part`, and `write_body_content` no longer pass Source. `CT_P::to_xml` passes Publication and `None`. Direct foreign-`w` paragraphs take Source only through `body_item_xml_with_ancestor_declarations`. Nested paragraphs inside a table or control still follow that owner's public writer, which was already true before this diff.

## Evidence

Recorded receipts were read and not re-run. They are tied to this review by the implementer's logs and by the current source hashes, not by a hash inside each log.

- DEEP-FRESH-01 RED then GREEN: `fresh01-red-01.log` (exit failure, 0.19s, the two default cases) and `fresh01-green-01.log` (1 passed, 877 filtered, 0.18s). Later `fresh03e` rerun reports 1 passed, 878 filtered, 0.14s.
- Publication split: `fresh03e-fx211_public_foreign_w_serialization_stays_publication.log`, 1 passed, 0.04s.
- Carrier: `fresh03e-fx211_note_and_field_profiles_reuse_existing_qualified_providers.log`, 1 passed, 878 filtered, 2.79s.
- `fresh03-fx211-all-02.log`: 28 passed, 1 failed, 3 ignored, 847 filtered, 4.23s. The failure is `fx211_selected_run_xml_replaces_checked_existing_owner_atomically` at `crates/rdocx/src/document.rs:62167`.
- `fresh03-before-runxml-01.log` panics with the same source string. Both actual snapshots are a canonical `w:p` whose `q:fldSimple` keeps the original instruction, `xmlns:x="urn:field"`, `x:key="keep"`, and an inserted Word `xmlns:q`. The `default=true` arm is not reached. This review did not rebuild the BEFORE tree. The matching panic text supports the no-regression claim for this assertion. It does not make the red test accepted.
- Assertion meaning: the test requires the imported field start tag as a substring (`crates/rdocx/src/document.rs:62118` and `:62167`). The writer keeps the instruction, stored text, and both `x` bindings, and inserts the Word URI on `q:fldSimple` after the paragraph becomes canonical `w:p`. That is declaration closure, and it fails the exact-byte rule the test states. No plan sentence in this increment authorizes deleting or weakening that assertion. It stays open debt. It is not a DEEP-FRESH regression.
- Scoped `cargo check` receipts `fresh03-check-rdocx-02.log` and the oxml check logs finish with the two existing `table.rs:304` and `table.rs:537` warnings. `fresh03-clippy-01.log` finishes with warnings outside the touched lines. This review did not re-run them. No full workspace, oracle, hash harness, or fmt gate was claimed or reviewed as passing.
- D1's `update_fields` discriminator was not executed. The byte split it depends on was executed by the publication test.
- Native claim in the handoff is stale. Progress at `.claude/scratch/F-X272-progress.md:4864` says native control fails during pipe startup before any application action. The same progress file at `:4476` already says Word 16.114.1 (26102031) launched and that the old pipe failure must not be repeated as a fresh observation. On disk, `/private/tmp/rdocx-s96-word-story-diagnostic-only/rdocx-q04-bookmark-probe-20261031T155151Z-omittedtype.word.docx` is 14064 bytes, mtime 2026-10-31 17:14:51, beside a 2563-byte work input from 15:51:51 and a Word lock file. `explicitnormal` has input and work copies and no `.word.docx`. This review did not open Word and did not read the package XML. The saved file shows an application save occurred. It is not structural retention evidence and it is not Issue430 acceptance. Historical `/private/tmp/rdocx-s96-run-refresh` was not recreated.

## Outcome for this increment

Not accepted. DEEP-FRESH-01 is repaired at the cited alias builder, and the ancestor URI clone of DEEP-FRESH-02 is removed. DEEP-FRESH-03 is implemented on the public body writer and broken again for field-cache association by D1. S1 leaves the span oracle too wide to trust as coverage. The pre-existing `fx211_selected_run_xml` failure remains red and is not waived. Pending aggregate admission, TOTAL growth, leaf and drawing consumers, D8's named union, Issue430, and pass 12 stay open. A later clean peer pass still cannot close them.

## Next remediation

Do not start it in this pass. The source lease returns to the implementer.

1. Change `updated_main_field_source` so an intentional Source versus Publication difference is not a mutation error. Keep public `to_xml` on Publication. Associate cache edits with the canonical part already checked against `body_item_xml` at `crates/rdocx/src/field.rs:4105`. Add a red test for `update_fields` on the foreign-`w` PAGE paragraph, and a second case whose only update is in a header while the main part has a direct foreign-`w` paragraph. Assert canonical `xmlns:w="urn:producer"`, public `rdocxProducer`, and a successful cache write.
2. Keep `crates/rdocx/src/field.rs:3760` from reparsing Publication bytes into a Source projection when the guard is narrowed.
3. Replace the whole-part needle at `crates/rdocx/src/document.rs:10456` with a check that still compares children which did not take the direct Source branch. Assert the foreign-`w` canonical spans against the assembled head, blocks, and tail.
4. Extend the alias test so default-before-q and default-and-nonempty keep expanded Word names, the producer attribute, and the existing nonempty alias, with `rdocxW` only when no nonempty Word alias exists.
5. Reconcile `fx211_selected_run_xml_replaces_checked_existing_owner_atomically` without weakening it. The q-prefixed case is red in BEFORE and AFTER with the same inserted `xmlns:q`. Make the imported field substring survive, including the `default=true` arm, or return a plan amendment that states the closure rule and replaces the substring check with expanded-name checks for the instruction, stored text, and both producer attributes.
6. Drop the pipe-startup sentence from the next handoff. Structural retention of the omitted-type `.word.docx` is still unread. Do not treat that file as Issue430 acceptance, and do not comment on the issue.
7. Leave the named pending story unimplemented until this defect and smell are cleared and the red run-xml expectation is reconciled.
