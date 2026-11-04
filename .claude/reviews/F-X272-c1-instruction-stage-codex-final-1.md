# F-X272 C1 shared instruction stage, Codex targeted final review 1

**Reviewed**: the stopped author's shared instruction prerequisite on `sprint/s96` at unchanged HEAD `91dfb0a7c6b48c054fc4a28df6bca7994f717fc9`, with dirty carry preserved. Scope is `SEMANTIC_SOURCE_STAGE_ONLY.patch`, the1705-line full `STAGE_ONLY.patch`, current corresponding source, HLD03, plan/progress additions,12-file formatting manifest and named scoped evidence. This is a targeted Root final audit after the existing implementation peers, not a new peer or reset.

**Verdict**: 0 new defects, 0 new smells, 0 nitpicks in this prerequisite. The stage is suitable to retain as the shared grammar/scalar foundation. Full C1 qualification is still incomplete and blocked by the explicitly preserved raw-fallback failure. This review grants no feature or provider acceptance.

## Defects

None new in the reviewed stage.

## Smells

None new in the reviewed stage.

## Nitpicks

None.

## Authentication and actual scope

The author handoff matches SHA-256 `845985dc90326727a9cdbb6e4c413260d8bd3177b506cfa1218177d13de26d44`. Final manifest matches `704c9508503d22c306ece465bfd719824543d1a3a187d547bc7888d4a83e5592`. Independently rehashed all57 manifest entries with zero missing files or mismatches. The12 formatting-only files match both their preserved before hashes and current after hashes. The semantic patch factors formatting out of text.rs and document.rs. Its leading removed snapshot-path annotations are comparison-artifact headers, not product source lines.

Read the complete shared-syntax/scalar changes and the full stage patch's semantic, formatting, documentation and lint changes. The additional carry edits are rustfmt presentation changes, with their separate manifest. The document.rs guarded-pop to pop_if changes retain the same depth predicate. Test literal simplifications and the local reused tuple alias do not change behavior. The existing eight-argument writer receives a locally explained Clippy allowance, without suppressing a validation or resource gate.

This is not a fresh audit of all dirty F272 carry or another story. F278 is outside scope. No source, test, plan, progress or metadata was modified by this reviewer.

## Shared grammar and admission ordering

`crates/rdocx-oxml/src/text.rs:20239` introduces a concrete `InstructionSyntax` carrying text or unique nested occurrence indices. `parse_instruction_syntax` at `:20297` moves actual child fields into a separate pool and uses the same lexer and classifier as ordinary instruction reads at `:20187`. No nested Field is cloned during token classification. The remaining operand clones copy text or indices, not recursively owned field payloads.

Complex close at `:8627` forms syntax, computes the existing source, runs `admit_word_xml` and `admit_initial_field_capture`, and only then materializes `FieldInstruction` at `:8653`. The nested-child model and normalization checks at `:1353` still inspect the actual populated fields in the pool. They are not replaced with an empty-cache assertion or an unproved XML dominance claim.

Materialization takes each selected nested occurrence once. Argument/switch grouping preserves the separate physical `nested_order`. Tokens are consumed once by the unchanged classifier decisions, so a selected occurrence cannot be both an argument and a switch operand. The first-token nested case still produces an empty name and drops that occurrence, matching the earlier classifier. Unselected pool entries are not charged as retained operands. Internal indexing and expect calls rely on this private constructor invariant, with no caller-supplied syntax object or new public panic surface.

The parsed child cache publication following Field construction is unchanged. Moving the child preserves its physical runs and other stored payloads. The new occurrence test retains and checks populated parsed runs and source identity. This is preservation evidence for moving existing fields, not a completed source-only model-summary proof.

## Exact scalar behavior

The shared arithmetic at `text.rs:1267` through `:1324` is used by both the existing model calculation and syntax calculation. Checked addition behavior is retained. It includes trimmed raw bytes, actual normalized name bytes, cached-result bytes, exact scope bytes,64per argument plus decoded text bytes, and64per switch plus normalized name and optional decoded text bytes.

The classifier retains existing Unicode uppercase name conversion, ASCII lowercase switch conversion, raw concatenation excluding nested operands and final trim. Lexer quote decoding, escaped quote handling, flags and switch-argument lookahead are unchanged. Existing authoring-token and bibliography consumers were updated to the new internal operand representation without changing their tests or rules.

The explicit new controls at `text.rs:23292` pin79bytes for `REF "a b"`, raw trim, the byte-changing Unicode uppercase case, decoded quoted text, cached/scope increments, switch flags/operands and arithmetic overflow. These expected values are independent of the shared helper equality checks. The occurrence control checks no eager source clone, argument/switch order, unique source identities, stored results and populated cached runs. Existing nested-switch grammar coverage remains.

No new trait, generic, source file, module, public model field, signature, budget or cutoff is introduced. The concrete carrier separates existing grammar from model publication and is directly used by two real construction routes.

## Recorded execution evidence

Inspected actual logs and adjacent receipts, including commands, semantic outcomes and output hashes. The deciding52-field and19-instruction subsets each record child0 and wrapper0. Both new syntax controls occur in the final instruction subset. Counts overlap and are not added into a unique-test total.

The final original-capture batch records child101 and wrapper101, with33passed and the same one failing `original_capture_keeps_supported_raw_control_fallbacks_over_the_typed_budget` assertion. Its constructor chooses raw, while admission still refuses. That failure is retained honestly and remains blocking outside this prerequisite.

Changed-crate all-target/all-feature Clippy with warnings denied records child0 and wrapper0 after the disclosed carry-lint corrections. Final fmt after the separately recorded carry formatting records0. Prose records0violations. Earlier nonexistent CT_R method and Rust literal escaping failures are setup errors, not semantic RED. New equivalence controls are coverage, not an invented old-source semantic failure. Semantic checks precede only the documented formatting-only carry transformation.

No build, test, benchmark, oracle, native app, network, Git mutation, cleanup or delegation was executed by this reviewer.

## Retained blockers and final boundaries

The stage does not implement exact owner/context constructor qualification before model and expanded-record construction. Raw fallback propagation, opaque/duplicate/ignored slot dispatch, namespace aliases and shadows, populated-cache weighted-owner/table-depth summaries, affine normalization summaries, single bounded qualification scheduling, all lexical/field routes and composed operation work remain unclosed.

Q-C1-01 exact fallback acceptance, Q-C1-02's narrow direct legacy-read exception, factor124, declaration pooling, duplicate charging and unused-URI/zero-growth controls remain unchanged. V11 remains rejected. No unproved257 example or empty-cache argument is accepted here.

Issue430 implementation, retained-reference native evidence, provider readiness, full C1 acceptance, independent remaining final review and integrated/native/schema/hash/installed/packaging/CI/release gates remain due. The existing fallback RED forbids completion. The reviewer stops after this report and returns the canonical writer lease to Root.
