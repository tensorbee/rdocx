# F-X274, all aspects, pass 3

**Reviewed**: Complete F-X274 implementation against immutable original Base `8cb33ef53717159bcf73f655552b99589272a274`, including the previously integrated implementation and current A1 amendment. Current HEAD is `55f5d9d5a45712bda00d7472dfc2ad0855e9d5d1`, branch `sprint/s96`. The approved A1 amendment is `4f0f836878723041d5d80265d19a436692d7866c`. The current working amendment has seven files, 560 additions and 59 deletions. The eleven F-X274 carrier paths against Base have 2381 additions and 256 deletions. The full intervening canonical Base diff inventories 42 files, 2659 additions and 307 deletions, including prior review records, metadata and separately reviewed F-X271 work. Shared Cargo.lock also contains that separate security change. This pass audits F-X274, not the independently reviewed F-X271 security amendment.

**Frozen identity**: Current working diff SHA-256 `3e8bb4f80dd06f038db3b53a5a45422ae7382254ffc143d491f0abbe6059f858`. Current plan SHA-256 `50975d9e5a181078314af147a0d32898136d4788b5a252c749a6b6dd34eb523a`. The227-path logical source mapping SHA-256 is `68d21bd1599a6262d72aee262e1594d2f66ec56e2c773c6ff1bcc63623c09253`. Its manifest file SHA-256 is separately `28ee9784faa0804cf46498b1485c5a3f3a515e6989205ec186761c80aef5eef8`. Independently authenticated all227 source paths with zero mismatches. The root's namespace-wording-only plan clarification accounts for the different historical handover plan/diff hashes. Every other handover changed-file hash matches current source.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. The amended scoped implementation is ready to return to the orchestrator. This is not final native, installed, geometry, hosted or sprint acceptance.

## Defects

Zero.

## Smells

Zero.

## Nitpicks

Zero.

## Contract and invariant review

Read the canonical microscope command, approved annotated F-X274 contract, current progress, actual native recovery facts, affected HLD08/HLD10 and prior full independent passes. F-X274 remains in-progress. The implementing owner acknowledged the A1 amendment before changes, retained the immutable feature Base and handed back the sole writer lease. Review writes are limited to this report. No tests, builds, native applications, product imports, helpers or source/progress changes ran during this pass.

The actual exact-original contradiction is preserved. Word displayed i,ii,1,2 while the prior library displayed i,ii,i,ii. Five native source/display controls distinguish explicit-attribute overlay from whole-last selection. The approved refinement applies that qualified recovery to a private rendering-effective projection. It does not change parser ownership or expand arbitrary malformed-input repair. The first-owner lower parser regression remains unchanged at `crates/rdocx-oxml/src/document.rs:6034`.

- `crates/rdocx-layout/src/engine.rs:1045` captures the typed first start as a compatibility fallback and distinguishes qualified absence from unsupported source. It reads a clone, retaining actual root-binding records while excluding raw child candidates during modeled-owner admission. A foreign/unqualified modeled owner cannot borrow a later qualified duplicate.
- `engine.rs:1081` then reads the complete serialized section only after admission. The depth predicate at `engine.rs:1105` admits direct numbering children, excluding nested prior sections. Element and attribute namespaces are resolved independently. Bound foreign and explicit no-namespace lookalikes, including unprefixed attributes, do not overlay Word facts. Unknown prefixes on consumed numbering reject qualification. Unknown unrelated opaque/nested names are not a new global namespace validator.
- `engine.rs:1128` folds explicit qualified fmt/start independently in source order. Omitted attributes and empty elements leave accumulated values unchanged. The first start continues to come from the typed model, including authored changes. Later starts use the existing u32 conversion and zero behavior at `engine.rs:1143`.
- An unsupported format is sticky at `engine.rs:1137`, so a later supported token cannot falsely rescue it. Chapter composition refuses at `engine.rs:1152`. Invalid/overflow later start, unknown consumed namespace or reader/writer failure discards the partial projection through `engine.rs:1078`, preserving the typed-start fallback with unqualified format. No source or parser state is mutated.
- The existing public doc-hidden format helper at `engine.rs:1029` and private start helper at `engine.rs:12844` consume the same private record. The supported-format predicate is shared with the formatter. No new public helper, model field, enum, trait, generic, module, file or dependency was introduced by the amendment.

## Consumer and lifecycle review

The initial full implementation remains covered by independent pass2 where unchanged, with the amended provider and its complete consumers challenged again.

- Explicit paragraph-ending sections at `engine.rs:4552` and the final/implicit section at `engine.rs:4655` both publish effective start to the existing paginator. `crates/rdocx-layout/src/paginator.rs:888` consumes that start or the previous displayed continuation. Physical page identity remains distinct from displayed numbering.
- Layout source/target formatting aggregates current section facts at `engine.rs:5415`. Ordinary standalone Engine layout obtains source identities for page-family fields, and provenance layout uses the same physical provider. PAGE follows the actual placement owner, PAGEREF the actual target section. Notes retain the existing reference-owner placement rule. NUMPAGES/SECTION/SECTIONPAGES default formatting is unchanged.
- Shared format/refusal maps still reach normal text, the direct multilingual arm, Group and MarkedContent at `engine.rs:9120`. Page classification remains independent of formatter support, with TargetPage cleared only for an unplaced target. The earlier D1 remains closed. Its actual Hebrew/Arabic visual-order input and per-wrapper execution assertions remain present and are included in current lower13.
- Reusable context clones actual section state at `engine.rs:3165` and compares it at `engine.rs:3263`, so duplicate-source changes cannot hide behind unchanged modeled start. Substitution identity includes current value/owner/target/format/instruction facts, and its owned capacity remains charged to the cache. New warm/cold/reopened controls change later explicit start and format without changing first-owner identity.
- Saved cache staging at `crates/rdocx/src/field.rs:16580` still preserves locked/unplaced/dynamic-furniture exclusions and uses the shared section-format provider. PAGEREF uses the effective target page/section supplied by layout. TOC qualification at `field.rs:12076` checks physical target and shared formatted text, rather than treating displayed text as a physical index. INDEX/TOF/TOA label publication at `field.rs:25901` shares the same target formatter. These callers receive the amended private start policy through layout without copied restart logic.
- Parser capture, namespace closure, opaque bytes, first-owner getters/setters and output ordering are unchanged by the amendment. The read-only provider's temporary writer output does not overwrite retained source. Native and lower assertions check first start12, authored first start19, raw numbering order/source bytes and reopen after layout/cache/navigation operations.
- Public Rust/Python/CLI signatures and defaults remain unchanged. HLD08/HLD10 describe effective projection separately from model ownership and preserve unsupported/refusal boundaries. The pending Python cases do not become execution evidence merely because the code is shared.

## Tests and actual evidence

Temporary evidence below is under `/private/tmp/rdocx-s96-run-refresh`.

- `final-word-page-recovery-native-facts.json` records five actual native exports, with input/PDF/tool/receipt hashes and four physical pages. Conflict yields27/28, format-only12/13, start-onlyxxvii/xxviii, emptyxii/xiii and correctly positioned conflict27/28. `final-word-page-recovery-saved-facts.json` records separately completed native save/reopen outputs with ordered singleton attributes agreeing with overlay. Exact original and its independently qualified save/reopen remain separate witnesses. The initially hypothesized whole-last policy is explicitly disproved, not silently adopted.
- `fx274-a1-current-red.json`, copied old executables and `fx274-a1-cli-originals-red-qualified/receipt.json` retain actual current-source RED. The original and four nonempty controls fail the old provider. Empty is an actual positive, not manufactured RED. The earlier unsupported CLI option attempt remains setup-only.
- Current lower tests at `engine.rs:30159` exercise independent attributes, Empty/Start, proper/late order, triple owners, aliases/default namespace/producer shadows, no-namespace/unprefixed positives, foreign/nested exclusions and zero. The refusal matrix at `engine.rs:30309` covers invalid/overflow later starts, consumed unknown element/attribute prefixes, first/later unknown formats and chapter state. Authored start/warm/cold/reopen coverage at `engine.rs:30379` distinguishes later omission from explicit override. The model and serialized source are asserted unchanged.
- `crates/rdocx/tests/regression_test.rs:411002` checks actual target physical2/display27/section1, PAGE/PAGEREF rendered labels, cache counts2/1, unchanged section source, TOC target label and reopen/rebuild. Adding Heading1 corrects the new test's style-based insert_toc setup without weakening target, source or label assertions. The earlier setup failure remains historical.
- `crates/rdocx-cli/tests/integration.rs:7877` checks four overlay combinations and invalid-start refusal through public fields update, saved cache text, numbering source and unchanged input package. The separately executed CLI render probe covers the exact original and five immutable native input files.
- Independently authenticated current named lower13 and native4 logs, source bindings and retained binary hashes in `fx274-a1-frozen-current-named.json`. Lower binary SHA-256 `f69c4acda2786f76dcd7e9ea5d841a1be632d15406d554fffbaf1072d0a7a013`, native binary `5709410107436fd34216ce3c5795d641803052a732297c33051fbeb29f6e548d`. Actual test runtimes are0.03s and0.11s respectively. Compile costs are not reported as numbering performance.
- `fx274-a1-scoped-successor.json` records CLI2, strict three-crate all-target/all-feature clippy, fmt, prose, adapter and whitespace outcomes, with log hashes independently matching. Correct candidate check is separately recorded from the accidentally continued old-source check. The unretained intermediate failing-native binary is not assigned a later executable's hash. Actual current named executions remove reliance on that intermediate artifact.
- Authenticated retained current CLI SHA-256 `d87b6d4cfb45a14259e1dd0a926e925016f53859d7bfdd1914e63e181aa83e05`. In `fx274-a1-cli-originals-green/receipt.json`, all six inputs have convert/extract exits0, four physical pages and the expected native label sequences. Independently checked each current input and generated PDF hash. Poppler26.01.0 is a reader of these library outputs, not a new Word execution or geometry/font acceptance.
- Five new Python nodes at `crates/rdocx-py/tests/test_core.py:15739` inspect rendered/cache/reopened text, held-handle success/refusal lifetime and both numbering owners' bytes. They are syntax-only so far, pending final fresh installed bindings. The old wheel and old Python51+1 observations do not qualify the changed provider. A rebuild and new installed execution remain required.

The test matrix rejects first-only selection, whole-last replacement, local-name-only attribute matching, partial recovery after invalid start, unqualified modeled-owner fallthrough and stale duplicate facts. Actual local assertions were inspected, not inferred from aggregate counts.

## A-ID disposition and pending final gates

| ID | Current scope and remaining obligation |
|---|---|
| A1 | Exact unchanged original and five controls now agree numerically between actual recorded Word and current CLI. Native saved-source facts qualify the overlay refinement. Complete native/library geometry, fonts, installed and output-source acceptance remain due. |
| A2 | Prior five-format/variant/source-owner controls apply through the shared provider. Current overlay controls add independent restart/format and qualified target coverage. Full native/installed variants remain pending. |
| A3 | Shared staged formatter and supported/unsupported explicit-switch rules are unchanged. Current refusal matrix retains stored display even with explicit Arabic. Final switch parity remains due. |
| A4 | Existing Roman/repeated-letter resource limits and rollover semantics remain unchanged. Later start uses u32 without widening/clamping. Native bounds and final installed union remain due. |
| A5 | Current alias/default/shadow/Empty/Start/foreign/nested/unknown/refusal controls supplement unchanged lower source-capture evidence. First-owner model and opaque source authority remain intact. Final native package/source qualification remains due. |
| A6 | Current native test checks eligible body and target caches with unchanged source. Prior dynamic physical footer policy is unchanged. Final fresh installed cache/render/save/reopen coverage remains due. |
| A7 | Current authored first start, later explicit/omitted attributes, warm/cold/reopen and pending Python held-handle cases retain lifecycle coverage. New Python execution and full installed/runtime/typing union remain pending. |
| A8 | Current TOC target27 is checked across cache/rebuild/reopen. Complete F-X272/F-X273/F-X274 composition with redline/furniture/geometry/native/PDF/PNG/font/hash evidence remains pending. |

## Not found and handback

**Correctness**: zero findings after provider, fallback, publisher, target and cache tracing.

**Contract**: zero findings. Native evidence triggered an explicit approved refinement, preserving A1 and first-owner authoring rather than waiving the original or inventing last-wins recovery.

**Panics**: zero findings. Qualified lookups, checked depth decrement, u32 parsing and existing numeric resource bounds remain intact. No new unsafe indexing or unchecked untrusted conversion was introduced.

**OOXML**: zero findings. Namespace identity, Empty/Start, opaque duplicates, nested prior exclusion and raw source authority were checked. The projection changes interpretation only, with no schema-order or general serializer mutation.

**Tests**: zero findings. D1 remains closed. Unsupported setup attempts and executable provenance limits are explicit. Pending Python execution is not relabelled green.

**Structure**: zero findings. One private concrete effective-numbering record and shared predicate serve existing consumers, without speculative public structure.

All final native/full/workspace, geometry/exact-font, fresh installed3.9/3.12, strict typing/stubs, WASM, hash49/pixel7 with reviewed deltas, README/current measurement, archive/size/supply-chain, hosted CI and integrated risk obligations remain pending. Previous full or hosted observations must be reassessed for applicability after the actual engine source change. No final acceptance, baseline approval or unrelated security self-review is implied.

This report ends the review pass. No remediation occurred. Return sole tracked-write control to the orchestrator.
