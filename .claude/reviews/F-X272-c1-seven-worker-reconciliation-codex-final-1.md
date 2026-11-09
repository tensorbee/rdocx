# F-X272 seven-worker source reconciliation

## Result and boundary

New source defects: 0. Smells: 0. Nits: 1.

Bounded read-only review of the recovered worker at canonical baseline1d2c5a803d346da4a7b65b1bf8ac7e590ad1492c plus its unfinished F272 carry. This is an integration source inspection, not whole C1 acceptance, a new implementation peer pass or execution evidence. The known F279a wrapper-policy interaction remains unresolved and blocks claiming final coherence.

## N1, stale chart replacement descriptions

`/private/tmp/rdocx-s96-f272-codex-recovery/crates/rdocx/src/document.rs:46211` still says replace_in_xml_parts handles chart parts, and the caller comment at45752 also lists charts. The joined function correctly omits chart traversal under F277c. Update those descriptions to match the retained exclusion. No code change is requested by this nit.

## Required known F279a interaction

The current joined source retains F279a's clean-Base policy at `crates/rdocx/src/revision.rs:1418`: transparent raw wrappers require canonical_w. Its existing regression at4054 expects `directincluded`, deliberately excluding an alias-only smartTag payload. C1's actual shared reader at `crates/rdocx-oxml/src/text.rs:13647` accepts actual qualified smartTag/customXml/fldSimple without that canonical-w prerequisite. Thus retaining canonical F279a source alone does not resolve the already-declared coherence obligation in F-X279a-design.md:15.

This is the previously disclosed mandatory interaction, not a newly discovered incremental regression or a demand to invent a new public policy. Root must reconcile the snapshot selector and its old gate-specific test against the completed actual C1 accepted-text behavior, preserving the real wrapper/run/control boundaries and unsupported/foreign outcomes. Do not merely change the expected string without asserting agreement through the actual native owner projection.

## Confirmed source joins

The joined replace_in_xml_parts at `crates/rdocx/src/document.rs:46216` retains Result<usize>, propagated raw text replacement failure, zero-count changed-copy handling, propagated document reparse and propagated note rewrite. The caller at45755 uses question-mark propagation. F277c chart import/loop is absent. This is the required source-level composition, not proof that every staging/refusal case passes.

The F280c shared XML helper is byte-identical to the canonical helper. Its heap-scoped namespace implementation is retained, not overwritten by the old carried source. Likewise F279a revision.rs is canonical before the pending policy reconciliation. No separate lexical validator or namespace reinterpretation was added during the join.

## Earned focused controls for Root

- Existing native `issue432_replacement_skips_chart_cache_and_workbook` and `issue432_literal_and_regex_keep_text_box_story_counts`, plus CLI `issue432_cli_text_and_expected_replacement_agree_without_chart_mutation`. Preserve exact chart/workbook bytes and replacement counts while exercising the final fallible C1 source.
- A failure case with a successful earlier typed or raw-text-box replacement followed by an actual later resource/reparse refusal. Assert public Document bytes, notes and chart companions remain unchanged. A helper-only error assertion cannot establish public atomicity.
- Existing F279a `issue434_snapshot_raw_wrapper_requires_base_canonical_writer_binding`, converted into an actual C1 coherence discriminator with alias-only, foreign-w plus Word alias, explicit canonical-w, foreign lookalike and empty wrapper cases. Compare native paragraph/current property-owner accepted text and owned revision snapshot through save/reopen. Retain stored-cache/instruction-poison, control-domain and hidden/deleted ancestor controls.
- Existing oxml-core `strict_xml_deep_scopes_preserve_aliases_undeclarations_and_declaration_limit` and representative actual C1 alias/shadow/escaped-URI cases. The already executed canonical F280ac18 and F2812 checks qualify their pinned canonical source only, not this dirty F272 worker.

No build, runtime test or source remediation ran in this review. Root remains sole source writer.

## Observed source pins

- `crates/rdocx/src/document.rs`: `ef92ea7aa9c5878952f4a1d4aee617f92072a052f8ca3f5375b2581684126143`
- `crates/rdocx/src/revision.rs`: `394a6eccc132c8eacc202a192dd3b768a6a80e9e575ee7f7538a16284887b11c`
- `crates/rdocx-oxml/src/text.rs`: `6d7d790581b2a2741db28a374241ecb991beb9cf7a912c3ca038555db223b416`
- `crates/oxml-core/src/xml.rs`: `5e71a6d5b2e2e4e9c3ee833c2c775a64120c16a6188bf94065b70d7c419d85d5`

## Root actual-consumer correction and observed coherence defect

The helper-only comparison above is insufficient to predict native accepted
text. Actual boundary_owners at the sampled worker text.rs:13574 first filters
raw bytes through is_run_wrapper at13591 using implicit w plus local raw-root
declarations. A document-only Word alias is therefore absent after detachment.
This corrects the earlier reasoning without declaring the coherence debt closed.

Root ran a native library diagnostic against the recovered combined source,
using actual CT_Document import, paragraph.accepted_text, to_xml/reopen and
story_part_revisions. The inherited-only a wrapper initially agrees as direct.
After save introduces canonical w on the document root, the source-based
snapshot gate at crates/rdocx/src/revision.rs:1418 returns directalias while
native still returns direct. An a:customXml declaring only canonical w locally,
with a inherited from the document, yields snapshot directcanonical and native
direct before and after save. Foreign lookalike and empty controls agree.
These are library observations, never Microsoft Word native-open evidence.

D1, existing mandatory F279a coherence debt is now concretely reproduced. The
source-ancestor canonical_w proxy at revision.rs:1236 and1418 does not reproduce
the detached raw-reader namespace role at text.rs:13591 and13879. Fix must
preserve actual typed Field source prefixes separately, raw-wrapper detachment,
local declarations, descendants and qualified marker/header selection. Removing
the gate globally or simply changing expected strings would not repair this
class. Preserve resolver/source inventory identity and all hidden/control/cache
boundaries. The snapshot role repair is not implemented.

The diagnostic initially failed to compile due to Root's nonexistent/unimported
StoryKind::Main and then disproved Root's assumed directalias native expectation.
Those are setup costs, not product RED. A subsequent compiled diagnostic collected
the actual three disagreements and exited101. Its source is preserved at
/private/tmp/rdocx-s96-reconciliation-wrapper-facts/ROOT_DIAGNOSTIC_revision.rs.
Worker revision.rs was restored byte-identically to HEAD before the mapped
batch writer took over. Original source-pinned observations remain historical.
No accepted C1/provider or final sprint status follows.
