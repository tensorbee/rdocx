# F-X272, namespace investigation, pass 3

**Role**: independent Codex design escalation for unchanged-owner identity. Read-only outside this report.
**Reviewed**: dirty `work/f-x272-codex` at Base `c8acc502b74f29ed7e9546939f17007709a9f22b`, 13 tracked files, 2563 added and 412 removed lines. Tracked binary diff SHA-256 `32426d4f6cf801acd0941f75131a13fe9d4739898fdb495ea31e154da6818d33` remained unchanged during this investigation. Canonical plan clarification `10d45b11` and the annotated worker plan/progress were read.
**Verdict**: 1 defect, 0 smells, 0 nitpicks in this focused pass. N3 is a distinct metadata-identity defect. This is not the full feature microscope or final acceptance.

## Actual evidence

The authenticated fresh dev wheel observation is retained in `/private/tmp/rdocx-s96-run-refresh/fx272-fresh-python-final.json`. Build and installation succeeded. Wheel SHA-256 is `592cc4ffe90f8778aca86958fc1daa7c9947fb22771e0d51ed7d58b95043cda6`, and both wheel native member and actual imported module hash to `f75d353a032792b5ddf9862c5fe069671480312f9ac3800a23a4d62f5d1e90ee`. Its source manifest stayed unchanged before/after build and execution. The initial cache-denied installation is a separate setup failure, not the later semantic outcome.

`fx272-python-green.log` records four selected nodes actually executed: three passed, one failed, 643 deselected, 3.84 seconds. `test_issue425_compare_keeps_revisions_beside_a_new_comment` adds a comment to body paragraph1 at `crates/rdocx-py/tests/test_core.py:2451`, then comparison refuses paragraph0 at `test_core.py:2460`. The failure is `comparison cannot change body/paragraph[0], which holds an existing tracked revision: the edited document changes it`. This is actual semantic RED, not a stale binary, cache warning or fixture setup failure.

The captured `fx272-comment-before.xml`, `fx272-comment-after.xml` and `fx272-comment-trace.log` show the only paragraph0 difference: redundant local mc/r/w/wp declarations disappear, while the same bindings remain inherited from the unchanged document root. The comment markers are added to paragraph1 as requested. Independent read-only XML declaration replay in this pass confirmed identical raw hashes for the two old revision subtrees: deletion `b1fcd47ae3578abce89d991e2221b2f5d652df1108c8fd6fe9bc2af3b6ac6afe` and insertion `886be4d4ed1fc6eb7796befe890bf8e7d5629098929b6b3bfd5405402215df55`. The complete encoded context entries are equal as an unordered collection, but their order differs. No product tests or builds were run by this investigator.

## Defects

### N3, unchanged-owner equality treats parser-context enumeration order as a content change

`crates/rdocx/src/comparison.rs:3808`
`crates/rdocx/src/comparison.rs:3826`
`crates/rdocx-oxml/src/revision.rs:42`
`crates/rdocx-oxml/src/revision.rs:53`
`crates/rdocx-oxml/src/revision.rs:130`
`crates/rdocx-oxml/src/revision.rs:162`

Comparison's exact unchanged-owner route uses complete `BodyContent` equality. That derives through `CT_P` and its revisions at `rdocx-oxml/src/document.rs:734`, `rdocx-oxml/src/text.rs:6975` and `rdocx-oxml/src/text.rs:6992`. `CT_Revision` derives `PartialEq`, including its private `parse_word_prefixes: Vec<String>`. Both empty and paired revision constructors store the supplied context vector verbatim.

`word_prefixes_at` starts with inherited entries, removes the old records for every locally declared prefix and appends that prefix's effective records at `rdocx-oxml/src/numbering.rs:607` and `numbering.rs:621`. Repeating the SAME ancestor binding locally therefore changes enumeration order without changing namespace meaning. In the actual captured input, the encoded entry order before comment insertion is mc,r,w,plain-Word-w,wp. After insertion it is w,plain-Word-w,r,mc,wp. All prefix/URI entries and both captured revision subtrees are unchanged.

The derived vector comparison makes the paragraph values unequal solely for this bookkeeping order, so the bounded revised-owner guard refuses an unchanged owner. A1/A2/A5 require comparison to retain the old redline while tracking an unrelated new comment, and A6 requires the publication boundary to distinguish actual owner changes. This case exercises that approved supported path. It is not permission to edit the revised paragraph and is not a defective comment-range fixture.

N3 is separate from N1. N1 concerned losing or rebinding namespace meaning when bytes were transported. Here the effective namespace context survives and only its representation order differs. The prior 22 native successes remain facts about their actual candidate and inputs, not proof of this new Python case or a clean feature verdict.

## Smallest coherent repair

Use a narrow equality-only policy for `CT_Revision`'s private parser context in the existing `rdocx-oxml/src/revision.rs`. Keep original parser vectors and their order intact. Compare the complete context entries without depending on enumeration order, while preserving the existing equality of EVERY other field: kind, id, author, timestamp, raw XML, typed content, content paragraph, nested revisions and require-author policy. Replacing the derived implementation with a local explicit implementation of the existing PartialEq trait is a concrete option. No new trait definition, public signature, source file, module, crate or generic XML comparator is needed.

An order-insensitive comparison of the complete vector entries, such as sorted borrowed entries, avoids copying or interpreting XML and preserves namespace information. Preserve multiplicity unless a separate proof justifies collapsing it. Do not retain only the plain Word-prefix membership. The encoded prefix/URI records at `numbering.rs:587` and `numbering.rs:595` carry foreign, default and empty namespace bindings that still matter to opaque content and future reparsing. A changed URI for the same prefix must remain unequal even when every revision byte and Word-prefix membership is the same.

Do not omit private context altogether. It is consumed by deferred physical content parsing at `revision.rs:207` and by accepted revision edit refresh at `revision.rs:348`. Do not omit raw/nested/projected fields from equality either. The public physical projection mutation at `revision.rs:197` and accepted-edit publication at `revision.rs:305` make such differences observable. Nested CT_Revision values should receive the same equality rule through normal recursive comparison, while retaining their locations and complete content.

Retain comparison's bounded refusal and actual old-revision byte check at `comparison.rs:3500`. Once the owner compares unchanged, the existing route copies its ORIGINAL source slice at `comparison.rs:3809`. That provides the requested original paragraph bytes, including its original redundant declarations. The repair must not normalize the original publication, accept arbitrary raw XML differences, or skip either reconstruction postcondition.

## Alternatives evaluated

- **Canonicalize only stored parse metadata**: smaller-looking constructor edits can remove this vector-order difference, but the stored vector is later reused for parsing and edits at `revision.rs:210` and `revision.rs:348`. This changes inputs to those consumers rather than only identity. It also does not automatically canonicalize already parsed nested/projected values. An equality-only rule preserves those original parser inputs directly and has the narrower behavioral effect.
- **Sort shared word_prefixes_at output**: rejected for this repair. That helper is used across paragraph, property, table and revision parsing. Enumeration order is also used to choose an available output Word prefix at `numbering.rs:764` to `numbering.rs:770`. Changing the shared scanner can change writer choices far beyond revised-owner identity. The remove/reappend scope update must retain local-shadow and undeclaration behavior.
- **Comparison-specific semantic body equality**: a recursive policy over paragraphs, tables, controls, properties and revisions would duplicate their existing equality and risk skipping opaque state or new fields. A generic semantic XML comparison is neither needed nor authorized. The representation problem belongs at the specific CT_Revision context equality boundary.
- **Ignore revision metadata or compare only author/id/text**: rejected. Raw opaque XML, nested revisions, source namespace meaning, property projections and parse policy remain part of value identity. A successful comment beside the revision must not weaken refusal of a real revised-owner mutation.

Record this internal equality clarification and its additional lower-level crate checks in the existing plan/progress before implementation. It is within approved preservation scope, with no material user product choice. The HLD03/HLD10 work list remains the same.

## Consumer and regression work list

| Consumer | Required invariant | Source citation |
|---|---|---|
| Unchanged body/table/control owner | Equivalent enumeration order may compare equal, all actual content/state still compares | `crates/rdocx/src/comparison.rs:3808` |
| Empty revision marker constructor | Same identity rule as paired revision, preserving require-author policy | `crates/rdocx-oxml/src/revision.rs:121` |
| Paired and nested revision parsing | Complete raw and nested/projected state remains authoritative | `crates/rdocx-oxml/src/revision.rs:153`, `crates/rdocx-oxml/src/revision.rs:768` |
| Deferred physical projection | Original parser context is consumed unchanged | `crates/rdocx-oxml/src/revision.rs:207` |
| Accepted revision refresh | Original context is retained through checked edit and reparse | `crates/rdocx-oxml/src/revision.rs:325`, `crates/rdocx-oxml/src/revision.rs:348` |
| Raw serialization | Equality changes no captured revision bytes or writer authority | `crates/rdocx-oxml/src/revision.rs:354` |

Before the repair, add and execute focused controls in existing test carriers. The fresh Python failure is already a genuine RED for the exact public reproduction, but it does not replace RED for a newly added low-level discriminator.

1. Construct otherwise identical revisions under complete context vectors with the same entries in different orders. Cover an empty marker and paired revision, plus a nested revision or retained projected paragraph. They must compare equal after repair, keep their raw bytes and original context vector, and retain the same parsed projections. This rejects continuing ordered-vector identity and catches incomplete handling of constructor variants or nested values.
2. Keep the raw revision bytes fixed while changing one non-Word prefix URI, a Word-versus-foreign binding, or default namespace/undeclaration context. Meaningful changed contexts must remain unequal. A foreign prefix used by opaque element/attribute content is a stronger control than an unused Word alias. Do not collapse encoded records to the set of Word prefix names.
3. Preserve unequal comparisons when changing actual raw opaque data, author/id/date, a nested revision, typed/projected content or parse policy. Keep the existing changed/revised table/control/move atomic-refusal controls. An explicit implementation must continue comparing every existing field.
4. Add a native PUBLIC comment reproduction matching the Python setup: create Reviewer A's redline, open independent copies, add Reviewer B's comment on the other paragraph with the checked public comment API, compare, save/reopen, then accept/reject B. Require the original paragraph and old revision bytes, authors and counts to survive, with the intended comment graph on each resolution side. Do not replace the public call with a hand-written XML comment edit as the only native evidence.
5. Rebuild and authenticate the successor wheel before rerunning the same four selected Python nodes, including the failing new-comment case and the priority atomicity control. Preserve the current authenticated 3-pass/1-fail result and original artifacts. No prior installed module can qualify changed lower-level source.

Add scoped check/clippy for the now-affected rdocx-oxml crate and consume or refresh applicable native/CLI local checks. Keep all 22 prior namespace/relocation/ignore/refusal controls in the bounded local batch after the equality change, with actual membership/outcomes. Include a parser projection or accepted-edit control proving the original context vector is still available and namespace-sensitive parsing is unchanged. Final complete suites, native Office, hosted CI, hash/pixel and installed typing remain at the integrated union boundary.

## Not found and remaining work

No evidence establishes another N1 transport failure in this comment reproduction. No need for a broader revised-owner editing policy or user choice was found. Zero independent panic, structure, smell or nitpick findings are reported for this narrow pass. This does not certify unrelated feature code.

N3 remains unresolved until implementation, actual focused GREEN, fresh artifact-bound Python results and independent review. The previous 22 native, one CLI, scoped check/clippy/fmt observations remain attributed to their frozen candidate, whose source hashes are recorded in `fx272-frozen-native-cli.json`. No source, test, plan or progress edits, builds, new product tests or commits occurred in this review. Return control to root after this report. Full feature microscope and final acceptance remain pending.
