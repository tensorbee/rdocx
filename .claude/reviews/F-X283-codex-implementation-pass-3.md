# F-X283, Codex implementation, pass 3

**Reviewed**: Frozen worker `work/f-x283-codex`, Base/HEAD `101c34c4a6d5fa62d9ba0e0ff745fd44c9aeaf13`. Twelve changed files, 2969 insertions and 315 deletions. Whole diff SHA256 `519c89907967c6903d0f1ddb017e354f25c013d1bf914369575cd6dac2016fab`.

**Verdict**: 0 defects, 0 smells, 0 nitpicks established. D1 is closed for the reviewed implementation. This is actual implementation peer 3 of 3. Counters are not reset. Root final, formal preparation/integration and final acceptance remain separate and pending.

## Defects

None found. Stable D1 disposition and evidence are below.

## Smells

None found.

## Nitpicks

None found.

## Authentication and scope

Read the operative plan including both private-publication and complete-membership amendments, prior peer findings, stopped handoff, complete remediation increment, affected tests and the current native/binding publication graph. Worker plan SHA256 is `f1dd048c5cec935c403e72135e80390975c133e0a462991ce089fe088b572ebb`. Plan-patch SHA256 is `9f382f897e8c880d766849de3b1d15056425c8e32c97ad15c46d91b91dc13569`. Remediation-only patch SHA256 is `a8cbc7ebf328c059286f057e1e08c3f392d3cc80d7de0029ddb0f851901ee139`.

All twelve source hashes, actual working diff and three patch hashes match the freeze. All 311 build inputs match current files, with manifest SHA256 `b27743032b5d745cbdbd8a0bb53988f2ba3d42c3a685ed2dbb932c10f6567193`. Installed extension SHA256 matches `af041c14a122b82a8219b83c1d4b7388621f187e892db0c8b82982bba48ce8be`. All 42 evidence-packet hashes match. Current diagnostic inputs and binary SHA256 `b394f4209b5fe64c4741555827841b189c5a74e936a54649b502f629871e6af2` match independently.

Unchanged source mechanisms were compared with their already authenticated earlier review inputs. The actual production repair remains confined to binding document.rs. No tests/builds, source edits, Git/network mutations or delegation were performed in this review. This report is the sole write. Canonical F281 work does not change this worker or its Base.

## D1, closed by complete physical membership before publication

The local matcher still uses its existing ordinal join. The repair now proves the topology for that join rather than treating one physical table as sufficient. WordOwnerFacts::has_complete_table_membership at `crates/rdocx-py/src/document.rs:5065` requires a standalone direct-body root table, exactly one physical Table fact, projected membership cardinality equal to the actual native physical cell count, matching table/part identity, TableCell story kind, contiguous physical ordinals starting at0, and unique valid direct coordinates. Every paragraph must belong to that same qualified cell-story range.

This closes both previously identified causes. Nested tables fail the physical-table condition. Controlled row/cell members omitted from direct-coordinate projection fail against native cell_stories, including empty omitted cells with no paragraph witness. A projected first cell with physical ordinal1 cannot be certified as standalone ordinal0. Thus the first indexed cell can supply the placement base only after complete source-ordered membership has been established. Every admitted paragraph then has a corresponding indexed cell parent, and its stored run descendants remain inside the selected-table indexed subtree.

The old piece is checked before local owner/index/revision preparation (`crates/rdocx-py/src/document.rs:11611`). Fresh candidate facts require the same certificate before placement and matching (`crates/rdocx-py/src/document.rs:11697`). Missing or ambiguous membership uses the existing checked publisher. It creates no new public refusal. Distinct CellText/ParagraphText/property lifetime policies are retained by that publisher. Controlled paragraphs inside an ordinary direct cell remain distinguishable through their per-story projection ordinals. Textboxes and links retain their separate existing fallback qualifications.

The recorded native-inventory diagnostic at `crates/rdocx-py/src/document.rs:20592` uses actual BodyItemOwners output. It observes2 physical cells versus1 projected cell for controlled rows/cells, both orders, including empty controlled cells. It verifies the selected physical ordinal, false certificate, direct empty-cell success and rejection of duplicate story/coordinate facts. This is materially stronger than asserting a synthetic boolean or comparing projection count with itself.

## Certificate constructors, custody and bounded qualification

All three WordBlockOwners constructors maintain the new fact. Full word_block_owners obtains cell_stories from native body_item_owners, then computes the certificate once from immutable standalone facts (`crates/rdocx-py/src/document.rs:5380`). plain_paragraph_owners stores false (`:5423`). table_owners computes it from the already shaped detached table inventory (`:5453`).

The detached constructor requires separate producer reasoning. Traced spliced_table_blocks and both production SplicedBodyItem::Table producers: Document.add_table and Document.insert_table create only Document::detached_table(rows,cols) candidates (`crates/rdocx-py/src/document.rs:17431`, `:17718`). They do not import arbitrary controlled tables into this shortcut. The existing spliced_table_blocks shape proof requires a standalone root, contiguous cell-story identities, direct indexed cell parents and paragraph/run membership before its piece is constructed (`:7243`). Its cells argument is justified by those concrete producers and shape proof, not a generic claim that projected cardinality equals native physical membership on imported input.

Reviewed full assembly, qualified Arc reuse, missing/incomplete cache fallback, prepared splice installation, active install_inner, property relocation and local replacement. New full pieces compute current facts. Reused pieces retain immutable block/fact identity under existing context/value qualification. Splices carry prepared pieces. Local publication installs the freshly certified replacement Arc after the last fallible revision commit. General unprepared installation retains the existing empty fallback. Dormant installation retains None. No path was found that relabels incomplete imported membership through a guessed positive certificate.

Setter eligibility reads one cached boolean in O(1). It performs no new model/table membership scan at that decision point. Candidate fact preparation remains table-local. Flat Issue445 tables, including direct empty cells, remain eligible. Existing flat structural work controls stay unchanged. This is bounded-work evidence, not release timing acceptance.

## Actual runtime and diagnostic evidence

Corrected RUNTIME_TEST_SOURCE.py matches the current test hash. It exercises controlled rows/cells before/after the selected direct cell with CellText, ParagraphText and property controls. Assertions cover physical story/XML routes, held/fresh paragraph and property observations, ordinary run/Font retirement, follow-up edits and save/reopen (`crates/rdocx-py/tests/test_formatting_tables.py:2701`). Actual Base records16PASS. The prior frozen artifact records8FAIL/8PASS, with all eight text cases failing the controlled-paragraph checked-lifetime discriminator and property cases passing. Repaired installed source records49PASS, comprising those16 cases plus the original affected21 and retained12 ancestry controls.

The initial probe incorrectly expected controlled held paragraphs to survive checked text publication. Its separate Base8FAIL/8PASS is fixture history, not a production RED. Initial frozen-route panics remain their own observations and are not relabeled as execution of the corrected fixture. The peer1 source finding and peer2 source finding likewise remain source review, separate from the later actual author runtime reproduction. Original nested-table, first-tcPr, distinct paragraph policy, field/control/textbox, related/direct-ordinal and comment observations are retained rather than replaced by counts alone.

Current membership diagnostic1PASS and C9 diagnostic1PASS have actual commands, child0/wrapper0 and authenticated executable/input identities. C9 still exercises current binding objects, an injected commit_scopes refusal, unchanged bytes/serial/held cell and successful retry. It is not installed-wheel fault injection. No installed hook is exported. That installed injection remains unperformed and must not be claimed by Root final. Current scoped check, all-target/all-feature warnings-denied Clippy, format and affected prose record0. Initial missing-HashSet qualification was a compile/setup failure, not semantic RED. Progress punctuation corrections do not rewrite historical results.

## Remaining graph and acceptance boundaries

Reexamined the affected local candidate, native source certificate, preparation/commit, owner/generation/route/parent indexes, reads and pieces. Distinct selected-owner publication checks and retirement remain. Disappearing owners, unqualified scopes and incomplete caches still use checked fallback. No additional correctness, panic/input-boundary, OOXML preservation/child-order, structure/dependency or HLD-impact finding was established.

Dormant construction, exact eager validation and discard, unconditional current-native activation, current-serial initializer, guarded factories, numeric publishers, unbound link removal, direct URL/image publishers, serializers and consuming replacement zero/mismatch/positive paths retain their authenticated earlier implementations. No retained dormant facts or pieces were introduced by remediation. Native/support sources remain identical to their earlier six actual cases, reused without claiming a new execution. HLD03/10/12 retain the checked/unproved-route boundary and original acceptance obligations.

All thirteen A-ID rows and original public bridge inventory remain. These are macOS arm64 dev CPython3.14.7 semantic/debug observations. Exact Linux x86_64 release qualification remains pending: bodies0/500/2000 with a20x5 table,100 acquired/reused handles, text/alignment/shading in order without warmup, each2000-body setter below1.0ms and at most1.5times empty-body plus0.05ms. Case2 retains4000 three-run paragraphs,1000 water replacements, three paired trials, lexicographic minimum Python total/open/replace/save tuple and minimum matching CLI command time including process start, ratio at most1.2. No counters or reduced workload substitute.

Installed3.9/3.12, strict typing/stubs, schema/native Word/applicable PowerPoint, hash49 unchanged, retained F26879/252/407/five-linear/count/memory/PowerPoint248 riders, WASM, full integrated union, documentation, archives/publication dry-run and hosted CI remain due. No F272 raw-admission/provider acceptance follows from this Base. No threshold, baseline, native gate or separate reviewed-SHA real-publication approval is waived.

This completes the final allowed implementation peer pass. No further peer or reset is authorized by this report. Return the review-directory writer lease to Root for the separate Root final. No remediation, formal prepare, integration or release approval was performed. STOP.
