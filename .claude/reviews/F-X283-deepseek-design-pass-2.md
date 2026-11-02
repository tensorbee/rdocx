# F-X283 independent design peer 2, DeepSeek

**Verdict: do not approve.** The plan stays draft. This is a design audit, not
implementation or remediation. No counter was reset. No other model's review was
impersonated.

## Source identity and protocol

Reviewed frozen snapshot `/private/tmp/rdocx-s96-f283-review-frozen`, detached at
`5820539cd277b417ae060e15010eb73bd1d00b01`, with all 20 pins in
`F283_DESIGN_PEER2_RELOCATED_PINS.json` authenticated. Plan SHA-256 is
`6dc77ac57636d88c4e4239b70d5762dcf7538a935f38991bd71973f2bfac469c`.
All relative citations below are from that snapshot. The prior canonical
worktree was not read. No Git mutation, build, test, benchmark, hash harness,
native app, network, credential or delegation was used. Runtime evidence is
absent and remains pending.

The plan is materially better than the earlier proposal that Grok peer 1
reviewed. It now names the setter route, the ordinal relocation, the candidate
and commit ordering, the width and margin distinction, and the CLI and Python
entrypoint difference. The remaining defects are in the callable closure, the
index and slot model, the refusal atomicity, the dormant authority activation,
and the bounded-work claim.

## Findings

Severity is assigned from the source, not from a passing test. Counts are
17 defects, 6 smells and 2 nits.

| ID | Severity | Finding | Trigger and expected outcome | Proposed or observed outcome |
|---|---|---|---|---|
| F1 | defect, high | `candidate.validate_resource_limits()` is not callable on `CT_Tbl`. Plan line 82. | A cloned `CT_Tbl` must be validated before return. Expected a callable native validation method. | `CT_Tbl` at `crates/rdocx-oxml/src/table.rs:2861` has no `validate_resource_limits`. Only `CT_Body` at `crates/rdocx-oxml/src/document.rs:5174` and `CT_Document` at `:5436` expose it. The plan's native route does not compile as written. |
| F2 | defect, high | `slot` is conflated between body content index and table ordinal. Plan lines 75, 92 to 96, 130 to 151, 241 to 243. | Case 1 has bodies 0, 500 and 2000 before one 20 by 5 table. The table ordinal is 0. The body content slot is 2000 at the largest body. | The native candidate and `WordPieces::splice` use body content index. `cell_paragraphs` is keyed by table ordinal at `crates/rdocx-py/src/document.rs:5639`. `PyCell` path `PathSeg::Body` is the table ordinal at `crates/rdocx-py/src/table.rs:748`. `direct_body_cell_slot` returns `body_slot` at `crates/rdocx-py/src/document.rs:10365`. The plan never states the conversion. A naive pass of the handle's table ordinal fails the case 1 workload. |
| F3 | defect, high | `tables_by_parent` does not give cell owners. Plan lines 139 to 144. | The preparation must find the affected cell owners of the selected table without scanning `self.owners`. | `WordOwnerIndexes::index_owners` indexes only `WordOwnerLocation::Paragraph` and `WordOwnerLocation::Table` by parent at `crates/rdocx-py/src/document.rs:4600` to `:4631`. `Cell` owners are never indexed by parent. `tables_by_parent` holds nested tables, not cells. The O(body) scan is not removed by the cited map. |
| F4 | defect, high | No incremental owner and index update exists. Plan lines 145 to 163 and checklist line 439. | The local path promises "re-indexes only the affected owners" and bounded per-call work. | The only reindex method is `index_owners_again` at `crates/rdocx-py/src/document.rs:4565` to `:4577`. It clears `by_generation`, `direct_body_paragraphs`, `paragraphs_by_parent` and `tables_by_parent`, then reindexes every owner. A full reindex is O(owners), which is O(body) for the case 1 workload. The plan does not name the exact per-map removals and insertions, including `by_generation` for newborn owners. |
| F5 | defect, high | `WordPieces::splice` mutates the live cache in place. Plan lines 171 to 174 and 176 to 185. | Preparation must not publish before `commit_scopes`. A refusal must leave the piece cache unchanged. | `splice` takes `&mut self` and mutates `self.blocks` at `crates/rdocx-py/src/document.rs:5526` to `:5547`. `WordPieces` is a `Vec<Arc<...>>` at `:5382` to `:5387`. Calling `splice` on `self.pieces` before `commit_scopes` leaves a spliced cache when the commit is refused. Cloning the pieces first is O(body). The plan states neither. |
| F6 | defect, high | `install_inner` takes a `rdocx::Document`, but the local path mutates `self.inner` in place. Plan line 185. | Commit step 3 says "calling install_inner" after `commit_direct_body_table`. Expected a callable install of owners, indexes and pieces. | `install_inner(&mut self, candidate: rdocx::Document)` at `crates/rdocx-py/src/document.rs:11236` to `:11243` replaces `self.inner` with a whole document. The local path has only a `CT_Tbl`, and `self.inner` was already mutated by `commit_direct_body_table`. There is no candidate document to pass and no piece-only install method. A clone to satisfy the signature is O(body). |
| F7 | defect, high | `unsavable_body_edit()` recomputes O(body) after the local commit clears its cache. Plan lines 96 and 224 to 229. | The qualification calls `unsavable_body_edit()` per setter. The case 1 timing requires bounded per-call work. | `commit_direct_body_table` calls `invalidate_layout`, which resets `unsavable_body_edit_cache` at `crates/rdocx/src/document.rs:24947` to `:24960`. `unsavable_body_edit` recomputes through `compute_unsavable_body_edit` at `:35188` to `:35230`, including `nested_modeled_namespace_owners(xml)` over the main part. The first call after open is cached, every later setter after a local edit is O(body). The plan's "no O(body) cache clear" claim is false. |
| F8 | defect, high | Dormant `expected_replacements` is not closed. Plan lines 272 to 276. | While dormant, `try_replace_text` with `expect=Some` and `replace_all` must not build inventory. | `expected_replacements` at `crates/rdocx-py/src/document.rs:11925` returns a candidate and counts. Its callers `try_replace_text` at `:15993` to `:16004` and `replace_all` at `:16035` to `:16047` unconditionally call `prepare_word_revision` when the count is nonzero. The plan's blanket statement that `expected_replacements` checks the flag does not install or suppress that caller work. The dormant path is not callable for these consumers. |
| F9 | defect, high | Handle binding activation is under-enumerated and not feasible at several sites. Plan lines 267 to 270. | Every handle binding site must call `ensure_word_authority(&mut self, py)` before `capture_owner`, `capture_snapshot` or `capture_prepared_owner`. | The plan says eleven sites. The frozen tree has 34 raw call sites, including 21 current-document `capture_owner` or `capture_snapshot` sites. `PyDocument::section` at `crates/rdocx-py/src/document.rs:14311` to `:14321` uses `slf.borrow()`, `PyDocument::stories` at `:15123` to `:15133` uses `&self` and `owned_story_snapshot` at `:17149`, and `PySectionHandle::section` at `:1571` to `:1579` uses `&self` with `borrow(py)`. These cannot call a `&mut self` activation without restructuring. The plan does not name the borrow changes. |
| F10 | defect, high | The retained related-part scan does not preserve open failure by itself. Plan lines 255 to 263. | A malformed related story part must still fail at open exactly as today. Expected the open path to propagate the parse error. | `WordPieces::assemble` swallows `story_part_owner_snapshots` errors with `.ok()` at `crates/rdocx-py/src/document.rs:5496` to `:5505`. `owner_snapshots_of` at `crates/rdocx/src/document.rs:32233` to `:32267` can return an error. Today the fallback `word_owner_locations` propagates it. The plan names the providers but not the error propagation. A lazy open that reuses `assemble` semantics would succeed where today it fails. |
| F11 | defect, high | Property edits can shift nested tables and content controls that are not on the "edited path". Plan lines 224 to 229 and 145 to 151. | A property edit that emits the first `w:tcPr` shifts every cell story item at or after the insertion. Expected every affected owner to be relocated or the cell to fall back. | `tcPr` is a `PreservedNode` item and paragraph paths use the story item enumerate index at `crates/rdocx/src/document.rs:11650` to `:11659`. Nested tables are `WordOwnerLocation::Table` with their own `location.index_path()` at `crates/rdocx-py/src/document.rs:5861` to `:5885`. The preparation relocates cell paragraphs by ordinal only. The qualification's "no nested table or content control on the edited path" is ambiguous. A cell with a nested table after the first paragraph can qualify for a property edit, and the nested table route goes stale. |
| F12 | defect, medium | Text edit revision state assignment is contradictory. Plan lines 152 to 157 and 166 to 170. | A text edit replaces runs and retires the changed paragraph and run scopes. Expected the changed paragraph owner to receive new content and descendant states. | Step 5 says changed owners keep their `content` state, following `relocated_property_owners` at `crates/rdocx-py/src/document.rs:7851` to `:7907`. That helper preserves property-edit states. A text edit cannot preserve the paragraph content state. `commit_plain_paragraph_xml` at `:10880` to `:10925` assigns newborn content states. The plan does not separate the property and text state paths. |
| F13 | defect, medium | `prepare_scopes` already adds the document scope, so the plan's wording can cause `DuplicateScope`. Plan line 166. | The preparation should retire the document scope once and birth its successor. | `RevisionCounter::prepare` pushes `self.document` into `retired` at `crates/oxml-py-support/src/lib.rs:452` to `:477` and rejects a duplicate scope. The plan says `prepare_scopes(operation, &retired, &born)` "retires the document scope". If the caller also puts the document scope in `retired`, the call fails. The document scope is automatic, not a caller entry. |
| F14 | defect, medium | The "same saved bytes" claim is not established by comment-anchor survival. Plan lines 214 to 222. | The local path skips `prepare_and_reopen_staged`. Expected proof that the skipped reopen is byte-neutral for the tested package classes. | The fallback `try_set_cell_text` runs `reconcile_comment_removal` and `prepare_and_reopen_staged` at `crates/rdocx/src/document.rs:35106` to `:35116`. The reopen serializes and reparses the package at `:25383` to `:25397`. Comment anchors surviving `Paragraph::set_text` prove the reconciliation is a no-op. They do not prove the reopen is a no-op for namespace declarations, opaque siblings or part ordering. A single comment-bearing byte-equality regression does not cover the opaque XML promised by A283-12. |
| F15 | defect, medium | The final rider list omits F-X268 252 and 407. Plan lines 407 to 410. | The Root prompt and sprint registration require F-X268 79, 252 and 407 plus the required final riders. | The plan names only "the existing F-X268 79 flat-cost gates" and hash 49. The 252 and 407 riders are absent from the test plan and final evidence. |
| F16 | defect, medium | The dormant representation of the non-Option authority fields is unspecified. Plan lines 265 to 270. | Lazy open must leave `PyDocument` in a valid state before `ensure_word_authority`. Expected explicit placeholder or `Option` semantics and the revision counter policy. | `PyDocument` holds `removal_scope: RevisionState`, `snapshot_scope: RevisionState`, `owners: Vec<WordOwner>` and `indexes: WordOwnerIndexes` at `crates/rdocx-py/src/document.rs:8139` to `:8149`. `from_document` at `:11488` to `:11526` creates the counter and initializes scopes. The plan does not say what the dormant scopes hold, whether the counter is created at open, or what happens to a previously read revision number when activation replaces the counter. |
| F17 | defect, medium | `word_block_owners` can return `None`, and the local path has no stated fallback. Plan line 130. | Preparation must either build the new block owners or fall back before commit. | `word_block_owners` returns `Ok(None)` when the standalone item's direct snapshot is not index 0 or `word_owner_facts` fails at `crates/rdocx-py/src/document.rs:5213` to `:5240`. The plan assumes `Some` and does not name the checked fallback. |
| S1 | smell | The dormant Python replace path clones the whole document twice. Plan lines 298 to 306. | Case 2 Python open, replace and save must stay within 1.2 times the CLI. Expected no unnecessary whole-document clone. | `counted_mutation` clones at `crates/rdocx-py/src/document.rs:11905` to `:11908`, then native `Document::try_replace_text` clones again at `crates/rdocx/src/document.rs:44280` to `:44283`. The CLI stages once through `try_replace_all_expected`. The plan does not address the extra clone. |
| S2 | smell | The debug `paragraph_route` assertion performs the whole-document serialization the local test is meant to exclude. Plan lines 246 to 249. | A283-5 local work-bound evidence must show no unrelated body traversal. | `paragraph_route` calls `location.native(document)` under `debug_assertions` at `crates/rdocx-py/src/document.rs:10215` to `:10219`. For a cell paragraph, `ParagraphLocation::native` calls `cell_paragraph_story_location` at `crates/rdocx-py/src/paragraph.rs:206` to `:212`. The plan calls this a diagnostic but does not provide a measurement path that disables it. |
| S3 | smell | `splice` clears the cache on a length mismatch, contrary to the plan's "does not clear" claim. Plan line 174. | The cache must remain aligned after a table replacement. | `splice` clears `self.blocks` when the length check fails at `crates/rdocx-py/src/document.rs:5535` to `:5539`. A dormant or previously cleared cache produces this branch. The plan's claim is conditional, not absolute. |
| S4 | smell | The local path has no stated fallback when `cell_paragraphs` is absent or incomplete. Plan lines 241 to 245. | A direct cell paragraph must resolve to a local route or fall back. | `WordPieces::inventory` can set `cell_paragraphs = None` and `routed = false` at `crates/rdocx-py/src/document.rs:5633` to `:5657`. The plan assumes the projection exists. |
| S5 | smell | The local error wording differs from the fallback for a missing cell. Plan lines 76 to 79. | Existing public errors are retained. | The local candidate would use "cell index is out of range". The `Cell.text` fallback uses `cell {row}:{cell} is out of range` at `crates/rdocx/src/document.rs:35109` to `:35111`. The Python closure uses "cell index is out of range" at `crates/rdocx-py/src/table.rs:2093`. The plan does not reconcile the two. |
| S6 | smell | The `Cell.width` `as i32` wrap is not pinned or tested. Plan line 324. | Field-specific numeric conversion must be preserved. | `Length::to_twips` divides by 635 then casts to `i32` at `crates/oxml-core/src/length.rs:73` to `:75`. The plan names truncation toward zero but not the i32 wrap for out-of-range EMU. The boundary cases 634, 635, 636, -634 and -635 do not exercise it. |
| N1 | nit | `commit_direct_body_table` writes `*existing = table`. Plan line 96. | The body child is `BodyContent::Table(Box<CT_Tbl>)`. | `*existing` is a `Box<CT_Tbl>` at `crates/rdocx-oxml/src/document.rs:1725` to `:1731`. The assignment needs `**existing = table` or `*existing = Box::new(table)`. |
| N2 | nit | The candidate adds a new "direct body slot is not a table" error string. Plan line 76. | Existing error families are retained. | `direct_body_table_xml` returns `Ok(None)` for the wrong kind at `crates/rdocx/src/document.rs:31526` to `:31534`. The existing table lookup returns "table index is out of range" at `:34754` to `:34758`. The new string is a small observable change for an invalid slot. |

## Peer 1 dispositions

Grok peer 1 findings are materially addressed. Residuals are new closure gaps, not
the same defects.

| Peer 1 finding | Disposition | Residual |
|---|---|---|
| 1, A283-5 targeted the read and missed the setter | Addressed. Plan lines 233 to 237 change `PyParagraph::set_text` to `paragraph_route`. | S2 debug assertion and F2 to F4 slot and index closure. |
| 2, first `w:tcPr` predicate was wrong | Addressed. Plan lines 145 to 151 and 194 to 205 use ordinal relocation and no properties-presence predicate. | F11 nested table and control shift. |
| 3, in-place property edit was not atomic | Addressed. Plan lines 63 to 98 and 128 to 192 use a cloned table candidate and a commit order. | F5 to F7 piece splice, install and cache clear. |
| 4, dormant replace was not closed over live authority | Addressed in intent. Plan lines 251 to 306 separate dormant and active authority and state the CLI entrypoint difference. | F8 `expected_replacements` callers and F16 dormant fields. |
| 5, A283-6 contradicted the setter | Addressed. Plan lines 323 to 325 state infallible width, negative storage and truncation toward zero by 635. | S6 i32 wrap not pinned. |
| 6, note and revision fallback had no distinct outcome | Addressed. Plan lines 209 to 213 state run content is replaced, including notes and revisions, with no fictional preservation. | F14 byte-equality claim. |
| 7, plain-content gate was delegated to F282 | Addressed. Plan lines 224 to 229 and 331 to 337 use an own qualification and no F282 bridge. | F11 ambiguity for property edits. |
| 8, C10 changed the case 2 timer | Addressed. Plan lines 363 to 370 keep the Python in-process timer and the CLI subprocess timer. | Zero defect found in the timer boundary. |

## A-ID dispositions

| ID | Disposition |
|---|---|
| A283-1 | Not acceptable as designed. Timing threshold exact. Mechanism blocked by F1 to F7 and F17. |
| A283-2 | Not acceptable as designed. Timing threshold exact. Blocked by F1 to F7, F11 and F17. |
| A283-3 | Not acceptable as designed. Timing threshold exact. Blocked by F1 to F7, F11 and F17. |
| A283-4 | Not acceptable as designed. Script, 1.2 ratio, paired trials and Linux wheel plus CLI are exact. Dormant path blocked by F8, F9, F10 and F16, with S1 as a cost risk. |
| A283-5 | Not acceptable as designed. Setter route is now named. Blocked by F2 to F4, F9 and S2. |
| A283-6 | Acceptable with a nit. The numeric contract matches `crates/rdocx/src/table.rs:2306` and `crates/oxml-core/src/length.rs:73` to `:75`. S6 asks for the i32 wrap to be pinned. |
| A283-7 | Not acceptable as designed. Correctness-only scope is right. Blocked by F4 to F7, F11 and F17. |
| A283-8 | Not acceptable as designed. Negative margin refusal matches `checked_table_twips` at `crates/rdocx/src/table.rs:398` to `:404`. Blocked by F4 to F7, F11 and F17. |
| A283-9 | Not acceptable as designed. Owner-view handles can survive a property edit. Shift and text delta are not proven because of F3 to F5, F11 and F12. |
| A283-10 | Not acceptable as designed. Atomic refusal is blocked by F5 to F7 and F13. |
| A283-11 | Partially acceptable. Comment, bookmark and permission anchors match `Paragraph::set_text`. Notes and revisions are dropped as stated. F14 blocks the byte-equality claim. |
| A283-12 | Not acceptable as designed. `w:tcPr` is written first and `extra_xml` slots stay. F11 and F14 block the shift and saved-byte claims. |
| A283-13 | Not acceptable as designed. Open and save can stay dormant. Live-authority replace, cache state and activation are blocked by F8 to F10, F16 and F17. |

## Callable route, caller map and atomicity closure

The native candidate and commit signatures are stated once. They are not
callable as written because `CT_Tbl` has no resource-limit method (F1) and
`commit_direct_body_table` assigns a `CT_Tbl` into a `Box` (N1). The binding
entry point is stated once, but its argument derivation is not: the body content
slot, the table ordinal and the `cell_paragraphs` key are three different values
(F2). The affected-owner lookup cites a map that does not hold cells (F3). The
incremental index update has no method (F4).

The caller map covers `Cell.text`, the four property setters, `Paragraph.text`,
`Document.open`, `try_replace_text`, `save`, `to_bytes` and the CLI `replace`.
It omits `replace_all` and the `expect=Some` caller of `expected_replacements`
(F8). It also omits the related-part open error path (F10).

The commit order is plausible only if every fallible step precedes the first
live mutation. The plan's order is not closed. `WordPieces::splice` mutates the
live cache before `commit_scopes` unless it is cloned, and cloning is O(body)
(F5). `install_inner` takes a document that the local path does not have (F6).
`commit_direct_body_table` clears the unsavable cache, so the next qualification
recomputes O(body) (F7). The document scope is added by `prepare_scopes`
automatically, so the plan's wording can duplicate it (F13). A refusal at
`commit_scopes` therefore does not publish nothing as the plan claims.

Held handle retirement is more precise than the plan states. `Cell.text` retires
the document scope, the snapshot scope, the changed paragraph scopes and the run
scopes. The selected cell owner scope is preserved, so held cell handles survive
a text edit. Paragraph and run handles do not. The plan's C10 line "text retires
them" should say which handles retire. This is a clarification, not a defect.

## Known open items

**Q-F283-01 remains open.** The plan correctly records that no minimal
native-valid input has demonstrated a binding-only mandatory body-open failure.
One source-derived narrowing is available. `admit_word_model` is monotone over
body children with the same starting cost and table depth at
`crates/rdocx-oxml/src/document.rs:174` to `:345`. If the whole body passes the
single-item resource check, each child alone passes it. The remaining witness is
standalone serialization or parse. The exact proof is a native-valid DOCX whose
direct body child fails `standalone_body_item_xml` or `owner_snapshots_of` while
`Document::from_opened_package` succeeds. The related-part scan error
propagation gap is F10, a separate mechanism defect. Root's four probe cases are
negative observations and do not prove equivalence.

**Q-F283-02 remains open.** The two additive hidden `rdocx::Document` methods
are a concrete contract assessment for Root, not an implicit user approval. The
plan's recommendation to approve is reasonable, but it requires the strict
SemVer review, archive check and publication dry-run rider. No F282 bridge or
provider edge is invented.

## Zero findings by category

No defect was found in the case 2 timer boundary. The plan keeps the Python
in-process start, the CLI subprocess start, the three paired trials, the
lexicographic minimum tuple, the minimum CLI command and the 1.2 ratio.

No defect was found in the width and margin sign rules. Width remains
infallible and negative values remain stored. Margins remain nonnegative and
refuse atomically.

No A-ID was dropped. All thirteen rows remain. No hash baseline movement is
authorized. No public properties-presence API is invented. No note or revision
preservation is promised. No F282 provider edge or future API is invented.

## Remaining local and final evidence

Local evidence still due: a RED/GREEN for every A-ID using existing test
entrypoints, with the actual old-source failure recorded separately. Bounded
work counters for the three setters at bodies 0, 500 and 2000, with the debug
`paragraph_route` assertion excluded from the measured window. An injected
`commit_scopes` refusal that proves bytes, owners, handles and the piece cache
are unchanged. Held, fresh, native and reopened observations for `None`,
`Some(empty)` and nonempty `tcPr`. Nested table, control and related fallbacks.
Width and margin boundaries including the i32 wrap. Comment, note, revision and
opaque XML outcomes. Dormant zero, nonzero and refused mutation, then first
handle activation.

Final evidence still due: the Linux x86_64 release wheel and matching CLI on the
exact case 1 and case 2 workloads, F-X268 79, 252 and 407, hash 49, saved bytes,
schema and applicable native riders, installed Python 3.9 and 3.12, strict
typing and stubs, WASM and the publish dry-run. F15 records the missing 252 and
407 names in the plan.

READ lease returned. No remediation. STOP.
