# S96 direct cell value certificate, design investigation 14

## Disposition

A direct-table-cell extension of the existing fixed standalone value certificate has a finite technical contract. The conditional shape below is supported. It remains unselected and is not approved for implementation by this report. Actual current candidate availability, both-side finite equality and complete root/selector/key persistence are required before recommending implementation. Those forthcoming observations were not supplied or consumed during this pass. A remaining-reader count is not evidence that its unexplained remainder is eligible table work.

Zero existing correctness defects, zero smells and zero nitpicks were established. Root remains sole product/Cargo writer. Only this report was written. No compiler, product import, test, runtime, instrumentation, benchmark or source mutation occurred. Earlier pass numbers are retained. The latest actual ordinary14e78/e490qualification still has reply and Document.add_picture failed. Nothing here predicts a Document picture repair, changes that small real miss or authorizes a retry.

## Source boundary

The Native source was read from the immutable `.claude/scratch/S96-final-evidence/f268-standalone-value-implementation-20261107-1/final-source/crates/rdocx/src/document.rs` while Root prepared temporary observations. Its SHA-25688abd485c574432bdc296a4c3ace85973c42c1061be157270df2a6a9a0c40bd2independently matches the Native entry in the Root27final4309input manifest. Native document.rs citations below refer to that unchanged reviewed source. No temporary diagnostic source was treated as adopted product policy.

Current CommentParagraphRead at Native document.rs:9192has one call-local candidate snapshot and direct-paragraph ordinal. Its read at9205recognizes only main Body owner0, non-end, one-component Paragraph locations. BodyBlockScan.standalone_value at9159is one lazy shared cell. StandaloneParagraphValue9186holds one complete XML Vec plus Weak<BodyContent>. Its hit checks exact key/live root identity, then clones the finite owned paragraph outside the lock. Its miss calls literal CT_P::from_xml_fragment at9252, then field-free/both-finite/full-equality checks before installation. This is the proof owner being extended, not Part-child or writer-read reuse.

The only consumer is CommentAnchorRead::from_prepared at `crates/rdocx/src/comments.rs:3085`. Existing parent ownership/unknown-parent checks and counts run first. Complete comment_story_range_paragraphs traversal runs at3095, then the snapshot is built. Per-paragraph comment_source_markers runs before values.read at3100. Flow, offsets, locations, marker ownership, namespace restoration and projection remain actual current-source work. This ordering must stay unchanged.

## Concrete leaf domain and restrictions

The stored value graph is ordinary owned data: BodyContent::Table owns a CT_Tbl, whose rows Vec is at `crates/rdocx-oxml/src/table.rs:3366`. CT_Row.cells is an owned Vec at2890. CT_Tc.content is an owned Vec at2399. CellContent at2388has Paragraph(CT_P), Table(CT_Tbl) and ContentControl(CT_Sdt) variants. A concrete direct selector is the physical `(row, cell, content)` triple into those vectors. It must use bounds-checked get operations and require CellContent::Paragraph at the final position.

Paragraph ordinal is not content index. Nested tables and controls occupy content entries even when a paragraph-only iterator hides them. Source locations and model ordinals are also different concepts. Do not use a Source index_path component, visual grid coordinate or merged-cell alias as the concrete model content index.

The closed selector domain may use one private enum with two real variants: existing root Paragraph, and direct TableCell paragraph with the three physical indices. A substantive same-file resolver may own the variant/bounds checks. It returns only an immutable CT_P reference from the live candidate root. No generic root provider, trait, module, source file, public selector or API is needed. Missing/wrong variants or bounds cause an optimization miss, never a new error before the original reader.

No nested table, content-control, raw, header/footer/note/textbox or visual merged-path support is selected. Do not recursively descend those branches. Root must name the initial direct structural admission predicate before implementation. A conservative first predicate may decline roots with nested/control/raw structure and merged-cell metadata entirely. The actual merge fields are CT_TcPr.grid_span, h_merge and v_merge at table.rs:2031. No new metadata restriction may reject source input: it declines only the candidate and leaves the old reader. Relaxing this candidate domain later is a separate contract decision, not implied support here.

The leaf language remains exactly paragraph_has_finite_plain_properties at Native document.rs:8830and paragraph_is_plain8804. Keep the lexical Field exclusion and check both selected leaf and actual returned CT_P before any complete leaf equality. Do not replace this with owner_facts_reusable_by_value. Do not compare a whole CT_Tbl or scan private sibling values with equality. Unselected siblings can contain private Field/read identities that the original one-paragraph call did not compare. Their actual Source/reader work elsewhere remains, but this certificate must add none.

The cold opened-paragraph writer-invariance grammar is unchanged. Comment marker/reference exceptions remain the separately unselected investigation10option. A cell leaf containing those markers, Fields, bookmarks, tabs, raw data, controls, revisions or nonfinite properties keeps the original reader. There is no predicate that strips unsupported data to manufacture eligibility.

## Once-per-factory advisory hints

Extend the existing snapshot at document.rs:32422after complete source traversal, retaining only already offered serialized-first block/root handles. Walk staged table rows/cells/direct content once per factory and record concrete direct selectors. Keep one root/block reference per root where practical, with compact hints referring to it, rather than another table clone or per-paragraph table walk. Resolving a selector is three checked vector lookups.

The source-facing key may be a main-part cell ordinal plus a paragraph ordinal within that owner. It is advisory. Source paragraph counters advance for every qualifying actual source paragraph, including ineligible leaves, absent candidates and failed finite matches. Model placeholders/ordinals must not collapse merely because a leaf is unsupported. Neither owner_index nor index_path is declared physically equivalent to a model slot by this map.

Only main-part TableCell non-end Paragraph locations in the approved direct domain obtain hints. Existing direct Body hints remain their separate actual case. Other sources keep literal reads. Missing staged roots, source-only blocks and skipped structural domains produce no positive certificate. Unknown source/model alignment must cause misses rather than new source discovery, parsing or fallible validation in the hint builder.

A wrong hint can select a different immutable value, but certification still requires the actual standalone reader first and complete finite CT_P equality. If it happens to select the same value, the proof is that owned value, not location, physical membership, source range, table identity, grid equivalence or writer permission. No Source authority is inferred from count, text, row index, staged equality or roundtrip bytes. Actual candidate observations must establish useful alignment rather than assuming it from this safety argument.

## Explicit bounded state and algorithm

Replace the current one-key payload with one active-root state per lazy shared cell. The state owns one Weak<BodyContent> and a map from valid concrete selectors to one complete XML Vec each. It owns no CT_P, CT_Tbl, row/cell model, source receipt or list of roots. A root Paragraph has one valid selector and therefore at most one entry. A direct table has at most its actual finite set of approved direct paragraph selectors. Arbitrary XML text cannot create a selector. Resolve/validate against the live root before installation and again on hit.

This explicitly changes the old one-key-per-BodyBlockScan contract. Direct paragraphs also pay whatever map/container overhead the chosen common representation introduces. The table's entry count is bounded by its current immutable model, not by a new user-chosen ceiling. Full XML keys can be large, and one key per selector is not a constant-memory claim.

Lookup requires all of: applicable advisory candidate, valid selector resolved on that root, exact complete XML byte equality for that selector, successful Weak upgrade and Arc::ptr_eq with that same live root. Do not use root identity alone, selector alone, a hash, text, a partial XML key or equal-but-independent root. Clone the selected CT_P outside every lock. The factory holds a legitimate strong block/root for the operation, so the reference cannot disappear between lookup and clone. Weak prevents ABA and does not retain the AST.

On miss, call the unchanged fixed ordinary CT_P::from_xml_fragment exactly once and propagate its exact first failure. Only actual success, unchanged lexical Field exclusion, both finite qualifications and complete selected-leaf equality may install a positive key. All reader, qualification, equality and returned clone work is outside locks. No actual standalone read is inferred from a staged model or a Part-child receipt.

For same root/selector success, replace that selector's sole key and drop the old Vec, including its old capacity. A successful nonmatch removes that selector's entry without clearing unrelated valid selectors. For root mismatch, expired root or model replacement, detach/reset the current block cell as described below, or atomically replace the entire active-root state before installing a new positive. Drop the old map/container and all keys rather than clearing entries while retaining a historical larger-root capacity. There is never a map per historical root.

An actual reader failure creates no entry or negative/error cache. For the same still-valid root it may leave the previously proved key available, matching the existing failure behavior. Invalid prior-root state may be discarded without suppressing that error. A malformed different key must still execute the real reader even when another key is warm. Unknown/missing candidates never skip actual reading.

Construct a complete replacement root state outside the mutex, then publish it as one intact state. Same-root key updates are likewise complete positive entries. Poison recovery follows the existing into_inner policy, but normal root/selector/full-key checks still apply. No partially installed state is accepted. Parsing, callbacks, full equality and leaf cloning never occur while holding the mutex.

Concurrent old/new snapshots may lose hits or replace one another's active state. That is safe. They may not append old-root maps or install a key under a different root from the one actually certified. Use the strong root captured by that actual miss, then recheck/update the locked active-root state consistently. Concurrent same-root different selectors can retain both valid keys. Concurrent different keys for one selector retain only the last complete positive entry.

## Producer, clone, mutation and carry closure

The actual producer audit gives the following complete staged-root/cell boundaries in the inspected production source:

| Path | Required handling |
|---|---|
| Writer creates new BodyBlockScan at10891 | Existing actual item clone creates a new immutable staged root. Start with an uninitialized certificate cell. No opening/eager standalone parse or key allocation. |
| MainPartBlocks::serialize_in10771reuses a block after its existing full writer qualification | Carry the same root and cell. Do not remove or replace the original qualification, including private/full fallback work. |
| BodyBlockScan derived Clone9150 | It may share initialized cell state only with the same staged root. Cloning a block is not proof for a later replacement root. |
| fresh_context_piece_from9665 | Current code clones staged and initialized cell together, clearing other context proofs. Sharing remains safe only with that identical root and exact full-key checks. Changed frame/namespace bytes miss. |
| MainPartBlocks::edited10405 | Changed blocks are new from_source records at10435with empty cell. Unchanged block/root handles may carry. Tail/frame-only copies use the preceding same-root rule. |
| qualify_document_children10594 | Outer vector and block use Arc::make_mut at10600/10602before read_source_child can replace staged. This is a real replacement path, not just a new record constructor. |
| read_source_child12156 | After actual checks it replaces staged with a newly allocated actual roundtrip root or None at12211. Detach/reset this block's certificate cell before/with replacement, including a previously shared cell from derived Clone. Weak mismatch alone prevents bad hits but does not release old retained keys. |
| producer_source_blocks28612 | Constructs from_source, performs the actual read, then explicitly clears staged28675. Keep the cell empty/reset. Source-only facts must not acquire value storage permission. |
| proved_unsavable_body_edit39489 | Temporary block at39520receives a new staged root. Its cell starts empty and stays ordinary temporary ownership. Do not seed certificates from that staging alone. |
| source_blocks_in28868and set_story_source_xml16665 | Replacement blocks use actual from_source/read producers. Keep all current Source, partition, resource and model installation checks. Carried unchanged roots retain only their own exact-key certificates. |
| clone_for_staging24869and reopen_staged28184 | Existing shared scan/record carry can keep the same immutable root/cell after original qualification. An independently equal newly parsed/rebuilt root cannot hit merely because it equals an earlier table. No history edge is added. |
| table_mut39009, edit_cell_paragraph36675, commit_direct_body_table39648and staged commit26867 | These mutate/replace the live owned document model, not an in-place staged root. Preserve their actual guards/atomicity. The subsequent writer supplies a newly qualified root or reuses the old immutable value through its existing proof. No direct copying of old certificate state onto a new root. |

Production search found no Arc::get_mut/make_mut/strong-count branch mutating the inner staged Arc<BodyContent> in place. Existing mutable vector/block paths use COW and then replace staged. Adding a Weak root therefore must not be used as a reason to remove a uniqueness check. If a new direct staged-root mutable access is introduced, it must dissociate Weak identity before exposing mutable data and reset certificate state. Cover sole-strong-plus-Weak and shared-root cases explicitly.

The table graph's vectors/paragraph leaves are ordinary owned immutable data through this Arc. There is no shared mutable projection of the selected finite leaf. Private sibling objects do not justify whole-table equality or retention. The call-local candidate snapshot and its strong handles die when the factory returns. No second AST, separate row/cell cache, opening-model history or delayed strong root may be retained in the certificate.

## Required actual factual closure before selection

The forthcoming diagnostic must use the current reviewed producer and exact original reply actions0/1/2, with no additional semantic getters between actions. Only after the original actions run may richer semantics/reopen checks inspect the result. Independently identify:

1. Actual fixed standalone-reader calls by body/cell/related owner and precise finite exclusion, rather than assigning an unexplained read remainder to tables.
2. Offered immutable table-root availability and each advisory hint's concrete valid selector. Count real missing/source-only/unsupported candidates.
3. Actual first standalone reader success, both-side finite eligibility and complete leaf equality before claiming a certifiable case. Do not execute whole-table equality as part of this observation.
4. Complete input XML key, root Arc and selector persistence from action1toaction2for every claimed useful warm candidate. An ordinal alone is not persistence. Wrong/shifted hints must be disclosed.
5. Mandatory public semantics, parent/anchor/range/text, untouched cells/neighbors, held/fresh behavior and reopen. Keep actual output inputs/packages and producer restoration custody.

If useful leaves fail finite qualification, lack offered roots, compare unequal or do not retain root/selector/key, this repair is not selected. Do not broaden grammar, private eligibility or Source rules to rescue the hypothesis. This report currently has none of those actual factual receipts and therefore makes no implementation recommendation.

## Implementation discriminators if later adopted

- Two different finite cell paragraphs under one root: real first reads followed by real hits for both, exactly two keys and no one-slot thrashing. A deliberately wrong selector with a different leaf must miss. Same XML under a separate equal root must not inherit proof.
- Zero/one/many rows/cells/paragraphs, missing staged root, invalid selector and shifted advisory ordinals. Resolve physical content positions correctly. Nested/control/raw/merged/related-story controls keep literal readers under the initial scope.
- Same-selector changed key, namespace/default/alias/shadow/root/suffix changes and malformed input. Error is uncached and preserves first cause. Successful alternate key replaces old capacity, successful nonmatch removes only that selector, unrelated valid entries survive.
- Private Field sibling next to a finite selected leaf, with actual private XML/equality work observed. No new whole-root/private-sibling comparison may appear. Marker/reference, Field, bookmark, tab, revision, raw and nonfinite-property selected leaves remain actual reads. Cold invariance grammar must remain unchanged.
- Changed model/text/properties, insert/remove/reorder, direct cell and source edits, staged None, actual read_source_child replacement, frame-only carry, clone/reopen/drop, independent equal root, Weak expiry, Arc::make_mut and refused/rolled-back candidate. New roots have no copied proof or retained prior-root keys.
- Parallel same-root selectors, parallel different roots, successful/nonmatching/error races and poisoned locks. Every hit still checks exact root/selector/key. One active root and one key per valid selector hold after every update. Old maps/keys and factory strong handles release.
- Genuine actual predecessor literal reads for the claimed current public cell category, followed by first successor actual reads and later real hits/reduced literal calls after successful semantics. Count actual full-key copies, retained selectors/keys/capacities and returned owned CT_P clones. Do not claim zero allocation or total reader removal.
- Pin the exact original inputs and compare entire predecessor/successor ZIPs and every member after full semantics and reopen. Keep unrelated malformed sources, marker/unknown-parent priority, resource refusal, callbacks and held/fresh state. No resource retry under a new role or input grant.

## Adoption and final riders

If actual facts select the proposal, amend the existing plan and HLD03/04/10explicitly. Name the expanded root/selector container, direct-domain exclusion, publisher reset paths, failure/replacement rules and source/value separation. This is a real expansion from one key per block, not an implementation detail already approved by design9. No marker grammar extension is adopted by association.

Update the existing memory preparation StandaloneValue endpoint to omit the complete new lazy container, Weak root, selector/map nodes, every complete key and all clone/carry/lazy installation paths. Include root state count, valid selector bound, map capacity, key lengths/capacities, header/control metadata and actual key/leaf copies in activation/lifetime proof. No sum of payload lengths is RSS or total allocation. The fixed font workload may not activate this layer, so an actual comment/table companion is required. The existing Q22policy and absence of an invented memory ceiling remain unchanged.

Actual implementation/scoped installed evidence and independent microscope/Root review must precede a fresh ordinary final-source qualification. Reply elapsed benefit is unproved and the separate Document.add_picture failure has no selected repair here. All final memory/Office/typing/hosted/archive/full-union riders remain due. No workload, threshold, policy, state/ledger, commit, push, tag or publication change follows. This conditional investigation ends here.
