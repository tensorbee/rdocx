# F-X283, Codex implementation, pass 2

**Reviewed**: Frozen worker `work/f-x283-codex` at Base/HEAD `101c34c4a6d5fa62d9ba0e0ff745fd44c9aeaf13`. Twelve changed files, 2770 insertions and 324 deletions. Current diff SHA256 `110ed68ae9d53c1aab68be97629db863b5fb1f491d9268ef5aaefcf9d3d86a0f`.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. D1 remains open on an additional physical-membership case. The original nested-table witness is repaired. This is actual implementation peer 2 of 3, not a counter reset, Root final, completion, prepare or integration approval.

## Authentication and scope

Read the operative contract and exact worker plan patch, peer1 report, stopped author handoff, new source/test/HLD delta and complete affected publication graph. Reused prior source review only after checking identical inputs. Plan SHA256 `7bfb825f75c34b7279648e6dfbc909a720a04bcfc12057d9a97815c871d2ce38` and plan-patch SHA256 `f27d47654175722f4f36b29bad2ea35f4140c63d42c36fb8a595ad7307dd76c1` match. All twelve manifest files, actual git diff and frozen patch match. All 311 current extension build inputs match. Only binding document.rs differs among production inputs from peer1.

Reconstructed the peer1 document.rs in memory from authenticated Base plus its frozen patch and confirmed its recorded hash before comparing with current source. The production remediation is exactly physical_table_count, its field in WordBlockOwners, initialization in all three constructors, and the early local eligibility check. No unrelated production repair is hidden in this delta.

Authenticated installed extension SHA256 `5e6e3c8b8e82da2067234e35dcc821d52eb8ff312d6f5c58f4a479a8f41ec724`, all sixteen new indexed log hashes, current C9 diagnostic inputs and its binary hash `1257e5b490a07aab3b65bcba18fecbd96b7ea9b71171b7d63e35b3058a81f081`. Inspected actual assertions and logs, not aggregate counts alone. No tests/builds, source/Git/network mutations or delegation occurred during review. This report is the sole write.

## Defects

### D1, high, one-table qualification still omits existing control-owned physical cells

`crates/rdocx-py/src/document.rs:11643`

The cached physical_tables==1 guard repairs the nested-table collision, but does not establish that the indexed table subtree contains every existing owner in that physical table. A row/cell control can own a physical cell without a direct row/cell coordinate. Native traversal explicitly assigns None to controlled rows and cells (`crates/rdocx/src/document.rs:21142`). The physical cell owner still exists, but story_table_cells_from_inventory emits a binding cell tuple only for Some coordinate (`crates/rdocx/src/document.rs:32908`). Paragraph facts for that cell story remain in the inventory. word_piece_owner_addresses consequently gives those paragraphs parent=None (`crates/rdocx-py/src/document.rs:5203`, `crates/rdocx-py/src/document.rs:5230`).

Concrete trigger: one direct-body table with a direct row containing a started row-level SDT that wraps a cell/paragraph `controlled`, followed by an ordinary direct cell/paragraph `clean`. No nested table, links, textboxes or nested namespace declaration is required. Acquire the clean direct cell, and optionally the controlled paragraph through document.story_items' paragraph factory (`crates/rdocx-py/src/document.rs:1197`). Set clean.text or clean.paragraphs[0].text. The table has exactly one physical table and the selected clean cell contains only a paragraph, so the new count guard and native selected-cell guard both qualify.

Two failures follow from the same incomplete table-membership proof:

1. The first indexed cell is clean, whose actual global cell-story ordinal is1 because the controlled cell is physical ordinal0. The helper uses that indexed cell's ordinal as the table's physical story base (`crates/rdocx-py/src/document.rs:11643`, `crates/rdocx-py/src/document.rs:11656`). Candidate standalone facts already assign clean ordinal1. Placement adds1 again (`crates/rdocx-py/src/document.rs:11687`), publishing clean owner/routes at ordinal2 even though current native table stories remain0 and1. Cached text and saved bytes can still appear correct while current physical routes are wrong.
2. The old indexed-subtree walk cannot reach the parentless controlled paragraph/run owners (`crates/rdocx-py/src/document.rs:11700`). Candidate copies are therefore treated as newborn. Only property mode rejects newborns, so text mode proceeds (`crates/rdocx-py/src/document.rs:11757`, `crates/rdocx-py/src/document.rs:11778`). Old omitted records/scopes remain installed while new copies are appended (`crates/rdocx-py/src/document.rs:11978`). This duplicates current authority instead of reconciling every existing table owner and preserving the checked publisher's lifetime policy. Putting the controlled cell after clean avoids the first-base shift but still exposes omitted-owner births.

The cell-paragraph completeness flag does not close this gap: the native projection skips cells without a direct address without setting complete=false (`crates/rdocx/src/document.rs:20919`). Equal physical cell-story counts before/after likewise prove no coordinate completeness.

Retain D1 as the same unproved table-wide physical-authority invariant on a different supported topology. Require a complete physical-to-indexed membership/base proof or checked fallback before revision preparation. Add controlled-row and controlled-cell neighbors before/after the direct selected cell, using both text publishers, held/fresh physical XML/routes and expected retirement, followed by another edit. The twelve new one-table controls wrap paragraphs inside an already indexed direct cell. They do not exercise cells/rows owned by controls. This is a source-traced defect. No runtime failure for this new input was executed or claimed during the review.

## Smells

None found.

## Nitpicks

None found.

## Original D1 repair and cache custody

physical_table_count counts actual Table facts, including nested physical tables, rather than facade projection count (`crates/rdocx-py/src/document.rs:5063`). word_block_owners and detached table_owners compute it once from prepared facts, while plain_paragraph_owners records0 (`crates/rdocx-py/src/document.rs:5347`, `crates/rdocx-py/src/document.rs:5390`, `crates/rdocx-py/src/document.rs:5420`). The local decision reads the cached integer, so it adds O(1) eligibility work rather than scanning each selected model on every setter.

Examined all piece publishers: full assembly builds fresh owner facts or reuses aligned qualified Arc pieces, splice packets supply constructed pieces, missing/misaligned caches fall back, general install consumes prepared pieces or clears to its existing empty fallback, and local commit replaces exactly one Arc after revision commit. Candidate properties/text cannot introduce another table. No additional stale-count or cache-invalidation defect was established. Flat issue tables retain eligibility and the unchanged flat bounded-work controls remain green. The original nested-table witness now selects checked publication before local owner/index/revision installation.

Actual frozen peer1 artifact records four failures with held outer/inner text exchanged while saved/reopened bytes are correct. Actual Base passes the same four cases. Repaired artifact passes them within the affected21-case batch. BEFORE_REPAIR_TEST_SOURCE.py hashes to the recorded earlier test source. The final file additionally appends the twelve ancestry controls, with no change to those selected21 cases. Their final source is separately pinned. This establishes real author-remediation RED/Base/GREEN, not a retroactive runtime claim for peer1's source-only finding.

The additional twelve control/field/textbox cases pass separately on Base, frozen and repaired artifacts. Per-cell paragraph projection ordinals distinguish multiple same-depth paragraph controls and field-bearing paragraphs. The existing nonzero text_boxes guard selects checked fallback for textbox owners. These passes are valid positive controls with no additional RED. They leave the distinct controlled-row/cell case in D1 uncovered.

## Other contracts and evidence limits

The unchanged local preparation/last-fallible-commit ordering, publication-mode owner checks, generation/route/parent deltas, distinct CellText/ParagraphText retirement and read/piece publication were reexamined with the count change. Dormant authority still discards all validation facts/pieces, retains exact eager checks and always validates current native state on activation. Construction, ten guarded factories, numeric publication, unbound hyperlink removal, direct URL/image publishers, serialization and consuming replacement zero/mismatch/positive branches retain their prior authenticated implementations. No additional correctness, panic/input, OOXML preservation/order, structure/dependency or HLD-impact defect was established.

The current cfg(test) C9 binding-object fault test actually runs1/1 PASS on repaired source, with unchanged bytes/serial/held cell and a successful retry. This supersedes the earlier compiled-only/current-source uncertainty for that diagnostic. It remains a test-profile binding-object injection, not installed-wheel injection. No installed fault transport is exported. Root final must preserve that evidence distinction when assessing formal coverage.

All thirteen A-IDs and original numeric workloads remain. The affected21 and separate12 observations are scoped macOS arm64 dev CPython3.14.7 semantics/debug work. Previous six native/support tests have unchanged source inputs and are reused, not newly rerun. Clippy/fmt/changed prose receipts are green. These observations do not establish original release timing or Office behavior.

Exact Linux x86_64 release wheel/CLI qualification remains pending: bodies0/500/2000 with20x5 table and100 reused handles, text/alignment/shading in order without warmup, three distinct setters below1.0ms and at most1.5times empty-body plus0.05ms. Case2 retains4000 three-run paragraphs,1000 water replacements, three paired trials, lexicographic minimum Python total/open/replace/save tuple and minimum matching CLI time including process start, ratio at most1.2. Counters and reduced workloads substitute for none of these criteria.

Installed3.9/3.12, strict typing/stubs, schema/native Word/applicable PowerPoint, hash49 unchanged, retained F26879/252/407/five-linear/count/memory/PowerPoint248 riders, WASM, full integrated union, docs, archives/publication dry-run and hosted CI remain due. No F272 raw-admission/provider acceptance follows from this Base. No baseline, threshold, public bridge inventory or final publication approval is waived.

Review ends with this report. No remediation performed. Review-directory lease returns to Root. STOP.
