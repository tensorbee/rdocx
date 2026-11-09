# F-X276, Codex final microscope, pass 2

**Reviewed**: the uncommitted working tree on `work/f-x276-codex`, Base and HEAD `9a3b2960ccda7c3208ecb9b753c7019b64c086f5`. Current tracked scope is 13 files, 2111 insertions and 133 deletions. This is a targeted continuation of the separate Root final phase after three completed DeepSeek implementation peers, not another peer or reset. The prior whole-story scope and dispositions are recorded in `F-X276-codex-final-1.md`.

**Contract**: approved plan SHA-256 remains `d414f1514710f450a473cd70de025f8fa54c2241e941b40f866ae2d7af7ee3b4`. Branch, HEAD, current progress acknowledgment and prior final report were checked. Root source writes were stopped for this report-only pass.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Final1 S1 is resolved in the inspected implementation and applicable local evidence. This is a clean local microscope outcome, not feature completion, integration or final sprint acceptance.

## Defects

None.

## Smells

None remaining in the inspected scope.

## Nitpicks

None.

## S1 disposition and composition

`crates/rpptx/src/lib.rs:311` now stores `Option<Arc<AutofitTable>>`. The table is allocated once outside the cell loop at `:27787`, and each cell context clones that immutable Arc at `:27888`. The full retained table payload is shared within that resolved table and consuming context, in both read-only and mutating inventories. This closes the newly introduced per-cell deep retained copy.

The local cache at `:4109` keys the result by table allocation identity and page number. The miss path at `:4198` evaluates grown geometry and retains an Arc with the result. Retention prevents allocator address reuse while the cache is alive. Distinct physical tables on the same page remain distinct, and separate resolved contexts have separate allocations. Page-number-sensitive text is not merged across page numbers. No global cache, unsafe dereference, trait, generic parameter or public API is introduced.

The exception still runs only after a successful supported non-empty measurement with stored width fit and stored height failure. Unsupported and empty paths still return before it. Stored width failure still overflows. Both hypothetical trial values remain governed by the unchanged candidate selector. The grown-table helper retains the shared painter growth, transformed rectangle and existing epsilon. An error propagates before cache publication. Reusing the result does not change saved XML or mutation policy.

The preexisting transient shape clone in `authoring_table_cell` and per-cell source fingerprint work are not claimed linearized. S1 addressed the new retained full-table duplication and repeated whole-table growth, not every cost in the historical inventory.

## Discriminating regression and recorded results

Inspected `issue431_table_contexts_share_storage_and_grown_measurement` at `crates/rpptx/src/lib.rs:36684`. It creates two physical tables on one slide with sixteen total cells. It counts retained table identities independently in mutable and read-only inventories, then counts actual calls to the growth helper. The expected tuple is `([2, 2], 2)`. The tables have different on-page outcomes, and all sixteen cell results are asserted against their physical table identity. Package bytes must remain equal.

This rejects sharing without memoization, memoization without sharing, and a page-only cache that mixes the two tables. The counter is incremented at the actual helper entry at `:27631`. It is not a constant inferred from the expected number of tables.

Read and authenticated the named logs and receipts under `/private/tmp/rdocx-s96-root-f276-repair`. Every inspected log hash matches its receipt. The corrected old-source attempt2 is genuine semantic RED, child101 and wrapper101, with actual `([16, 16], 16)` versus `([2, 2], 2)`. Attempt1 instead failed on nonpositive usable fixture bounds and is not resource RED. The post-repair Issue431 batch records seven passing tests, child0 and wrapper0. It retains the prior geometry, pair, width, duplicate and empty-notes controls alongside the new resource case. Four-crate all-targets/all-features clippy with warnings denied, fmt and prose each record child0 and wrapper0 with agreeing log content. No execution counts are summed into unique coverage.

## Source and Python artifact identity

Rehashed all 326 tracked inputs in `/private/tmp/rdocx-s96-root-f276-python-table-repair/BUILD_MANIFEST.json`. All match current source. Comparing current inputs with the preceding build manifest identifies only `crates/rpptx/src/lib.rs` as changed. The other 325 build inputs, including Markdown, renderer, CLI and binding sources, remain equal. Prior unaffected evidence and the prior whole-story review remain applicable within their stated limits.

The actual imported extension hashes to `e88cd792ee94cb109288f1ea253b4eab3c92a9fb5d6f65e3941995a1240ad73c`, matching the new manifest and runtime log. Inspected the new focused runtime script and binding-control script plus their recorded output. Python3.14.7 reports three selected tests passed with 1031 deselected and cache disabled. The separate binding controls pass both empty-notes variants and both public Markdown punctuation markers, with source preservation. Both receipts record child0 and wrapper0. This fresh evidence supersedes the pre-table extension for affected production code. It establishes neither installed wheels nor the final Python version matrix.

## Aspects and retained boundaries

Correctness, contract, panics, OOXML, tests and structure produced no new finding. The cache uses pointer identity only as an opaque hash key and retains its owner. No new XML serialization or schema ordering path is added. Both inventories use the same shared representation. A1 through A5 retain the functional outcomes and local coverage adjudicated in final1, with A2 now adding the deterministic resource discriminator. Earlier stale-accessor and admitted-pair contract corrections remain operative.

This continuation inspected the changed representation, inventory assignment, fit-report cache, growth helper instrumentation, complete new regression, progress, prior report, manifests, scripts and named logs. It did not repeat a full review of unchanged giant files or replay author histories. No tests, builds, benchmarks, network, Git mutations, native control, cleanup or source edits were performed by this reviewer.

Full integrated union, hash49/pixels7, native/PDF/PNG/LibreOffice, installed Python3.9/3.12 typing/stubs and archives, hosted CI, full public Markdown export benchmark and release approval remain pending. No native PowerPoint promise is made. The reviewer stops after writing this report and returns the sole writer lease to Root.
