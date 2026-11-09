# F-X272, imported payload implementation, Codex pass 1

**Reviewed**: The frozen imported-payload increment in `/private/tmp/rdocx-s96-root-c1-implementation/word-grammar-carriers/imported-payload-increment`. Six before/final file pairs, 772 added and 50 removed lines, and the complete 1061-line unified patch. This is a bounded implementation review, not a whole C1 or provider review.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the reviewed increment.

## Contract and source boundary

Read canonical `AGENTS.md`, `CLAUDE.md`, `.claude/WORKFLOW.md`, `/microscope`, the approved F-X272 source/format/rewrite signatures, the finite imported XML payload and publication-coverage contract, and the positive ordinary root-event clarification. F-X272 remains in progress. The HLD03, HLD04 and HLD10 changes describe the bounded implementation and retain its exclusions.

The reviewed source is the sealed snapshot, with these identities:

- Source boundary: `0ba51dbe6a91245a9bdeb1bc50b4bde03fec94fceeab0637c7efd5a5d3571fbc`.
- Patch: `b74d7d8d4fb8c2e43918644674f2c491020aaee81181dd499c090c3de6cee664`.
- Final document source: `eba4909de470f9ebb706bdc85b707a53a80f849f7c2901123cbdb911f761cc11`.
- Final text source: `83fc01bc7aa26203fe79201cb5daf32ce76de6d934d8cc7bc0b17f2f9bb8d11c`.
- Retained test executable: `9799476ac4f42e70ec599fd62d0e1f737215e64d7d29dcdf3544df3e4b944af5`.

Independently checked all six before hashes against `BEFORE.json` and all six final hashes against the sealed boundary. Reconstructed the patch exactly from those pairs in the manifest's file order. Authenticated the retained executable, listed logs and four preserved range-cost RED artifacts. The after-run seal does not supply an invented pre-run workspace inventory.

## Defects

None found in this bounded increment.

## Smells

None found in this bounded increment.

## Nitpicks

None.

## Implementation examined

- `crates/rdocx-oxml/src/document.rs:4985`: classified sources reuse their actual constructor packet. Unclassified payloads and changed records invoke admission only at the real Word reader. The source, paragraph and wrapped-row consumers preserve their respective existing parser boundaries.
- `crates/rdocx-oxml/src/document.rs:5009`: positive Field coverage may answer directly. Ordinary classification uses the combined mapped and current facts with Field priority. Unknown and mixed owners are not inferred ordinary.
- `crates/rdocx-oxml/src/document.rs:5037`: the private receipt holds a Weak parent occurrence, nonempty owner bounds and exact optional Field prepayment. It retains no parent source/model buffer and cannot be reconstructed from matching bytes or lengths.
- `crates/rdocx-oxml/src/document.rs:6070`: the shared assembly records actual prefix, payload and suffix pieces. Ordinary composition still admits the complete piece view before destination byte allocation. Imported rewrite deliberately leaves XML admission to its later consumer. Materialization creates a fresh occurrence and drops payload receipt authority.
- `crates/rdocx-oxml/src/document.rs:6602` and `crates/rdocx-oxml/src/text.rs:4539`: scalar production validates its actual operation and owner before exact reservation and allocation. Raw output is not an XML validator. The existing five-character escape policy preserves carriage returns and the existing raw-byte behavior. Field demand does not depend on whether another consumer activated the metric flag. No input or template-copy credit is created.
- `crates/rdocx-oxml/src/document.rs:6649`: rewrite validates parent identity, operation, edit geometry, owner containment, current role and payload length before reservation or output bytes. Sorted nonoverlapping edits make total selected Field bytes minus removed selected intersections equivalent to retained Field prefix/suffix demand. Only the exact moved Field payload avoids a second debit. Empty values retain a nonempty owner witness. Empty insertion is strictly interior, while an empty edit list moves the original source unchanged.
- `crates/rdocx-oxml/src/document.rs:534` and `crates/rdocx-oxml/src/document.rs:669`: numeric owned ranges normalize before binary selection. Slicing and concatenation preserve their sorted, unioned representation. Physical publication pieces and capture windows remain distinct, including overlapping collector seams. No binary selection assumption was added to those overlapping tapes.
- `crates/rdocx-oxml/src/document.rs:6755`: repetition copies constructor classification state without manufacturing a payload receipt. Generic composition, closure and changed format outputs likewise do not duplicate prepayment. Copied Word origins remain inert correspondence, not selected input or new allowance.

No new trait, feature, product module or public receipt type is introduced. The concrete assembly has two real policy consumers and owns substantive byte/map materialization. Arithmetic, indexing, locking and the new Option branches were examined without finding a new panic or lock-across-product-write path.

## Evidence examined

Read the inputs and assertions of all 12 new `scoped_import_` tests, starting at `crates/rdocx-oxml/src/document.rs:27190`, rather than relying on their aggregate count. They distinguish normalized ranges without filling gaps, selected-query cost, out-of-order edits, interior and boundary insertions, raw versus escaped output including Unicode/CR/NUL, deferred admission, retained-copy demand after deletion, equal-byte production with new parent identity, rejection of copied receipts, no-op moves, foreign anchors, unknown owners, overlap and sticky preallocation refusal.

The final raw log records 12 named passes. The six related logs contain 76 distinct named passes across ordinary roots, document factories, publication, Word membership, repetition and namespace closure. These are 88 distinct local controls on the retained executable. Earlier 8-test and 9-test logs are historical subsets, not additional passes.

The preserved range-cost RED has exit 101 and reports 65536 examined owned ranges for 256 selected slices. The final test bounds the actual numeric range predicates. It does not claim an end-to-end Native importer complexity measurement or qualification of every physical collector query.

The retained strict Clippy log shows successful lower and Native all-target/all-feature completion in 21.60 seconds. This is ordinary compilation and lint evidence, not Native runtime evidence. Formatting, prose, adapter drift and whitespace checks are reported as zero by the author session. Their raw metadata receipt was not supplied, as the seal explicitly records. This reviewer did not rerun builds, tests or metadata gates.

## Scope limits and unavailable criteria

The two lower methods are reviewed as callable implementation seams with explicit protocol witnesses. Rejecting an unproved owner in these controls is not evidence that routing a valid Native import through an unproved witness is compatible. Actual Native scanner handoff, scan/error order, complete ordinary and Field owner closure, mixed-owner producer subranges and the named Native import writers remain pending.

Native clone/preparation lineage, deferred and nested current-role producers, complete constructor/context pre-growth qualification, external compile/SemVer reconciliation, serial integration and all final union, resource/history, schema, hash, installed and Office acceptance remain outside this review. The local range test does not waive remaining metadata or caller complexity work. No whole C1, provider checkpoint or sprint completion follows from this verdict.

## Not found

No additional correctness, contract, panic, OOXML byte/namespace policy, test-assertion or structural finding was found within the frozen delta. No rendering, runtime Native parity or final acceptance claim was made or adopted.

Review ended after writing this report. No source, progress, plan, tracker or build state was modified by the reviewer.
