# F-X279b, all, implementation pass 2

**Reviewed**: Frozen working candidate against Base and HEAD `391280577110c9b7b92ace0420a4a1d98c707889`, branch `work/f-x279b-codex`. Thirteen tracked changed paths including the operative plan, 2543 insertions and 85 deletions. This is implementation pass 2 of the original cap of three.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Pass-1 D1 and D2 are closed. Source review is clean. Fresh successor Python execution and final integrated gates remain pending, so this is not completed feature or whole-union acceptance.

## Source and scope

The exact candidate is `/private/tmp/rdocx-s96-f279b-evidence/peer2-frozen/source`. `SOURCE_PINS.json` SHA256 is `a40ca895d24e25fe5f966131dbd884bf8b327f0e91971b2b56056b40faa5b3d9`. All 16 source/configuration/plan snapshots and working files authenticate. The three `CONTEXT_PINS.json` entries authenticate the plan, progress and unchanged pass-1 report. Every source and context entry was checked again after the review and remained unchanged.

Revisited the complete B diff against Base and compared each successor file with the authenticated pass-1 snapshot. Read the current approved plan and readiness amendment, progress, microscope/workflow requirements and affected HLD03/04/10. The operative A/C/D exclusions and the genuine integrated F272 provider remain as recorded in pass 1. This pass does not restart their reviews or reset B's original review cap.

## Pass-1 closure

### D1 closed, actual reply replay supplies both model and Source

`crates/rdocx/src/comparison.rs:972`, `crates/rdocx/src/comparison.rs:1132`, `crates/rdocx/src/comparison.rs:1140`

The graph-only qualifier now returns the actual original-plus-reply replay with the added identities. The caller installs that replay as its edited operand before relationship remapping, XML extraction and comparison Metadata construction. Thus later direct body/related publication consumes the replay's actual model and Source together. Ignored body formatting, whitespace and related-story values come from the original, while the newly authored reply graph is retained. A failed qualifier continues through the existing comparison path. Existing revised-comment snapshot and protected-owner guards remain.

The new regression at `crates/rdocx/tests/regression_test.rs:413167` combines a crossing reply with ignored bold-to-italic changes or whitespace changes, including related header whitespace and an unrelated protected revision. It checks selective accept and reject, original body/header values, original formatting, graph counts and ranges. Its assertions distinguish the old normalized-proof/direct-edited-copy error. The corresponding Python controls cover both ignored modes. Recorded actual predecessor Rust and authenticated Python failures are retained, rather than inferred from the passing successor.

### D2 closed, typed table/row control collections share concrete traversal

`crates/rdocx/src/field.rs:5306`, `crates/rdocx/src/field.rs:5318`, `crates/rdocx/src/field.rs:5331`

Table traversal visits both direct rows and table-owned controls. Row traversal visits both direct cells and row-owned controls. The same concrete row/cell functions are used for their SdtContent variants. The actual enclosing table-style context is forwarded to paragraphs and nested block controls, while entering a nested table selects that table's own context. The change does not include the excluded inline-run control, revised raw payload or conditional-table-formatting branches.

The regression at `crates/rdocx/tests/regression_test.rs:413261` first proves actual typed row/cell control delivery. It covers row SDT, cell SDT and nested row/cell/block control shapes, source docDefaults28 versus Normal22, source-effective bundled rendering, saved/reopened insertion and unchanged destination formatting. The source family's deterministic substitute is compared with the actual source result, with stored Georgia checked independently. Python additionally inspects the saved selected style or direct properties. This closes the omitted real-owner branches without changing style precedence or the approved font frontier.

## Whole-candidate checks

Rechecked coalesced equality and marker/shell boundaries, physical-owner logical-word alignment, paired nontext fallback, ordinary ignore-formatting shell selection, separate edited content Source, and one enclosing word revision with every original owner still consumed. No new finding in the distinction between ordinary comparison and F272 paired leaf capture. The earlier six actual Source failures and L1 investigation remain historical observations, not discarded test outcomes.

Reviewed native body-only insertion, Python full companion closure, exact insertion-policy constructors and unchanged direct import defaults. Effective source defaults remain limited to ASCII/high-ANSI family and half-point size. Table, paragraph, character and direct property precedence, source-style selection, destination-default preservation, restart versus explicit numbering reuse and saved authoritative fragment bytes remain consistent with the contract. No new public signature, stub, module, trait, generic abstraction, quota or Source authority is introduced by the repairs.

Checked reply snapshot rejection against the actual added identity list and existing staged resolution. P1/P2 success does not waive revised-comment replacement or protected-body refusal. HLD04 describes the actual replay producer and HLD10 names typed row/cell controls within the existing body-font promise.

## Evidence and pending execution

`peer1-green/RECEIPTS.json` records actual public Rust18, Native29, CLI4 plus9 and strict all-target/all-feature Clippy exit0. Its source pin map equals the frozen pin map exactly. Retained runtime binaries were independently hashed against their receipts. The public assertions include both repairs, logical-word splits on both operands, ignore-formatting left bias, token neighbors, protected-owner behavior, P1/P2, deterministic list SVG and source-effective fonts after save/reopen.

The predecessor Rust evidence preserves actual failures for the ignored-value and control-font cases. The corrected font discriminator proves the typed source owner and distinguishes stored Georgia from bundled Gelasio. The lost initial fixture compilation log is explicitly recorded in `peer1-red/PRODUCT_APPLICABILITY.json` and is not represented as retained semantic RED. `python-provider-gate/run-vbwpqhqi/PEER1_PREDECESSOR_RUNTIME.json` records three actual new discriminator failures on the authenticated predecessor module, not on the repaired producer.

The Python equal-split assertion now proves original paragraph bytes, zero revisions and the fresh value. If the package changed, it permits only the observed inherited namespace closure on sectPr and requires established changed-publication staleness. If bytes did not change, it requires the held handle to remain usable. This follows the Root-adopted paragraph versus whole-operation distinction. It does not relax F272's genuine whole-operation no-op or refusal controls. Four existing F272 installed nodes, including exact-byte and live paragraph/run no-op assertions, are selected in the fresh successor run.

The companion test now proves delivery of a supported inline chart reference before import and resolves absolute OPC targets only for ZIP lookup. The predecessor diagnostic for that valid fixture is distinct from the original invalid-fixture failure. No Python runtime claim is manufactured from the Rust or source review.

At report writing the fresh optimized successor wheel and its 17 selected Python nodes are still pending actual results. Earlier six Source cases retain their separately qualified per-case unchanged-branch applicability, not a new execution of those six on this successor. Full integrated suites, hash measurement, installed compatibility/typing matrix, WASM, packaging and fresh native comment-candidate acceptance remain final riders. This reviewer ran no Cargo, executable, network or oracle operation.

## Defects

0. D1 and D2 from pass 1 are closed as described above.

## Smells

0.

## Nitpicks

0.

## Not found

No additional correctness, contract, panic/indexing, OOXML schema/namespace, testing-structure, binding/stub or repository-structure finding in the full B candidate. Pending required execution is kept separate from source acceptance and no final provider, sprint or publication claim is made.

The pass ends here. Only this review file was written, with no product, test, progress, plan or HLD changes by the reviewer.
