# F-X283, Codex implementation, pass 1

**Reviewed**: Frozen working diff against Base and HEAD `101c34c4a6d5fa62d9ba0e0ff745fd44c9aeaf13`, branch `work/f-x283-codex`. Twelve files, 2596 insertions and 327 deletions. Diff SHA256 `f8fa946e2c6ad3adf9ebeba1b71e390ce3361e32c9e836a60b736ee75bea9d12`.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. Not clean. This is actual independent implementation peer 1, not a Root final or acceptance pass.

## Scope and authentication

Read canonical AGENTS.md, CLAUDE.md, WORKFLOW.md, microscope, the approved operative plan and paragraph-publication amendment, cited HLD sections, Base producers and all changed implementation/test/HLD files. Confirmed F-X283 is in-progress in both delivery records. Worker plan hash `c85a318e124518f97e52b7c5f6f364567066b249af7bbec59957af92e11000da` and separate plan patch hash `30daf0907711ba131a153830d8f8dd74a9c58e753f6587bb68c112da6197a811` match. Historical draft language does not supersede operative approval.

Authenticated all twelve changed-file hashes, both patches and the actual current git diff against FINAL_SOURCE_MANIFEST.json. All 311 REVIEW_BUILD_SOURCE_INPUTS.json inputs match current files. The installed extension hash matches `d9884b17b0bb0d0e26fedcf06b61ff282dda41c75862a33e732cd750a6e42656`, and its build manifest hash matches. All 63 indexed evidence-log hashes match. Identity is not correctness by pass count.

Reviewed native local candidate/commit and consuming replacement, certificate creation/clone/open/main-source/flush custody, local matching/index/generation/retirement/piece publishers, optional authority construction/initialization, ten first factories, numeric structural publishers, unbound hyperlink removal/retarget, direct images and serializers. Comparison consumers include existing full publication and property relocation, plus named nested/related/direct-ordinal and Font controls. No tests, builds, network operations, Git mutations or source edits were performed during this pass. Only this report was written.

## Defects

### D1, high, local matching transfers authority between nested and enclosing cell paragraphs

`crates/rdocx-py/src/document.rs:11724`

The local matcher keys paragraphs and runs by property_owner_key plus paragraph projection ordinal. InCell keys contain story kind, part, path depth and nearest row/cell coordinates, but omit physical cell-story and enclosing table/cell identity (`crates/rdocx-py/src/document.rs:7907`, `crates/rdocx-py/src/document.rs:7926`). Paragraph ordinals restart per cell story (`crates/rdocx-py/src/document.rs:5201`). An enclosing row0/cell0 and nested row0/cell0 therefore give their first direct paragraphs the same key. Their first runs also collide.

Previous is collected by LIFO traversal. Table children are pushed after paragraph children, so nested descendants precede their enclosing cell's paragraphs (`crates/rdocx-py/src/document.rs:11681`). New addresses retain native physical owner order, enclosing cell before nested cell (`crates/rdocx/src/document.rs:32539`, `crates/rdocx-py/src/document.rs:5192`). The bucket cursor joins different physical owners by occurrence (`crates/rdocx-py/src/document.rs:11737`). Equal cardinalities pass the disappearance guard.

Concrete trigger: an ordinary direct-body one-row two-cell table, with no nested namespace declarations or links. Cell0 contains a direct paragraph `outer` and a nested one-cell table whose paragraph says `inner`. Cell1 contains an ordinary paragraph. Hold both cell0 paragraphs and their runs, then assign cell1.shading or cell1.vertical_alignment. Native qualification excludes nested content only in selected cell1 (`crates/rdocx/src/document.rs:35679`). Publication walks the entire table. Its completeness guard does not reject this topology, because unprojected nested cells are skipped without setting complete=false (`crates/rdocx/src/document.rs:20919`).

Property publication preserves the mismatched states and installs their new locations/read facts (`crates/rdocx-py/src/document.rs:11850`, `crates/rdocx-py/src/document.rs:11963`). Held outer/inner authority is exchanged. The cached Paragraph.text getter exposes the other paragraph's text through that reassigned generation (`crates/rdocx-py/src/paragraph.rs:950`). Text publication in clean cell1 has the same unrelated-owner problem. Native bytes can remain correct while held views address another physical owner.

This violates A283-9 and prepared physical-owner identity. Join through proven containing-owner identity, or take checked fallback for ambiguous selected-table topology before publication. A discriminator must edit a clean cell while holding enclosing and nested neighbor paragraph/run/property views with distinct values. Existing nested-target fallback and separate-direct-table controls do not cover this combination. This is source-traced, not an executed runtime failure in this pass.

## Smells

None found.

## Nitpicks

None found.

## Acceptance evidence and limitations

All thirteen A-ID rows remain. A283-1/2/3 have installed semantic/bounded-work controls, including actual comparisons and Document.new with unrelated growth, using reduced local bodies and a 2x2 table. Their original three Linux release timing passes remain pending. A283-4 retains ordered counts, zero/mismatch/positive behavior, one dormant serial commit and current-model validation. Its exact 4000-paragraph same-build wheel/CLI ratio is pending. Counters substitute for neither workload.

A283-5/9 distinguish CellText, ParagraphText and properties across absent/empty/nonempty tcPr, selected/later handles and save/reopen. Intermediate C17 failures and Base policy passes remain separate from the corrected artifact. D1 is an uncovered composition, so these passes do not close matching.

A283-6 covers signed truncation and i32 wrap boundaries. A283-7/8 consume installed border/margin/color refusal and round-trip controls, plus supplemental local/fallback diagnostics. A283-10 has native candidate refusal and installed ordinary refusal/later-success evidence. The current internal commit fault is at `crates/rdocx-py/src/document.rs:20201`. Earlier supplemental Rust diagnostics executed a fault, but the frozen installed artifact exports no hook and the final native/support run does not execute this binding test. A current installed commit-fault injection is not established. Retain this distinction before acceptance.

A283-11/12 compare direct/controlled cross-cell comments, unchanged comments-part bytes and native checked/local bookmark, note/revision-bearing run, opaque property and saved-byte behavior. These are library comparisons, not Office qualification. A283-13 covers eager missing-link refusal, transient discard, fresh activation, ten factories, current-serial initializer, unbound links, images and failed-stream serialization. Internal invariant assertions and installed public observations are distinct evidence.

Actual installed observations are 52 selected cases plus four existing controls on a macOS arm64 dev Python3.14.7 extension. Fresh native/support evidence is five native tests and one support test. Assertions and inputs were inspected. Historical Base RED retains its earlier counter meaning. Reconstructed old test-source provenance is not a contemporaneous hash receipt. Invalid namespace probes and the table-row typo remain test defects, not product RED. The valid actual-flush certificate witness and fixed-source identity are separate. Historical artifacts do not qualify release performance.

Installed Python3.9/3.12, strict typing/stubs, exact Linux release workloads, schema/native Word/applicable PowerPoint, hash49 unchanged, retained F-X268 79/252/407/five-linear/count/memory/PowerPoint248 riders, WASM, full integrated union, documentation, archives/publication dry-run and hosted CI remain pending. No F-X272 provider or raw-XML admission acceptance is inferred from this Base. No commit, integration, completion or publication is approved.

## Not found

No additional defect found in dormant fact/piece absence, eager validation/error timing, fresh activation, guarded factory borrows, consuming zero/refusal/positive replacement, certificate allocation identity/current-output classification, distinct selected-publication retirement or last-fallible-commit ordering. No additional panic/input-boundary, OOXML child-order/opaque-byte, structure/dependency or HLD-impact finding found. These statements are bounded by inspected paths and pending gates.

Review ends with this report. The author remains stopped until Root reassigns the sole writer lease. No remediation was performed.
