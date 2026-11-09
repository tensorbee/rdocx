# F-X279d, microscope, pass 2

**Reviewed**: Frozen working source against `ff246c8b8b76ef573e8438b582259ba7045124ba`, with pass 1 as the predecessor. The complete working diff spans 12 paths, 1,034 added and 20 removed lines, including the separately reviewed C2 cleanup neighbor and the approved D1 plan clarification. All 12 hashes match `REVIEW_SOURCE_FREEZE_2.json`.

**Contract**: Approved F-X279d authoring and initials story, including its exact D1 projectable-reference clarification. This is pass 2 of the original review cycle. No comparison, import-composition, C1 provider or final native acceptance claim is made.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. D1 is closed in this bounded implementation.

## Defects

0 defects.

### D1 closure

`crates/rdocx/src/comments.rs:1920` obtains the selected physical parent end and reference from the already qualified source inventory. The helper rejects a parent end after its reference. At `crates/rdocx/src/comments.rs:1945`, the reference must also occupy the parent's actual projected end location and run boundary. This independently excludes an early reference and a separate intervening text run. The new id is allocated only after these proofs, at `crates/rdocx/src/comments.rs:1961`.

`crates/rdocx/src/comments.rs:2333` checks the entire selected reference run after removing only the selected parent reference from its clone. Typed nonempty Text, Fields, other content and alternate drawings refuse. Root-attribute records and empty typed Text remain supported. When retained raw children need classification, the existing run writer supplies their real retained namespace context. The resulting check accepts only Word run properties and empty Word Text, while unrelated opaque children and foreign Text refuse. This closes the mixed-run suffix and opaque-payload modes without introducing a splitting grammar or changing imported read/save shape.

The physical start-adjacency splice remains on the same staged authoring path. Existing parent markers and payload stay in place. The caller still commits only after successful authoring and related-story flush, so refusal discards the candidate rather than changing the original document or Python revision.

The new recorded predecessor discriminator exercises four separate bad shapes: early reference, suffix Text in the reference run, intervening Text and opaque reference-run payload. The exact predecessor extension records four failures of the intended refusal assertions and two formatted-reference positives. The final extension records passing refusal and positive cases. The earlier first-guard failure on the supported empty-Text positive is preserved separately and is not presented as final green evidence.

The local-run alias fixture remains a compatibility limitation of the existing namespace transport. The archived predecessor also refuses it atomically. It is not treated as a newly supported form, a repair regression or authority to widen namespace policy. Root-level alias and formatted producer-attribute positives remain covered.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Whole-story compatibility readback

The unchanged public Rust shims, additive initials method, Python keyword and stub, CLI flag and atomic publisher retain their pass 1 caller flow. Omitted, empty and nonempty initials remain distinct. The existing controls still cover repeated and selected-parent replies, reply-to-reply behavior, immediate physical start adjacency, separate references, multi-paragraph and related-story endpoints, supported control and table-cell owners, save/reopen, reply removal, thread resolution, protected movement, imported opaque preservation and missing or ambiguous range refusal.

The D1 clarification narrows authoring to the exact supported endpoint and preserves refusal for nonprojectable imported arrangements. It does not reinterpret imported point, orphan or rangeless replies. The separately reviewed C2 cleanup remains a neighbor, not acceptance of comparison composition or completion of its formal provider obligations.

## Evidence and limits

The final Native comments log records 31 passes and 1 ignored test. The public Rust regression and actual CLI test each record 1 pass. The fresh Python log records 30 passes, including the four bad-shape refusals and formatted/empty-Text positives. Their assertions check unchanged imported bytes and live held views on failure, as well as preserved producer attributes, anchor text and successful reopen on supported input.

The actual final extension hash is `56919a62e0d2ef8d5472bfbf7bbf5ff87f7954a5ed3e960dcda96729e3c558c5`. It matches `D1-metadata-python-artifact.json`, whose Native, binding, stub and test-source pins match this freeze. The source pin check matched all 12 paths. Recorded strict all-targets, all-features Clippy completed successfully for rdocx, CLI and Python. The final formatting and prose receipts are retained alongside the other metadata results. No reviewer build, test or oracle execution was performed.

The historical self-hash test and constant remain unchanged and were excluded only from the local Native cohort. They are not newly ignored or claimed passed. Fresh final Word execution, Root final review and integration, comparison/resolution composition and the relevant integrated gates remain pending.

## Not found

No remaining correctness or contract defect was found in this bounded authoring repair. No new panic, OOXML preservation, namespace, schema-order, binding or structural-rule finding was found. The supported-reference classification and endpoint checks address the pass 1 fault without a new public type, parser API, quota or comparison implementation.
