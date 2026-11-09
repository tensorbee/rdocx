# F-X272 body seed membership implementation review, pass 1

Result: zero defects, zero smells and zero nitpicks in the sealed increment. This review qualifies neither complete Native input lineage nor C1/provider completion.

## Scope and custody

Reviewed only body-seed-membership-increment under the existing word-grammar-carriers evidence directory, against canonical32044373 and its approved seven-argument constructor. SOURCE_BOUNDARY.json SHA256 is `1aac0490fce698fd5ea583578804e6fa0122ab12e7d08aaa7cbe064437b547da`. The exact six-path patch is `55ce0247180fcca83a345d239da798436b4ec7ab624bb3c0f94f008d77232676`. Read CLAUDE.md and WORKFLOW.md. Moving worker source was not used as authority for the changed implementation.

Authenticated all six before/final snapshot pairs and applied the entire frozen unified patch strictly in memory, checking each context/deletion and full reconstructed final bytes. Authenticated the retained final binary, GREEN log, related receipts/logs and both controlled RED source/binary/log sets against the boundary. Final lower document.rs is `394f0457cf040eb078cc401dc25bd5f5cb2d8a8f6c22aab0207ccdb6105a15ac`, Native template.rs is `0f1241d58fc68db510ad70292ae6924afc4b9c5332a2bc222e070f1cf1345ae6`, and content_control.rs is `20cce6da8f789f8521a8092ff1e11837a9fb44107472e041da94f24ffbb76585`.

## Implementation checks

The frozen lower `crates/rdocx-oxml/src/document.rs:5187` adds actual borrowed bodies immediately after paragraph seeds. The remaining callable order is preserved. At5208 the constructor combines Body, already supplied paragraph and Field roots, then expands the existing physical resource graph. All existing call sites in this patch receive the new argument. This is membership enrollment, not a new public owner type or admission API.

The private collector at `crates/rdocx-oxml/src/document.rs:5265` deduplicates the concrete tagged owner key before inspecting leaves or expanding children. Existing WordOwner::key at1118 distinguishes owner variants and their live borrowed identities. The graph at1149 reaches stored Field owners, revision content, Ruby, current controls and retained original control runs. Inert enrollment does not call a display selector, normalized paragraph projection or parser. Reference catalogs and work-list storage are distinct from scalar/model materialization.

At5291 actual paragraph membership and reached raw records are borrowed. Run leaves at5302 enroll both Text and DeletedText by their actual CT_Text identity. Field enrollment at5311 keeps the actual frozen Field for its existing private source/current relationships. Body/table/row/cell/control/revision/drawing raw records are borrowed at5279 through5333. These explicit raw sites do not certify every property or root record in OOXML. The approved scope still leaves records outside this traversal and actual relevant consumer joins pending. Membership alone changes neither writer/cache eligibility nor input allowance.

Control retention at `crates/rdocx-oxml/src/document.rs:5316` consumes one existing-storage iterator. The new crate-private iterator at `crates/rdocx-oxml/src/content_control.rs:1347` walks inline_run_sources once and returns each actual absolute content index alongside its borrowed raw record. The public indexed getter at1335 retains its original find/map behavior, adding only a cfg(test) probe increment. It still neither chooses writer eligibility nor interprets the index among only Run entries. Dormant retained captures are enrolled inertly, with later selection authoritative.

Native `crates/rdocx/src/template.rs:1353` supplies the actual main body and each glossary body before clone_for_staging at1376. It separately borrows glossary/docpart retained root XML. Existing note paragraph collection and existing paragraph visitor behavior are unchanged. The operation still flows through the preexisting evaluation route. At1582 the evaluation borrow is explicitly dropped before commit_template at1583, and no later operation use follows. This does not implement TemplateWordLineage or derive candidate correspondence from the newly enrolled original identities.

HLD03 at3318 and HLD10 at5161 accurately describe physical membership, seven arguments, read-visitor compatibility and lack of funding or clone correspondence. The final external v5 client is preparation only, not a compiled client result supplied by this packet. Its actual compilation and the strict public API inventory remain final obligations.

## Evidence and limits

The retained final log contains29 distinct named successful lower tests, including three new body-membership cases at `crates/rdocx-oxml/src/document.rs:23209`,23260 and23361. The assertions cover overlapping body seeds, current versus stored original control Text, absolute retained capture index, nested stored Field Text, revisions/Ruby, zero allowance/copy state and foreign cloned paragraph refusal. The32-record case observes32 physical retained-source probes after duplicate roots, rather than asserting only a wall-clock bound.

The exact earlier cost regression is retained as an introduced implementation failure:560 probes versus32, exit101. The separate enrollment-omission source actually produces two named failures with exit101. Neither is a historical claim-Base defect. Related repeat4 and stored-accessor1 successful logs are separately pinned. The earlier Native6 log predates the linearity repair and is explicitly not final-source Native execution evidence. SOURCE_BOUNDARY does not independently encode the final GREEN producer exit, and does not pretend otherwise. Its raw log and retained binary association are preserved separately from Root's reported successful command.

The strict lower/Native all-target/all-feature Clippy log ends successfully after17.76seconds. Root reports fmt, prose, adapter drift and whitespace checks successful. No new confirmation build or test was run by this reviewer. Final external-client, Native semantic execution, raw-scope/relevant-consumer closure, actual producer lineage, final integrated source, resource/history, SemVer and all sprint riders remain pending.

## Prose and review boundary

Checked this new review prose before writing. No em dash, en dash or prose semicolon occurs. Only this review file was written. No source, plan, support packet, compiler, Native operation or oracle was changed or executed. Review pass ends here and returns control to Root.
