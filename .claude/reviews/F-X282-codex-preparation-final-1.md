# F-X282, Codex final preparation pass 1

REJECTED PREPARATION pending a corrected validation recipe and behavior account. This is actual Root final after three completed design peers. No fourth design peer, implementation approval, claim or feature acceptance. A later implementation phase has its own user-authorized cap of three peers.

## Identity and coverage

Reviewed Grok V3 at /private/tmp/rdocx-grok-harness/f-x282-preparation/GROK_PROPOSAL_V3.md, SHA-256 866467ea72935db43ccf3ed168e3859287db960ca140050da73dd04f172b056d, 86497 bytes. Native handoff is 0937c90d394e86663bd4eed31d69a7b3bbfb9d63d56e5378825a37e8f6ff58dd. Actual DeepSeek third peer accepted the preparation with three nits. Root read the proposal, original Issue437 scripts, the peer report, relevant canonical verification requirements and scoped current source for styles publication, the cell editor, Python mutation and append consumers, and packaging support.

The 919-file packet is historical. No clean Base, accepted Source provider or new runtime evidence is implied. Root ran no build, product test, timer, native export or network fetch for this audit. Reading installed cargo publish help established that the pinned Cargo supports workspace publication, so there is no finding about that option itself.

## D1, medium, final verification commands do not describe executable binding and packaging gates

GROK_PROPOSAL_V3.md:309 names cargo test -p rdocx-py among the final commands, followed by the vague qualification about workflow exclusions. This does not implement the promised Python binding regressions. CLAUDE.md:72 explains that the extension-module configuration cannot link a Rust test binary. The canonical workspace gate explicitly excludes both binding crates at .claude/commands/verify.md:124. An exclusion cannot turn a targeted binding-crate cargo test into an installed Python test.

The packaging row at proposal :312 also substitutes a plain cargo publish --workspace --dry-run --allow-dirty --locked for the canonical gate. .claude/commands/verify.md:183 requires every publishable internal crate to be patched to the reviewed local source. That matters for the prepared, unpublished 0.15.0 and 0.17.0 families. Workspace support in Cargo does not waive the repository's local-source recipe.

Correct the test plan to consume the canonical full workspace and patched packaging recipes once, with the bindings excluded from Rust tests. Name the actual release Python3.12 binding producer and pytest/installed-binding execution for the declared test_core.py cases and exact issue scripts. Do not claim source-only compile checks prove installed runtime or typing agreement. Preserve the existing final union, artifact identity and no-publication boundaries. This is a validation-contract defect, not a new observed linker or package failure.

## S1, behavior-accounting smell, the stated only possible delta omits an explicit planned refusal

Proposal :347 says the only possible behavior delta is a parseable but unserializable styles part. However :208, :232 and the C6 row at :247 already require malformed replacement bytes to return an error rather than be rewritten. Current flush code at crates/rdocx/src/document.rs:26178 converts from_xml failure into the changed branch, then may retain the root and rewrite the part. The new cache helper instead returns that parse error before set_part.

The plan already includes a discriminator for this state, so the behavior account must include it. Establish the actual before/after result on the implementation Base, distinguish publicly reachable states from a private package-injection fixture, and label any real intentional behavior change in its own commit with the expected result, as AGENTS requires. Do not describe a changed refusal as a byte-identical performance-only transformation. A parseable-unserializable proof failure is deliberately preserved at save, so that case alone does not establish a changed outcome.

This confirms the substance of peer nit N3 but makes the implementation and commit obligation explicit. No new user policy question is necessary for that factual accounting.

## Remaining nits and current facts

Peer N4 is valid. Make the current styles part bytes Arc identity the first explicit cache-hit condition in the helper bullets at proposal :92. Model pointer equality alone must never suppress inspection after a part replacement. The behavior and C9 rows already require the right result, so this is a clarity correction rather than a second algorithm defect.

Update dated coordination wording at canonical adoption. The proposal and peer predate the user's Q-S96-07 answer. It is now bundled default with system opt-in, unrelated to F282. F272 remains in progress, with Root final4 rejecting the current C1/C4 increment and a design clarification pending. Do not turn an old as-of sentence into an allegation of fabrication or a new provider dependency. F276 remains a draft. Approval, a claim and the clean shared Base remain distinct from safe draft adoption.

Clarify that F282 adds no hash delta relative to the reviewed integrated expectations. Other sprint stories can own separately declared behavior changes. F282 must neither absorb nor undo their intentional baseline work. Its own exact issue thresholds and F-X268 gate definitions stay unchanged.

## Prior mechanisms accepted as a preparation basis

Zero additional defects in the following reviewed mechanisms. This is design acceptance of an approach, not implementation acceptance or measured performance.

- The unchanged styles cache is conditional on a successful serialization proof obtained at open, outside the original timed edits. Unserializable records retain the later writer failure. Changed publication updates the installed bytes Arc and proof only after success.
- Shared CT_Styles custody uses Arc::make_mut for a real change, and old staged documents retain their own model. The documented shared-part accessor preserves the existing get_part signature and enters the shared0.15 API inventory.
- The cell bridge clones the selected paragraph and publishes only on Ok. The reader/mutable sequence proof explicitly includes controlled paragraphs and control-owned nested tables, with the checked route retained when counts differ. Both current generic return types are named: unit and the append pair containing AcceptedRunPath.
- Python run binding uses the existing prepared index with the checked fallback preserved. The proposal correctly drops the unrelated native lazy-table rewrite.
- Original case1 uses ten calls including the first, without warmup. Case2 retains cold n=3 cell-place sampling and the separate one-call body warmups. Case3 retains the actual thirty and one-hundred-twenty cell loops and the ratio of per-call costs. No threshold or workload was changed.

No dropped A-ID, invented provider edge, new product file/module/crate/test binary/trait/feature flag, donor timing acceptance or fabricated execution found. The additive generic has two actual current instantiations. All performance, F-X26879/252/407, saved-byte, hash49, native, typing, wheels, WASM, archive and hosted gates remain unexecuted.

## Handoff

Return to orchestration for original Grok author corrections, then affected Root final recheck. No fresh design peer cycle. Correct canonical source and head citations before eventual draft adoption, then follow the ordinary approval and clean-Base claim gates. No commit, push, release or completion follows. This pass ends here.
