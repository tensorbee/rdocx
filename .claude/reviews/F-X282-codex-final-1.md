# F-X282, Codex final implementation microscope 1

**Reviewed**: Root-delegated independent final implementation review of work/f-x282-codex, Base aff3f2b8744ed86755db8520c9bf6abeed5f2038. Twelve tracked files, 1166 insertions and 96 deletions. Read the working diff, approved contract, progress, both actual implementation peer reports, author handoff, manifests, actual outcome logs and affected HLD contracts. This is Root final review, not another numbered implementation peer or a counter reset.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Clean local implementation review. This is not release performance acceptance, integrated source acceptance or feature completion.

## Defects

None found.

## Smells

None found.

## Nitpicks

None recorded.

## Production and consumer trace

Styles custody checks byte allocation identity before model identity or equality at `crates/rdocx/src/document.rs:26155`. Opening records an actual serialization proof, including failure, at `crates/rdocx/src/document.rs:24662`. Staging shares the model and cache at `crates/rdocx/src/document.rs:22722`. In-memory reopen carries the cache at `crates/rdocx/src/document.rs:25508`, while the ZIP branch reopens newly read bytes. Changed styles serialize and retain root attributes before installation and cache refresh at `crates/rdocx/src/document.rs:26235`. Failed proof retry does not mark a failed writer successful. New allocation identity cannot be hidden by an unchanged model pointer.

Inspected copy-on-write default-font, paragraph-default, numbering-link, style add/edit/remove, merge and fragment import publishers. Representative paths are `crates/rdocx/src/document.rs:38521`, `crates/rdocx/src/document.rs:39828`, `crates/rdocx/src/document.rs:40022`, `crates/rdocx/src/document.rs:43963` and `crates/rdocx/src/field.rs:5793`. HTML and layout retain owned model clones at `crates/rdocx/src/document.rs:46454` and `crates/rdocx/src/document.rs:47068`. The F-X278 font policy is preserved. The public accessor uses the existing deterministic part lookup at `crates/oxml-opc/src/package.rs:543`.

The immutable table selector and mutable selector enumerate the same body/control/cell variants at `crates/rdocx/src/document.rs:20533` and `crates/rdocx/src/document.rs:20582`. Cell paragraph eligibility compares the reader against direct paragraphs before callback execution at `crates/rdocx/src/document.rs:32943`. This matches the different reader and mutable projections at `crates/rdocx-oxml/src/table.rs:2154` and `crates/rdocx-oxml/src/table.rs:2161`. A direct nested table contributes to neither paragraph sequence. A control-owned table can contribute reader paragraphs and fails direct eligibility.

The callback edits an isolated paragraph clone. Serialization, section refusal, strict sufficient qualification, relationship inspection, comment-marker comparison and retained-owner qualification precede assignment at `crates/rdocx/src/document.rs:32963`. Failed qualification uses the already edited fragment with actual source facts at `crates/rdocx/src/document.rs:33014`. It does not rerun the callback. The direct branch reacquires and compares the original paragraph before assignment and invalidates only afterward at `crates/rdocx/src/document.rs:33044`.

Actual typed fallback is `replace_story_content` at `crates/rdocx/src/document.rs:31277`, through relationship admission at `crates/rdocx/src/document.rs:15039`, source replacement, comment reconciliation and prepared reopen at `crates/rdocx/src/document.rs:31344`. It is not the raw-owner replacement guard. C1 is excluded from this Base. Integration must reconcile its actual typed publication policy with this bridge. No speculative strict raw-owner refusal is required here.

Python cell mutation calls the bridge at `crates/rdocx-py/src/paragraph.rs:266`. Run binding uses the prepared route at `crates/rdocx-py/src/run.rs:72`, whose indexed lookup and legacy fallback remain at `crates/rdocx-py/src/document.rs:10209`. Existing signatures and owner authority remain. The generic has actual unit and run-insertion return consumers.

## Prior findings and evidence

Implementation peer1 D1 is closed. The four actual retained lexical inputs at `crates/rdocx/src/document.rs:49881` are serialized inside both callbacks, and the emitted candidate's strict result is asserted at `crates/rdocx/src/document.rs:49910`. Both checked publication and bridge succeed with one callback. Strict-error cases select checked fallback, retain raw bytes and match exact saved ZIP bytes at `crates/rdocx/src/document.rs:49964`. The ordinary positive selects direct publication. This rejects a blanket strict-error refusal.

Implementation peer1 D2 is closed. The control-owned nested-table fixture at `crates/rdocx/src/document.rs:49974` establishes three reader paragraphs against two direct paragraphs. All three ordinals preserve the actual checked resolver's preflight refusal, zero callbacks, model/source/save bytes and reopened order. The public fixture at `crates/rdocx-py/tests/test_core.py:16083` checks held paragraph/run and fresh/reopened values. The same three public controls pass on the authenticated old artifact, so this preserves existing policy rather than narrowing a supported mutation.

Independently hashed all twelve current files against FINAL_SOURCE_MANIFEST.json, all indexed logs against EVIDENCE_INDEX.json, and compared current git diff bytes with FINAL_DIFF.patch. All match. The extension hashes to bcd1116683d73c37241ed0d178e2ca55127db73d1901fd601b22797bf9efa74b. The current production prefix before the native test module hashes to c282062b435a09a8052545f2f5cecc8847d368f7d1c039a778c1cf9ad87fb86f, matching the recorded reviewed prefix. Of 326 compiled-input manifest entries, only the declared native test carrier and interpreted Python test source differ from current source. Read the test-only changes and their applicability records. Current interpreted tests hash to 123eb7f6d36585909d883026c5a0dfb97a0385802580ccbea4600a01ccc5a29d.

Read actual named outcomes: the final fourteen native cases pass, the earlier OPC case passes, the earlier 42 public cases pass, and the three added public controls pass separately on current and old artifacts. Indexed child/wrapper exits and log contents agree for the successful observations. Earlier failed fixture and setup receipts remain failures, not invented production REDs. The old public direct-setter work assertion fails while the new runtime passes. Debug Python 3.14 observations establish semantic behavior only. No combined newly executed 45-case run is claimed.

## Not found and retained limits

Correctness, contract, panics, OOXML, tests and structure produced no remaining finding in this final pass. No new trait, file, module, crate, feature flag or test binary is introduced. The two public additions are the approved ones. Current shared status records identify F-X282 as in-progress.

The synthetic failed styles proof at `crates/rdocx/src/document.rs:50103` proves guard behavior, not discovery of a real parseable-open writer-error source. The strict-qualified foreign-extension fixture explicitly records inherited-scope versus redundant namespace placement at `crates/rdocx/src/document.rs:49873`. Semantic equality does not establish exact output-byte equality. The malformed styles replacement refusal at `crates/rdocx/src/document.rs:50206` remains an intentional behavior delta. Root must retain its isolated patch and dedicated labelled integration commit.

All exact A282-1, A282-2 and A282-3 Linux release Python 3.12 ratios remain pending, with original input sizes, cold first-call inclusion, no cell warmup, the single permitted body warmup and unchanged thresholds. Counters do not replace those gates. Final integrated saved bytes/hash49, F-X268 and 79/252/407, full workspace/native/oracle, installed runtime and typing, WASM, wheels, archives and hosted CI remain pending. Integration composition with C1 is explicit debt. No local review converts that debt into acceptance or authorizes publication.

No product edits, tests, builds, benchmarks, network operations or Git mutations were performed. This report is the only write. Review ends and the review-directory lease returns to Root.
