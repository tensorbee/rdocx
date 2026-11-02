# F-X278, Codex final microscope, pass 2

**Reviewed**: uncommitted work/f-x278-codex, Base and HEAD5820539cd277b417ae060e15010eb73bd1d00b01. This continues Root final1 after the three actual implementation peers, without another peer or counter reset. Approved seven-row plan remains cacd4844683b63f7c9fb2fbcde658d7b3b03fcdc36eb747768ac46fe9af02bae. Author explicitly stopped after its updated handoff.

**Verdict**:0defects,0smells,0nitpicks. RF278-01 is resolved. Local implementation is ready for final-acceptance preparation, with final integrated qualification still pending.

## Defects

None remaining in the inspected source and applicable local evidence.

## Smells

None.

## Nitpicks

None.

## RF278-01 closure

The existing FontContributionScope now stores a private word_diagnostic_start beside the old diagnostic_start at crates/oxml-layout/src/font.rs:845. Begin captures the actual new channel length at :1446. Both successful capture and trace-overflow return truncate that channel to the captured prefix, at :1470 and :1517. Existing stale/foreign/nesting validation remains ahead of any rollback. Replay still publishes real contribution facts through the existing shared recording helper. No public API, cap, new type or abstraction is added by this repair.

The prefix representation preserves already committed facts and nested outer facts. A repeated request already present before a scope stays in that prefix. Distinct requests created only in the inner or outer scope disappear when that scope is captured, then return only when its contribution events are replayed. Replay remains deduplicated. Both channels reset at begin_layout. The Word list remains separate from the PowerPoint-facing list.

Inspected the extended existing fx210 success and overflow tests at font.rs:6552 and following. They use distinct outer and inner missing requests, preserve a preexisting committed diagnostic, inspect the actual new channel after each nested capture, replay staged events and assert deduplication. The overflow control exceeds the actual trace-entry limit, captures the inner and then outer scope, and checks that committed facts survive while provisional facts disappear. This rejects adding a reset only at layout start, fixing only successful capture, losing committed messages, or publishing facts twice on replay. Existing scope/refusal and seventy-fact controls remain.

## Evidence identity and outcomes

Independently rehashed all171 current FINAL_PINS.json entries, no missing file or mismatch. Independently checked all326 BUILD_SOURCE_INPUTS.json entries against the current worker, all equal. The new staged extension hashes to9bd8cc2ad5037562441719870e391d7e98a61ee97ea6cb5ecfce5737a30169bf, matching EXTENSION_ARTIFACT.json and the actual runtime import line. Manifest b6c5b64253e21778b9fa1d8795d1880eac301f49268022fd7e7b3fae134a8aff binds the successful build inputs. Prior artifact and source are frozen separately.

Read and rehashed each adjacent log/receipt in /private/tmp/rdocx-s96-root-f278-rf278-01/logs. The old-source regression records child101 and wrapper101, with actual new-channel counts4instead of3 on nested capture and4instead of2 on overflow. It is genuine semantic RED before the production edit. The later scope batch records8passing tests, child0/wrapper0, including both new controls, four existing scope/replay/refusal controls, bounded trace and seventy distinct real contributed requests. Counts overlap prior evidence and are not summed into unique coverage.

Affected oxml-layout all-target check, all-target/all-feature Clippy with warnings denied, fmt, prose and diff checks record child0/wrapper0. The successful fresh extension build precedes actual import and execution of the two named public policy/pagination definitions. Runtime prints the exact new artifact and dcda9c18bd904053fd26cb2651df2b3f3d892705a6c88a28b25bf32eb3b07b3e interpreted test hash, then reports both definitions passed. This is the documented AST runner, not full pytest collection, an installed wheel or the final Python matrix. All deciding log hashes match their receipts.

The previous six-crate Clippy and unchanged functional controls remain applicable to their unchanged production paths. No source-dependent build failure fell through into testing an old extension. Historical malformed fixtures, broader failures and later greens remain separately attributed.

## Contract and composition assessment

The source/contract reads from final1 and the actual third independent whole-story peer remain applicable beyond this narrow repair. Root additionally inspected the approved route tables, feature-off probe clause, all seven acceptance rows, matrix, exact HLD impact, compatibility inventory, residuals and final gates. Coherent source reads covered Rust selection/cache behavior, Python diagnostic/page routes, shared symbol and diagnostic contribution paths, paragraph-property merging, table border/shading forwarding and row barriers, TOC run/stored-segment filtering and tbRl Owned retention.

A278-1 retains stored/generated TOC and body Hyperlink controls. A278-2 retains both merge providers, public/EPUB/ODT projection and painted tbRl baseline with list/anchor ownership. A278-3 retains table-owned borders and the exact3by3/0by0 controls. A278-4 retains the original split tolerance and cantSplit barrier. A278-5 retains cell/control forwarding, nested-cell isolation and explicit/inherited colour precedence. A278-6 retains bounded Word-only mappings, both numbering and explicit-symbol consumers, unknown facts and PowerPoint separation. A278-7 retains one selected tier across Rust/Python/CLI siblings, caller priority, exact rendered-page selection, default SVG, cache identity and feature-off refusal. Declared residuals and native calibration are not silently counted as passed.

The preliminary omission-layout probe in the four nonempty system routes is explicitly prescribed in the approved contract. Its extra first-call work is not a new defect or a claimed measured failure. The new render_pages delegation avoids a separate selection layout. Retaining Owned in the tbRl arm addresses the actual list/anchor assertion without weakening the ordinary shared-paragraph invariant.

Correctness, contract, panics, OOXML, tests and structure have no additional finding in this inspected scope. No new dependency inversion, forbidden trait/generic, source file or test binary is introduced. Public changes remain the approved Word0.17.0/shared0.15.0 inventory. This review did not reread every unrelated line in the giant files or claim native parity from internal agreement.

## Pending final boundary and handback

Full integrated suites, deterministic hash49/pixels7 and the reviewed actual14-key delta account, native Word, pinned LibreOffice/pdftoppm, installed Python3.9/3.12 typing/stubs, WASM, wheels, archives, hosted CI, publication dry-run/size and separate release approval remain pending. F-X273 and F-X274 affected output qualification stays in that union. Historical integration462pass/10fail/8ignored and library391pass/17fail/1ignored remain failed observations. No deterministic baseline or oracle was recorded.

No source, test, plan, metadata, Git or native application was changed in this review. No build, test, benchmark, network, cleanup or delegation ran. The report ends this audit. Return to orchestration for worker preparation and later semantic integration against the preserved F-X272 carry.
