# F-X278, Codex final microscope, pass 1

**Reviewed**: uncommitted work/f-x278-codex at Base and HEAD5820539cd277b417ae060e15010eb73bd1d00b01,21changed tracked files,2970insertions and348deletions. This is Root final after two DeepSeek implementation peers and the actual Codex third peer. No peer counter is reset.

**Contract**: approved plan cacd4844683b63f7c9fb2fbcde658d7b3b03fcdc36eb747768ac46fe9af02bae. The current sprint and backlog retain in-progress status. The author and third reviewer explicitly stopped. Source has not been edited in this review.

**Verdict**:1defect,0smells,0nitpicks. This pass stops for remediation of the new shared diagnostic transaction boundary. It does not grant final acceptance.

## Defects

### RF278-01, the new substitution channel escapes contribution-scope rollback

`crates/oxml-layout/src/font.rs:1444`, `crates/oxml-layout/src/font.rs:1466`, `crates/oxml-layout/src/font.rs:1509`, `crates/oxml-layout/src/font.rs:1622`, `crates/oxml-layout/src/font.rs:1867`, `crates/oxml-layout/src/font.rs:1951`

The new Word substitution list is appended by record_font_approximation during ordinary contribution recording and replay. The existing contribution scope snapshots and rolls back only layout_font_diagnostics, in both the successful capture and overflow paths. It never snapshots or truncates word_substitution_diagnostics. A completed capture therefore removes its provisional contribution events while leaving their new substitution facts visible through the newly public Word method.

Concrete API trigger: begin a deterministic layout, record one committed nonempty missing-family contribution, begin a contribution scope, record a different missing-family contribution, and finish the scope without replaying its returned events. The committed diagnostic should remain and the provisional diagnostic should be absent, matching the existing transactional channel. Source currently returns both. Replaying captured final events should publish the second fact then, with deduplication. Nested captures and trace overflow need the same boundary, without losing preexisting committed facts. The existing fx210_owned_diagnostics_roll_back_provisional_scopes_and_replay_final_events at font.rs:6547 verifies that boundary only for the old channel, so it does not discriminate this change.

This is a source-derived public component counterexample, not an executed failing test or a claim that the current Word engine itself uses contribution scopes. Current scope consumers include PowerPoint rendering and authoring, which do not read the Word list. The new public shared channel still needs to compose correctly with the existing scope API. Preserve channel separation and actual contribution recording. A private second snapshot length in the existing scope carrier and matching rollback branches can address it without a new abstraction. Add discriminating controls in the existing test module and establish actual old-source RED before remediation.

## Smells

None found in the inspected scope.

## Nitpicks

None.

## Inspected scope and evidence

Read the approved seven-A-ID plan, author handoff and complete third-peer dispositions. Bounded source reads examined Rust policy/cache selection, Python public render/page routes and bindings, shared font resolution and contribution replay/transaction boundaries, TOC/style consumers, indent pair merging, cell shading propagation, table-border ownership, retained row split guards and tbRl Owned semantics. Read the changed paragraph-property controls and existing contribution-scope regression. The new source defect was found by following the added diagnostic state through existing transaction publishers, rather than only begin_layout reset.

Independently rehashed all128 entries in FINAL_PINS.json, with no mismatch or missing entry. The exact imported artifact hashes to cadb117962e6d93dfeee048493dedae6e8d3da77b93388cd65e032789e87dbf7. The third peer already authenticated build receipts,326-input applicability and the post-build interpreted test change. This pass does not repeat that execution or call AST execution a full pytest collection. The author handoff and peer evidence retain real pagination RED, malformed fixtures, actual child/wrapper results and later greens separately.

The preliminary omission layout for a nonempty Python system-font request is explicitly required by the approved plan's Alternative A feature-off gate. It is not a new finding in this pass. The render_pages fix shares its diagnostic sibling and derives indices from the rendered layout. The tbRl repair retains Owned and its existing semantics, preserving the list/anchor structure path. The seven A-ID criteria and declared residuals remain intact. No additional issue in those inspected regions is asserted here.

Correctness and contract have RF278-01. Panics, OOXML, tests and structure have no additional finding within the inspected scope. This pass does not assert that every unrelated line in the giant files was reaudited or that absence of a finding proves native parity.

## Pending and handback

No source, test, plan, metadata, Git state or native application was modified. No test, build, oracle, benchmark, network or cleanup ran. The review ends with this report. Remediation is a separate phase and the clean final microscope remains due.

All final integrated workspace, deterministic hash49/pixels7 and reviewed14-key deltas, native Word, pinned LibreOffice/pdftoppm, installed Python3.9/3.12 typing/stubs, WASM, wheels, archives, hosted CI, publication dry-run/size and release approval remain pending. Historical broader suite failures remain failed. No integration or completion follows from peer3's clean report alone.
