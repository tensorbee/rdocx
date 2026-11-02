# F-X276, Codex final microscope, pass 1

**Reviewed**: the uncommitted working tree of `work/f-x276-codex` against Base and HEAD `9a3b2960ccda7c3208ecb9b753c7019b64c086f5`. The tracked diff contains 13 files, 2033 insertions and 133 deletions. The three existing independent DeepSeek implementation reports were also inspected. This is the separate non-author Root final-review phase after those three passes, not a fourth implementation peer or a counter reset.

**Contract**: approved `.claude/plans/F-X276-design.md`, SHA-256 `d414f1514710f450a473cd70de025f8fa54c2241e941b40f866ae2d7af7ee3b4`. Both sprint records identify F-X276 as in-progress. The worker's historical DeepSeek owner entry is not treated as current implementer authority. The current progress records the authorized Codex takeover and acknowledges the current contract.

**Verdict**: 0 defects, 1 smell, 0 nitpicks. The new resource smell requires remediation or an explicit disposition before a clean microscope gate. The scoped semantic evidence does not constitute final sprint acceptance.

## Defects

None found in the inspected scope.

## Smells

### S1, each table cell retains a deep copy of the entire table

`crates/rpptx/src/lib.rs:27867`, `crates/rpptx/src/lib.rs:317`, `crates/rpptx/src/lib.rs:27767`, `crates/rpptx/src/lib.rs:4190`, `crates/rpptx-render/src/lib.rs:1149`, `crates/rpptx-layout/src/lib.rs:441`

Severity: medium, resource growth and duplicated work. Trigger: a resolved table with many physical cells. The inventory constructs one `AutofitTable` containing a full owned `ResolvedShape` and then executes `cell.and(table.clone())` inside the cell loop. `ResolvedTable` owns its rows, cells and text in vectors. Its derived clone is deep. Every retained cell context therefore holds another complete table, including empty cell contexts. For C cells, the newly retained table payload grows as C squared rather than C. This also happens for the mutating inventories, because creation of this payload is not restricted to the read-only caller. This is a source-derived allocation finding, not a claimed measured out-of-memory failure.

The comment at line 314 describes a table shared by its cell contexts, but the representation does not share it. In addition, when multiple cells have width-qualified stored height failure, line 4190 recomputes the same whole-table grown height for each such cell. The growth helper visits all table cells. That adds repeated full-table shaping to the same resource problem. The existing one-cell and two-by-two functional fixtures do not discriminate this behavior.

Expected remediation is one immutable table payload per resolved table and consuming context, shared by its cell contexts, with its grown page result computed at most once for that context when required. Preserve distinct slide/page-number contexts and the existing width qualification. A concrete shared owner or equivalent local representation is sufficient, without a new trait. Add a deterministic sharing or operation-count discriminator to the existing test carrier. Do not claim that merely sharing the payload also removes repeated full-table growth.

## Nitpicks

None.

## Scope and evidence identity

Read AGENTS guidance, CLAUDE.md, the applicable workflow and microscope rules, current plan and progress, both in-progress sprint rows, all three prior implementation reports, and the canonical peer2 contract adjudication. Bounded source reads covered:

- `rpptx/src/lib.rs`: changed Display, fit-report and mutation callers, candidate selection, transformed grown-table bounds, inventory and duplicate handling, notes fallback and both report publishers, and the Issue431 regressions around lines 36290 to 37030.
- `rpptx-render/src/lib.rs`: public table-height helper, extracted growth loop, painter row offsets, exact span-delta and fill-position controls. The unchanged owned table representation in `rpptx-layout/src/lib.rs` was also inspected for S1.
- `rpptx/src/html.rs`: all changed production regions and the expanded long-line regression, including paragraph assembly, marks and links, heading/list entry policy, notes, alt text, newline modes and escaping state.
- CLI commands and changed integration assertions for fit formatting, JSON, exit policy, picture diagnostics, contextless SmartArt, notes collision and validation Display.
- Python validation conversion, live slide validation and new index method, the additive stub and changed public binding tests.
- The HLD08 and HLD10 changes and their named contract sections. No unrelated Word field region was changed.

Rehashed the 99 entries in `F276_REPAIR2_FINAL_FROZEN_PINS.json`. Exactly 94 remain equal. Only the five disclosed Root-edited files differ: plan, progress, `rpptx/src/lib.rs`, `rpptx/src/html.rs` and HLD10. No pinned file is missing. Earlier CLI, geometry, paint and span evidence remains applicable where its production paths and assertions are unchanged. The historical CLI stack abort remains child101 with an incorrect wrapper0, separate from the later genuine green. It is not semantic RED.

Independently rehashed all 326 tracked inputs in the fresh Python `BUILD_MANIFEST.json`. All match current source. The imported extension hash equals `a1f93089d079346f39f14387775f3f681a1059f16da58588d36cd59f6c2e9ce8`. Inspected the authenticated runtime log and binding-control script and output. They identify the actual extension and Python3.14.7, with three selected tests passing and 1031 deselected. The additional public binding controls exercise keyed and unkeyed empty notes, unchanged package bytes and both Markdown punctuation markers. This is local extension evidence, not installed wheels or the final Python matrix. The old b92 extension is historical after Root production edits.

Read the named Root repair logs and receipts for the real empty-notes RED and six-test GREEN, actual Markdown work RED and one-test GREEN, post-repair fx218 21-pass run and Issue431 integration four-pass run. The Markdown semantic RED reports 4196351 prefix visits for 4096 characters. Its earlier expected-string compile failure is not semantic RED. Final four-crate all-targets/all-features clippy, fmt and prose receipts record child0, and their logs agree. Counts are per execution and are not summed into unique coverage. No test, build, benchmark, oracle or native application was run by this reviewer.

## Prior findings and Root repairs

Pass1 Display conversion and CLI assertions are fixed in both consumers. Both duplicate text leaves now have actual measured-scale assertions, picture diagnostics remain report-only, and mutation refusal remains intact. The original admitted table fixture pins `Some(40000)`. Transformed geometry, exact span delta, short neighbor, named existing paint, empty-frame, collision/export and contextless controls address the previous local gaps.

Peer2's two disputed clauses are resolved by the explicit canonical adjudication and current approved plan. Pair selection still admits only pairs fitting all contexts in both measurements. No pair yields both needed values and both trials absent. An admitted pair publishes both fitting trials. The stale slide accessor remains `StaleElementError`, invalid numeric lookup remains `IndexError`, and rejected index arguments become the approved `ValueError`. No synthetic non-fitting trial or fabricated public absent-owner path is required.

Peer3 S1 is resolved: the notes inventory attaches its unresolved cause only when the owner is non-empty. Both contextual and contextless empty paths preserve empty status and false overflow. The new native regression covers keyed and unkeyed empty notes, diagnostic absence and byte equality while retaining non-empty, type, key-collision and mutation-refusal controls. Duplicate diagnostics retain their distinct channel.

The Root digit-prefix repair is equivalent to the prior prefix predicate in the inspected cases. Digits extend the prefix until the first non-digit, whitespace clears it, hard breaks reset it, literal newlines invalidate it, and a false first-line content-start flag still suppresses ordered-marker escaping. State spans runs. The counter increments at the real digit predicate, and both punctuation markers have deterministic bounds plus newline and whitespace controls. Lazy line-end scanning remains limited to the first-column setext checks. The full public export benchmark is still pending.

## Acceptance criteria and microscope aspects

| Criterion | Inspected result | Remaining boundary |
|---|---|---|
| A1 | Physical duplicate leaves, picture-only report diagnostics, unchanged mutating refusal, Display parity, both unsupported publishers and notes cardinality/empty cases conform in inspected code and scoped evidence. | Final integrated union and applicable render/native riders. |
| A2 | Width-qualified exception, full transformed page rectangle, exact40000 admitted pair, absent-pair values, independent span delta and painter extraction conform functionally. S1 remains for table resource behavior. | Hash49, pixels7, final render and integrated gates. |
| A3 | One plain line per frame, shape/cell/context identity, diagnostic channel, unsupported exit1 and JSON reason, checked-error exit2 and validation Display conform. | Final integrated CLI union. |
| A4 | Assembled spaces and punctuation, hard-break state, notes prefix, alt and labels conform. Current native and exact-extension evidence exercise the new repair. | Full Markdown export benchmark and final integrated HTML/export checks. |
| A5 | Identity uses the same presentation and live path. Held and fresh wrappers after move, foreign/twin/reopened/non-slide/empty/removed cases and preserved accessor errors are represented in the public test and fresh runtime evidence. Stub signature matches the additive method. | Installed Python3.9/3.12 suites, typing, stub and archive qualification. |

Correctness: no additional finding. Contract: no additional finding after the explicit adjudications and empty-owner clarification. Panics: no new input-dependent panic found. The Markdown slices and unwraps follow the same assembled character stream and markdown branch, and table span bounds retain the existing checked helper. OOXML: no finding. The notes geometry changes affect measurement clones, existing raw preservation and schema child order are not rewritten, and package equality controls remain. Tests: no additional finding beyond the resource discriminator requested in S1. Structure: S1 above, with no new trait, generic, feature flag, module, crate or test binary introduced.

## Exclusions and pending final gates

This pass did not audit unrelated source or entire giant files, rerun suites, replay giant author JSONL histories, change code or metadata, or perform network, Git mutations, cleanup or native control. It did not establish native parity from self-produced fixtures.

The final integrated workspace union, hash49/pixels7, native/PDF/PNG/LibreOffice, installed Python3.9/3.12 typing/stubs and archives, hosted CI, full Markdown export benchmark and release approval remain pending. Issue431's unreduced layout/master/chart refusals are not silently removed by these local examples. No native PowerPoint guarantee is made. This report ends the audit phase. Remediation belongs to Root after the reviewer stops.
