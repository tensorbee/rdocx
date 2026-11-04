# F-X277c, Codex self-review pass1

**Reviewed**: 7 files changed, 794 insertions(+), 23 deletions(-). Seven product/test/HLD paths against bd48016f2ef48ed77c63505b2db21ba9e8100382. Candidate diff SHA25694b5a06871292a3e67cda0a935ba91d78396b34b088bd0d4d652ff710189c771. Identified review of Root's own implementation, not independent acceptance. No builds, tests or source/progress edits in this phase.
**Verdict**: 0 defects, 0 smells, 0 nitpicks. Actual independent microscope remains required before preparation.

## Defects

Zero findings.

## Smells

Zero findings.

## Nitpicks

Zero findings.

## Aspects examined

Correctness: merge distinguishes lone-empty shape from text-empty semantic content and preserves the origin on all-empty selection. Last paragraph and direct last-run fallback have explicit zero/one/many and mark-priority controls. Split uses the unchanged cell-only helper. The authored TOC receives a separate end paragraph only with no headings. Literal replacement removes only chart traversal. Existing regex/chart-data/template paths remain separate.

Contract: all four original A-IDs remain with C. Both native row/column publishers and actual Python staged wrappers reach the repaired native path. Python regression source uses the real replace_all_regex method, not a similarly named native API. Native, CLI plain/JSON and expected-count agreement use the actual supported-story projection. Parent, F272 and F280 barriers are unchanged. No public signature, dependency, module/file or quota was added.

Panics: existing checked topology and selected-row bounds precede merge indexing. New helper uses checked Option iteration and indexes only values obtained by enumerate from the same immutable vector. No user-derived arithmetic or unchecked slicing was added. Raw reader errors preserve the unknown carrier rather than panic.

OOXML: existing property writer retains sequence responsibility. The new paragraph starts without source identity. Numbering, section and paragraph-property revision ownership are removed, including raw qualified duplicates. Canonical implicit w fragments and explicitly rebound foreign names are distinguished. Other opaque children keep bytes and source positions. Source cell remains unchanged. Schema/native qualification is still pending, never inferred from reopen.

Tests: actual old producer fails merge, inheritance, TOC-before-headings and corrected chart-cache regression. Setup failures are separately classified in progress/receipts. Final scoped observations are authenticated in FINAL_SOURCE.json. New positive controls distinguish nontext content, multiple empty paragraphs, fallback priority, source identity, foreign names, split defaults, opaque refusal, imported TOC atomic refusal and paired text-box copies. Existing four native and three CLI controls pass. JSON counts only text projections rather than duplicated run-detail strings.

Structure: one substantive private paragraph helper and inline merge classification in the existing table implementation, with paragraph inheritance used by two real row/column consumers. No forwarding abstraction, new trait/generic, crate, feature flag or file. HLD03/HLD10 describe the implemented table, authored TOC and chart-exclusion behavior without changing other owners' sections.

## Evidence limits and handoff

All reviewed scoped Rust/CLI, lint, format, prose, whitespace and AST observations are local. Python tests are source-only here. Installed supported Python/runtime and typing, SDK/raw preservation, final full/hash, WASM, archive/dry-run, native residual and hosted CI remain mandatory final debt. No provider completion, baseline, commit, reporter comment or publication is claimed. End this self-review here and request actual independent review of the frozen candidate.
