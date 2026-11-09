# F-X272, Root non-emitting functional review, pass 1

**Reviewed**: The sealed template.rs increment at `ROOT_TAKEOVER/root-nonemitting-1`, 581 additions and 52 deletions including tests. Final source SHA256 is `fe01954ce4def74bb0069922c400141a1d46aad9f4194b7db9ccf675d771bff7`. The preserved executable authenticates as `1f52f31434cd01226215266dbd871ba5411cb3f210b9e66d8ce00b33a7e00e18`. This is independent bounded Root review of the author handoff, not whole F-X272 acceptance or another full peer cycle.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this bounded functional increment.

## Defects

0 defects.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Functional assessment

The false versioned expression path at `crates/rdocx/src/template.rs:1132` uses the shared evaluator in validation mode. The false legacy path at `crates/rdocx/src/template.rs:1169` retains its actual resolver and validates the selected scalar without scalar-copy activation. Deferred expressions retain the existing earlier skip. Unused dependency collections receive no selected operands.

The shared driver at `crates/rdocx/src/template.rs:2962` preserves actual per-filter sequencing. Emitting join still selects and materializes each scalar before examining the next member, and emitting case conversion retains its reservation before later filter type errors. Default, borrowed scalar and converted-string states remain distinct. A null after trim does not become a later fallback. Join after string conversion still refuses. The real emitted lowercase implementation remains std str::to_lowercase.

The borrowed validation representation keeps actual scalar, fallback or joined inputs and their remaining filter suffix. It checks all actual join member scalar kinds before final XML validation, preserving a later member type error ahead of an earlier illegal character. It never calls scalar_text, scalar input activation or output reservation to validate dropped text. Numbers use their actual Display implementation through the validation sink.

The character sink at `crates/rdocx/src/template.rs:3135` projects case conversion without constructing replacement Strings. Contextual sigma changes identity but not XML legality or character count. Trim preserves leading and pending trailing whitespace behavior. Pending internal invalid whitespace is reported only once a later non-whitespace character makes it internal. Final error positions count transformed characters rather than bytes, using the shared XML character predicate and the existing error format. The actual emitting fallback clone now reserves its real copy before growth under the existing Word-copy metric.

The type predicate at `crates/rdocx/src/template.rs:3906` is shared with scalar_text and preserves its path error. The additional private states and validation sink have current emitting and non-emitting consumers. No new public carrier, trait, generic parameter, module, file, quota or input baseline is introduced.

## Evidence assessment

The actual false-branch reproduction at `crates/rdocx/src/template.rs:4894` failed before repair because 600 selected 10k scalar replacements exhausted copy work despite no emitted branch. The final case passes in legacy and versioned modes. The zero-test selector attempt is preserved as an invocation error, not semantic RED.

The reviewer inspected the five non-emitting controls and their actual assertions. They cover final XML error positions after expansions, boundary versus internal invalid whitespace, joined separators, numbers, scalar/filter type state, deferred empty-loop paths, no input credit and actual emitting reservation-error precedence. The distinct fallback-copy case verifies eventual refusal without renewed allowance. Source and binary hashes authenticate. The seal records 50 passing local test observations, including overlapping named cohorts, rather than 50 independent acceptance criteria. Scoped check, warnings-denied all-targets/all-features lower/Native Clippy and metadata observations are separately recorded. No reviewer build or test rerun was needed.

## Explicit boundaries

This review does not prove general adversarial complexity or overflow properties and does not run deferred deeper security checks. It does not accept rebuilt revision installation, comments part readers/writers, the planned four additional carriers, external client compilation, HLD or canonical source reconciliation, the genuine provider gate, dependent stories or final integrated/native/oracle/release checks. Current implementation remains42carriers/7accessors, with46/7planned. The immutable source reviewed here is distinct from the subsequent installer worktree edits.

## Not found

No additional correctness, panic on the exercised ordinary inputs, namespace/XML preservation, filter error-order, input ownership, compatibility or structural-rule finding was found in this exact bounded increment. Rendering and native Word behavior were not claimed or independently observed.
