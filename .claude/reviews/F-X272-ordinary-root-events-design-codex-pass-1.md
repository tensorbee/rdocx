# F-X272 ordinary root events, Codex design pass 1

**Scope**: Only `Positive ordinary root events from the actual document reader`
at `.claude/plans/F-X272-design.md:6131`. Source citations refer to the
sole-writer worktree `/private/tmp/rdocx-s96-f272-codex-recovery` at inspection.
This is bounded design review, not implementation acceptance.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Checked boundaries

- The same-reader event contract at `.claude/plans/F-X272-design.md:6133`
  matches the actual factory. A single Reader is created at
  `crates/rdocx-oxml/src/document.rs:11811`. Qualified document Start,
  body Start, body Empty and document End have actual separate branches at
  `crates/rdocx-oxml/src/document.rs:11832`,
  `crates/rdocx-oxml/src/document.rs:11858`,
  `crates/rdocx-oxml/src/document.rs:11887` and
  `crates/rdocx-oxml/src/document.rs:11898`. Same-reader before/after positions
  can witness those events without another parser or a recovered QName search.
- Body opening end must be captured before the delegated child reader at
  `crates/rdocx-oxml/src/document.rs:11863`. The amendment states that exact
  boundary at `.claude/plans/F-X272-design.md:6136`. It does not extend body
  opening ownership across children or claim a body-closing event that this
  outer loop did not consume directly.
- Whole-factory success remains the commit boundary at
  `.claude/plans/F-X272-design.md:6137`. The actual reader can fail on late
  XML events, a missing document close or a missing body at
  `crates/rdocx-oxml/src/document.rs:11912` and
  `crates/rdocx-oxml/src/document.rs:11924`. The ordinary factory currently
  calls the admitted reader without origins at
  `crates/rdocx-oxml/src/document.rs:11795`. Keeping its companion disabled
  preserves that path without publishing partial successful-prefix facts.
- Unknown body contents, backgrounds, wrappers and the Field complement remain
  unproved at `.claude/plans/F-X272-design.md:6142`. This is consistent with
  the existing partial publication contract at
  `.claude/plans/F-X272-design.md:5339` and source_document consumer limits at
  `.claude/plans/F-X272-design.md:6014`. Positive root-event intervals do not
  certify all bytes, an empty Field list or later relevant consumers. Later
  actual Field ownership takes priority over ordinary intervals.
- The source-map requirement at `.claude/plans/F-X272-design.md:6149`
  preserves only genuine copied interval images across slice, composition,
  insertion, repetition and predecessor release. Generated gaps stay unknown.
  Ordinary facts are separate from Field-copy demand and add no allowance.
  This does not introduce a generic XML-copy quota or infer credit from
  unchanged/equal byte values.
- Required tests at `.claude/plans/F-X272-design.md:6157` distinguish qualified
  events, aliases and shadowed names, failing factories, actual interval
  transport, predecessor drop and later Field priority. Existing HLD03/04/10
  and all pending consumer obligations remain explicit at
  `.claude/plans/F-X272-design.md:6160`. No public signature, type, carrier
  count, model field or binding changes are authorized. Neither imported-payload
  method, actual Native lineage nor provider completion is supplied or waived.

## Verification limits

Read-only design and actual reader inspection. No source, plan or private
artifact edits, builds or tests. Only this review record was written. Actual
implementation and validation remain due, with no peer-cap reset.
