# F-X272 ordinary property events, Codex design pass 1

**Scope**: Only the strengthened27-line `Positive ordinary property events from actual
readers` amendment at `.claude/plans/F-X272-design.md:6164`. Recovery-worktree
source citations refer to `/private/tmp/rdocx-s96-f272-codex-recovery`.
This is technical design review, not implementation or provider acceptance.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Checked boundaries

- The selected branches exist in the substantive readers. Qualified Empty
  pStyle decodes its value at
  `crates/rdocx-oxml/src/paragraph_properties.rs:902`, while qualified Empty
  rStyle decodes at `crates/rdocx-oxml/src/run_properties.rs:986`. Reporting
  only after the existing decoder succeeds, as required at
  `.claude/plans/F-X272-design.md:6166`, preserves branch, namespace and error
  authority. A missing optional value remains the decoder's existing outcome,
  rather than a new validation rule. Start forms and foreign/shadowed names
  acquire no fact from matching local spelling.
- Actual paragraph/run factories currently capture property XML and invoke
  scoped parsing at `crates/rdocx-oxml/src/text.rs:20672` and
  `crates/rdocx-oxml/src/text.rs:9930`. Their actual scoped readers are
  `crates/rdocx-oxml/src/numbering.rs:1389` and
  `crates/rdocx-oxml/src/numbering.rs:1489`. The plan's private optional
  handback at `.claude/plans/F-X272-design.md:6169` extends those existing
  consumers rather than introducing a parallel classifier. Ordinary None
  calls remain without a new transcript.
- Capture transport must use the existing actual source/output transcript at
  `crates/rdocx-oxml/src/document.rs:3088`. Its generated gaps and copied
  images are distinguished at `crates/rdocx-oxml/src/document.rs:3095`.
  The amendment therefore does not authorize a constant offset or a blind
  shift across normalized capture bytes. The earlier root contract retains
  that same copied-image boundary at
  `.claude/plans/F-X272-design.md:6149`.
- Publication remains staged through the actual retained constructor-return
  chain at `.claude/plans/F-X272-design.md:6173`. A decoded style event in a
  discarded run/owner or raw-fallback candidate supplies no committed ordinary
  fact, even if the document later succeeds. This strengthened rule excludes
  successful temporary-property observations that never survive model delivery.
  Actual Field coverage takes precedence, including cache properties. Unknown
  children, Start variants and the entire property container remain unproved
  at `.claude/plans/F-X272-design.md:6178`.
- Retained smartTag/customXml wrappers can obtain facts only when their own
  later actual property reader runs, with captured-map transport at
  `.claude/plans/F-X272-design.md:6180`. This preserves the existing partial
  publication rule at `.claude/plans/F-X272-design.md:5339`. Neither the
  initial reader nor an empty Field list supplies a generic ordinary
  complement or later-consumer closure.
- The evidence matrix at `.claude/plans/F-X272-design.md:6183` covers decoder
  and parent failures, namespace and Start distinctions, capture normalization,
  aliases/BOM, Field priority and generated gaps. HLD03/04/10 remain explicit
  impact files at `.claude/plans/F-X272-design.md:6188`. Public APIs, valid
  acceptance, quota and allowance remain unchanged. Relationship, drawing and
  other identity-owner handoffs and all final riders remain pending.

## Verification limits

Read the actual reader/capture seams, previous root-event contract and canonical
design/microscope command rules. No source, plan, progress or private artifact
was edited. No builds or tests were run. Only this review record was written.
Native implementation remains a separate increment. No peer cap is reset.
