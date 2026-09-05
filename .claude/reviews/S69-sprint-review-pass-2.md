# S69 sprint review, pass 2

**Reviewed**: completed dependency prefix on `sprint/s69` at
`7fde4033b7cdf17f7c6e309dfccf7d1b9a6b1d44` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 31 files and 5,258 changed
lines, comprising 4,405 additions and 853 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`
**Verdict**: 1 blocking, 1 should-fix, 0 nice-to-have

## Blocking

### B1, the fail-closed subresource scan still omits valid fetching forms
`crates/rdocx/src/html.rs:1358`
`crates/rdocx/src/html.rs:1617`
`crates/rdocx/src/html.rs:1682`

The remediation added `video[src]`, `audio[src]`, `track[src]`, responsive
image attributes, and `url(...)` scanning, but the preflight is still a set of
partial selectors. It does not inspect resource-valued `background` attributes
such as `<body background="https://outside.test/a.png">`. Its CSS scanner also
recognizes only `url(...)`, so valid string-form imports such as
`<style>@import "https://outside.test/a.css";</style>` produce no reference.
Both inputs can therefore reach projection and publish a document after their
external resource is silently ignored. This is the same fail-closed contract
gap as pass 1 B1. The fix must cover every resource-bearing form admitted by
the bounded HTML parser, or reject unsupported fetching forms conservatively,
with mutations for `background` and string-form `@import` in addition to the
forms already tested.

## Should-fix

### S1, the completion record attributes the remediated oracle to the old SHA
`docs/sprints/AS_BUILT.md:11777`
`crates/rdocx/tests/integration_test.rs:215`
`.claude/reviews/S69-sprint-review-pass-1.md:38`

The AS_BUILT entry says the exact Word oracle passed at `7462b36`, but pass 1
found that the differential at that SHA never consumed a Word-produced record.
The common-input record and acceptance predicate now live at `7fde403`, whose
full verification is recorded in the sprint state. Update the tracked evidence
to name the remediated SHA so the completion record does not describe the
pre-remediation test as the exact integrated oracle.

## Nice-to-have

None.

## Pass 1 remediation

- **B1 remains open.** The newly named `video`, `audio`, `track`, responsive,
  and CSS `url(...)` cases fail closed, but the selectors and CSS scanner still
  omit the forms described in B1 above.
- **B2 is closed.** Import and export explicitly admit only
  `ImageFormat::Png` and `ImageFormat::Jpeg`
  (`crates/rdocx/src/html.rs:438`, `crates/rdocx/src/html.rs:1655`). The focused
  matrix accepts PNG and JPEG and rejects GIF, BMP, TIFF, WebP, SVG, EMF, and
  WMF in both directions
  (`crates/rdocx/src/html.rs:3375`).
- **B3 is closed in code.** The ordinary differential consumes a pinned
  Word-produced record and compares it with the rdocx result through one
  acceptance predicate (`crates/rdocx/tests/integration_test.rs:215`). The
  ignored exact-version regeneration test sends the same source-built MHTML to
  Word and rdocx, normalizes both results identically, authenticates both pinned
  records, and calls that predicate
  (`crates/rdocx/tests/integration_test.rs:272`). The seven perturbations cover
  every record dimension.
- **B4 is closed.** The exporter now walks direct run items and hyperlink
  metadata and children in source order
  (`crates/rdocx/src/html.rs:575`, `crates/rdocx/src/html.rs:599`). The regression
  proves ordered diagnostics for deleted text, fields, notes, comments, raw XML,
  legacy VML, hyperlink metadata, revisions, and nested raw XML while retaining
  supported siblings (`crates/rdocx/src/html.rs:3505`).

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold at this dependency-prefix checkpoint.
The modern package-variant clause remains assigned to pending F-238
(`docs/sprints/CURRENT_SPRINT.md:36`). F-X077's shared validator and all three
owner error surfaces passed their focused regression matrices. F-239's PNG and
JPEG, nested-loss, ordinary differential, integration, and parser-limit tests
pass, but its complete fail-closed resource condition remains blocked by B1.

## Not found

- **Interaction**: no cross-feature conflict was found. The strict XML refactor
  remains below the MHTML facade, and F-239 does not create a second lexical
  policy.
- **Duplication**: no sprint-local duplicate helper remains. Glossary,
  embedded-content, and package-story owners all call
  `validate_strict_xml_1_0` and retain local structural validation and error
  mapping.
- **Layering and dependencies**: no manifest changed, and no `oxml-*` crate
  gained a dependency on `rdocx-*` or `rpptx-*`.
- **Harness**: both completion entries declare an unchanged 49-entry harness,
  and the focused pass-2 check reports all 49 entries match.
- **Docs**: apart from S1, the approved HLD impact lists describe the implemented
  shared lexical and bounded MHTML surfaces without opening legacy `.doc`, Word
  2003 XML, network fetch, or permissive recovery.
- **Public surface**: the concrete XML lexical error and validator and the MHTML
  result, diagnostic, error, and document methods are the additive pre-1.0 API
  shapes approved by the two design plans. No unplanned binding, CLI, feature,
  module, crate, trait, generic, wrapper, or dependency was added.

Focused evidence passed for all five MHTML unit tests, both ordinary MHTML
integration tests, the shared XML lexical matrix, all 42 glossary tests, the
embedded owner mapping test, the five package-story tests, and the 49-entry
hash harness. These results do not close B1 because its omitted forms are not
part of the asserted matrix.
