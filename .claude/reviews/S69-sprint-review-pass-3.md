# S69 sprint review, pass 3

**Reviewed**: completed dependency prefix on `sprint/s69` at
`3b8bff5228fb10c3b145fdcb4db3c81e2746a4f2` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 32 files and 5,440 changed
lines, comprising 4,587 additions and 853 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, non-image drawings are still silently lost during MHTML export
`crates/rdocx/src/html.rs:585`
`crates/rdocx-html/src/emitter.rs:311`
`crates/rdocx/src/run.rs:24`

The remediated nested-loss walk treats every `RunItemRef::Drawing` as fully
supported and emits no diagnostic. The outbound HTML backend only writes a
drawing when it has an image relationship whose bytes exist in the image map.
The public run model expressly distinguishes images from anchored shapes and
other drawing constructs such as charts. Those non-image drawings therefore
produce no HTML and no `MhtmlDiagnostic`, so an unsupported safe fact can still
disappear beside supported text without its required path-aware loss record.
The walk must diagnose drawing kinds and relationship states the emitter does
not serialize, while leaving successfully emitted embedded PNG and JPEG images
undiagnosed. A source-built regression must cover a shape or chart beside
supported run content and prove ordered loss plus sibling survival.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding closure

- **Pass 1 B1 and pass 2 B1 are closed.** Direct resource selectors cover the
  ordinary fetching elements, responsive attributes cover `srcset`, source,
  poster, and input forms, and the legacy selector checks every `background`
  attribute (`crates/rdocx/src/html.rs:1656`,
  `crates/rdocx/src/html.rs:1721`, `crates/rdocx/src/html.rs:1755`). The CSS
  scan handles both `url(...)` and quoted string-form `@import`, rejecting
  unsupported or escaped import forms conservatively
  (`crates/rdocx/src/html.rs:1358`). The focused matrix makes external and
  unresolved `background` and quoted-import mutations fail
  (`crates/rdocx/src/html.rs:3363`, `crates/rdocx/src/html.rs:3375`).
- **Pass 1 B2 is closed.** Both import and export admit only PNG and JPEG after
  byte sniffing (`crates/rdocx/src/html.rs:438`,
  `crates/rdocx/src/html.rs:1694`). The focused matrix accepts those two
  formats and rejects every other recognized image format in both directions
  (`crates/rdocx/src/html.rs:3444`).
- **Pass 1 B3 is closed.** The ordinary differential compares the rdocx record
  with the pinned Word record through one acceptance predicate, and seven
  independent perturbations cover every record dimension
  (`crates/rdocx/tests/integration_test.rs:215`). The ignored regeneration gate
  sends the same source-built MHTML through Word and rdocx, normalizes both
  outputs identically, verifies both pinned records, and authenticates the exact
  Word build (`crates/rdocx/tests/integration_test.rs:272`).
- **Pass 1 B4 remains open only for B1 above.** Direct runs and hyperlinks now
  receive source-ordered nested diagnostic walks for deleted text, fields,
  notes, comments, raw XML, legacy VML, hyperlink metadata, revisions, and
  nested run content (`crates/rdocx/src/html.rs:542`,
  `crates/rdocx/src/html.rs:575`, `crates/rdocx/src/html.rs:599`). The drawing
  arm is the remaining unsupported nested loss.
- **Pass 2 S1 is closed.** The completion record now attributes the remediated
  common-input Word oracle and its full verification to
  `7fde4033b7cdf17f7c6e309dfccf7d1b9a6b1d44`, where that evidence was added and
  run (`docs/sprints/AS_BUILT.md:11779`).

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold at this dependency-prefix checkpoint.
The modern package-variant clause remains assigned to pending F-238
(`docs/sprints/CURRENT_SPRINT.md:36`). F-X077's shared validator and three
owner error surfaces have executable focused evidence. F-239's resource,
image-format, ordinary differential, exact Word, integration, and nested-loss
tests pass, but the complete loss-diagnostic contract remains blocked by B1.

## Not found

- **Interaction**: no cross-feature conflict was found. F-X077 remains the
  lowest shared XML lexical owner, while F-239 stays within the HTML and MIME
  facade and uses the established DOCX save and reopen path.
- **Duplication**: no sprint-local duplicate lexical helper remains. Glossary,
  embedded-content, and package-story owners all route shared lexical policy
  through `validate_strict_xml_1_0` (`crates/oxml-core/src/xml.rs:37`).
- **Layering and dependencies**: no manifest changed, and no `oxml-*` crate
  gained a dependency on `rdocx-*` or `rpptx-*`.
- **Harness**: both completion entries declare the harness unchanged
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`). The
  sprint state records a passing full verification at the reviewed exact HEAD
  with all 49 entries unchanged.
- **Docs and ledgers**: apart from B1's remaining implementation gap, the
  approved HLD impact lists describe the implemented prefix. The backlog,
  current sprint, feature tracker, and completion log consistently record
  F-X077 and F-239 as done.
- **Public surface**: the shared lexical and MHTML APIs match the approved
  native pre-1.0 shapes. No unplanned binding, CLI, feature, module, crate,
  trait, generic, wrapper, or dependency was added.

All five focused MHTML unit tests, both ordinary MHTML integration tests, the
shared XML lexical matrix, all 42 glossary tests, the embedded owner mapping
test, and all five package-story lexical tests pass. `git diff --check` also
passes. Those results do not close B1 because no asserted MHTML loss case
contains a non-image DrawingML shape or chart.
