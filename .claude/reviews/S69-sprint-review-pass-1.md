# S69 sprint review, pass 1

**Reviewed**: completed dependency prefix on `sprint/s69` at
`0ebae29751e56ec5307356c787f1e6d4f1f2d86a` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 30 files and 4,765 changed
lines, comprising 3,914 additions and 851 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`
**Verdict**: 4 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, common HTML subresources bypass fail-closed preflight
`crates/rdocx/src/html.rs:1519`
`crates/rdocx/src/html.rs:1574`

The two resource selectors omit standard fetching attributes such as
`video[src]`, `audio[src]`, and `track[src]`. An MHTML root can therefore carry
one of those elements with an external or unresolved URL and still reach HTML
projection, where the unsupported element is dropped or retained as text. The
approved contract requires every external or unresolved subresource to reject
the complete import before publication. Preflight must enumerate or otherwise
recognize all accepted HTML subresource forms, and a regression must prove that
each omitted form fails atomically.

### B2, MHTML accepts image formats outside the specified PNG and JPEG boundary
`crates/rdocx/src/html.rs:1550`
`crates/rdocx/src/html.rs:436`
`docs/hld/04-opc-and-packaging.md:502`

Import and export accept every format recognized by
`oxml_media::ImageFormat::sniff`, including GIF, BMP, WebP, TIFF, SVG, EMF, and
WMF when the declared type matches. The current HLD limits this interchange
surface to contained PNG and JPEG. Some additional raster forms also pass the
native-size path, so this is an effective input and output surface rather than
dead code. Both directions must enforce the declared two-format boundary, with
positive PNG and JPEG cases and rejection cases for the other sniffed formats.

### B3, the named differential gate never compares rdocx with Word
`crates/rdocx/tests/integration_test.rs:126`
`crates/rdocx/tests/integration_test.rs:201`
`docs/sprints/AS_BUILT.md:11772`

`mhtml_conversions_match_the_pinned_word_structure` derives its expected record
from rdocx's own import, then compares other rdocx imports against that record.
It asserts the oracle version as a string but consumes no Word-produced result.
The ignored test does invoke Word, but it uses a different fixture and checks a
small set of literals without comparing the same normalized record from both
implementations. This does not establish the approved differential gate, even
though the AS_BUILT entry reports that it does. One common source-built input
must be converted by Word and rdocx, compared through the same normalized
structure, and exercised by the declared mutation-sensitive acceptance
predicate before the feature can remain complete.

### B4, export loss diagnostics skip nested run and hyperlink losses
`crates/rdocx/src/html.rs:537`
`crates/rdocx/src/paragraph.rs:61`
`crates/rdocx/src/run.rs:269`

The diagnostic walk treats every run and hyperlink as wholly supported and
continues without inspecting their children. Hyperlinks can contain revisions,
unmodelled XML, tooltip and document-location metadata, while runs can contain
fields, notes, comments, and unmodelled XML. The outbound HTML projection can
drop or reduce those facts without producing an `MhtmlDiagnostic`, contrary to
the stable path-aware loss contract. The walk must classify nested run and
hyperlink facts against what the emitter preserves, and tests must prove each
unsupported sibling yields an ordered diagnostic while supported siblings
survive.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold at this dependency-prefix checkpoint.
The modern package-variant clause remains assigned to pending F-238
(`docs/sprints/BACKLOG.md:444`). F-X077 has executable evidence in
`strict_xml_1_0_validator_rejects_every_shared_lexical_class`
(`crates/oxml-core/src/xml.rs:419`), and the three owner adapters retain their
local error contracts. F-239's ordinary integration and parser tests pass, but
its required Word differential evidence does not hold for the reason in B3.

## Not found

- **Interaction**: no cross-feature conflict was found. F-X077 changes strict
  XML validation for glossary, embedded-content, and package-story owners.
  F-239 remains in the HTML and MIME boundary and publishes its generated DOCX
  through the existing save and reopen path.
- **Duplication**: no sprint-local duplicate helper remains. F-X077 removes the
  three lexical stacks and routes all owners through
  `validate_strict_xml_1_0` (`crates/oxml-core/src/xml.rs:37`).
- **Layering and dependencies**: no manifest changed, and no `oxml-*` crate
  gained a dependency on `rdocx-*` or `rpptx-*`.
- **Harness**: both AS_BUILT entries declare the harness unchanged
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11782`). The
  independent checkpoint run confirms all 49 entries match.
- **Docs and ledgers**: apart from the F-239 contract and evidence mismatches in
  B2 and B3, the approved HLD impact lists were executed and the backlog,
  current sprint, feature tracker, and completion log consistently record
  F-X077 and F-239 as done (`docs/sprints/BACKLOG.md:445`,
  `docs/sprints/CURRENT_SPRINT.md:34`, `docs/sprints/CURRENT_SPRINT.md:37`,
  `docs/sprints/SPRINT_TRACKER.md:391`).
- **Public surface**: the new `XmlLexicalError`, MHTML result, diagnostic, and
  document methods match the approved native pre-1.0 API shapes. No unplanned
  binding, CLI, feature, module, crate, trait, generic, or wrapper was added.

The focused named F-239 integration test and parser-limit unit test pass, as do
`python3 scripts/hash_harness.py --check` and `git diff --check main...HEAD`.
Those passing commands do not close B1 through B4 because the missing cases and
oracle comparison are outside their asserted predicates.
