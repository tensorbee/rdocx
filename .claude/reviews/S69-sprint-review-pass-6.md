# S69 sprint review, pass 6

**Reviewed**: full integrated dependency prefix on `sprint/s69` at
`3ee9c3f7d1efcb986a8a9418e874debbc72d555c` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 36 files and 5,951 changed
lines, comprising 5,097 additions and 854 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`, `rdocx-py`
**Pass authority**: pass 6 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, a quoted URL can still validate only the prefix before an inner parenthesis
`crates/rdocx/src/html.rs:1455`
`crates/rdocx/src/html.rs:1457`
`crates/rdocx/src/html.rs:1461`
`crates/rdocx/src/html.rs:1868`
`docs/hld/04-opc-and-packaging.md:557`

The pass 5 remediation safely rejects CSS resource syntax containing escapes,
but the remaining literal URL parser is not quote-aware. A closing parenthesis
is ordinary data inside a quoted CSS string, so
`url("https://example.test/index.html)outside")` refers to the complete URL
ending in `outside`. The scanner instead stops at the inner parenthesis, strips
the unmatched leading quote, and returns only
`https://example.test/index.html`. That prefix resolves to the MHTML root in
the standard fixture, so CSS preflight accepts it even though the actual
resource is unresolved. The normal HTML projection can then drop the
unsupported declaration and publish a document. This is the same
prefix-validation failure as pass 5 B1 without an escape. URL extraction must
recognize quoted-string boundaries before the function's closing parenthesis
and validate the complete reference. Add prefix-collision regressions for a
quoted `url(...)` declaration and quoted `url(...)` form of `@import`.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding closure

- **Pass 5 B1 remains open only for B1 above.** Identifier escapes, escaped URL
  delimiters, escaped quoted-import delimiters, and CSS line continuations all
  enter the conservative escape guard. It decodes them only to identify
  resource syntax, then rejects before literal extraction
  (`crates/rdocx/src/html.rs:1444`). The matrix covers escaped `url` and
  `@import` identifiers plus escaped URL and quoted-import delimiters
  (`crates/rdocx/src/html.rs:3476`, `crates/rdocx/src/html.rs:3486`). Ordinary
  literal `url(...)`, `@import url(...)`, and quoted `@import` forms retain the
  non-escape extraction path (`crates/rdocx/src/html.rs:1452`,
  `crates/rdocx/src/html.rs:1469`). A parenthesis inside a quoted literal URL
  is the remaining demonstrated boundary error.
- **Pass 4 B2 remains closed.** `Error::Mhtml` and
  `InvalidEmbeddedMutation` map to the existing `RdocxError` class
  (`crates/rdocx-py/src/lib.rs:66`), and the exact class regression covers both
  variants (`crates/rdocx-py/src/lib.rs:137`). The adapter compiles across all
  targets.
- **Pass 3 B1 remains closed.** The loss walk distinguishes shapes, charts and
  other drawings, linked images, unresolved images, and supported embedded
  images (`crates/rdocx/src/html.rs:597`). Its source-built regression asserts
  the lossy drawing states in source order while supported siblings survive
  (`crates/rdocx/src/html.rs:3689`). The independent reader regression pins
  embedded, linked, shape, filled-shape, and chart classification
  (`crates/rdocx/src/run.rs:969`).
- **Pass 1 B1 and pass 2 B1 remain closed outside B1 above.** Direct fetching
  elements, responsive attributes, legacy `background`, literal resource
  functions, and string-form imports remain covered
  (`crates/rdocx/src/html.rs:1749`, `crates/rdocx/src/html.rs:1814`,
  `crates/rdocx/src/html.rs:1848`, `crates/rdocx/src/html.rs:1861`).
- **Pass 1 B2 remains closed.** Import and export admit only sniffed PNG and
  JPEG resources (`crates/rdocx/src/html.rs:446`,
  `crates/rdocx/src/html.rs:1787`), and their two-direction format matrix
  passes (`crates/rdocx/src/html.rs:3558`).
- **Pass 1 B3 remains closed.** The ordinary differential uses one shared,
  mutation-sensitive acceptance predicate for the rdocx and pinned Word
  records (`crates/rdocx/tests/integration_test.rs:102`,
  `crates/rdocx/tests/integration_test.rs:215`).
- **Pass 1 B4 remains closed.** Direct runs, hyperlinks, tables, and drawing
  states have source-ordered nested loss handling
  (`crates/rdocx/src/html.rs:548`, `crates/rdocx/src/html.rs:627`,
  `crates/rdocx/src/html.rs:684`).
- **Pass 2 S1 remains closed.** The completion record attributes the corrected
  common-input Word oracle and full verification to
  `7fde4033b7cdf17f7c6e309dfccf7d1b9a6b1d44`
  (`docs/sprints/AS_BUILT.md:11779`).

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`).

The full milestone gate does not yet hold at this dependency-prefix checkpoint.
The modern package-variant clause remains assigned to pending F-238
(`docs/sprints/CURRENT_SPRINT.md:36`). F-X077's shared validator and owner
adapters retain executable focused evidence. F-239's Python consumer,
differential oracle, image formats, nested diagnostics, escape guard, and
ordinary MHTML tests pass, but complete quote-aware resource preflight remains
blocked by B1.

## Not found

- **Interaction**: zero additional cross-feature conflicts were found. F-X077
  remains the lowest shared XML lexical owner, and F-239 remains in the native
  HTML and MIME facade.
- **Duplication**: zero sprint-local duplicate lexical helpers were found. The
  glossary, embedded-content, and package-story owners use the shared
  `validate_strict_xml_1_0` policy (`crates/oxml-core/src/xml.rs:37`).
- **Layering and dependencies**: zero findings were found. No manifest changed,
  and no `oxml-*` crate gained a dependency on `rdocx-*` or `rpptx-*`.
- **Harness**: zero unexplained output deltas were found. The independent check
  reports all 49 entries match, consistent with both completion records
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`).
- **Docs and ledgers**: zero additional findings were found beyond B1. F-X077
  and F-239 remain consistently recorded as done
  (`docs/sprints/SPRINT_TRACKER.md:391`).
- **Surface**: zero unplanned API findings were found. MHTML remains native
  only, and the Python adapter retains its existing generic error surface
  (`docs/hld/10-bindings-spec.md:302`, `crates/rdocx-py/src/lib.rs:76`).

Focused evidence passed for all five MHTML unit tests, both ordinary MHTML
integration tests, the Python generic error-class test, `cargo check -p
rdocx-py --all-targets`, the drawing classification regression, the shared XML
unit matrix, all 42 glossary tests, the embedded owner mapping test, all five
package-story lexical tests, the 49-entry hash harness, and `git diff --check`.
The ignored Microsoft Word regeneration test was not run in this pass.
