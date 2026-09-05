# S69 sprint review, pass 7

**Reviewed**: full integrated dependency prefix on `sprint/s69` at
`60128b59ee1fe77b9fc442478c10abdcc448ada8` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 37 files and 6,331 changed
lines, comprising 5,378 additions and 953 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`, `rdocx-py`
**Pass authority**: pass 7 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding closure

- **Pass 6 B1 is closed.** The dedicated URL extractor enters a quoted branch,
  finds the closing matching quote before the function delimiter, preserves
  inner parentheses, permits only whitespace before the closing parenthesis,
  and rejects trailing syntax (`crates/rdocx/src/html.rs:1444`,
  `crates/rdocx/src/html.rs:1450`, `crates/rdocx/src/html.rs:1460`). The common
  URL scan uses that extractor for ordinary declarations and the URL form of
  `@import`, while the import scan skips the already processed URL form and
  retains its separate string form (`crates/rdocx/src/html.rs:1498`,
  `crates/rdocx/src/html.rs:1507`, `crates/rdocx/src/html.rs:1516`). Updating
  the scan position to the complete function end preserves multiple-resource
  traversal (`crates/rdocx/src/html.rs:1501`). The regression rejects the
  quoted prefix collision and independently asserts the complete URL with an
  inner parenthesis (`crates/rdocx/src/html.rs:3535`,
  `crates/rdocx/src/html.rs:3573`).
- **Pass 5 B1 remains closed.** CSS containing escapes is decoded only to
  identify resource syntax, then conservatively rejected before literal
  extraction (`crates/rdocx/src/html.rs:1488`). The malformed matrix covers
  escaped `url` and `@import` identifiers, an escaped URL delimiter, and an
  escaped quoted-import delimiter (`crates/rdocx/src/html.rs:3513`,
  `crates/rdocx/src/html.rs:3523`, `crates/rdocx/src/html.rs:3529`).
- **Pass 4 B2 remains closed.** `Error::Mhtml` and
  `InvalidEmbeddedMutation` map to the existing `RdocxError` class
  (`crates/rdocx-py/src/lib.rs:66`), and the exact class regression covers both
  variants (`crates/rdocx-py/src/lib.rs:137`). The adapter compiles across all
  targets.
- **Pass 3 B1 remains closed.** The loss walk distinguishes shapes, other
  drawings, linked images, unresolved images, and supported embedded images
  (`crates/rdocx/src/html.rs:597`). The source-built loss regression and the
  independent reader regression cover those classifications
  (`crates/rdocx/src/html.rs:3739`, `crates/rdocx/src/run.rs:969`).
- **Pass 1 B1 and pass 2 B1 remain closed.** Direct fetching elements,
  responsive attributes, legacy `background`, literal resource functions,
  quoted URL functions, and string-form imports enter bounded preflight
  (`crates/rdocx/src/html.rs:1786`, `crates/rdocx/src/html.rs:1851`,
  `crates/rdocx/src/html.rs:1885`, `crates/rdocx/src/html.rs:1898`).
- **Pass 1 B2 remains closed.** Import and export admit only sniffed PNG and
  JPEG resources (`crates/rdocx/src/html.rs:442`,
  `crates/rdocx/src/html.rs:1824`). The positive and negative two-direction
  format matrix passes, including successful embedded PNG and JPEG round trips
  (`crates/rdocx/src/html.rs:3609`).
- **Pass 1 B3 remains closed.** The ordinary differential uses one shared,
  mutation-sensitive acceptance predicate for the rdocx and pinned Word
  records (`crates/rdocx/tests/integration_test.rs:102`,
  `crates/rdocx/tests/integration_test.rs:215`). The source is built in code,
  the Word version is pinned, and the intentional image divergence is asserted
  rather than hidden (`crates/rdocx/tests/integration_test.rs:16`,
  `crates/rdocx/tests/integration_test.rs:120`).
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
(`docs/sprints/CURRENT_SPRINT.md:36`). Within the completed prefix, F-X077's
shared validator and owner adapters retain executable focused evidence, and
F-239's resource preflight, Python consumer, differential oracle, image formats,
nested diagnostics, and ordinary MHTML tests pass.

## Not found

- **Interaction**: zero cross-feature conflicts were found. F-X077 remains the
  lowest shared XML lexical owner, and F-239 remains in the native HTML and MIME
  facade.
- **Duplication**: zero sprint-local duplicate lexical helpers were found. The
  glossary, embedded-content, and package-story owners use the shared
  `validate_strict_xml_1_0` policy (`crates/oxml-core/src/xml.rs:37`).
- **Layering**: zero dependency-direction violations were found. No manifest
  changed, and no `oxml-*` crate gained a dependency on `rdocx-*` or `rpptx-*`.
- **Harness**: zero unexplained output deltas were found. The independent check
  reports all 49 entries match, consistent with both completion records
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`).
- **Docs**: zero HLD or ledger inconsistencies were found. The sprint contract
  records F-X077 and F-239 as done and retains F-238 as pending
  (`docs/sprints/CURRENT_SPRINT.md:34`, `docs/sprints/CURRENT_SPRINT.md:36`).
- **Dependencies**: zero new dependency or feature findings were found. The
  integrated prefix changes no manifest.
- **Surface**: zero unplanned API findings were found. MHTML remains native
  only, and the Python adapter retains its existing generic error surface
  (`docs/hld/10-bindings-spec.md:302`, `crates/rdocx-py/src/lib.rs:76`).

Focused evidence passed for all five MHTML unit tests, both ordinary MHTML
integration tests, the Python generic error-class test, `cargo check -p
rdocx-py --all-targets`, the drawing classification regression, the shared XML
unit matrix, all 42 glossary tests, the embedded owner mapping test, all five
package-story lexical tests, the 49-entry hash harness, and `git diff --check`.
The ignored Microsoft Word regeneration test was not run in this pass.
