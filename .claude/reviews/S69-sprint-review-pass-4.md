# S69 sprint review, pass 4

**Reviewed**: completed dependency prefix on `sprint/s69` at
`aece74b41d6432d1b44ad67c64b195846b311518` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 33 files and 5,610 changed
lines, comprising 4,756 additions and 854 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`
**Pass authority**: pass 4 extends the default three-pass bound under the
user's explicit instruction on 2026-09-05 to run as many passes as required.
**Verdict**: 2 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, escaped CSS resource tokens bypass fail-closed preflight
`crates/rdocx/src/html.rs:1397`
`crates/rdocx/src/html.rs:1401`
`crates/rdocx/src/html.rs:1416`
`docs/hld/04-opc-and-packaging.md:557`

The resource preflight searches a lowercased byte string only for literal
`url(` and `@import` spellings. CSS identifiers allow escapes, so a valid
resource function such as `u\72l(https://outside.test/a.png)` is not returned
by `css_resource_references`. The normal HTML projection then diagnoses or
drops the unsupported `background-image` declaration and can publish the
document, even though the MHTML contract requires every resource reference to
pass bounded preflight before publication. This reopens the original resource
coverage boundary from pass 1 B1 beyond the literal forms added after pass 2.
The fix must tokenize CSS resource constructs or conservatively reject escaped
resource-bearing CSS, with a regression proving an escaped external and an
escaped unresolved contained reference both fail before projection.

### B2, the new MHTML error variant is absent from the Python error adapter
`crates/rdocx/src/error.rs:38`
`crates/rdocx-py/src/lib.rs:67`
`docs/hld/10-bindings-spec.md:310`

F-239 adds `Error::Mhtml`, but the in-workspace Python adapter exhaustively
matches `rdocx::Error` without an `Mhtml` arm. `cargo check -p rdocx-py` at the
reviewed SHA reports E0004 for `Mhtml` in addition to a separate pre-existing
unmatched variant. MHTML correctly remains native-only and needs no Python
method or exception class, but the existing adapter must still compile against
the facade it consumes. Map `Mhtml` to the existing `RdocxError` class without
adding a binding entry point or error type.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding closure

- **Pass 3 B1 is closed.** The diagnostic walk now distinguishes embedded and
  linked images, unresolved image relationships, shapes, and other drawings
  (`crates/rdocx/src/html.rs:597`). The source-built regression asserts the
  linked, unresolved, shape, and chart records in source order
  (`crates/rdocx/src/html.rs:3667`). The reader regression independently pins
  embedded and linked relationship classification, filled-shape
  classification, and chart classification (`crates/rdocx/src/run.rs:969`).
  Supported embedded images with package bytes take the explicit no-diagnostic
  branch (`crates/rdocx/src/html.rs:602`), and the PNG and JPEG export matrix
  plus integration path return empty export diagnostics while preserving the
  image (`crates/rdocx/src/html.rs:3483`,
  `crates/rdocx/tests/integration_test.rs:182`).
- **Pass 1 B1 and pass 2 B1 are reopened only for B1 above.** Direct fetching
  elements, responsive attributes, legacy `background`, literal `url(...)`,
  and quoted string-form `@import` remain covered. Escaped CSS token spellings
  are the remaining demonstrated bypass.
- **Pass 1 B2 remains closed.** Import and export admit only sniffed PNG and
  JPEG resources (`crates/rdocx/src/html.rs:442`,
  `crates/rdocx/src/html.rs:1733`), and the two-direction format matrix passes
  (`crates/rdocx/src/html.rs:3483`).
- **Pass 1 B3 remains closed.** The ordinary differential consumes the pinned
  Word record through one mutation-sensitive acceptance predicate
  (`crates/rdocx/tests/integration_test.rs:215`).
- **Pass 1 B4 is closed by the pass 2 nested walk and the pass 3 drawing
  remediation.** Direct runs, hyperlinks, tables, and drawing states now have
  source-ordered loss handling (`crates/rdocx/src/html.rs:548`,
  `crates/rdocx/src/html.rs:627`, `crates/rdocx/src/html.rs:684`).
- **Pass 2 S1 remains closed.** The completion record attributes the remediated
  Word oracle and full verification to `7fde4033b7cdf17f7c6e309dfccf7d1b9a6b1d44`
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
adapters retain executable focused evidence. F-239's drawing-loss remediation,
ordinary MHTML tests, and unchanged hash harness pass, but its resource
preflight and binding-consumer integration remain blocked by B1 and B2.

## Not found

- **Interaction**: zero additional cross-feature conflicts were found. F-X077
  remains the lowest shared XML lexical owner, and F-239 stays in the native
  HTML and MIME facade.
- **Duplication**: zero sprint-local duplicate lexical helpers were found. The
  three owners route shared lexical policy through
  `validate_strict_xml_1_0` (`crates/oxml-core/src/xml.rs:37`).
- **Layering**: zero dependency-direction violations were found. No manifest
  changed, and no `oxml-*` crate gained a dependency on `rdocx-*` or
  `rpptx-*`.
- **Harness**: zero unexplained hash deltas were found. The independent check
  reports all 49 entries match, consistent with the completion entries
  (`docs/sprints/AS_BUILT.md:11736`, `docs/sprints/AS_BUILT.md:11785`).
- **Dependencies**: zero new dependency or feature findings were found. The
  prefix changes no manifest.
- **Docs and ledgers**: zero additional documentation or ledger findings were
  found beyond B1 and B2. F-X077 and F-239 remain consistently recorded as done
  (`docs/sprints/BACKLOG.md:445`, `docs/sprints/SPRINT_TRACKER.md:391`).
- **Surface**: zero unplanned native API findings were found beyond the
  downstream adapter omission in B2. The XML lexical and MHTML values and
  methods match their approved pre-1.0 designs.

Focused evidence passed for all five MHTML unit tests, both ordinary MHTML
integration tests, the drawing classification regression, the shared XML unit
matrix, all 42 glossary tests, the embedded owner mapping test, all five
package-story lexical tests, the 49-entry hash harness, and `git diff --check`.
The ignored Microsoft Word regeneration test was not run in this pass.
`cargo check -p rdocx-py` failed as recorded in B2.
