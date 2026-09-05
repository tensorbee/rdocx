# S69 sprint review, pass 5

**Reviewed**: full integrated dependency prefix on `sprint/s69` at
`1dae52ca2366caa4ec2d9e2f1803bfaba501e44a` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 35 files and 5,809 changed
lines, comprising 4,955 additions and 854 deletions, crates: `oxml-core`,
`rdocx-oxml`, `rdocx`, `rdocx-py`
**Pass authority**: pass 5 extends the default three-pass bound under the
user's explicit authorization on 2026-09-05 to run as many passes as required.
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, decoded CSS delimiters can validate only a resource prefix
`crates/rdocx/src/html.rs:1397`
`crates/rdocx/src/html.rs:1427`
`crates/rdocx/src/html.rs:1452`
`crates/rdocx/src/html.rs:1487`
`docs/hld/04-opc-and-packaging.md:557`

The pass 4 remediation decodes every CSS escape into an ordinary character
before it identifies URL and quoted import boundaries. That loses whether a
closing parenthesis or quote was escaped. For example, the valid CSS URL token
`url(https://example.test/allowed.png\)outside)` refers to
`https://example.test/allowed.png)outside`, but decoding produces an ordinary
inner `)`. The scanner stops there and validates only
`https://example.test/allowed.png`. If that prefix names a contained part, the
preflight passes even though the actual CSS resource is unresolved. A quoted
`@import` with an escaped quote has the same prefix-validation path because the
post-decode loop treats the produced quote as the string terminator. The
focused matrix covers escaped `url` and `@import` keyword spellings, but it does
not cover escapes that produce their delimiters. CSS token boundaries must be
recognized before escape decoding, or escape provenance must otherwise remain
available through reference extraction, with external and unresolved prefix
collision regressions for URL and string-form import.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding closure

- **Pass 4 B1 remains open only for B1 above.** Identifier escapes in `url` and
  `@import` are decoded and the new focused cases reject their external and
  unresolved forms (`crates/rdocx/src/html.rs:1397`,
  `crates/rdocx/src/html.rs:3471`). Escapes that produce syntax delimiters are
  the remaining demonstrated bypass.
- **Pass 4 B2 is closed.** `Error::Mhtml` and the previously unmatched
  `InvalidEmbeddedMutation` both map to the existing `RdocxError` class
  (`crates/rdocx-py/src/lib.rs:66`). The exact class regression covers both
  variants (`crates/rdocx-py/src/lib.rs:149`), and the adapter compiles across
  all targets.
- **Pass 3 B1 remains closed.** The loss walk distinguishes shapes, other
  drawings, linked images, unresolved images, and supported embedded images
  (`crates/rdocx/src/html.rs:597`). Its source-built regression asserts those
  losses in source order while supported siblings survive
  (`crates/rdocx/src/html.rs:3672`).
- **Pass 1 B1 and pass 2 B1 remain closed outside B1 above.** Direct fetching
  elements, responsive attributes, legacy `background`, literal `url(...)`,
  and quoted string-form `@import` are covered
  (`crates/rdocx/src/html.rs:1744`, `crates/rdocx/src/html.rs:1809`,
  `crates/rdocx/src/html.rs:1843`, `crates/rdocx/src/html.rs:1856`).
- **Pass 1 B2 remains closed.** Import and export admit only sniffed PNG and
  JPEG resources (`crates/rdocx/src/html.rs:446`,
  `crates/rdocx/src/html.rs:1782`), and their two-direction format matrix
  passes (`crates/rdocx/src/html.rs:3542`).
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
differential oracle, image formats, nested diagnostics, and ordinary MHTML
tests pass, but complete resource preflight remains blocked by B1.

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
  only, and the downstream Python change restores its existing generic error
  mapping without adding an entry point or exception class
  (`docs/hld/10-bindings-spec.md:302`, `crates/rdocx-py/src/lib.rs:76`).

Focused evidence passed for all five MHTML unit tests, both ordinary MHTML
integration tests, the Python generic error-class test, `cargo check -p
rdocx-py --all-targets`, the shared XML unit matrix, all 42 glossary tests, the
embedded owner mapping test, all five package-story lexical tests, the 49-entry
hash harness, and `git diff --check`. The ignored Microsoft Word regeneration
test was not run in this pass.
