# S69 sprint review, pass 16

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`ee16be81eb3bf34046f5c009c966b01618730f85` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 102 files and 11,219 changed
lines, comprising 10,084 additions and 1,135 deletions. The 20 touched crates
are `oxml-chart`, `oxml-cli-support`, `oxml-core`, `oxml-drawing`,
`oxml-layout`, `oxml-media`, `oxml-opc`, `oxml-pdf`, `oxml-sml`, `rdocx`,
`rdocx-oxml`, `rdocx-py`, `rdocx-wasm`, `rpptx`, `rpptx-chart`, `rpptx-cli`,
`rpptx-layout`, `rpptx-oxml`, `rpptx-render`, and `rpptx-wasm`.
**Pass authority**: pass 16 extends the default three-pass bound under the
user's explicit authorization to run as many review and remediation passes as
required.
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Pass 15 closure

- **B1 is closed.** The composed M22 predicate now proves that TOC rebuilding
  removes the stale cache, writes the heading, and creates the page-reference
  field (`crates/rdocx/tests/integration_test.rs:411`). It proves that two
  sectioned merge records produce exactly one next-page boundary
  (`crates/rdocx/tests/integration_test.rs:445`). It also proves that the body
  addition and body insertion markup survive the final Flat OPC round trip,
  separately from the asserted header insertion
  (`crates/rdocx/tests/integration_test.rs:507`). Each operation identified by
  pass 15 can now independently break the gate predicate.
- **B2 is closed.** Namespace selection recognizes both prefix-list MC
  attributes, all three QName-list MC attributes, and unprefixed `Requires` on
  an MC `Choice` element (`crates/rdocx/src/flat_opc.rs:674`,
  `crates/rdocx/src/flat_opc.rs:711`). It resolves each value prefix through the
  payload-local, ancestor, and inherited scopes before materializing a wrapper
  binding (`crates/rdocx/src/flat_opc.rs:731`,
  `crates/rdocx/src/flat_opc.rs:751`). The focused regression exercises
  `Ignorable`, `MustUnderstand`, `ProcessContent`, `PreserveElements`,
  `PreserveAttributes`, and `Choice Requires`, then compares the reopened part
  with the complete expected bytes and all six bindings
  (`crates/rdocx/tests/integration_test.rs:274`).
- **B3 is closed.** The canonical state records `/verify --full` passed at the
  exact remediation implementation SHA
  `754c117af6cf8d1cb26e87023c1da9a78e018651`, with all 49 harness entries
  unchanged (`.claude/scratch/S69-run.json:273`). The only later commit changes
  the AS_BUILT evidence binding. Current-prefix skill synchronization, prose,
  and `git diff --check` also pass.
- **S1 is closed.** The completion record now attributes the pass-15 mutation
  and namespace work to the remediation and binds the consolidated test, WASM,
  rustdoc, dependency-direction, packaging, archive-size, and supply-chain
  evidence to exact verified SHA `754c117af6cf8d1cb26e87023c1da9a78e018651`
  (`docs/sprints/AS_BUILT.md:11906`, `docs/sprints/AS_BUILT.md:11930`). It no
  longer claims that the later composite inventory ran at the original F-238
  feature SHA.

## Milestone gate

The M22 gate requires one representative modern document to author and render
equations, rebuild fields and a table of contents, perform advanced merge and
comparison, inventory embedded content, and round-trip its modern package
variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`).

The gate holds. The source-built
`representative_m22_document_composes_the_complete_milestone_gate` test authors
and deterministically renders a fraction, rebuilds and inspects the TOC cache,
updates the merge field, creates and inspects the sectioned merge, inventories
the retained VBA project, compares body and header edits, and checks DOTM
identity, equations, exact VBA bytes, and unsupported XML after Flat OPC reopen
(`crates/rdocx/tests/integration_test.rs:360`,
`crates/rdocx/tests/integration_test.rs:403`,
`crates/rdocx/tests/integration_test.rs:411`,
`crates/rdocx/tests/integration_test.rs:423`,
`crates/rdocx/tests/integration_test.rs:445`,
`crates/rdocx/tests/integration_test.rs:464`,
`crates/rdocx/tests/integration_test.rs:468`,
`crates/rdocx/tests/integration_test.rs:483`). It passed in this review and in
the exact-SHA full gate cited above. The HLD records both the complete operation
set and its mutation-sensitive predicates as current testing policy
(`docs/hld/12-testing-strategy.md:1266`).

This clean dependency-prefix gate makes F-X078 eligible. It does not claim that
the sprint's stable v0.13.0 publication definition of done is already complete.
F-X078 remains correctly pending and requires its own reviewed preparation,
full verification, separate final release approval, publication, and external
verification (`docs/sprints/CURRENT_SPRINT.md:39`,
`docs/hld/14-development-backlog.md:3779`).

## Prior closures and integrated records

- Pass 14 B3 remains closed. Transitional and Strict `aFChunk` relationships
  resolve to opaque targets before content-type classification for import and
  export, while unrelated XHTML remains XML
  (`crates/rdocx/src/flat_opc.rs:416`,
  `crates/rdocx/src/flat_opc.rs:453`,
  `crates/rdocx/tests/integration_test.rs:296`).
- Passes 1 through 7 remain closed. The full prefix retains fail-closed MHTML
  resource preflight, quote-aware CSS scanning, escaped-resource rejection,
  nested loss reporting, supported image handling, and the pinned common-input
  differential. The focused unsafe and over-limit MHTML regression passed.
- F-X077 remains the sole shared strict lexical policy. The exact declaration,
  character, name, namespace, reference, comment, and processing-instruction
  matrix passed, while Flat OPC reuses the same public validator
  (`crates/oxml-core/src/xml.rs:37`, `crates/rdocx/src/flat_opc.rs:54`).
- F-X080 remains intact. The package workflow enumerates all 24 bundled fonts
  and six legal files (`.github/workflows/ci.yml:533`,
  `.github/workflows/ci.yml:567`). The Pandoc installer retains its reviewed
  digest, download and extraction bounds, two authenticated aliases, and
  fail-closed member policy (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:79`). The three selected CI reconstruction
  regressions passed in this review.
- Passes 8 and 9 S1 remain closed. The aggregate CI gate names only the filtered
  jobs it checks and does not claim ownership of the independent package job
  (`.github/workflows/ci.yml:656`).
- F-X079's published state remains consistent across the HLD, changelog, and
  completion record. The exact 15-package family remains at 0.10.0 from
  immutable annotated `rpptx-v0.10.0` tag
  `1e409c553b950eb8029e3e78e39ff775f18ba3ab`, while stable source remains at
  0.12.0 with current shared 0.10.0 pins
  (`docs/hld/03-architecture.md:841`, `CHANGELOG.md:35`,
  `docs/sprints/AS_BUILT.md:11838`).
- The sprint ledgers agree. F-X077, F-239, F-X080, F-X079, and F-238 are done,
  F-X078 is pending, and M22's 12 stories are complete
  (`docs/sprints/CURRENT_SPRINT.md:34`, `docs/sprints/BACKLOG.md:40`,
  `docs/sprints/BACKLOG.md:532`, `docs/sprints/SPRINT_TRACKER.md:391`).

## Not found

- **Interaction**: zero findings. The composed M22 gate now tests the relevant
  F-228 through F-238 interactions, and F-239, F-X077, F-X079, and F-X080 do not
  conflict with its Flat OPC path.
- **Duplication**: zero findings. Flat OPC continues to reuse `OpcPackage`,
  package limits, signature invalidation, atomic publication, and the shared
  lexical validator.
- **Layering**: zero findings. No `oxml-*` crate gained a forbidden `rdocx-*` or
  `rpptx-*` dependency, and Flat OPC remains private to `rdocx` as specified
  (`docs/hld/03-architecture.md:893`).
- **Harness**: zero findings. The verified remediation reports 49 of 49 entries
  unchanged (`.claude/scratch/S69-run.json:273`).
- **Gate**: zero findings. The complete M22 predicate and exact implementation
  full gate both pass.
- **Docs and ledgers**: zero findings. The package and test HLD describe current
  behavior, the AS_BUILT evidence is SHA-bound, and story and release status
  agree across live records.
- **Dependencies**: zero findings. The prefix adds no unapproved external
  dependency or feature flag, and the only manifest changes are the reviewed
  0.10.0 incubating-family carriers.
- **Surface**: zero findings. F-238's additive native package-class and Flat OPC
  API, F-239's additive native MHTML API, and F-X077's shared validator are the
  surfaces called for by their approved stories. Binding surfaces remain at
  their documented boundary.
- **Limits and security**: zero findings. Flat OPC rechecks the decoded part
  bound after namespace insertion (`crates/rdocx/src/flat_opc.rs:309`), and
  external relationships cannot classify an opaque alternative-format target.
- **CI and release**: zero findings. Package inventory, Pandoc extraction,
  Python error mapping, selected-family notes, version carriers, and published
  F-X079 records remain coherent. F-X078 is scheduled work, not an omitted
  completed-prefix record.

Focused evidence passed all 11 runnable F-238 package-class tests, including
the composed M22, inherited MC namespace, and two-family `aFChunk` regressions.
The pinned Word 16.104 GUI acceptance remained correctly ignored because its
successful observation is already recorded at `docs/sprints/AS_BUILT.md:11928`.
The F-239 unsafe and over-limit MHTML regression, F-X077 shared lexical matrix,
three selected F-X080 and F-X079 workflow regressions, skill synchronization,
prose validation, and full-prefix `git diff --check` also passed.
