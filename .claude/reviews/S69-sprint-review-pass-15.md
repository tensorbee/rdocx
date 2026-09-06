# S69 sprint review, pass 15

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`ba96a887bec544106bedaf5b156b6ce878d84d1a` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 101 files and 10,912 changed
lines, comprising 9,777 additions and 1,135 deletions. The 20 touched crates
are `oxml-chart`, `oxml-cli-support`, `oxml-core`, `oxml-drawing`,
`oxml-layout`, `oxml-media`, `oxml-opc`, `oxml-pdf`, `oxml-sml`, `rdocx`,
`rdocx-oxml`, `rdocx-py`, `rdocx-wasm`, `rpptx`, `rpptx-chart`, `rpptx-cli`,
`rpptx-layout`, `rpptx-oxml`, `rpptx-render`, and `rpptx-wasm`.
**Pass authority**: pass 15 extends the default three-pass bound under the
user's explicit authorization to run as many review and remediation passes as
required.
**Verdict**: 3 blocking, 1 should-fix, 0 nice-to-have

## Blocking

### B1, the composite M22 gate is not mutation-sensitive for three required operations

The milestone requires the representative document to rebuild a table of
contents, perform advanced merge, and perform full document comparison
(`docs/hld/14-development-backlog.md:2079`). The new composite test calls
`rebuild_toc`, but asserts only the returned entry count before serializing
(`crates/rdocx/tests/integration_test.rs:390`). It never checks that the stale
cache was replaced in `post_toc_bytes` (`crates/rdocx/tests/integration_test.rs:392`).
A regression that computes the report without applying the cache update would
therefore pass this end gate.

The same test proves that two merge values occur, but does not assert the
section boundaries promised by its sectioned merge call
(`crates/rdocx/tests/integration_test.rs:417`). It also gives comparison both a
header edit and a body addition, then asserts only that some revision exists
(`crates/rdocx/tests/integration_test.rs:435`) and that the header contains an
insertion (`crates/rdocx/tests/integration_test.rs:479`). A comparison regression
that drops the body addition still passes. Make the single composed gate inspect
the rebuilt TOC cache, merged section structure, and body comparison result so
each required operation can independently break the predicate.

### B2, inherited namespaces used by markup-compatibility values are still lost

The remediation finds required inherited declarations only from element and
attribute qualified names (`crates/rdocx/src/flat_opc.rs:684`,
`crates/rdocx/src/flat_opc.rs:696`). Attribute values are never inspected before
the selected declarations are copied to the payload root
(`crates/rdocx/src/flat_opc.rs:705`). That is insufficient for OOXML markup
compatibility, where `mc:Ignorable`, `mc:MustUnderstand`, `mc:ProcessContent`,
`mc:PreserveElements`, and `mc:PreserveAttributes` carry namespace prefixes in
their values. The existing reader resolves those values against the active
namespace scope and treats an unbound prefix as invalid
(`crates/rdocx/src/embedded.rs:1935`, `crates/rdocx/src/embedded.rs:1960`).

A valid Flat OPC wrapper can therefore supply `xmlns:w14` on `pkg:xmlData` while
the payload root contains only `mc:Ignorable="w14"`. Import materializes
`xmlns:mc`, because it occurs in an attribute name, but drops `xmlns:w14`,
changing the standalone part's compatibility semantics. The focused regression
covers a prefix used in element names and a default namespace used in element
names, but no namespace used only by a QName-valued attribute
(`crates/rdocx/tests/integration_test.rs:223`). Preserve every inherited binding
needed by namespace-bearing values and add an exact-byte round trip for this
case.

### B3, full verification is not recorded at the remediation HEAD

Every verification and review record is bound to its exact HEAD, and checkpoint
evidence cannot close a later changed HEAD (`.claude/commands/run-sprint.md:224`).
The canonical state ends its successful full-verification history at
`4264cc3ccaee23d50d91f414a167d4732d88fb58`
(`.claude/scratch/S69-run.json:260`), before the production parser and composite
gate changed in `ba96a88`. The focused tests below are green, but they are not
the workspace, WASM, rustdoc, package, archive, dependency, and supply-chain
gate required after an integrated change (`.claude/commands/run-sprint.md:244`).
Run and record `/verify --full` at the post-remediation candidate before another
clean review or release approval.

## Should-fix

### S1, AS_BUILT attributes the post-remediation test inventory to the old SHA

The completion record now names the composite M22 test and the inherited
namespace and alternative-format coverage
(`docs/sprints/AS_BUILT.md:11915`), but its adjacent verification sentence still
claims that all tests passed at pre-remediation `f4c5bb0`
(`docs/sprints/AS_BUILT.md:11927`). The named composite test now begins at
`crates/rdocx/tests/integration_test.rs:340` and was introduced by the later
remediation. After B3 is satisfied, bind this consolidated evidence to the
actual verified remediation SHA rather than describing the earlier feature
integration as having executed tests that did not yet exist.

## Nice-to-have

None.

## Pass 14 remediation

- **B1 remains open.** The new source-built DOTM test executes all named M22
  APIs and preserves equations, executable bytes, unsupported XML, and class
  identity through Flat OPC (`crates/rdocx/tests/integration_test.rs:340`). B1
  identifies the remaining mutation-insensitive assertions.
- **B2 remains open.** Wrapper declarations used by element and attribute names
  are now materialized and the part-size limit is rechecked after insertion
  (`crates/rdocx/src/flat_opc.rs:590`, `crates/rdocx/src/flat_opc.rs:309`). B2
  identifies the remaining namespace-bearing attribute-value case.
- **B3 is closed.** Both Transitional and Strict internal `aFChunk`
  relationships classify their resolved targets as opaque before the MIME
  heuristic (`crates/rdocx/src/flat_opc.rs:992`). Export and import use that same
  set (`crates/rdocx/src/flat_opc.rs:416`, `crates/rdocx/src/flat_opc.rs:453`).
  The two-URI regression verifies byte-exact XHTML through ZIP, Flat OPC, and
  ZIP while an ordinary `application/xhtml+xml` part remains XML
  (`crates/rdocx/tests/integration_test.rs:270`).

## Milestone gate

The M22 gate is a representative modern document that authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads
(`docs/hld/14-development-backlog.md:2079`). The new test supplies one composed
representative and passes, but the gate does not yet hold as tested because B1
shows that TOC mutation, sectioned merge structure, and the body-comparison
branch can regress without failing its predicate.

## Prior finding closure and integrated records

- Passes 1 through 7 remain closed. The full prefix retains fail-closed MHTML
  preflight for `background`, CSS imports, escaped and quoted resource syntax,
  and unresolved resources (`crates/rdocx/src/html.rs:1488`,
  `crates/rdocx/src/html.rs:1885`). The pinned common-input differential and its
  perturbation matrix remain in place (`crates/rdocx/tests/integration_test.rs:1021`).
- F-X077's shared strict XML validator still checks namespace bindings and
  duplicate expanded attributes (`crates/oxml-core/src/xml.rs:154`). The F-238
  remediation reuses that validator and adds no second lexical policy
  (`crates/rdocx/src/flat_opc.rs:54`).
- F-X080 remains intact. The package job names all 24 bundled fonts and six
  legal files (`.github/workflows/ci.yml:533`, `.github/workflows/ci.yml:567`).
  The Pandoc installer retains its pinned digest, download and extraction
  bounds, exact authenticated aliases, and fail-closed member handling
  (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:30`,
  `scripts/install_pinned_pandoc.py:56`). The Python adapter still maps both new
  native error variants exhaustively (`crates/rdocx-py/src/lib.rs:66`).
- Passes 8 and 9 S1 remain closed. The aggregate CI gate names the filtered jobs
  it actually checks without claiming ownership of the independent package job
  (`.github/workflows/ci.yml:656`).
- F-X079's released state remains truthful in the HLD, including the immutable
  annotated `rpptx-v0.10.0` tag (`docs/hld/15-build-and-toolchain.md:343`). The
  live sprint correctly leaves F-X078 pending after the five completed prefix
  stories (`docs/sprints/CURRENT_SPRINT.md:34`).

## Not found

- **Interaction**: zero additional interaction findings beyond B1 and B2.
- **Duplication**: zero findings. Flat OPC still projects into `OpcPackage`, and
  F-X077 remains the sole strict lexical validator.
- **Layering**: zero findings. No `oxml-*` crate gained a forbidden `rdocx-*` or
  `rpptx-*` dependency.
- **Harness**: zero output-delta findings. The focused hash check reports 49 of
  49 entries unchanged. B3 covers the missing exact-HEAD full gate record.
- **Gate**: two gate findings, B1 and B3. Zero additional gate findings.
- **Docs and ledgers**: one should-fix finding, S1. The Flat OPC HLD correctly
  records relationship-owned alternative-format opacity and inherited binding
  materialization (`docs/hld/04-opc-and-packaging.md:39`). B2 is a code and test
  mismatch with that current-intent statement.
- **Dependencies**: zero findings. The remediation adds no dependency, feature,
  crate, module, trait, or generic parameter.
- **Surface**: zero findings. The remediation adds no public API.
- **Limits and security**: zero additional findings. Decoded part and total
  bounds are rechecked after namespace materialization
  (`crates/rdocx/src/flat_opc.rs:309`), and external relationships remain
  excluded from opaque-target classification
  (`crates/rdocx/src/flat_opc.rs:999`).
- **CI and release**: zero new CI topology, package inventory, Pandoc, Python
  mapping, version-carrier, release-note, or published-state findings beyond B3
  and S1.

Focused evidence passed all 11 runnable F-238 package-class integration tests,
including the composite gate and the new namespace and aFChunk regressions. The
pinned Word GUI test remained correctly ignored. The 49-entry hash harness,
agent-skill synchronization, prose check, and full-prefix `git diff --check`
also passed. These focused results do not close B1 through B3.
