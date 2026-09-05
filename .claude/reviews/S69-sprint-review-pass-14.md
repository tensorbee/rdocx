# S69 sprint review, pass 14

**Reviewed**: full integrated sprint prefix on `sprint/s69` at
`4264cc3ccaee23d50d91f414a167d4732d88fb58` against merge base
`c8908d077f0bb6a1649aa1265548e67fb6342c4b`, 100 files and 10,216 changed
lines, comprising 9,081 additions and 1,135 deletions. The 20 touched crates
are `oxml-chart`, `oxml-cli-support`, `oxml-core`, `oxml-drawing`,
`oxml-layout`, `oxml-media`, `oxml-opc`, `oxml-pdf`, `oxml-sml`, `rdocx`,
`rdocx-oxml`, `rdocx-py`, `rdocx-wasm`, `rpptx`, `rpptx-chart`, `rpptx-cli`,
`rpptx-layout`, `rpptx-oxml`, `rpptx-render`, and `rpptx-wasm`.
**Pass authority**: pass 14 extends the default three-pass bound under the
user's explicit authorization to run as many passes and remediations as
required.
**Verdict**: 3 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, the required representative M22 end gate does not exist

The milestone contract requires one representative modern document to combine
equation authoring and rendering, field and table-of-contents rebuilding,
advanced merge and comparison, embedded-content inventory, and a modern package
round trip that retains unsupported XML and executable payloads
(`docs/hld/14-development-backlog.md:2079`,
`docs/sprints/CURRENT_SPRINT.md:76`). The full prefix has no source-built test,
fixture, script, or recorded acceptance run that composes those operations in
one document. The completed F-238 record names only its package-class gate and
focused Flat OPC matrices (`docs/sprints/AS_BUILT.md:11913`), and its source
fixture covers package identity, VBA, and unrelated XML without the other M22
behaviours (`crates/rdocx/tests/integration_test.rs:53`).

This is not satisfied by the green workspace gate. That gate proves the
individual feature suites pass, while the end gate is explicitly a composed
interaction test. F-X078 is forbidden to begin until this gate passes
(`docs/sprints/CURRENT_SPRINT.md:51`), and the release contract repeats that
ordering (`docs/hld/14-development-backlog.md:3779`). Add and record one
source-built representative M22 test that executes the complete required path
before version preparation.

### B2, Flat OPC import loses inherited payload namespace declarations

The parser permits namespace declarations on `pkg:package`, `pkg:part`, and
`pkg:xmlData` as non-semantic structural attributes
(`crates/rdocx/src/flat_opc.rs:121`, `crates/rdocx/src/flat_opc.rs:190`,
`crates/rdocx/src/flat_opc.rs:545`). `read_xml_data` then copies only events
inside `pkg:xmlData` into a new standalone buffer and validates that buffer
(`crates/rdocx/src/flat_opc.rs:238`, `crates/rdocx/src/flat_opc.rs:253`,
`crates/rdocx/src/flat_opc.rs:284`). It never transfers namespace bindings
inherited from any Flat OPC wrapper. A valid package such as
`<pkg:xmlData xmlns:w="..."><w:document/></pkg:xmlData>` therefore enters with
a bound `w` prefix but is revalidated as the invalid standalone fragment
`<w:document/>`. The shared validator correctly rejects that unbound prefix
(`crates/oxml-core/src/xml.rs:154`).

The named namespace regression aliases only the `pkg:` wrapper and adds unused
local declarations, so it does not exercise a payload root that uses an
inherited declaration (`crates/rdocx/tests/integration_test.rs:165`,
`crates/rdocx/tests/integration_test.rs:180`). This misses the design's promised
prefix-alias and default-namespace coverage (`.claude/plans/F-238-design.md:115`).
Preserve the in-scope inherited bindings on the extracted payload root, with
positive tests for both a used inherited prefix and a used inherited default
namespace.

### B3, XML media-type classification corrupts or rejects alternative-format import parts

Both Flat OPC directions select `xmlData` or `binaryData` solely from the part
content type (`crates/rdocx/src/flat_opc.rs:136`,
`crates/rdocx/src/flat_opc.rs:421`). The classifier treats every `+xml` media
type as XML (`crates/rdocx/src/flat_opc.rs:766`). Word alternative-format import
parts are relationship-defined opaque parts and may legitimately use
`application/xhtml+xml`. Word requires those `aFChunk` targets to use
`pkg:binaryData` in Flat OPC. The current importer instead demands `xmlData`
(`crates/rdocx/src/flat_opc.rs:171`), while the exporter emits `xmlData` and
parses the payload as package XML (`crates/rdocx/src/flat_opc.rs:455`). That
rejects a Word-compatible Flat OPC input and produces a repair-required output
for a valid DOCX carrying a well-formed XHTML chunk.

The source-built package fixture covers generic `application/xml`, `text/xml`,
VBA, and octet-stream parts but has no alternative-format relationship
(`crates/rdocx/tests/integration_test.rs:53`). Microsoft Open XML SDK issue 525
and its shipped fix 659 document this exact interoperability rule and repair
failure. Classify `aFChunk` relationship targets as opaque before applying the
media-type heuristic, cover both Transitional and Strict relationship forms,
and add ZIP to Flat OPC to ZIP tests with a byte-sensitive XHTML payload. The
HLD's unqualified XML-versus-opaque sentence must record the relationship-owned
exception (`docs/hld/04-opc-and-packaging.md:45`).

## Should-fix

None.

## Nice-to-have

None.

## Prior finding closure

- Pass 13's pass 11 B1 closure remains intact. The state records successful
  full verification at the exact reviewed HEAD with all 49 hashes unchanged
  (`.claude/scratch/S69-run.json:253`).
- Pass 11 B2 and S1 remain closed. The selected-family release notes keep the
  stable family outside F-X079 publication authority and identify its current
  shared pins without claiming a stable release (`CHANGELOG.md:46`).
- Passes 8 and 9 S1 remain closed. The CI gate still describes its actual
  aggregate dependencies without claiming that the independent package job is
  included (`.github/workflows/ci.yml:656`).
- Passes 1 through 7 remain closed. F-238 and its ledger integration do not
  alter the remediated MHTML resource preflight, quoted CSS scanning, escape
  rejection, nested loss diagnostics, supported image boundary, shared lexical
  mapping, or pinned differential record.
- F-X080 remains intact. The package job enumerates the 24 fonts and six legal
  files (`.github/workflows/ci.yml:533`, `.github/workflows/ci.yml:567`), and the
  Pandoc installer retains its reviewed digest, download and extraction bounds,
  exact two authenticated aliases, and fail-closed member handling
  (`scripts/install_pinned_pandoc.py:19`,
  `scripts/install_pinned_pandoc.py:26`,
  `scripts/install_pinned_pandoc.py:29`,
  `scripts/install_pinned_pandoc.py:79`,
  `scripts/install_pinned_pandoc.py:88`).

## Integrated prefix and release records

- F-X077, F-239, F-X080, F-X079, and F-238 are consistently completed across
  the live sprint, backlog, tracker, and completion log
  (`docs/sprints/CURRENT_SPRINT.md:34`, `docs/sprints/BACKLOG.md:444`,
  `docs/sprints/BACKLOG.md:532`, `docs/sprints/SPRINT_TRACKER.md:391`,
  `docs/sprints/AS_BUILT.md:11882`). F-X078 remains correctly pending
  (`docs/sprints/CURRENT_SPRINT.md:39`).
- The M22 story ledger is 12 of 12 complete (`docs/sprints/BACKLOG.md:40`). That
  status does not waive the separate end-of-milestone gate described in B1.
- F-X079's selected 15-package 0.10.0 release record, annotated tag, hosted
  publish run, registry ownership, release body, exclusions, and empty
  contribution notification inventory remain unchanged from the independently
  verified pass-13 boundary (`docs/sprints/AS_BUILT.md:11838`).
- Stable source and published-state boundaries remain distinct. The changelog
  records stable v0.12.0 as outside F-X079 authority
  (`CHANGELOG.md:46`), while F-X078 owns the future exact seven-package v0.13.0
  preparation and publication (`docs/hld/14-development-backlog.md:3770`).

## Not found

- **Interaction**: zero additional interaction findings were found beyond B1,
  B2, and B3. F-X077's validator, F-239's MHTML boundary, F-X080's CI repairs,
  and F-X079's published state otherwise remain compatible with F-238.
- **Duplication**: zero duplication findings were found. F-238 reuses
  `OpcPackage`, package limits, signature invalidation, and atomic publication.
- **Layering**: zero layering findings were found. No `oxml-*` crate gained a
  forbidden stable-format dependency, and Flat OPC remains a private `rdocx`
  module.
- **Harness**: zero harness findings were found. The exact-HEAD full gate records
  49 of 49 deterministic hashes unchanged
  (`.claude/scratch/S69-run.json:253`).
- **Gate**: one gate finding was found, B1. Zero additional gate failures were
  found.
- **Docs and ledgers**: zero separate ledger inconsistencies were found. B3
  requires the Flat OPC HLD sentence to be made precise as part of the code fix.
- **Dependencies**: zero dependency findings were found. The integrated prefix
  adds no unapproved external dependency or feature flag.
- **Surface**: zero public-surface findings were found. F-238's additive native
  API stays within its approved plan, and bindings remain unchanged.
- **CI and release**: zero separate CI-repair or F-X079 publication findings were
  found.

Focused evidence passed the nine F-238 package-class integration tests, with
eight passed and the pinned Word GUI acceptance test correctly ignored, the
F-239 unsafe and over-limit MHTML unit regression, and the F-X077 strict lexical
matrix. The repository records `/verify --full` green at the exact reviewed
HEAD with all 49 hashes unchanged (`.claude/scratch/S69-run.json:253`). These
results do not cover the absent composite gate or the two Flat OPC cases above.
