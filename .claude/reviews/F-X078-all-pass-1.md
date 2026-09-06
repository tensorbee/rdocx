# F-X078, all, pass 1

**Reviewed**: Current uncommitted F-X078 working-tree diff on `sprint/s69` at
base HEAD `3e150ec3b8af7fbb89f1449e78199172d2570dc3`, 26 files and 304 changed
lines
**Verdict**: 0 defects, 0 smells, 0 nitpicks

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Not found

- **Correctness**: The workspace version is 0.13.0 and all nine stable-group
  dependency pins are 0.13.0, while every shared OOXML and PowerPoint pin
  remains at 0.10.0 (`Cargo.toml:34`, `Cargo.toml:55`, `Cargo.toml:78`). The
  carrier regression enumerates all eleven inherited packages, validates their
  lock entries, checks both Python project versions, and derives exactly seven
  publishable crates (`scripts/test_sprint_workflow.py:4964`,
  `scripts/test_sprint_workflow.py:5021`, `scripts/test_sprint_workflow.py:5046`,
  `scripts/test_sprint_workflow.py:5054`).
- **Contract**: The stable tag route publishes exactly `rdocx-opc`,
  `rdocx-oxml`, `rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and
  `rdocx-cli` in dependency order (`.github/workflows/publish.yml:61`). The
  separate incubating route remains at the 0.10.0 family boundary
  (`.github/workflows/publish.yml:78`). Python support, both Python bindings,
  and the rdocx WASM carrier remain ineligible for crates.io publication
  (`crates/oxml-py-support/Cargo.toml:6`, `crates/rpptx-py/Cargo.toml:6`,
  `crates/rdocx-py/Cargo.toml:6`, `crates/rdocx-wasm/Cargo.toml:14`).
- **Release notes and contribution inventory**: The notes describe the M22
  Word outcomes, name the exact seven-package stable family, identify the
  separately published shared 0.10.0 dependency, and preserve binding
  publication exclusions (`CHANGELOG.md:11`, `CHANGELOG.md:43`,
  `CHANGELOG.md:46`, `CHANGELOG.md:52`). The reviewed empty inventory is stated
  explicitly (`CHANGELOG.md:56`). GitHub Issue 68 was inspected and is a future
  JSON-like DSL request, not an implementation of the selected M22 changes.
  Its exclusion agrees with the durable inventory contract
  (`docs/hld/12-testing-strategy.md:1855`).
- **Tests**: The full 97-test sprint-workflow module passed with one expected
  skip. The release-note truth contract rejects contribution links and the
  wrong family tag (`scripts/test_sprint_workflow.py:4927`,
  `scripts/test_sprint_workflow.py:4959`). Publication routing has negative
  coverage for swapped predicates, extra packages, failure-tolerant commands,
  and preflight authority mutations (`scripts/test_sprint_workflow.py:7087`,
  `scripts/test_sprint_workflow.py:7109`, `scripts/test_sprint_workflow.py:7126`,
  `scripts/test_sprint_workflow.py:7153`). The README archive and doctest gate
  also passed.
- **Registry proof**: The opt-in proof passed against crates.io. It packages
  `rdocx-layout@0.13.0`, inspects the normalized archive for an exact
  path-free `oxml-layout@0.10.0` dependency, and resolves the packaged manifest
  without an `oxml-layout` patch (`scripts/test_sprint_workflow.py:5142`,
  `scripts/test_sprint_workflow.py:5174`, `scripts/test_sprint_workflow.py:5182`,
  `scripts/test_sprint_workflow.py:5188`).
- **HLD discipline**: The specifications distinguish prepared 0.13.0 source
  from the still-published 0.12.0 stable family, retain the published shared
  0.10.0 boundary, and keep bindings unpublished
  (`docs/hld/03-architecture.md:850`, `docs/hld/10-bindings-spec.md:1156`,
  `docs/hld/15-build-and-toolchain.md:449`). External actions remain owned by
  `/release`, which requires clean exact-SHA verification and review plus a
  separate final approval before mutation
  (`docs/hld/15-build-and-toolchain.md:479`,
  `docs/hld/15-build-and-toolchain.md:512`).
- **Panics, OOXML, and structure**: The diff changes release metadata, tests,
  release notes, and current-intent documentation. It adds no runtime parsing,
  OOXML serialization, trait, generic, wrapper, crate, module, or feature
  surface (`.claude/plans/F-X078-design.md:62`). No issue was found in these
  aspects.
