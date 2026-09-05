# F-X080, Restore CI release readiness

**Status**: approved
**Sprint**: S69
**Size**: S
**Depends on**: F-X077, F-239

## Problem

The latest `main` CI run has three persistent independent failures. The
`package-oxml-layout` job rejects four bundled Noto fonts and two Noto legal
files because its hard-coded expected inventory predates those assets
(`.github/workflows/ci.yml`, "Check package inventory"). The test job rejects
the authenticated Pandoc 3.10 archive because its 162,406,703 extracted bytes
exceed the 128 MiB ceiling (`scripts/install_pinned_pandoc.py`,
`MAX_EXTRACTED_BYTES`). The `rdocx` Python job cannot compile because the
exhaustive native error adapter omitted current error variants
(`crates/rdocx-py/src/lib.rs`, `impl From<rdocx::Error> for PyErr`).

These failures have persisted since S58, S65, and S68 respectively. They block
the aggregate `CI gate`, so neither S69 release should begin while the tracked
hosted contract is known to be red.

## Spec reference

- `docs/hld/12-testing-strategy.md`, "Binding tests" and "What CI runs".
- `docs/hld/14-development-backlog.md`, "F-X080, Restore CI release readiness".
- `docs/hld/15-build-and-toolchain.md`, "Toolchain pinning", "Packaging", and
  "CI job matrix".
- `docs/hld/10-bindings-spec.md`, "CI".
- `.claude/commands/release.md`, "Preconditions".

## Approach

Update the existing `package-oxml-layout` workflow inventory to the complete 24
TTFs and six legal files already declared by the build specification. Keep the
inventory explicit so an unexpected package asset remains a reviewed change.

Raise only the authenticated Pandoc extracted-size ceiling from 128 MiB to 160
MiB. The exact archive sums to 162,406,703 bytes, while 160 MiB is 167,772,160
bytes. Keep the existing 40 MiB download, 256-member, SHA-256, root-layout,
member-type, path, and executable-identity checks unchanged.

Retain the current generic Python exception surface while mapping both
`Error::Mhtml` and `Error::InvalidEmbeddedMutation` explicitly. Add standard
library regression coverage for the workflow inventory and Pandoc bound, and
use the existing Rust mapping regression for the adapter. Do not change a
version, dependency, feature, release allowlist, public method, or package
asset.

## Rejected alternatives

- Derive the expected font inventory from the package output. That would make
  the check tautological and stop detecting unexpected assets.
- Remove the Pandoc size ceiling. Digest authentication does not replace
  bounded extraction.
- Raise the ceiling far above the reviewed payload. A 160 MiB ceiling admits
  the exact archive with less than 4 percent headroom.
- Add Python MHTML or embedded mutation APIs. The binding contract requires the
  generic exception mapping only.
- Treat the failures as transient and rerun GitHub Actions. All three causes
  are deterministic in the tracked source.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression | `test_ci_oxml_layout_package_inventory_matches_bundled_assets` | The workflow lists all 24 TTFs and all six required legal files, and removing any Noto entry fails. |
| regression | `test_pinned_pandoc_installer_accepts_authenticated_archive_with_bounded_headroom` | The ceiling is exactly 160 MiB, exceeds the authenticated 162,406,703-byte payload, and the existing safe extraction checks remain. |
| unit | `import_errors_map_to_the_generic_public_error_class` | MHTML and embedded mutation errors map to the established Python exception class. |
| integration | reconstructed `package-oxml-layout` commands | The exact packaged font and legal inventories compare cleanly and the archive is below 10 MiB. |
| integration | pinned Pandoc installer against the reviewed archive | Download hash, extraction limits, and executable identity pass together. |
| integration | `cargo check -p rdocx-py --all-targets` | The exhaustive Python adapter compiles against the current native error enum. |

The **test gate is regression**. Deleting any Noto inventory entry, lowering
the Pandoc ceiling below the authenticated payload, or omitting either native
error arm makes a named matrix fail. All three reconstructed CI paths pass at
one reviewed SHA.

## HLD impact

- `docs/hld/12-testing-strategy.md`
- `docs/hld/14-development-backlog.md`
- `docs/hld/15-build-and-toolchain.md`

## Risk routing

- **Bundled fonts**. Keep each packaged font family adjacent to its real legal
  files, compare the explicit package inventory, run the verified archive
  build, and enforce the 10 MiB ceiling.
- **WASM or PyO3 bindings**. Compile and test `rdocx-py`, keep workspace
  all-feature tests excluding both Python crates, and run both WASM target
  graphs through the full gate.
- **Release scripting**. Change no version carrier or release allowlist. Run
  `/verify --full`, require a clean sprint review, and complete F-X080 before
  release preparation.

## Hash harness

Expected unchanged. The story changes CI and test infrastructure only. Any
output delta blocks completion.

## Implementation checklist

- [ ] Add failing workflow inventory and authenticated Pandoc-bound regressions.
- [ ] Update the explicit package inventory and Pandoc extraction ceiling.
- [ ] Confirm both current native errors retain the generic Python mapping.
- [ ] Reconstruct all three failed hosted commands locally.
- [ ] Run focused checks, the risk riders, `/microscope`, and `/verify --full`.
- [ ] Update exactly the three HLD impact files and complete the sprint record.

## Open questions

None. The hosted logs, current package contents, authenticated archive size,
and existing binding policy determine the repair without expanding product
scope.
