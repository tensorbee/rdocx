# F-X275, all aspects, pass 6

**Reviewed**: Current-source archive measurement refresh only. Four existing files, 6 additions and 6 deletions, 51 incremental patch lines. Approved plan remains SHA256 `bb9ee3466e77c10ba1e28d7a0ff84630f29e5219604a6b8e9573a178a2761110`.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this bounded increment. Measurements match the genuine retained archives. Compilation-verified packaging, README examples and final hosted/native/full acceptance remain pending.

## Defects

0 findings.

## Smells

0 findings.

## Nitpicks

0 findings.

## Scope and authentication

Read canonical workflow/microscope requirements, current F275 plan/progress, prior all-pass5 and the applicable HLD10/14/15 release/packaging contract. F275 remains in-progress in CURRENT_SPRINT and BACKLOG. Prior runtime and actual-CI review scope is unchanged and is not re-reviewed here without a new source reason.

Packet: `/private/tmp/rdocx-s96-run-refresh/fx275-current-archive-refresh-v1`. Independently authenticated all 337 sealed members. Seal SHA256 is `e90ce16875fea703dc3bc6c0bac5a493f3756b13e7913799c445c97ada53d2e6`. Handoff SHA256 is `520844893f6384d43728929ecb512ba15e2cf35ded48d038e3866c52c0700634`. Incremental patch SHA256 is `23895637d8da85fab60d4e93e44a95b10ff2d0dd4cbf3598912e7650137b1d67`.

Independently checked all 3,855 current source/configuration inventory entries. Exactly README.md, crates/rdocx-layout/README.md, crates/rdocx-oxml/README.md and scripts/readme_doctests.py differ from the actual before map. The other 3,851 inputs remain exact, including production repairs, baselines, versions, plans, prior reviews and unrelated work.

Current root README SHA256 is `185987efa328f6d4cb0a102cea06ac61e048f5e20df1f0212772f5ca6d6e6b21`. Layout README is `bbe9446eb6b13eab7634801b6a64fd434a81f3566dbe11111c280ce8eea0d411`. OXML README is `65d27835ea1485c9de1653a9f34e9c0d9b9e5e8fa029aa9d77b2891145ee6f04`. Measurement script is `481b58b61f514342a8e61d3fdb6fa9dfd5b9975e787fc69acea867a413f81304`.

Only this review file was written. No TEMP writes, product imports/executions, tests, Cargo, builds, network, native application or fixes ran. Read-only hashing, AST/JSON inspection and tar member reading did not execute product code or extract files to disk.

## Metadata and unchanged policy

The full patch changes only three archive tuples at `scripts/readme_doctests.py:405` and their corresponding README rows at `README.md:44`, `crates/rdocx-layout/README.md:23` and `crates/rdocx-oxml/README.md:22`.

| Package | Recorded compressed bytes | Normalized member bytes | Members |
|---|---:|---:|---:|
| rdocx | 4,270,946 | 55,071,275 | 37 |
| rdocx-layout | 448,683 | 2,463,941 | 15 |
| rdocx-oxml | 618,434 | 3,777,791 | 32 |

The root README is the facade package's actual packaged README. Its row is not incorrectly duplicated into a separate facade path. Versions, measured date, platform labels, build recipe and provenance columns remain unchanged. Earlier performance observations are not relabelled as current measurements.

Independent AST comparison after removing ARCHIVE_MEASUREMENTS finds the entire before/after script exact. Within that dictionary, only the three intended tuples change. Other19tuples, SPEED rows/constants, helpers, normalization, inventory, compression tolerance and package-size policy remain unchanged. No new HLD behavior is introduced by this metadata refresh.

## Genuine archive evidence

Inspected the acquisition wrapper and original helper. The wrapper delegates to real subprocess.run, retains raw outputs, copies an archive only after actual cargo package exit0, and invokes the unchanged selected helper or validate_inventory. It does not manufacture a passing measurement or substitute a preexisting target archive. Preexisting files are explicitly retained as unqualified.

The first acquisition fails public-index DNS and produces zero archives. That failure remains preserved. The authorized public-read retry obtains three archives using the original locked/allow-dirty/no-verify recipe with all22local patches. Three subsequent self-reference acquisitions supply the recorded rows. The unchanged full inventory then rebuilds22archives and validates27README sources. This is measurement acquisition, not compilation-verified publication packaging.

Independently opened every final tar archive and checked its file SHA, every regular member's raw size/SHA, unique member names, exact package/version root, normalized Cargo name/version and actual README bytes. The original VCS normalization removes dirty and zeros only the Git SHA before counting member bytes. No other member normalization was introduced. All22normalized totals and counts exactly match the recorded rows. Actual compressed deviations are at most2bytes within the unchanged64byte tolerance. The largest archive is oxml-layout at4,916,050bytes, strictly below10MiB.

Comparing actual prior and final member inventories confirms the other19archives' members/hashes remain exact. The three changed archives carry precisely their changed README and actual current source payloads: comparison/document for rdocx, engine for layout and text for OXML. Script baseline manifests are not packaged. An older archive's original producer is not silently transferred to these changed packages.

## Checks and limits

The actual six-test metadata/version invocation exits0 and raw stderr reports6PASS with no failures. Inspected the existing assertions: exact normalized footprint and clean/dirty VCS equivalence, enforced performance floors, dated provenance refusals, exact approved row membership, deferred-measurement tracking and current metadata-derived root requirements. The expected README error diagnostics are negative controls, not failed deciding tests. Actual prose, adapter drift and diff checks pass at stable source.

Post-acquisition Cargo/Rust1.97.1 and platform observations remain retained. The optional CPU-brand read was refused and is not presented as successful verification. No unchanged platform label or performance claim was revised to infer evidence from that refusal.

Correctness, contract, panics, OOXML, tests and structure: 0 additional findings in this metadata increment. No API, runtime policy, module, file, trait, generic, dependency, test binary or unsupported package inventory expansion was introduced.

Actual compilation-verified packaging, README example/rlib compilation, successful reviewed-source hosted CI/wheels and all final native/full/risk/installed/hash/raster/release gates remain due. This review does not complete F275, any product story, Issue430, the sprint or a release. No baseline change, tracker completion, commit, push, dispatch, public comment, yank, tag or publication is granted.

End this pass and return the canonical writer lease to root. No unchanged-scope confirmation pass is requested.
