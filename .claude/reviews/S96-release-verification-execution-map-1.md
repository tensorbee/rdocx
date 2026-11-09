# S96 release verification execution map 1

2026-11-08. Bounded read-only execution preparation at the adopted Q-S96-16
scope. No tests, Cargo, runtime, network or source changes ran in this pass.
Numerical performance and unfinished RSS/history/calibration are carried to
S96.1. This map adds no gates and does not qualify an artifact.

## Launch order and scope

1. Consume the frozen bounded namespace observations and independent review.
   The actual packet is
   `.claude/scratch/S96-final-evidence/s96-release-freeze-20261108-1/`:
   `FREEZE.json`, `SOURCE-inputs.json`, `SCOPED.json`, `SELECTED_TESTS.json`,
   `TEST_BINARIES.json` and individual logs. `SCOPED.json` records the actual
   affected check and test-build exit zero. This is focused evidence, not a full
   workspace or complete installed-suite receipt. Root's active batch is not
   duplicated here.
2. Run the missing ordinary local union once at the frozen source. Independent
   Python artifact preparation can proceed alongside workspace work, subject to
   Root's existing Cargo lease. Authenticate each installed payload before tests.
   Stop a dependent chain after a failed build or install.
3. Review/fix actual failures and rerun affected categories only. Preserve the
   earlier failures. Commit/push/normal CI and the subsequent rehearsal remain
   controlled by the existing workflow, not authorized by this map.
4. At the reviewed source consume actual complete hosted/native authority, then
   build-only rehearsal and artifact qualification. Final clean-tree verified
   packaging, final sprint acceptance/close and separate reviewed main-SHA
   publication approval remain their existing boundaries.

The consumption rule is `.claude/commands/verify.md:24`: retain exact source,
command/configuration, tools, input and result, and do not reuse an older green
over a relevant later failed attempt. No currently absent receipt is silently
reconstructed from a progress summary.

## Existing ordinary commands

Commands below are an execution map, not a claim that they ran in this pass.

| Category | Existing command or controller | Prior evidence and current disposition |
|---|---|---|
| Complete Rust workspace | `RUST_MIN_STACK=8388608 cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py` | Canonical `.claude/commands/verify.md:124`. The 8MiB value is the native thread-stack budget in HLD15:1194, not a runtime or input-depth change. S96-run's four full records all remain false. Current complete union is missing, rather than inferred from scoped tests. |
| Lint/format | `cargo clippy --workspace --all-targets --all-features -- -D warnings` and `cargo fmt --all --check` | Canonical local verify command. H6 full lint success is recorded in F275 progress:1227, but its raw temporary packet is absent. Current freeze has scoped clippy/fmt logs, which do not become complete workspace lint proof. |
| Hash 49 | `python3 scripts/hash_harness.py --check` | Canonical verify:128, CI:540. H6 progress:1227 records unchanged49matches at its source. The later namespace allocation source requires actual current command or explicit unchanged-input applicability, not a new baseline. |
| Deterministic pixels 7 | `python3 scripts/golden_png_harness.py --check` | CI:279. `scripts/golden_png_harness.py:23` names all seven sample PDFs and `:399` defines check mode. Prior full record says fresh checks were pending. Never use update mode to erase an unexplained delta. |
| No default features | `cargo test -p oxml-layout --no-default-features` | Canonical verify:168. Consume only an actual unchanged applicable receipt or run once in the missing final union. |
| WASM | `cargo check --target wasm32-unknown-unknown -p rdocx-wasm -p rpptx-wasm` | Canonical verify:174. No present final-source success receipt was established by this audit. |
| Rust docs | `RUSTDOCFLAGS='-D warnings' cargo doc --workspace --no-deps` | Canonical verify:177. H6 progress:1227 records historical Rustdoc zero, not current final-source runtime proof. Hosted CI additionally uses its existing all-feature/exclusion command at CI:668. Keep those configurations distinct. |
| README27 and inventory22 | `python3 scripts/readme_doctests.py` | Canonical verify:178. Actual main dispatch at script:1686 first calls `validate_inventory`, then compiles README examples. The inventory packages with `--allow-dirty --no-verify` at :1262. It is not a substitute for verified publication dry run. Root's new source invalidates affected package identities. |
| Verified archives22 | Exact existing 22-local-patch `cargo publish --workspace --dry-run` block in `.claude/commands/verify.md:184`, with the size assertion at :214 | Run on the clean integrated source without `--allow-dirty`. Keep all22 archives below10MiB. H6 actual zero/246.80s is recorded at F275:1267, referencing the now absent `/private/tmp/rdocx-s96-h6-verified-publication-dryrun-1/RESULT.json`. Historical summary alone does not supply reusable archive bytes or current source identity. Same-source verified archives can later be consumed by release without a second dry run. |
| Supply chain | `cargo deny check` | Canonical verify:230. Current full success not established here. |
| Final repository policy | `python3 -m unittest scripts.test_sprint_workflow`, `python3 scripts/prose_check.py`, `python3 scripts/sync_agent_skills.py --check` | Canonical verify:137. Root runs the already identified affected metadata checks after Q16 edits, then consumes/runs the unique final policy union without replaying whole product verification for metadata. |

## Installed Python, typing and external object-model oracle

The ordinary complete suite commands are the authenticated installed Python3.9
and Python3.12 interpreters running:

```text
<installed-word-python> -m pytest crates/rdocx-py/tests
<installed-powerpoint-python> -m pytest crates/rpptx-py/tests
```

Use the actual approved environment/dependency pins and authenticate the imported
module against the wheel payload first. Bind the wheel, extension, test source,
interpreter and configuration in each receipt. Pinned object-model dependencies
remain `python-docx==1.2.0` and `python-pptx==1.0.2`, as in CI:323 and wheels:171.
The Word suite includes the existing `test_python_docx_parity.py` oracle carrier.
The existing native font-tool preparation remains the separate CLI in
CI:220, rather than a presumed effect of collecting its Python file with pytest.

The H6 complete suites are recorded at `.claude/scratch/F-X275-progress.md:1306`:
1415pass/25debug-onlyskip/3quietdeselected on each interpreter, with both mandatory
F286 controls unskipped. Their referenced `word-h6-release-successor-1` runner
and `COMPLETION_RECEIPT.json` are not present in the inspected durable S96 tree
or referenced temporary root. The exact original three quiet selectors were not
recovered here. Do not invent those selectors, reproduce the historical counts
as current, or add a numerical gate. Root must retain the actual current suite
membership and explicit non-deciding disposition of carried timing tests.

Existing durable S96 `qualify_release.py` and `authenticate_and_check.py` in
`f268-guard-iterator-release-qualification-20261108-1/` run focused ordinary
functional controls plus eight exact original-state observations. They are useful
prior inputs, but their explicit node list is not a complete installed runner.
Do not execute their hard-coded old artifact configuration as final acceptance.

The existing whole-package typing commands, copied from wheels:198 to205, are:

```text
<installed-word-cp312> -m mypy --strict crates/rdocx-py/tests/typing_smoke.py crates/rdocx-py/python/rdocx
<installed-word-cp312> -m mypy.stubtest rdocx
<installed-powerpoint-cp312> -m mypy --strict crates/rpptx-py/tests/typing_smoke.py crates/rpptx-py/python/rpptx
<installed-powerpoint-cp312> -m mypy.stubtest rpptx
```

The workflow pins mypy2.3.0. H6 typing applicability recorded at F275:1308
transferred original H4 observations under unchanged relevant source and same
interpreter with new payload authentication. That is an applicability claim,
not a new mypy execution. Absent raw H4 custody cannot be replaced with a label.
Actual unchanged binding/stub/smoke/interpreter inputs can support reuse only
with their concrete authentic receipts.

The wheels workflow's `WHEEL_TEST_NODES` at `.github/workflows/wheels.yml:29` is
the existing downloaded-wheel priority subset, invoked at :179. It is not the
complete suite above. F275 plan:264 requires42Word plus44PowerPoint calls for
each of3.9/3.12,172total, and3.12typing/stubs for the hosted26payload rehearsal.

## Rendering oracles, native Office and CLI

Existing generic executable entrypoints are:

```text
python3 scripts/docx_authoring_conformance.py --public
python3 scripts/docx_ssim_harness.py --check --corpus-dir <actual-docx-corpus> --output-dir <owned-output>
python3 scripts/pptx_ssim_harness.py --check --corpus-dir <actual-pptx-corpus> --output-dir <owned-output>
```

These are the existing CI controls at :263, :620 and :579. Use deterministic
fonts and the pinned tools already specified there. These generic corpora do not
replace the exact approved Issue423 to445 composition/native/SDK rows. The
workspace suite covers its ordinary Rust/CLI binaries, but neither generic
corpora nor workspace success proves a missing native Office action.

Historical current-source applicability for native inputs is explicitly recorded
at F275 progress:1300: H6 issue430 generator emitted50byte-identical H4 SDK
inputs, and B6 retained six native/SDK input bytes. Those are historical input
equivalences, not new UI operations. The approved Q-A430-28 one-release exception
remains bounded at CURRENT_SPRINT:76. Do not reopen it or call it passed.

The referenced `/private/tmp/rdocx-s96-run-refresh`,
`/private/tmp/rdocx-s96-prepared-integration-readiness`,
`/private/tmp/rdocx-s96-H11-checkpoint` and H6 verified-archive packet are absent
locally in this pass. Thus the exact old native/SDK/CLI helper commands and raw
Office output custody cannot honestly be recovered from those paths. Recover
actual retained packets from their authentic source or run only genuinely
missing/invalidated approved rows. Do not invent a replacement native script or
claim that recorded historical success is present raw evidence.

## Normal CI, native authority and rehearsal controllers

The durable implementation is `scripts/sprint_workflow.py`, not a missing
scratch helper. Canonical verify:254 directs the existing private
`_prepare_release_semver_evidence('all', owned_directory)` controller. Its live
collector is `_collect_release_semver_authority` at script:2909 and retains a
current complete catalog under `.claude/scratch/semver-authority/`. It mutates
authority state and makes actual API requests, so it was not invoked by this
read-only pass. The subsequent existing command is:

```text
python3 scripts/sprint_workflow.py release-semver all --check-evidence DIR [DIR ...] --refresh-registry
```

Native execution, where not actually reusable, is the existing separate native
producer command at verify:277, with its exact supported host/tool pins:

```text
python3 scripts/sprint_workflow.py release-semver all --target HOST_TRIPLE --tool VERIFIED_TOOL_PATH --output NEW_DIR [--reuse-evidence DIR ...]
```

H11 run42368584662 genuinely succeeded with24jobs and two expected optional
skips for exact cec87f95, recorded at S96-run:4048 to4074. It excludes the later
dirty source. Its referenced signed handoff is absent locally. Current complete
source proof and custody must come from the actual collector, not arbitrary old
green selection. Existing signed native packets may be reused only under the
controller's source/dependency/features/baseline/tool/attempt applicability.

F275 plan:595 requires successful exact-source normal CI before the build-only
wheels rehearsal. Normal dispatch retains `owner_compatibility=false` and
`f283_benchmark=false`. F275 plan:264 defines the unchanged12CLI/12wheel/2sdist
inventory. The durable routing and install/typing commands remain in
`.github/workflows/wheels.yml`. Four publication jobs stay skipped. The absent
old scratch dispatch/download helpers are not runnable assets discovered here.

The later selected-family publication consumer is the existing release:161
`release-semver TAG --check-evidence DIR [DIR ...] --release-candidate`, after
its own live authority collection and fresh mutable registry checks. It does not
authorize tags or publication and cannot replace separate exact-main-SHA approval.

## Minimal remaining evidence disposition

Current focused namespace controls are present. Historical full union entries
remain false. Prior local full, complete installed/typing, native/SDK and verified
archive successes are recorded, but their named temporary raw packets are absent
here. Root should consume any authentic independently retained unchanged-input
receipts it actually holds, launch the missing ordinary final union once, and
retain actual affected reruns. No performance pause, 79/79 retry, RSS overlay or
new threshold is needed for this execution preparation or Q16 release closure.
