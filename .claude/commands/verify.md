---
description: The gate. Runs formatting, lints, tests, the hash harness and the prose rules.
---

# /verify [--fast | --scoped F-XXX | --full]

There are three verification scopes:

- `--fast` runs formatting, changed-crate linting and the unique affected local
  regression tests. It does not run whole-crate or complete installed suites.
- `--scoped F-XXX` is the feature-completion gate for local evidence. It binds
  the fast checks and the approved plan's named local RED/GREEN regressions to
  the claimed Base and actual source. Independent `/microscope` remains required.
- `--full` is the final integrated sprint gate. It runs the workspace suites
  and the union of every approved plan's final oracle and risk checks once.

Whole-crate suites, ordinary complete installed suites, external oracle
comparisons, native Office controls, hash/pixel corpora, whole-package typing
and stub checks, WASM, documentation, publication dry runs and supply-chain
checks belong to the final union. A small affected typing or compile check can
be a local regression when the change needs it. Do not promote it to another
whole-suite run. Keep distinct configurations and artifact identities separate.

The caller consumes an applicable successful observation instead of running
it again. Record exact source, command/configuration, tools, inputs and results.
A later failed attempt invalidates an older green result for the same check.
After a source fix rerun affected checks, preserving earlier failures. Never
call an unperformed final oracle a passed local gate.
Record in-flight evidence and pending checks in the feature's existing
`.claude/scratch/<F-ID>-progress.md` before handing over.

## Acceptance applicability

### Build, artifact and semantic outcome

Dependent batches stop after a failed prerequisite. Use `set -e` in a shell
build/test batch or check the actual child status before the next stage. A
failed extension build must not fall through to tests or timings on an older
installed module. Preserve the failed attempt and authenticate the replacement
artifact after a successful build.

Keep the wrapper/transport exit, checked child exit and semantic acceptance
outcome separate. A wrapper that successfully records Cargo exit 101 has
recorded a failed build. Exit zero with a reported target of false is successful
measurement execution, not a satisfied target. Classify an approved trend as a
trend and a required threshold miss as a failed criterion.

Do not measure or test a leftover binary after a failed rebuild. Record the
successful build command, actual child exit, source identity, profile/features,
artifact path and content hash. Link the observation to that artifact and
source, including the actual imported extension path for Python. A diagnostic
such as enum size or body_content_bytes alone does not identify a binary.
Retain failed attempts separately. A later successful build and run can
supersede their applicable result, but must not relabel them as successful or
claim to qualify newer source without an applicability check.

For benchmarks, retain every required operation, small/large workload sizes,
raw sample values and units, aggregation, ratios, threshold and pass/fail.
Preserve relevant runtime/load observations. Near-threshold success remains
success under the approved protocol. Do not retry to select a favorable result,
change the denominator or relax the threshold after seeing it. Repeat only for
an explicit invalidation or a predeclared sampling rule. Keep structural work
bounds, timing, functional correctness and final native evidence distinct.
Measure the acceptance consumer itself. Native timing does not qualify a
Python API workload: retain the imported release extension path and hash, and
include binding lookup, validation and publication in the timed call. Trace
work counters through the actual loops they claim to bound. A constant marker
beside an uncounted inventory scan is not a structural work-bound control.

### Selecting applicable checks

For scoped Python checks, select the A-ID tests and their affected controls
before launching a broad file. Preflight only their required dependencies,
oracle pins, installed extension identity and environment configuration. Avoid
running a large known-unready suite to rediscover missing imports or settings.
Final installed-product and oracle suites remain mandatory at their scheduled
boundary. If a selected run fails for an environment reason, retain its failure
and repair the environment or leave that check pending. Do not turn a required
case into a skip. A broader failed command is still failed even when selected
cases passed inside it. Cite those cases individually, and do not infer
correctness from the absence of AssertionError or a particular error class.

Before choosing commands, map each A-ID to the observation that proves it.
Check selected test membership and discriminators, not just exit status. Cases
that promise consumer agreement compare the same input across those paths.
Keep internal agreement, round-trip correctness and native/oracle parity as
separate claims. Preserve expected output, actual output and the failing
assertion for semantic RED. A compile failure or self-catching reporter exit
cannot replace that observation.

After completion/status/capability edits, run only the metadata-dependent checks
whose inputs changed, even when the earlier full product gate remains valid.
At minimum run the existing live-capability-owner policy after status changes,
plus classification checks when capability rows change. The metadata-adoption
helper enforces the owner invariant but does not cover every possible delivery
policy. Record exact post-edit applicability. Do not repeat unrelated product
suites, native UI or corpora for metadata alone.

## Steps

1. **Format.** `cargo fmt --all --check`.

2. **Lint.** For `--fast` and `--scoped`, run `cargo clippy -p <crate>
   --all-targets --all-features -- -D warnings` for each changed crate. For
   `--full`, run `cargo clippy --workspace --all-targets --all-features
   -- -D warnings`. Include both PyO3 binding crates in this single workspace
   lint. Only cargo test excludes their un-linkable extension test binaries.

3. **Run the affected local tests.** Under `--fast` and `--scoped`, use the
   approved plan's unique regression cases against the F-ID Base, not the sprint
   base. Retain actual old-source RED and fixed-source GREEN evidence. Reuse
   applicable observations already produced during implementation. Bindings
   use focused fresh installed tests when needed, never a PyO3 Rust test binary.
   Under `--full`, skip this category when the workspace or complete installed
   suite already covers the same tests and configuration. Retain a distinct
   focused configuration only when the broader suite does not cover it.

   Batch related local cases sharing a test binary when practical, using exact
   libtest filters or a bounded named group. Retain per-case outcomes and check
   the selected membership. A successful batch does not cover filtered-out
   cases. Reuse unaffected observations without another binary invocation.

4. **Test the workspace.**
   `RUST_MIN_STACK=8388608 cargo test --workspace --all-features --exclude rdocx-py --exclude rpptx-py`.
   Use the native test-thread budget documented in HLD15 and CI. It does not
   change product runtime or input-depth limits. `--full` only.

5. **The hash harness.** `python3 scripts/hash_harness.py --check`. Run once under
   `--full`, not under `--fast` or `--scoped`.

   **Mandatory for every story in M1 through M6.** An unexplained delta fails
   the gate. An expected delta must be declared in the design plan's
   `## Hash harness` section, and the reported delta must match what was
   declared. A delta that is real but undeclared is a failure, not a prompt to
   update the baseline.

6. **The repository-policy checks.** Run once under `--full`. A workflow-only
   change runs its unique affected policy tests locally, without product gates.
   Newly written completion/review Markdown and commit messages receive their
   cheap prose check after the full gate. Also rerun affected completion-policy
   assertions against final metadata, including live capability ownership.
   Metadata adoption enforces prose and owner consistency without repeating
   product verification.
   Do not run the entire policy suite after each product story.

   Run `python3 scripts/prose_check.py` over tracked Markdown and the commit
   message. No em-dash, no en-dash, no prose semicolon.

   Then `python3 scripts/sync_agent_skills.py --check`. The `.agents/skills/`
   adapters are generated from `.claude/commands/` and `.claude/skills/`, and
   drift means Codex and Claude are following different rules while believing
   they agree. Regenerate with the same script and commit the result.

   Then `python3 -m unittest scripts.test_sprint_workflow`. This holds the
   release family preflights that `.github/workflows/wheels.yml` invokes by
   name as the publication gate, and the assertions over the pinned CI
   toolchains. A failure means a version carrier moved without its assertion
   moving with it. The fix is the carrier or the assertion, never deleting the
   test.

   Without this step the preflights run for the first time on a tag, after the
   sprint is closed. S42 is the demonstration: F-X022 moved every version
   carrier under `crates/`, passed the entire local gate, and left the
   incubating preflight and the `ci.yml` WASM literal asserting the old
   version. The standalone README gate in step9 owns actual README compilation
   and archive inventory checks. Policy tests retain their unique mutation and
   contract assertions without replaying that full integration gate.

7. **The no-default-features path.**
   `cargo test -p oxml-layout --no-default-features`. This is the only thing
   that exercises system font discovery being off while bundled fonts remain
   available. Run once under `--full`. The plan records its
   applicability there. Skip it under `--fast` and `--scoped`.

8. **The wasm targets.**
   `cargo check --target wasm32-unknown-unknown -p rdocx-wasm -p rpptx-wasm`.
   Run once under `--full`, not under `--fast` or `--scoped`.

9. **Docs.** `cargo doc --workspace --no-deps` with `RUSTDOCFLAGS=-D warnings`,
   then `python3 scripts/readme_doctests.py`. Run once under `--full`, not under `--fast` or `--scoped`.

10. **Packaging.** Run the workspace dry run with every publishable internal
    crate patched to its reviewed local source:

    ```bash
    cargo publish --workspace --dry-run \
      --config 'patch.crates-io.oxml-core.path="crates/oxml-core"' \
      --config 'patch.crates-io.oxml-drawing.path="crates/oxml-drawing"' \
      --config 'patch.crates-io.oxml-layout.path="crates/oxml-layout"' \
      --config 'patch.crates-io.oxml-media.path="crates/oxml-media"' \
      --config 'patch.crates-io.oxml-opc.path="crates/oxml-opc"' \
      --config 'patch.crates-io.oxml-pdf.path="crates/oxml-pdf"' \
      --config 'patch.crates-io.oxml-sml.path="crates/oxml-sml"' \
      --config 'patch.crates-io.oxml-cli-support.path="crates/oxml-cli-support"' \
      --config 'patch.crates-io.oxml-chart.path="crates/oxml-chart"' \
      --config 'patch.crates-io.rdocx.path="crates/rdocx"' \
      --config 'patch.crates-io.rdocx-cli.path="crates/rdocx-cli"' \
      --config 'patch.crates-io.rdocx-html.path="crates/rdocx-html"' \
      --config 'patch.crates-io.rdocx-layout.path="crates/rdocx-layout"' \
      --config 'patch.crates-io.rdocx-opc.path="crates/rdocx-opc"' \
      --config 'patch.crates-io.rdocx-oxml.path="crates/rdocx-oxml"' \
      --config 'patch.crates-io.rdocx-pdf.path="crates/rdocx-pdf"' \
      --config 'patch.crates-io.rpptx.path="crates/rpptx"' \
      --config 'patch.crates-io.rpptx-cli.path="crates/rpptx-cli"' \
      --config 'patch.crates-io.rpptx-chart.path="crates/rpptx-chart"' \
      --config 'patch.crates-io.rpptx-layout.path="crates/rpptx-layout"' \
      --config 'patch.crates-io.rpptx-oxml.path="crates/rpptx-oxml"' \
      --config 'patch.crates-io.rpptx-render.path="crates/rpptx-render"'
    ```

    Cargo rewrites packaged path dependencies to the registry. The local
    patches keep archive verification on this exact reviewed source graph,
    including versions that have not yet been published. They do not enter
    any generated archive. Then assert every generated archive is below the
    crates.io 10 MiB limit:

    ```bash
    oversized=$(find target/package -name '*.crate' -size +10485760c -print)
    if [ -n "$oversized" ]; then
      echo "package exceeds 10 MiB: $oversized" >&2
      exit 1
    fi
    ```

    Run once under `--full` on the clean integrated tree without
    `--allow-dirty`. Do not run publication dry runs under `--fast` or `--scoped`.
    Its retained verified archives also supply the release preflight packaging
    evidence at the same source. No second publication dry run in that preflight.

11. **Supply chain.** `cargo deny check`. Run once under `--full`, not under
    `--fast` or `--scoped`.

## Reporting

Report each step as pass or fail with its command. On failure, show the actual
output rather than summarising it, and stop at the first failing step unless
the failures are independent.

**Never report a pass you did not observe.** If a step was skipped, say it was
skipped and why.

## Refused situations

- **`--fast` as the gate for `/complete-feature`.** Run `/verify --scoped
  F-XXX` instead.
- **Dropping a declared final rider or oracle.** Local scoped readiness does
  not waive final acceptance. Preserve named pending checks in the handoff.
- **Updating the hash baseline to make step 5 pass.** The baseline changes only
  through a declared, reviewed delta.

## Strict native compatibility evidence

`--full` consumes the three-host strict SemVer union once at the integrated
source. First use the reviewed private `_prepare_release_semver_evidence('all',
owned_directory)` controller in `scripts/sprint_workflow.py`. It collects the
complete current public CI run, attempt, job and artifact catalog into the
canonical current sprint state, outside every supplied evidence tree, and
returns fixed proof directories. Then run:

```bash
python3 scripts/sprint_workflow.py release-semver all --check-evidence DIR [DIR ...] --refresh-registry
```

Each deciding caller acquires its own authority context. A saved catalog, packet
label, caller-supplied digest or an absent context cannot substitute. Frozen
`--check-evidence` without refresh uses only that independently acquired cutoff,
local pinned gh, unchanged bundles and the embedded approved root. It performs
no API, TUF or Cargo acquisition. Newer relevant failed or incomplete attempts
block older success. The running CI gate consumes completed native and signing
jobs, while external full verification also requires completed successful CI.
Owner-only dispatch, forks without signing permission and scheduled jobs do not
qualify this union. Missing real hosted proof remains pending.

Execution is the separate native acquisition interface:

```bash
python3 scripts/sprint_workflow.py release-semver all --target HOST_TRIPLE --tool VERIFIED_TOOL_PATH --output NEW_DIR [--reuse-evidence DIR ...]
```

Use the tracked official archive and executable pins for the genuine native
host. Repeated reuse options each take one directory. Only authenticated rows
with unchanged applicable source, baseline, features, compiler, tool and Cargo
configuration are reused. Missing original provenance or invalidated rows are
acquired again. Each baseline still undergoes fresh registry applicability and
checksum authentication. Default and all-feature library views, binary-only
applicability and exact compiler witnesses remain distinct. First unpublished
minor lines have explicit empty-line applicability, which is not old-line API
compatibility. Existing behavioral, CLI, Python, WASM, hash and native gates
remain required.
