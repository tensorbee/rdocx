# F-X271, all aspects, pass 5

**Reviewed**: Frozen working A7 diagnostic increment on `sprint/s96`, under approved plan `d45df33dc838a8abcfce9944967af6e3eb72dc7f8697b514d1f0f9fec4d763cb`. The real-before increment changes2existing files,268additions and10deletions,335unified-diff lines. Prior pass4 retrieval/retry observations remain applicable only to unchanged scope.

**Verdict**: 0 defects,0 smells,0 nitpicks. This is a clean review of bounded diagnostic implementation. It does not fix or close the historical MSRV timeout, A7 or final CI acceptance.

## Defects

0 findings.

## Smells

0 findings.

## Nitpicks

0 findings.

## Scope and authentication

Read canonical workflow/microscope rules, current approved F-X271 A7 contract, progress, prior review and actual Test/MSRV/public integration callers. Inspected the full incremental patch, real before/current source, formatter, public exception path, generated client markers, self-test assertions and HLD15change. F-X271 remains in-progress. No unrelated F275 implementation or current release status is adopted as completed work.

Packet: `/private/tmp/rdocx-s96-run-refresh/fx271-public-diagnostics-implementation-v1`. Independently authenticated all34sealed members, including raw byte counts/hashes. Seal SHA256is `0b31b5e1fe618c17a8b013a86d08cedba1890912f7c23d27f13fc1b185875333`. Handoff SHA256is `c6cdaf9e5455aed81ed8056d72cf8a0d7c6aa89cd3f074ec26c9d3e70590413f`. Patch SHA256is `bbb9e1685f8d0a49e71c7b8ac74714d82116bc268ddce34d262ba4da4c9c45da`.

All3,846after-source entries match actual files before this report. Current script SHA256is `24ce9d9378896becc04badf2a043f038856381d96bcdb202c91a6e0c054b715f`. HLD15SHA256is `7d187ab88703f910b035e6a370badfc962819fbefa9c7623323b660056583796`. The2changed files account for the increment, and3,844protected entries remain exact. Only this review file was written during review. No product script import, test, build, native, network or process-inventory action ran.

## Public failure semantics

`scripts/docx_authoring_conformance.py:603` receives the actual public TimeoutExpired or CompletedProcess outcome. It retains the exception command separately from the actual invocation argv, the actual timeout or nonzero return code, and both captured streams. Type labels plus repr preserve bytes, text, empty values and None distinctly. Non-UTF8 bytes survive as reconstructable escaped bytes rather than replacement-decoded text. Command strings remain data and are never interpreted or executed by the formatter.

The formatter reports exact cwd, the Python host platform and explicitly allowlisted inherited target/toolchain/flag/dev-profile settings. Missing environment entries are shown as unspecified. Unrelated environment values are not dumped. The ordinary command selects dev, while actual compiler version remains explicitly unobserved at this call site. The output does not pretend that a directory or selector string establishes the actual hosted compiler version.

At `scripts/docx_authoring_conformance.py:677`, argv remains cargo run --quiet --offline followed by the original output argument. Cwd, copied environment, existing setdefault target fallback, capture behavior, text mode and timeout300are preserved. The timeout catch at line687converts only that known public child failure into the existing ConformanceError path. Nonzero completion at line691remains a deciding failure with retained diagnostics. The constructed message escapes TemporaryDirectory cleanup and reaches the unchanged main failure handling and existing integration carrier's stderr display.

The three stderr markers occur at actual generated main entry and immediately before each existing deterministic render. The first render still precedes the second marker and second render. Marker presence/absence narrows observed progress without proving compilation, a Cargo lock wait or another wait cause. The formatter explicitly states that no wait cause is inferred. No extra subprocess, retry, target relocation, threshold change, stage skip or concurrency policy was introduced.

## Preserved consumers and privacy

Independent AST comparison finds only run_public and the existing HarnessSelfTests changed among prior top-level functions/classes. Private generators, private execution, argument parsing and main remain unchanged. The public generated source differs only by the three marker lines. Public post-child evidence checks remain intact: required nonempty output files, four package-class profiles, normalized package graph, opaque-part preservation, modeled-state report and repeated deterministic PNG agreement.

The existing integration test still invokes the same public mode and requires success. No Rust integration or CI YAML change is present. The ordinary Test standalone warm-up versus MSRV nested cold-consumer asymmetry remains an observation, not a diagnosed cause or an authorized resource adjustment. The public route remains distinct from F275's SemVer clients.

Private TimeoutExpired handling retains its existing generic non-disclosing message. The added privacy control supplies a secret in private command/output/stderr and requires that exact generic main result. Public diagnostics do not install a global timeout formatter over private operations. HLD15adds only the public diagnostic/marker contract near line91. Its existing F275SemVer intent and all unrelated sections remain unchanged.

## Actual tests and limitations

Authenticated each of the six command receipts and its raw stdout/stderr. The prior-runtime/test-only execution genuinely reports4tests and8assertion failures, covering six typed timeout subcases, nonzero diagnostics and the missing marker. The incomplete-output refusal already passes there. Those are synthetic diagnostic REDs on the actual old runtime, not actual MSRV/Cargo failures.

The fixed source genuinely reports12named harness self-tests PASS with no ignored tests. Inspected assertions distinguish list/tuple commands, fractional/whole timeout values, text/bytes/None, non-UTF8/empty bytes, fallback versus inherited target settings, exact unchanged invocation options, one child call and cleanup. Nonzero completion retains its actual return code and remains failure. Exit0without required public evidence remains failure. Markers are unique and ordered, and both deterministic render calls remain required. Private mode and non-disclosure controls stay active. The optional-private diagnostic line is not a unittest skip.

The existing CI/privacy wiring test separately reports1PASS. Recorded prose, adapter drift and diff checks pass. No Cargo or actual generated public client ran. Synthetic exception/output observations prove diagnostic retention and refusal only. No native compiler version, target-lock contention, render stall, successful source-qualified public conformance or clean final CI is inferred. The denied read-only process inventory remains recorded as unavailable, not a system-wide no-process assertion.

## Not found and remaining acceptance

Correctness, contract, panics, OOXML, tests and structure:0additional findings in this scope. No public product API, command mode, receipt schema, process manager, generic logging framework, module, file, crate or dependency was introduced. No parser/serializer, child ordering, source XML or rendering policy changed.

Actual final-source MSRV public conformance must either pass all original stages with its real source/toolchain/profile evidence or produce diagnostics that support the separately bounded cause repair. The old timeout remains failed. A7requires that result and the clean final CI union. Diagnostic GREEN is not timeout closure. Earlier installer, corpus transport, advisory and all original F271hosted/native/fidelity/full obligations remain due, together with applicable F272/F273/F275and sprint native/hash/raster/installed/WASM/risk gates.

No source-availability, feature-completion, baseline, push, dispatch, comment, tag or publication status is granted. End this pass and return the sole canonical writer lease to root. No unchanged-scope confirmation pass is requested.
