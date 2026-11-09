# F-X271, all, pass 3

**Reviewer**: Codex agent `/root/fx271_microscope`, independent reviewer of this F-X271 implementation amendment. This agent provided read-only scope advice and did not implement these four changes. Its earlier F-X274 implementation is outside this review. No independence is claimed for reviewing that other feature.

**Reviewed**: Working amendment against HEAD `4f0f836878723041d5d80265d19a436692d7866c`, containing the committed approved plan amendment. Four files,5 additions and5 removals. Binary diff SHA256 `2ffe7b8a486fdd7ea5f18b0fd63b6ccbf4366d52cfd92909639ccea749ad77dd`. All four uncommitted changes inspected, with relevant existing F-X271 retrieval/policy/importer context. F-X272/F-X273/F-X274 product implementation and later numbering-recovery design are excluded.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Scoped amendment review is clean. This is not final integrated acceptance or a new-dependency importer/native parity claim.

## Defects

Zero.

## Smells

Zero. Prior S1 readiness finding remains closed under pass2. The explicit final gate remediation amendment retains final acceptance debt and overrides the original narrower tooling-only scope without claiming that the new dependency was already qualified by old artifacts.

## Nitpicks

Zero.

## Contract and implementation review

Read canonical microscope, current approved F-X271 design including its final gate amendment, prior pass2, relevant HLD15, and the actual four-file diff. F-X271 remains in-progress. The root stopped tracked implementation writes during this pass. This reviewer wrote only this review report and ran no tests, builds, product imports, native applications, helper generation or temporary writes.

- `crates/rdocx/Cargo.toml:37` changes only the exact existing dependency pin from0.13.4 to0.13.5. Default features remain disabled and only html is selected. `Cargo.lock:1371` changes only that package version and checksum. Dependency members, other package versions and graph edges are unchanged. The checksum matches the authenticated official157218byte archive and retained preparation receipt. Its package Rust minimum1.71.1 is below workspace1.93. No new API, module, source file or feature flag is introduced.
- `docs/hld/15-build-and-toolchain.md:1230` changes only the dependency pin. Its adjacent description continues to match manifest feature selection, runtime TABLES/STRIKETHROUGH options and existing facade ownership. No source intent is broadened.
- `scripts/test_sprint_workflow.py:1362` replaces only the stale stable-release macOS URL with the already authenticated archive route. The operative CI URL at `.github/workflows/ci.yml:309` and the existing consumer contract at `scripts/test_sprint_workflow.py:1067` agree. Digest, readonly mounting, exact version, matrix condition, ordering and failure propagation remain checked. No expected result or mutation is deleted.
- Existing Linux `scripts/install_pinned_libreoffice.py:54` retains bounded streamed download and digest validation. Its extraction at81 retains member/expanded-size bounds, expected root, traversal/type refusal and required package inventory. The macOS CI retains335544320byte curl cap, checksum before readonly mounting and exact executable identity. The roadmap expectation and negative controls are unchanged by this amendment. Prior pass1/pass2 archive and safety evidence remains historical and applicable to those unchanged mechanisms, not a new installation claim.

## Advisory and consumer boundaries

The actual old-source `final-supply-chain/stderr.log` reports RUSTSEC-2026-0414 for pulldown-cmark0.13.4. Its advisory describes exponential event growth with ENABLE_WIKILINKS. Actual runtime options at `crates/rdocx/src/html.rs:588` enable only tables and strikethrough. The recorded old public-importer probe returned successfully and is not a semantic importer RED despite its misleading historical filename. No additional vulnerability identity is inferred from this run or from unrelated existing advisory exceptions.

The changed runtime dependency is still significant even without that option: patch-release parser changes can affect existing Markdown parsing. `html.rs:575` projects bounded parser events and source ranges, and `html.rs:656` renders them through the same dependency before native HTML import. `html.rs:891` uses that projection for existing public Markdown import, template/resource and diagnostic association. Python, CLI and WASM inherit the facade, rather than adding an alternate parser or external executable. Old extension bytes cannot prove new-dependency behavior. The amended plan correctly retains fresh affected importer/installed/native, WASM/MSRV, README/archive and hash/pixel qualification in the final union. No advisory allowlist or product option change bypasses that work.

## Actual evidence consumed

- The original final policy run at H a859c19f observed193tests with4 failures,7 error records and2 existing opt-in skips. The four failures genuinely reached the stale baseline URL expectation before its mutation loops. Process/DNS errors remain prerequisite failures and are not silently counted as passed.
- `final-policy-url-repair-green.log` records the same four full named tests passing in0.059seconds: test_debug_optimized_binding_jobs_keep_all_debug_safety_guards, test_python_binding_ci_pins_the_deck_viewer, test_python_pr_job_builds_both_extensions_before_pytest and test_python_pr_job_rejects_failure_swallowing_and_incomplete_cells. The source assertions were inspected, not merely their aggregate count.
- The repaired baseline permits actual mutation loops to execute. At `scripts/test_sprint_workflow.py:1571` these reject disabled debug assertions/overflow checks, changed optimization and a release substitution. At2610 the viewer tests reject missing/wrong/skipped cells, swallowed failure, early success, unpinned digest and commented-out digest check. At2640 the broader contract rejects its existing38 trigger, permission, tool pin, environment, matrix, build, collection, exclusion and failure mutations. The literal replacement does not weaken these discriminators.
- `final-security-patch-deny-green.log` ends with advisories, bans, licenses and sources all ok. Root reports checked child exit0 for the exact0.13.5 candidate. This is the genuine supply-chain GREEN paired with the retained old deny failure, not proof of importer output equivalence.
- `final-security-patch-check.log` records compiling pulldown-cmark0.13.5 and checking rdocx all targets, ending Finished dev profile in1m24s. Root confirms actual child exit0 through its retained child session. Compilation is not functional/native parity.
- The earlier full workspace execution at H was interrupted with SIGINT, child-2, for required remediation. It is not a passing whole suite and cannot be silently carried as one. The earlier local story observations retain only their demonstrated applicability.

## Aspects with no findings

Correctness: zero. The manifest/lock pair agrees, operative and asserted authenticated URLs agree, and original guards remain intact.

Contract: zero. The committed approved amendment authorizes exactly the security patch, lock checksum/version, HLD pin and stale literal, with separately labelled security and test increments still required.

Panics and unsafe input: zero new findings. No authored input loop, unchecked arithmetic, extraction operation or runtime parser option changes here. Actual safety of changed upstream parsing still requires the named final checks.

OOXML: zero. No source XML mechanism, model ownership or serialization code is changed by these four files. No review of unrelated feature mechanisms is claimed.

Tests: zero. Actual stale-assertion RED and four-case GREEN are genuine, existing mutations remain executable, and deny RED/GREEN is correctly distinguished from the passing old public importer probe.

Structure: zero. No new dependency edge, trait, generic, wrapper, file or public operation is introduced.

HLD intent: zero. Listed HLD15 pin matches source selection without making an unsupported fidelity or performance claim.

## Final debt and handback

Fresh affected HTML/Markdown/importer correctness and current installed consumers, full policy with actual process/network prerequisites, full workspace/native suites, hosted Test/MSRV/fidelity/Python/CI jobs, WASM, README and verified archives/size, hash49/pixel7 and all applicable S96 native/font/geometry/risk composition remain pending. Exact prior failures and successful old artifacts remain retained, not relabelled as successor acceptance. Expected existing hashes stay unchanged unless an actual delta is separately explained and reviewed. No baseline update or delivery completion is authorized by this report.

Return control to root for separately labelled dependency-security and stale-literal commits and the pending final union. End this review pass without remediation or a confirmation review.
