# F275 advisory: retired macOS ARM64 rehearsal runner repair

Zero required changes were found in the bounded runner-selector repair. It replaces the two retired macOS ARM64 runner values, updates the existing CLI platform expectation and adds a meaningful retired-selector negative. Targets, artifact formats, wheel deployment-tag requirements and publication guards remain unchanged. This advisory is separate from the completed F287 metadata advisory and does not reset any formal review lineage or authorize a release.

## Retirement cause authenticated

All four retained GitHub annotation records for run42524582945 at bbdd4d651588099e27cea5266c274422a3de781c state that the macos-14-arm64-us-default configuration is unsupported. The four job IDs are128865250617,128865250679,128865250728 and128865250752. READ_ONLY_FACTS.json authenticates the complete retained annotation bytes and source associations. These are failed hosted qualification, not successful artifact builds.

Both official sources were independently read. GitHub's [October1 retirement notice](https://github.blog/changelog/2026-10-01-github-actions-macos-14-runner-image-retirement/) gives November2,2026 as the macos-14 retirement date and lists macos-26 among supported ARM replacement labels. The official [runner-images retirement issue13518](https://github.com/actions/runner-images/issues/13518) confirms that date for GitHub Actions Intel and ARM, while distinguishing the later Azure DevOps Intel extension. The extension does not make macos-14 available to this GitHub Actions workflow. Reusing a supported explicit macos-26 route is consistent with those primary sources.

## Actual reviewed scope

The complete working wheels.yml equals its HEAD preimage after exactly two replacements of runner macos-14 with macos-26 beside target aarch64-apple-darwin. There are no other workflow differences. The changed wheel cell is at .github/workflows/wheels.yml:25 and the CLI cell at line336. Matrix labels, package/family selection, ordering, targets, installation roles, archive types, build/profile flags, SDK/action pins and artifact identities are byte-preserved. Current ordinary Python and native SemVer CI already use macos-26 at ci.yml:314 and958, respectively. That prior route does not substitute for a new release artifact attempt.

The selected wheel product still has two packages across six platforms, plus two source distributions. Manual CLI rehearsal still has both families across six platforms. The expected26 unique payload count is unchanged. The aarch64 Apple Rust target is unchanged, and no product crate or Cargo file is modified.

The macOS ARM wheel platform remains `macosx_11_0_arm64` in scripts/sprint_workflow.py:107. validate_python_release_artifacts constructs the exact expected filenames from that unchanged list at lines3902 to3909, rejects missing/unexpected names at lines3913 to3921 and enforces the exact cp39-abi3 platform tag at lines3952 to3958. The runner OS label does not change that required wheel platform, prove the new output has that tag or qualify an unbuilt new artifact. A higher-floor output cannot be accepted by merely relabelling the runner migration successful.

Because the workflow diff is exactly two runner scalars, all existing native wheel clean-install/runtime, installed typing/stub, CLI smoke/archive/readme/license, attestation, exact-main/sprint-tag, selected-family and tag-only publication conditions remain unchanged. The review introduces no skip, error swallow, alternate publishing branch or credential-scope change. The new macOS toolchain/environment still requires actual successor artifact validation at the original requirements.

## Policy mechanism and prose

scripts/test_sprint_workflow.py:5387 updates the exact CLI expected platform tuple to macos-26 without changing the other tuple fields or platforms. The new mutation at line9206 changes only that real CLI row back to macos-14. The existing assertions at lines9237 to9239 require a genuine source mutation and rejection by assert_rust_release_assets_contract. This is a meaningful regression discriminator, not a weakened allowlist or a no-op mutation. Existing wheel, release-asset and security mutation checks remain in place.

Root reported the two focused policy invocations exited0 and the full policy suite was queued. This advisory inspected their source mechanism and did not run or independently relabel those tests. Actual final policy results remain Root's own retained verification evidence.

The F275 addition at .claude/plans/F-X275-design.md:916 to920 accurately attributes the old scheduler failure, limits the repair to two selectors and preserves exact-platform/check/publication requirements and failed-attempt history. HLD15 at docs/hld/15-build-and-toolchain.md:1103 to1109 records the supported route with an official source and explicitly retains the macosx_11_0_arm64 contract. Neither addition broadens F287, waives native layout qualification, changes product API/support scope or records a hosted pass.

## Remaining execution and boundaries

Root's final review and affected policy gates precede its workflow-authorized clean checkpoint and push. The exact successor normal CI and subsequent26-payload build-only rehearsal remain required. The old rehearsal must first reach a terminal state, and the fresh rehearsal must authenticate its own complete selected artifacts and skipped publication jobs. This advisory did not query or assert the old whole-run terminal state, dispatch a run, consume signed authority, commit, push, publish or repeat Cargo/UI/product tests.

READ_ONLY_FACTS.json binds the current four reviewed paths, unchanged runtime policy and normal CI references, exact two-scalar workflow comparison and retained annotations. REVIEWED_DIFF.patch preserves these hunks, including F275's already-reviewed carry text as context only. The earlier ten-file F287 advisory remains its own completed phase, not another F272 peer or a reopened source review. No automatic remediation follows. End this advisory and return control to Root.
