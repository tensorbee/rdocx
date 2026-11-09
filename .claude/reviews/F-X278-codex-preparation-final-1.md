# F-X278, Codex final preparation pass 1

REJECTED PREPARATION. Actual Root final after three completed peer passes. No peer 4, source implementation, canonical adoption, approval or claim. The reviewed Grok V3 is a draft only. Q-S96-07 remains unanswered.

## Identity and scope

Proposal /private/tmp/rdocx-grok-harness/f-x278-preparation/GROK_PROPOSAL_V3.md SHA-256 9f77eeb94e2668d5fd05d3d5ac243d888242bd539b27512f3eaf685ab33a0945, 56784 bytes. Native handoff b5bb5c740c9f8f909c4bf9fd86d8e7d8d1173014ae4497412663c65f1f8e02b9. Actual DeepSeek third review a14b68246d07582a1e02c38e2184a30daf762477e77bf65a67d2727749f8a05e. Root read the complete proposal, actual third review, original registered criteria and approved F-X275 release policy. Independently traced Python nonempty-font branches, CLI layout calls, Arc cache types, cell paragraph callers and vertical-text merges against scoped source. The proposal's 919-file packet is historical, not a clean acceptance Base.

No source edits, builds, tests, native actions or network calls. Report writing waited until the separately owned four-file source lease had returned. That is coordination, not a provider edge for F278.

## D1, high, nonempty Python font requests have no concrete policy-preserving route

GROK_PROPOSAL_V3.md:256 promises one selected tier for nonempty caller fonts, and :334 promises diagnostic parity. The explicit non-diagnostic Rust methods listed at :218 ignore the selector. Python page_count still names its additional-font branch at crates/rdocx-py/src/document.rs:12516, and to_pdf calls to_pdf_with_fonts_and_options at :13514. The diagnostic to_pdf branch calls the method being changed to consume the selector at :13540. The private policy helper cannot be invoked directly from the Python crate. A keyword and a broad same-tier assertion do not close this callable gap.

Name the exact existing public call selected for every nonempty branch under A and B, including PDF, PNG, SVG, page lists, all pages, layout and layout_page. Keep diagnostic and non-diagnostic siblings on the same tier. One caller font with both policies is the discriminator. A nonempty font list must not silently select the normal fallback engine under A. This confirms peer3 finding 1 and reopens the claimed complete R-A278-7c closure.

## D2, high, CLI flags are ignored by their named layout calls

GROK_PROPOSAL_V3.md:281 adds selector flags while :218 says the explicit bundled-fallback methods ignore that selector. Actual convert calls layout_with_fonts_and_bundled_fallback_and_options at crates/rdocx-cli/src/commands.rs:833. Render uses it at :2248, and layout uses the no-options bundled-fallback method at :438. Merely setting RenderOptions.use_system_fonts therefore leaves A opt-in and B omission on bundled fallback.

Give convert, render and layout concrete policy-aware call or branch contracts. Add acceptance and local/final discrimination for all three flags, not only the four issue entry points at proposal :354 and :366. Confirm peer3 findings 2 and 10 as one caller-contract defect. Inspect stays explicitly deterministic.

## D3, medium, helper owned return contradicts shared-cache contract

GROK_PROPOSAL_V3.md:190 declares Result<WordLayoutResult>, while :198 selects cached methods returning Result<Arc<WordLayoutResult>>. Those signatures are present at crates/rdocx/src/document.rs:24995 and :25020. Proposal :432 promises shared Arc results for accepted omission. Specify one compatible concrete return, preserving the cached Arc and wrapping uncached owned caller-font results when necessary.

Correct peer3's wording: Arc::clone preserves sharing. Cloning the inner WordLayoutResult is the operation that would break the promised shared result behavior. This is a contract-type defect, not a new executed compiler failure.

## D4, medium, Alternative B Default is unspecified

GROK_PROPOSAL_V3.md:208 keeps Default and :301 requires a true boolean default under B. The current struct derives Default at crates/rdocx/src/document.rs:73, which would initialize a newly added bool to false. State the manual Default implementation for B and its exact feature-off behavior. Test omission separately from explicit bundled selection. Peer3 finding 8 is confirmed. Q-S96-07 still selects neither alternative.

## Remaining smells

S1, cell shading consumer inventory. Proposal :124 omits the direct paragraph calls at crates/rdocx-layout/src/table.rs:1735 and :1869 and the non-table wrapper call at engine.rs:11215. Name the forwarded cell shading or explicit None at each call. Include a control paragraph and nested table control. Peer3 finding 5 confirmed.

S2, vertical first-line merge. Proposal :86 omits merge_direct_ppr callers in the tbRl path at crates/rdocx-layout/src/engine.rs:4802 and :4805. Add this consumer and a discriminating inherited first-line/hanging test without changing unit rounding. Peer3 finding 6 confirmed.

S3, WASM engine terminology. Proposal :444 calls Alternative B output deterministic, but crates/rdocx-wasm/src/lib.rs:85 calls to_pdf and B chooses the normal engine/cache. With system discovery disabled that engine uses bundled fonts, which alone does not make it the deterministic engine. State bundled normal-engine output, or explicitly specify and justify a binding route change. Peer3 finding 7 confirmed.

S4, wrong no-default-features evidence owner. Proposal :463 assigns the new rdocx selector error to cargo test -p oxml-layout --no-default-features. That command cannot test a policy helper in the rdocx facade. Retain the shared rider and name the applicable rdocx or Python feature-off invocation for the explicit error. Peer3 finding 9 confirmed.

## Refuted peer finding and public API inventory

Peer3 finding 4 does not establish a required oxml-layout 0.16.0 bump. Approved .claude/plans/F-X275-design.md:31 prepares the shared family at 0.15.0, and :49 explicitly identifies 0.15.0/0.17.0 as first unpublished minor lines. publish=true in crates/oxml-layout/Cargo.toml:12 proves publishability, not an existing publication at that version. The three public additions still belong in the shared 0.15.0 API inventory and exact published-baseline/compiler checks. Do not bump the prepared family again based solely on that manifest or describe shared API as unchanged. PowerPoint behavior remaining unchanged is a separate claim.

## Nits

Proposal :163 should say existing behavior remains unchanged when a private helper is shared, rather than promising unchanged bodies. The explicit-method inventory at :220 omits layout_with_fonts_aliases_and_bundled_fallback and its options sibling, present at crates/rdocx/src/document.rs:45983 and :45996. HLD10 replacements at proposal :432 and :436 should explicitly retain layout_page's route through layout_with_options. These are peer3's three nits.

## Coverage and gates

All seven A-IDs, original unreduced issue examples and named residuals remain present. No silent Q-S96-07 choice, new product source file, module, crate, trait, generic, compile feature flag or test binary was introduced by this preparation. The proposal retains deterministic hash baselines, no donor-output acceptance, the deliberate tint and rounding rules, Word-only symbol scope and named PowerPoint controls. No additional findings in those categories.

Rendering claims remain unexecuted. Native and pinned LibreOffice evidence, original scripts, actual 14-key hash effects, installed Python typing and wheels, WASM, archives and hosted CI remain due in the integrated union. Missing qualification is not a new product choice or manufactured test pass.

Return to orchestration. Grok author fixes the full contract, then Root affected final recheck. No new peer cycle. Do not approve, claim or implement F278 while Q-S96-07 remains unanswered. This pass ends when its required review artifact is written.
