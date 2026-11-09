# S96 bounded collector reconstruction review, pass 1

Verdict: 1 defect, 0 smells, 0 nitpicks. Source review of the explicit new private collector and retained scoring transport only. This does not reset an F-ID review phase or provide product timing acceptance.

## Defect D1, complete new collector receipts cannot reach scorer publication

The new collector initializes its schema at `.claude/scratch/S96-final-evidence/f268-collector-reconstruction-20261107/s96_scaling_receipts.py:137` and completes/serializes it at `:183-207`, but never emits `all_pinned_timing_assertions_passed_without_xfail`. The unchanged scorer unconditionally reads that field at `.claude/scratch/S96-final-evidence/f268-collector-reconstruction-20261107/score_flat_linear.py:221`. A complete valid raw252 receipt and extra8 receipt therefore reach a KeyError before result publication, and the scorer reports incomplete/exit1 instead of the actual scoring results. The existing synthetic harness at `selfcheck.py:65-66` checks terminal raw completion and preserved pytest1, but does not feed the resulting schema through that scorer publication boundary.

Repair the transport field from actual recorded timing-call outcomes, preserving failed ratio results, actual pytest exit and separate original API-audit reports. Add synthetic end-to-end new-collector-to-scorer positive and failed-ratio controls. Do not infer success from wrapper exit alone or change clock/statistic/threshold bodies.

## Other inspected boundaries

The exact recovered costs_ms function retains its original clocks, factories, CALLS, min assignment and early stop. The AST transport captures the existing elapsed operand and observes after the assignment. Its stripping check reconstructs the original whole function AST. Original252 order plus final audit and strict raw phase coverage are retained. Raw pytest1 can remain a complete observation, with no xfail masking. Four/six-sample early-stop validation matches the original minimum-over-two-sizes condition.

The specification contains the complete252 catalogue,76 reused rows,3 extra reads and5 separate linear calls. The unchanged V3 scoring bodies retain inclusive79 bounds, five-linear bounds, strict107 historical non-growing regressions,69 growing residuals and explicit13 post-pin surface dispositions. No raw cases are intentionally dropped. The scorer identity-expression change correctly refers to the explicit new spec rather than pretending the missing historical spec was recovered.

All manifest entries and all15 immutable inputs, instrumenter and qualification-code pins matched at review. Missing10 historical metadata inputs are explicitly disclosed rather than authenticated. The22 retained synthetic observations are tool controls, not product timings. Current H5 context quiet flags remain false and later source/artifact admission is pending. No product import, Cargo, benchmark or source fix was run by this reviewer. The pass ends with this report.

## Key pins

- s96_scaling_receipts.py: `182b73ad16a251c9bfe27dc9154ad83f47d99ffcee048dd6b681aecfafd91db1`
- score_flat_linear.py: `0a00663a880923a95403633111441244edac127026a7abd2c13038791c619d93`
- qualification_common.py: `7e63b074e75d7fa464648a4973b845cbf8cb7c63d1dfeff3f7914e32b93715a8`
- qualify_extra8.py: `0bc8e41ef470b272ba6ee5d8c751400399143470b3669b50ddc34009affaed4a`
- QUALIFICATION_SPEC.json: `368dcaf846cb77ac00242252c86fd0210b71673a292f67c39a8085e7451a28ef`
- MANIFEST.json: `2a772c231dbe933ff46afec7e602545c9631061759777ab09707b838b5384be2`
