# F-X273, generated-number and K1 provider prefix, pass 3

**Reviewer**: independent Codex reviewer, not an implementer of this remediation.
**Scope**: focused D1shared-invariant review after prefix pass2. Current font.rs SHA256b855fca056cf6f625098d5f195102bb8052c6a5308e1705ef2a0e837e7a84e09. Independently reconstructed the prior pass2font in memory by removing only the added empty-occurrence matrix and reverting the empty constructor's break flag, obtaining23f2973bd1b8e2487925bfd5a01be46eb71f44196915d3960b8d320ecb2d2bb7. The other six prefix files retain their pass1pins. F-X272and unrelated sprint changes remain excluded.
**Verdict**:0defects,0smells,0nitpicks. D1is addressed for the reviewed public provider invariant. The scoped prefix microscope exit condition is met. This does not complete main Word K1/R4/cache integration, F-X273or the sprint, and does not supply native qualification.

## Defects

Zero. D1is addressed as described below.

## D1closure, no invented empty occurrence breaks

`crates/oxml-layout/src/font.rs:2625`
`crates/oxml-layout/src/font.rs:2641`
`crates/oxml-layout/src/font.rs:2661`
`crates/oxml-layout/src/font.rs:2971`
`crates/oxml-layout/src/font.rs:5222`
`crates/oxml-layout/src/line.rs:1405`

The explicit-kerning factory now retains an empty styled occurrence with break_after=false. Its logical occurrence, language, kern override and legacy context remain assigned through the same construction path. It is not dropped, shaped as ordinary visible content or inferred to be a forced structural break.

The nonempty branch retains the earlier legacy_shaping && !no_wrap correction. No-wrap requests therefore use the empty paragraph-break set for either legacy choice, while wrapped legacy requests keep their approved per-segment Unicode opportunities. The final shaped-span constructor derives break_after from that selected set. The empty branch and this nonempty constructor are the two public factory publication routes. Existing context reshaping, hanging-space splitting and bidi reconstruction preserve the source span's break/context facts rather than independently inventing a no-wrap opportunity.

The extended existing regression crosses both legacy flags, both no-wrap flags and zero/one/many styled occurrences: [], [empty], and [ab, empty, cd]. It checks complete logical text, retained empty occurrence count, absence of invented empty breaks and all-false breaks when no_wrap=true. The earlier nonempty ab-space-cd controls remain, including the wrapped interior offset3positive, so the repair cannot pass by suppressing every break opportunity. Existing request-count, memo identity, capacity and replacement validation assertions remain in the same body.

## Actual RED/GREEN evidence

Independently authenticated all ten seal members in each of fx273-nowrap-empty-provider-red-v1and green-v1, exact source-before/source-after equality, successful actual builds, frozen binary hashes and exact named child outcomes. Green's original executable still matches its frozen artifact. The stopped-source REDbinary remains separately preserved.

The actual REDchild exited101at the FIRST failing empty occurrence, legacy=false and no_wrap=false. The later true/no-wrap or multi-occurrence cases were not executed on that failed child and are not falsely credited as separate observed REDs. The corrected child exited0with1passed/0failed/0ignored, exercising the complete matrix and all earlier assertions.

| Observation | Actual result |
|---|---|
| REDreceipt |6dd53e34c0fe9e73ef42b2af3ed6698b379908623c61cd67da1a38589e01ebeb |
| REDseal |056bd2f93d937bfa6a5bb6218fc186bd99ebe936e9099ee74fabda24d7a58f6d |
| GREENreceipt |c6873d66ba286c7b2c43046a18d2e3729ab2c88ea57e450b8263e5c1ae66fccf |
| GREENseal |defdb47ae0d6cbed056db0307c739354863a24fd7befb0dd951c4dd1cf485c01 |
| Strict scoped lint | Two crates, all targets/features, warnings denied, exit0in7.23seconds, source stable |
| Lint receipt |22bfdc3989f7e5bda4cad17c674dfce7ce333524a5f518b987b005672327c410 |

The original ten-test prefix observation remains evidence for unchanged source/assertions. The two targeted successors are separate actual single-body observations, not a fabricated repeat of all ten tests. No final suite or native observation is inferred from this evidence.

## Smells

Zero.

## Nitpicks

Zero.

## Not found and retained limits

Zero additional correctness, contract, panic, structure, test or OOXML findings in this focused shared-branch remediation. The mixed ordinary-script/itemization concern remains unproved and is not manufactured into a defect. Previous source-qualified prefix conclusions remain bounded to their inspected files and observations.

Main Word run/annotation/field-fallback K1wiring, exact positioned field companions, R4transport and positive private producer association for cache admission remain unfinished. Warm/cold, section/restart/cache composition, retained native repaired-artifact comparisons, full workspace/installed/typing/WASM/hosted risks and reviewed49hash/7pixel deltas remain due. Existing LOCAL provider completion metadata does not waive that debt.

Only this report was written under root's released review-file lease. No source, tests, plans, HLD, progress, TEMP or other file was changed. No build, test, diagnostic, native call, oracle, commit or push ran during review. Return control to root. Stop this focused review after its clean result, without a confirmation pass.
