# F-X273, generated-number and K1 provider prefix, pass 2

**Reviewer**: independent Codex reviewer, not an implementer of this remediation.
**Scope**: focused D1remediation after provider-prefix pass1. Current font.rs SHA25623f2973bd1b8e2487925bfd5a01be46eb71f44196915d3960b8d320ecb2d2bb7. Independently reconstructed the prior font prefix in memory by removing only the new test block and reverting the one branch condition, obtaining the pass1hash aeb73aeaef1ca1cb81e65dc86af8c988906efb28ea34e12a0981b7ec855e14f0. The other six prefix files match their pass1pins. F-X272changes remain excluded.
**Verdict**:1defect,0smells,0nitpicks. The original nonempty D1case is repaired, but the same no-wrap invariant still fails at an empty styled occurrence. Prefix readiness is not granted. No whole-feature or native acceptance claim.

## Defects

### D1, the empty explicit-kerning occurrence still invents an interior no-wrap break

`crates/oxml-layout/src/font.rs:2625`
`crates/oxml-layout/src/font.rs:2641`
`crates/oxml-layout/src/font.rs:2537`
`crates/oxml-layout/src/line.rs:1405`
`crates/oxml-layout/src/font.rs:5192`

The nonempty branch now uses legacy_shaping && !no_wrap before collecting per-segment Unicode break opportunities. This correctly repairs the observed "ab cd" legacy no-wrap case and retains the ordinary wrapped interior-space positive.

However, the preceding empty-segment branch still passes literal true as break_after when constructing a MultilingualTextSegment. It applies whenever explicit kerning is requested, independently of both no_wrap and legacy_shaping. A valid ordinary empty TextSegment with no glyphs, no advances and zero width is admitted by the existing constructor. It is not evidence of a forced break or of a known structural marker producer.

The concrete source-derived input is three LTR styled occurrences containing "ab", "" and "cd", with three enabled entries and no_wrap=true. The two nonempty occurrences acquire no soft breaks. The middle occurrence takes the empty branch and carries break_after=true. split_into_segments then appends that occurrence and flushes the current group on its break_after value. It therefore exposes a soft boundary between the two halves of the otherwise unbroken word. The same branch is reachable with either legacy flag, while the existing generic no-wrap provider skips empty text. The approved empty-placeholder/context retention rule does not authorize deriving forced structure from an arbitrary empty text value.

The new regression loops both legacy values but supplies one nonempty occurrence only. It cannot distinguish this zero-length styled-occurrence boundary. Retain its successful nonempty/wrapped controls and close the shared no-wrap invariant without dropping empty occurrence context. This is the same D1fault class on another admitted input, not an unrelated new issue. No new runtime execution was performed by the reviewer, so this remaining case is source-derived rather than relabelled as an actual RED.

## Actual remediation evidence

Independently authenticated all ten seal members in each of fx273-nowrap-public-provider-red-v1and green-v1, both source-before/source-after identities, actual successful builds, exact child execution and frozen binaries. Green's original executable still matches its frozen hash. The old REDartifact is retained separately, not compared to the overwritten original path as if it were current.

| Observation | Actual result |
|---|---|
| Stopped-source public-factory regression | Build exit0, child exit101. Raw panic identifies legacy=true exposing a no-wrap soft break |
| One-condition successor | Build exit0, exact existing body1PASS/0failed/0ignored, including legacy/generic no-wrap and wrapped offset3controls |
| REDreceipt |0202d6f503e058b54df8a0067033cdc999f881ae5d483eff1623a5444e4d6efb |
| GREENreceipt |22143ab08665bd4fbdeb51a981889e15f142e0773fbec9a554ad76f1795b7491 |
| GREENseal |eace378cc5557197eb68104120badeba7fff675ed4c1f29bded6ca9e3500694e |
| Strict scoped lint | Two crates, all targets/features, warnings denied, exit0in4.20seconds, source stable |
| Lint receipt |751fb55eacf60afa4be67a90342a316aca94068e31ca5473ad7d350d5dd8a12b |

The earlier ten-test prefix observation remains applicable to unchanged source/assertions, with this fresh changed-body execution recorded separately. It is not described as a repeated actual10test suite. The newly passing body still does not establish the empty-input branch or complete provider no-wrap conformance.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Not found and retained limits

No additional independent correctness, contract, panic, structure, test or OOXML finding beyond D1. The mixed ordinary-script/itemization concern remains unproved and is not manufactured as a second defect. Other prefix conclusions from pass1remain bounded to their inspected source and evidence.

The full main Word/annotation/field-fallback K1wiring, private positioned companions, R4transport, producer-associated cache admission, warm/cold and section/cache composition, native artifacts and final integrated risk/hash/installed gates remain pending. Existing LOCAL provider completion metadata does not waive this debt.

Only this report was written under root's released review-file lease. No source, tests, plans, HLD, progress, TEMP or other file was changed. No build, test, diagnostic, native call, oracle, commit or push ran during this pass. Return control to root for distinct remediation with D1still blocking prefix readiness.
