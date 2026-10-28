# F-X274, all aspects, pass 2

**Reviewed**: Complete frozen working tree against immutable Base and HEAD `8cb33ef53717159bcf73f655552b99589272a274`, branch `work/f-x274-codex`. Eleven implementation/contract files, 1852 additions and 244 deletions, plus the retained historical pass1 review. No implementation commit or untracked product file is omitted.

**Identity**: Full binary diff SHA-256 `8c3f6bf4ab2536b6893d00e1e39bde13c516e4a58cdfe33f0b51a4ce78b0b689`. Independently authenticated every changed-file hash, progress and prior review against `fx274-d1-freeze.json`, with zero mismatches. Full227 source manifest `d3d54971d69df2c2ec2d6a6f76a59fa7ac9b2bf906b1699cac2efd4fee8e2543`. Engine SHA-256 `9e24de8b01cadfadab18c6aabd6458a1d3e8cd5d5be5ea56c571d26e0aa6795e`.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. D1 closed. The scoped implementation is ready to return to the orchestrator. This is not final native or sprint acceptance.

## Defects

Zero open defects.

### D1, closed, explicit multilingual visual-order evidence is now exercised

`crates/rdocx-layout/src/engine.rs:30246`, `crates/rdocx-layout/src/engine.rs:30252`, `crates/rdocx-layout/src/engine.rs:30420`, `crates/rdocx-layout/src/engine.rs:30453`.

The matrix now actually admits the Hebrew/Arabic stored display. Both wrapper variants still execute the real FontManager-shaped direct carrier, shared formatting/refusal map and production substitution consumer. The explicit assertion requires the Arabic span origin to precede the Hebrew span origin. It is independent of the logical-index sorting used to reconstruct text. The new terminal execution-count assertion requires exactly `[1, 1]`, covering Group and MarkedContent independently, so deleting the relevant input or skipping either assertion fails.

The prior zero/one/many, Latin/Arabic, logical-text, metadata, baseline, validity and adjacent natural-advance checks remain intact at `engine.rs:30247`, `engine.rs:30382` and `engine.rs:30435`. No assertion was weakened. The actual recorded focused child exited0 with1passed in0.08s, using retained binary SHA-256 `ff0b925621ccb26631d71083893d0fae25817320cd1e920c3d6d99fce05970fa`. Its log hash matches `fx274-d1-checks.json`. This supplies the previously missing visual-order observation. There is no newly observed visual-order product RED, and the progress correction retains that distinction.

## Smells

Zero.

## Nitpicks

Zero.

## Full aspect reassessment and applicability

Consumed the complete pass1 contract, consumer and A-ID analysis, with the unchanged canonical instructions, approved annotated plan and HLD08/HLD10 intent. Independently authenticated the retained before-engine file against pass1 SHA-256 `ab3dd26b79f210c96d712ffe6f2d86f73c8a662ef0fd3a04e96b90ccbda46138`. Compared its actual bytes with the successor. The only changes are four insertion regions inside the existing cfg-test direct multilingual function: counter initialization, the new matrix value, counter increment and terminal assertion. All other changed-file hashes match pass1. The production-prefix and allother226 applicability proof is retained in `fx274-d1-test-only-proof.json`. This permits reuse of the prior independent source review without inventing new product executions.

- **Correctness**: zero findings. Reassessed the new direct-consumer discriminator and retained the unchanged analysis of `engine.rs:1000` shared formatting, `engine.rs:5396` per-placement owner/target context, `engine.rs:9076` ordinary substitution and `engine.rs:9144` rich substitution. PAGE identity remains independent of formatter support. Only an unplaced PAGEREF target loses TargetPage qualification. Group/MarkedContent and zero/one/many publication now have the previously missing explicit ordering proof.
- **Contract**: zero findings. Approved two doc-hidden concrete helpers, five formats, stage precedence, numeric bounds, dynamic cache policy, source nonmutation, consumer signatures and HLD impact remain unchanged. The original A1 input and separately labelled intended numbering control remain separate, with native recovery pending under design lines153 to161. No new public model, authoring API, default or policy expansion is introduced.
- **Panics**: zero findings. The new counter indexes are `usize::from(bool)` into a two-element array. The positive-control span lookups deliberately fail the test if either required script is absent. Production checked-depth, guarded one-span pop and bounded numeric allocation analysis is unchanged.
- **OOXML**: zero findings. The actual namespace source closure at `document.rs:1542`, both Empty/Start page-number capture branches, first-owner/opaque duplicate ordering and temporary writer projection at `engine.rs:1034` are unchanged. Exact ordinary canonical source, expanded aliases/default namespaces, producer-w shadow and opaque values remain covered by the applicable pure lower matrix. No general serializer or model repair is inferred.
- **Tests**: zero open findings. D1 is closed with actual assertion execution. Prior actual PAGE/CardText, unresolved PAGEREF, implicit final section, owner/note placement, raw source, generated targets, cache lifecycle and save/reopen assertions remain applicable. Styled/bidi Word inputs remain distinguished from the direct multilingual carrier because `engine.rs:14770` excludes dynamic fields from that producer.
- **Structure**: zero findings. The correction adds no source file, trait, generic, wrapper, dependency or public operation. The previously reviewed shared cache/formatter ownership and single direct quick-xml dependency edge remain unchanged.

Generated TOC physical target qualification at `field.rs:12076`, INDEX/TOF/TOA labels at `field.rs:25901`, cache staging at `field.rs:16580`, implicit final section aggregation at `field.rs:16747`, per-page reuse context and its capacity charge at `engine.rs:5462` and `engine.rs:7525` retain the pass1 analysis. The correction does not broaden direct-rich producer claims or change any of these publishers.

## A-ID and evidence disposition

| ID | Current disposition |
|---|---|
| A1 | Exact original witness retained. Local source interpretation and intended-control formatting are distinct. Actual Word recovery/output and native source save remain pending. |
| A2 | Applicable five-format, actual owner/target, selected first/even/default and warm/cold assertions retained. Native variants and final consumers pending. |
| A3 | Shared numeric/character/general-stage and CLI/cache/render switch discriminators retained. No formatter change in this pass. Native switches pending. |
| A4 | Repeated-letter rollover and exact bounded output/refusal controls retained. Library bounds remain separate from native limits. |
| A5 | Applicable qualified namespace, source nonmutation, Empty/Start, duplicate/nested exclusion, start mutation and reopen controls retained. Native source qualification pending. |
| A6 | Dynamic footer render versus physical saved cache, eligible body cache, held-handle lifetime and save/reopen controls retained. Exact pre-existing combined simple-field closure/dirty delta remains attributed to predecessor/current evidence. |
| A7 | Same-owner format/instruction and target-format warm/cold controls and capacity checks retained. Full installed/runtime/typing union pending. |
| A8 | Qualified provider/navigation/generated-target local controls retained. Final integrated prior-redline/furniture/merged-heading/format composition, native/PDF/PNG/text and exact fonts remain pending. |

Evidence directory: `/private/tmp/rdocx-s96-run-refresh`.

- Current focused D1 test, changed-crate strict clippy, fmt, prose, adapter and whitespace records are in `fx274-d1-checks.json`. Independently checked all recorded log hashes, with zero mismatches. No check was rerun during review.
- Earlier native19, CLI4, other lower15 and pure lower4 retain their actual original source/binary identities and the prior cfg-test alias applicability proof. They are not relabelled successor executions. The replacement direct-consumer case is the actual focused1pass above.
- Fresh wheel `7fd7c78096da1c34db3371a8a21eaffc1a779a868a3c165374ba817bbd4235d9` and isolated imported native `36dc3825120653caa756767ee0ac782d92b4e2d8c292d94a0210906b084f5e06` remain product-applicable because only cfg-test code changed. Python remains51 unchanged observed passes plus1 corrected-node observed pass on that same artifact, never a claimed single52pass run. Root composite authentication preserves the exact test reconstruction and per-operation footer predecessor/current bytes.
- Failed helper/source-capture/implicit-section/classification/direct-consumer histories remain historical. The old unreachable ordering check is not relabelled successful. The retained252-field result is carried only under its established applicability, without a successor execution or performance claim.

## Final debt and handback

All final native and integrated gates remain pending: source-qualified Microsoft Word and pinned LibreOffice/Poppler output, original423 recovery and saved source, exact font and geometry measurements, full workspace/native suites, installed Python3.9/3.12 runtime and strict typing/stubs, WASM, hash49/pixel7 with reviewed intentional deltas, archive/size/supply-chain, README/current measurements, hosted CI and the integrated A8/risk union. This zero-finding local review does not waive those obligations or claim native parity.

Only this review report was written. No tests, builds, native applications, product imports, temporary helpers, source/plan/progress changes or remediation occurred. Review pass ends here and sole worker control returns to the orchestrator.
