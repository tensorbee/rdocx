# F-X273 provider prefix, pass 4

**Reviewer**: independent Codex reviewer, not an implementer of this repair.
**Scope**: bounded legacy ordinary shaping-chunk correction. Three product/document files, 326 added and 11 removed lines against the actual pre-lease copies. Read the current approved plan, canonical microscope, previous prefix pass 3, actual incremental patch and affected consumers. This is not full F-X273 or main Word readiness.
**Verdict**: 0 defects, 0 smells, 0 nitpicks in the examined repair scope. The earlier mixed ordinary-script concern is resolved by the source-qualified witnesses, approved amendment and this finite repair. Prior D1 no-wrap and empty-occurrence fixes remain intact.

## Identity and evidence authentication

Plan SHA256 is 464fee768d79181b3580db9a5b126d3b03f671ce1af04daf67f808b441b9cb24. Product incremental patch SHA256 is 19ed2218fbe932395c51ac9f5eecd6f56c2d73b49931ca2d6751dc270a5fc4d3. Handoff SHA256 is cead06a16cb4c237e30eb3f32616f05dea089e8a9e35ae79d2933b18eaf0da1a. Independently authenticated all 106 packet seal members with no mismatches and checked the actual current after pins.

| Input | Current SHA256 |
|---|---|
| crates/oxml-layout/src/font.rs | 957a4102debb6a6c996553b3aca6897835d9a693a9c6f37a28bf4e755734ed7e |
| crates/oxml-layout/src/line.rs | 2346018b8b260633c1a07d9bb9960b183dde04914c42d4b91317add4447c4987 |
| docs/hld/08-rendering-spec.md | 4fddb05d6d2cc66a3c00f02dbf9efffe317e73a92e73a0661c3c6a38132c36ee |

Root additionally authenticated all 3839 protected file pins. The review does not attribute unrelated sprint changes to this repair. The local progress after pin also matches the handoff. No source changed during the frozen acquisition/review handback.

## Production review

At font.rs:2661, legacy shaping now chooses actual per-segment Unicode cuts regardless of no_wrap. At font.rs:2672, published opportunities use the empty paragraph set for no-wrap while retaining those shaping cuts. Wrapped legacy mode uses the same shaping/published offsets. The private existing level helper takes a pair of existing set references, without a new public type, method or argument. Generic callers supply the same set twice.

At font.rs:2740, only legacy mode suppresses coarse script-only splitting. Actual bidi-level changes remain barriers. Explicit-direction levels and their validation are unchanged, and do not pretend to be ordinary Auto equivalence. The legacy branch selects the existing nominal face and guessed-properties shape_legacy call at font.rs:2896 with the occurrence's explicit kern. It does not spread a pre-break occurrence's guessed script across old independently shaped Unicode chunks. Retained script metadata does not override that guess in legacy context reshaping: shape_in_context at font.rs:3019 and reshape_context_segment at font.rs:3076 use the same legacy policy. Generic/rich shaping still consumes its explicit script, language and direction.

Logical/source character ownership, clusters, logical-index increments and explicit per-occurrence kerning remain associated with each occurrence. Empty styled inputs retain their context and false break. No-wrap nonempty ends now publish no break while still shaping old Unicode chunks. New ranges are formed only at existing grapheme and actual level boundaries. The introduced set pair does not add an unvalidated byte-offset source. Existing exact-capacity normalization remains on legacy outputs.

The nonlegacy condition is algebraically unchanged: script or level splits remain, shaping and publication offsets still coincide, and no-wrap still supplies empty offsets. Generic language/fallback/caps paths retain their existing caller and font selection. The generic direct helper passes the same existing offsets twice. line.rs production is unchanged. No new PPT-specific policy, public export, owner/cache authority, module, trait or feature flag was introduced.

HLD08 records the bounded provider behavior and explicitly retains unfinished main Word/K1, positioned field, R4 and native obligations. It does not describe the provider correction as completed product adoption.

## Assertion and measured-result review

Inspected all five extended existing bodies rather than accepting aggregate pass counts:

- font.rs:5295 adds both exact mixed Greek/Latin/Cyrillic/combining/fi witnesses, with and without the em dash, both kern states and both wrap modes. It checks actual legacy Unicode chunks, reference glyphs/advances/width, source intervals, context and normalized capacities. Equal and empty occurrences retain distinct kerning/logical indices.
- line.rs:2239 repeats the existing hyphenation observation against the same raw owner, checking deterministic context-preserving output without changing its producer.
- line.rs:2458 exercises explicit RTL/LTR and numeric level barriers with source slices and both kern states, keeping every no-wrap break false.
- line.rs:3108 compares actual old plain and candidate line results for the two retained witnesses at 468pt fitting, 20pt wrapping/overflow and no-wrap. It checks line count/width and per-chunk source, glyphs and advances. The no-dash 20pt case remains an unbreakable overflow, not a claimed wrap.
- line.rs:3518 exercises center, right and decimal tab alignment using context-shaped reference widths for both kern states.

These discriminate the witnessed extra script boundary and lost fi substitution while retaining actual Unicode chunking and later reshaping. They do not establish the future Word integration's positioned/rendered behavior. The earlier actual 15-route witness packets remain prospective low-level old/candidate evidence, not main Word or native parity.

## Actual acquisition accounting

The packet is /private/tmp/rdocx-s96-run-refresh/fx273-legacy-chunk-implementation-v1. The original 18-attempt batch has 17 actual passes and one ignored F206 physical small-caps oracle child. Its receipt remains NONPASS. Exit 0 from that ignored child is not credited as a test pass.

Root approved a separate existing nonignored Arabic-joining test from the same frozen binary. Its actual passing supplement gives corrected local membership of 18 nonignored passing tests from 19 total attempts, with the one ignored attempt retained. Frozen executable SHA256 is 57b62fbe07724cd4421eaebe3cf3e8f0dd0edad3a79e36aa4c2795971f42469e. The actual artifact is oxml-layout0.15.0 with opt1, debug assertions and overflow checks enabled, test=true, and the declared 8MiB test-stack environment. No rebuild or waived oracle is inferred from the membership correction.

Current check2 passed in 1.88 seconds after the retained check1 failure. Strict all-target/all-feature lint passed in 3.00 seconds. Format, prose and diff checks passed. The exact test receipts, raw logs, handoff and artifact bindings are sealed. No product execution was performed by this reviewer.

## Defects

Zero additional correctness, contract, panic, source ownership, compatibility, structure or test defects established in this bounded repair. No OOXML transformation is introduced here. Do not create a finding merely because full integration/native work is still pending.

## Smells

Zero.

## Nitpicks

Zero.

## Retained limits and handback

The F206 independent physical small-caps oracle remains pending. Main Word run/annotation/field-fallback K1 wiring, exact positioned field companions, R4 transport, positive private cache-producer association, warm/cold composition, final allocation accounting, repaired-artifact native observations and the integrated workspace/installed/typing/WASM/hosted/hash/risk gates remain due. Existing LOCAL provider metadata does not waive them. No whole F-X273, sprint or release completion is asserted.

Only this report was written under root's released review-file lease. No source, tests, plans, HLD, progress, TEMP or other file was changed. No build, test, diagnostic, native operation, oracle, commit or push ran. End this clean focused pass without a confirmation pass and return control to root.
