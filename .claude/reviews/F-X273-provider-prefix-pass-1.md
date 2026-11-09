# F-X273, generated-number and K1 provider prefix, pass 1

**Reviewer**: independent Codex reviewer, not an implementer of this prefix.
**Scope**: seven-file incremental prefix on sprint/s96 at6506fa660372b2cfe7dd620cfaaf0959a62ba640,1428added and262removed lines. Exact patch SHA256d4b3ac05f7dc106384827a49337fdd0c9b0de6e8f576e801982c7eea0d58228e. The real pre-lease copies and seven current source/plan/HLD hashes were independently checked against handoff.json before inspection and again immediately before this report. They stayed unchanged. Concurrent F-X272changes are outside this prefix scope.
**Contract**: read canonical repository/workflow/microscope guidance and the complete applicable F-X273K1, R4, generated-number and allocation amendments, alongside the existing rendering spec and delivery context. The existing done cell denotes the earlier approved LOCAL provider checkpoint. It does not assert completion of this remediation, native riders or whole F-X273acceptance.
**Verdict**:1defect,0smells,0nitpicks. The provider prefix is not yet microscope-clean. This is not a full-feature completion review.

## Defects

### D1, explicit legacy shaping ignores the no_wrap argument

`crates/oxml-layout/src/font.rs:2537`
`crates/oxml-layout/src/font.rs:2611`
`crates/oxml-layout/src/font.rs:2661`
`crates/oxml-layout/src/font.rs:2738`
`crates/oxml-layout/src/font.rs:2971`

The new approved public shape_paragraph_with_kerning method accepts and forwards no_wrap, but the legacy branch bypasses it when constructing break_offsets. The generic path uses an empty paragraph-break set for no_wrap=true. Legacy mode instead always collects unicode_linebreak opportunities from each segment and exposes those opportunities through the returned spans' break_after values.

A concrete source-derived discriminator is one ordinary LTR segment containing "ab cd", one enabled entry, no_wrap=true and legacy_shaping=true. The offset after the space enters break_offsets, splits the segment in the logical-range loop and makes its first returned span break_after=true. The corresponding generic no-wrap request does not expose that soft break. This is a provider-contract fault on a callable new public method, independently of the still-unwired main Word consumers. A later line-break parameter cannot make the provider's explicit argument effective or repair its returned span facts.

No approved legacy no-wrap exception was found. K1requires preservation of the existing Word line-break conventions and private context through reflow. Respect the no-wrap request while retaining the approved legacy versus rich shaping distinction, and discriminate it in an existing test body. The existing unwrapped hyphenation control exercises LineBreakParams.wrap=false, which does not distinguish this provider branch. This reviewer performed no runtime probe and does not label the source-derived case as an executed RED.

## Evidence authentication

The packet is /private/tmp/rdocx-s96-run-refresh/fx273-final-remediation-v1. Independently authenticated all31members of prefix-tests-v2/seal.json, the exact10named child outcomes and raw summaries, both original/frozen executable hashes, and the actual opt1/debug-assertion/overflow-check Cargo profiles. Every selected child exited0and reported1passed/0failed/0ignored. Source before/after agrees. This is finite prefix semantic evidence, not a full library suite or native comparison.

| Observation | Actual evidence |
|---|---|
| Exact test selection | Six neutral font/line bodies and four Word style/engine bodies,10PASS |
| Semantic receipt |2549a386866e6473a12d2a869d757ff61218d860a26e97b20d0e510e297241b1 |
| Execution seal |555ae2bab2356791562afffda3c3d72c9487b8e82d60dab7d31349912d559c10 |
| Source manifest |4aa16e3297ef07292ebeb80cd292c0e0977bfc59a083f8675a76aec2359459d5 |
| Scoped clippy | exit0,20.19seconds, source stable, receipt db363ad7d7828e8c5d20c710a15c1df4f3c52c163e7483bdbb85556ea9efa204 |

Inspected actual assertions rather than accepting counts alone: memo default/off/on identity and forced fingerprint collision, count mismatch and invalid base/advance replacement, exact private vector/language capacities, hyphen context/source boundaries, hanging spaces, tabs/spacing, existing rich line order, named style inheritance/toggles, font slots, generated nominal glyph properties and line counter boundaries. None of these assertions exercises the D1no-wrap legacy provider request. Earlier failed check/lint attempts remain historical and are not promoted to success.

## Prefix implementation audit

The six additive doc-hidden methods retain the approved signatures and existing public types. Generic shape_text still requests default shaping, explicit requests add only the kern feature, and memo equality/fingerprints distinguish None/Some(false)/Some(true). Existing caps feature requests compose with explicit kern. Font data, IDs, face indices and variations are not repurposed as policy identifiers. No ambient Word default is installed into FontManager.

MultilingualTextSegment keeps private kerning, legacy and hyphenation context. Compatible base replacement checks text/font/size/direction/glyph identity, finite advances and retained cluster validity. Context survives examined source rebinding, splitting, hanging spaces and hyphen reflow. Private Word-legacy vector/language capacity normalization uses boxed conversions at approved factory/replacement boundaries. Public base allocations retain their actual exposed capacities. Generic/PPT constructors retain their defaults and do not enter the legacy normalization path.

Generated numbering resolves the correctly typed trimmed primary names Normal and line number, following actual same-type basedOn edges. Missing, wrong-type, ignored-ID and cyclic edges terminate with bounded diagnostics. The nearest properties within each level feed the approved cascade. Bold/italic use the explicit document-default dominance and paragraph/character XOR policy without modifying generic run-property resolution. Source XML is not renamed or rewritten.

The Word label context is prepared once after font loading and begin_layout, then carried privately through section/pass/pager ownership. The resolved nominal family/size/weight/kerning tuple drives shaping, alignment width and painted glyphs. Failed face resolution remains Some(Word context) with no face, so it cannot silently become neutral9pt. Invalid size paints no glyph while numbering advances. Number leaves have neither field_kind nor field_source and do not create field-policy members.

Word omitted countBy disables numbering, explicit zero retains compatibility as one, and source start is a skipped count. The private u64counter covers all admitted u32starts plus one. Neutral public pagination keeps its previous start and9pt/default shaping path. Existing source schema, public geometry structs and pagination signatures remain unchanged. The HLD addition describes these implemented number rules without claiming the main Word K1or R4wiring complete.

The legacy factory still passes through script/bidi logical itemization before per-fragment guessed-property shaping. Mixed ordinary-script compatibility requires attention when completing the approved full K1consumer union. This review did not establish a concrete glyph/output counterexample for that separate concern, so it is not manufactured as an additional finding or recorded as a runtime failure. D1is the independently concrete ignored-argument path above.

## Explicitly unfinished work

This prefix does not wire main Word run/annotation/field fallback kerning, exact private field-only paint companions and late substitutions, full R4layer/row/anchor/note/endnote transport, or positive private producer association for cache admission. Those items are explicitly retained in handoff/progress. Their absence is not misrepresented as completed behavior and is not counted as a new defect in this bounded prefix.

Full restart/continuous-section/cache and warm/cold composition controls, native repaired-artifact comparisons against the retained outputs, the full risk/installed/WASM/hosted/render union and reviewed49hash/7pixel deltas remain due. The existing LOCAL provider checkpoint exception and done metadata do not waive them. No new native campaign, baseline update or final acceptance claim follows from10selected local tests.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Not found and handback

Zero additional independent correctness, contract, panic, structure, test or OOXML findings beyond D1. No unsupported public model/field, seventh public method, new module/crate/file, speculative trait/generic, font substitution constant, source clamp or numeric-unit change was introduced by this prefix. The missing D1discriminator is part of that defect, not a duplicate finding.

Only this review file was written after root released the report write. No source, plan, HLD, progress, test, TEMP or other file was changed. No build, test, native call, oracle, commit or push ran during review. Return control to root for a distinct remediation phase with D1blocking prefix readiness and all full-feature debt retained.
