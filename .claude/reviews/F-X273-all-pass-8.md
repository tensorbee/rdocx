# F-X273, all, pass 8

**Reviewed**: Focused working review of pass7 D1 remediation on sprint/s96 at HEAD `6506fa660372b2cfe7dd620cfaaf0959a62ba640`. The real-before incremental patch changes 2 tracked files, 62 added and 3 removed lines, with 100 total diff lines. Reviewed the three production guards, the extended existing test body, HLD08, approved plan, actual shared callers and retained RED/GREEN evidence. Earlier pass6 observations outside this increment remain historical and are not repeated acceptance claims.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. D1 is closed for this reviewed local remediation.

## Defects

0 findings.

## Smells

0 findings.

## Nitpicks

0 findings.

## D1 remediation and consumer inspection

At `crates/oxml-layout/src/line.rs:1112` both MultilingualText trailing-space arms now require uses_legacy_shaping. Genuine rich spans reach the prior trailing=0 fallback. At line1176 the decimal parser likewise admits only legacy MultilingualText, preserving the genuine opaque in_number.then_some(0) fallback, including an already active number. The inner prefix-width call is reachable only after that admission. Plain Text and HyphenatedText handling is unchanged. Legacy spans still use their own context-qualified prefix measurement. No ambient shaping default, owner context or new public API changes.

Traced PendingTab::add through LineState::push at `crates/oxml-layout/src/line.rs:843` and width_with at line869. Right, Center and Decimal creation at line970 share the repaired trailing rule. Existing public genuine shaping and explicit-context shaping remain separate from legacy admission. The change restores the approved generic/default boundary in `.claude/plans/F-X273-design.md:325` and existing-entrypoint output contract at line818. `docs/hld/08-rendering-spec.md:953` records that boundary. Its tab statement is scoped to the surrounding legacy-versus-genuine carrier discussion and does not remove existing plain Text tab behavior.

The existing test at `crates/oxml-layout/src/line.rs:3835` retains its prior plain and legacy Center/Right/Decimal assertions unchanged. The added block at line3926 constructs actual genuine owners for default None and explicit off/on, asserts legacy=false and exact override provenance, requires actual positive trailing advances, and compares unclamped Right/Decimal gaps against the original opaque fallback derived from actual glyph advances. It rejects all six D1 cases without invented width constants or weakened legacy positives. Center default restoration is source-mapped through the same trailing branch. There is no new Center runtime RED/GREEN claim.

The current normal PowerPoint route remains excluded from this PendingTab fault: `crates/rpptx-render/src/text.rs:1334` supplies no explicit stops, and the default stop at `crates/oxml-layout/src/line.rs:957` is Left. This does not assert that every public generic tab route is unreachable or that the full font/line increment is neutral.

## Independently authenticated evidence

Read the actual packet at `/private/tmp/rdocx-s96-run-refresh/fx273-tab-default-remediation-v1`. Authenticated all 181 seal members by byte count and SHA256. Seal SHA256 is `37ae949d54327d13ea868d75100628afeb4e1dd649337552b370ac591de6ef01`, handoff is `c3dc87f823680df0de7fecbd65a925ebc7030b901ca0d16cf07acf3e5eb5e93f`, and incremental patch is `e42cd1d3c32dfe85a5168ccaf524bb93a12deb556ac56ea0ff053fd4996c741a`.

All 3848 recorded current source inputs match actual files. Only line.rs and HLD08 differ between the real before/after inventories, leaving 3846 protected inputs unchanged. Current line.rs is `ee4795ba197620e4fac18033f3e84ae9749c940b06bfeb82eb969232e33289ea`, HLD08 is `4f711a2d714ad60a77b25385969bdaa317be6caf682f64ff83d009f5bb3a6101`, and approved plan remains `f39cf22966952420d61a45b8d88e26179442b3b2dc2ffb418b4ba55ef77f130c`.

The actual old-source test child exits101. Its raw stderr reports all six None/off/on Right/Decimal mismatches. This is an executed RED, unlike pass7 source-only reasoning. Current actual47 distinct named children, 19 shared and 28 Word, each have raw one-test PASS with zero failed or ignored tests. Inspected their raw named outcomes and hashes. The extended tab body passes alongside the unchanged legacy positives.

CargoJSON records genuine oxml-layout0.15.0 and rdocx-layout0.17.0 test executables and successful build-finished, with opt1, debug assertions and overflow checks enabled. Independently matched both original and frozen binary hashes: shared `57d6c55c4058df98a0f9735e2c00807400074ec60dded89e982aa270340be44e`, Word `b2135f622d331f0045f7fdab8c765baaf98f68ad92cef7bf47e870e6d0c078e5`. Build and deciding static-check before/after source maps match the final current input map. Retained check, both-crate all-targets/all-features strict clippy, fmt, prose, adapter and diff receipts have actual exit0 and raw streams. The transient Center edit/lint is explicitly withdrawn and excluded from deciding evidence. Final source matches the acquired build, without a second binary acquisition.

## Not found and limitations

No additional correctness, contract, panic, OOXML, test or structural finding in this bounded repair. No new traits, generics, files, modules or dependencies were introduced by the implementation increment. This review ran no builds, tests, imports or product experiments and changed only this review file.

The provider checkpoint remains historical. This report does not complete F-X273 or the sprint, establish native parity, authorize a baseline or prove broad output neutrality. Retained native/font/caps-oracle, full/hash/raster/installed/WASM/API/hosted/risk and final sprint acceptance debt remains pending. No completion status is changed.

End review and return the sole canonical writer lease to root. No confirmation pass or remediation follows inside this review.
