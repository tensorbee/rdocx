# S96 eager owner and portable fixture support, pass 1

**Reviewed**: Frozen canonical working delta against H9 e76ea82b010c4053a8b0212ac641b34f3b363e97, six paths,170 insertions and6 deletions. This is a bounded integrated microscope support pass, including both production repairs, their tests and operative contract changes. Earlier private candidate analyses are supporting inputs rather than this formal review. Root retains final review and remediation authority.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Source and contract review

Constructor validation at `crates/rdocx-py/src/document.rs:12295` now uses the authoritative Native whole-story snapshot provider and the same selected paragraph parser. Native `crates/rdocx/src/document.rs:34821` retains story source production, and its owner snapshot implementation at35010 retains qualified scopes, physical owner admission and link resolution. Old optional table/piece projections were acceleration rather than an alternative XML admission authority, as shown by `crates/rdocx-py/src/document.rs:5010`,5130,5561 and8309. I traced both successful piece inventory and its full fallback. Selected paragraph parse errors and Native provider errors retain the public conversion, related-story and missing-link refusal. This does not make an unchecked typed model authoritative or introduce Source/Work credit. The changed constructor retains no owner/piece/snapshot authority, with first activation rebuilding and allocating current authority at12323 and positive dormant replacement still validating at12778. No new supported ingress, API, quota or error deferral was found.

Plain piece admission at `crates/rdocx-py/src/document.rs:5350` accepts only successful existing checked Native facts. Err and None take the old full route. `crates/rdocx/src/document.rs:34333` checks plain/no-section eligibility, resources, actual namespace-complete writer and reparse. Its predicate at8495 excludes Fields, controls, hyperlinks, revisions, drawings, equations, bookmarks and raw paragraph children. Property metadata remains under the real writer/parser. The resulting paragraph location uses the actual body context and same unfingerprinted body route used by piece inventory at `crates/rdocx-py/src/document.rs:5700`. Placement still rebases physical body slots and source context. The other block-owner consumer remains table-only, and the shared activation/fallback paths retain their authority. No alias/raw bypass, foreign-prefix rebinding, resource error suppression into successful facts, route drift or hidden retained Source was found.

The added facts test at `crates/rdocx-py/src/document.rs:20401` compares original full metadata and owner addresses across plain/rich/aliased/property-raw/control cases, requires actual checked direct-fragment producer selection, and checks actual invalid caller error equivalence. Constructor tests at20456 compare the old validator across nested tables, block control, Fields, alias/raw, links and related shared/case-equivalent targets. The malformed selected header must actually error. Empty dormant stores and zero piece work discriminate transient construction removal, while existing factory and activation tests remain. I inspected inputs and assertions, rather than treating the25 total as proof of every combination.

Portable SmartArt setup at `crates/rpptx/tests/integration.rs:498` calls a deterministic common theme helper, not the host Office resource loader. `add_smartart_theme` at6776 supplies only the existing default theme/relationship. The existing resource-dependent renderer setup at6792 still calls the shared helper then performs all original resource/data preparation. The unsupported, overflow, exact single local cause, absent duplicate global cause and byte-preservation assertions at504-530 remain. This removes an environment prerequisite from an intentionally unresolved contextless fixture without pretending it became native-renderable or silently skipping it. No new generic/module/forwarding abstraction was introduced.

Plan amendments preserve operative A283-13 and cold407 contracts at `.claude/plans/F-X283-design.md:1303` and `.claude/plans/F-X268-design.md:717`, with actual timing failures still pending and no threshold changes. HLD10:4987 matches the implementation. User-approved Q-S96-NATIVE-01 at `.claude/plans/F-X267-design.md:396` preserves legacy and16.114.1 trust, six SmartArt families, nine exact-time timeline samples, existing font/capture provenance and thresholds. It authorizes requalification16.114.4, explicitly not an already passed capture or global retarget. No new native pins are fabricated in this delta.

## Evidence and limits

The retained combined binding-unit stdout records25 passed,0 failed,0 ignored in249.59s. Root reports actual command terminal0. The retained portable fixture v3 stdout records1 passed,0 failed,0 ignored. Prior v1 missing-theme failure and v2 omitted8MiB-stack failure remain predecessor observations. This pass performed no tests or compilation. The current Clippy stderr now contains a completed dev profile line in29.08s, but this review does not infer terminal process exit from that text. Root owns its terminal receipt.

Successful units do not qualify an ordinary current shipping wheel, cold/open timing, flat-cost thresholds, native application comparisons, final union, sprint completion or release. Those explicit pending criteria remain intact. No native capture or performance measurement ran in this review.

## Authentication

Diff SHA256: `ca368975db6fc2543c7d857714b0fa6cd50ef859834a17f590edde9ea238b8fe`.

| Frozen path | SHA256 |
|---|---|
| .claude/plans/F-X267-design.md | `bb2423cc6bd6f8a381b22d36574746a43e67acbc4e9ff53df29af43aa427f945` |
| .claude/plans/F-X268-design.md | `123c7f3c96392d5b13d10ab4f5bd711b25f0cc7bde0becdb7ec82ce5eb3acf29` |
| .claude/plans/F-X283-design.md | `8720f509a4b396ec5d77854dcf2cdd06fe263c1e2d99c0e4964aee63a72bbbaf` |
| crates/rdocx-py/src/document.rs | `77a956f4ed195d2c18bafb0e08a2f531aff2dd6eb5ba3857bd2f73dd318f9bd9` |
| crates/rpptx/tests/integration.rs | `a6653fa088a3c8aa0651773c41702332bb94e56f8cbcaf89c0ea46c26801435c` |
| docs/hld/10-bindings-spec.md | `8e66be501acd1c9e6d29ea469c47fc442c7ea0015af48f71a55b4578ac4d71dd` |

Retained outcome logs read in this review:

- `/private/tmp/rdocx-s96-prepared-integration-readiness/FX237-H9-open-source-triage-1/ACTUAL_COMBINED25_GREEN.stdout` SHA256 `e38eece8f1d4e2a2ed390e238031e6c6fbebd2b08d0659f1ddfee415c0a78dc5`.
- `/private/tmp/rdocx-s96-prepared-integration-readiness/H9-ordinary-timing-queue-preparation-20261107/admitted-successor-1/smartart-msrv-fixture-proposal-1/ACTUAL_GREEN_v3.stdout` SHA256 `f43fd9ad1203f24b31193b4f5e31753cbbd8b5ccbea72dd7d00ab792f849eaea`.

All six source hashes were checked again immediately before report write and remained unchanged. Only this review file was written.
