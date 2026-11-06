# F-X279b, Root final, pass 1

**Reviewed**: Frozen working diff against Base391280577110c9b7b92ace0420a4a1d98c707889,13 tracked paths,2543 insertions and85 deletions. This is Root final review after actual independent implementation peers1and2. It does not reset the original three-peer cap.

**Verdict**: 1 defect,0 smells,0 nitpicks. Return for bounded remediation.

## Defects

### D3, ignored formatting widens protected-owner admission

`crates/rdocx/src/comparison.rs:10627`
`crates/rdocx/src/comparison.rs:10631`
`crates/rdocx/src/comparison.rs:13055`
`crates/rdocx/src/comparison.rs:13190`

The new coalesced paragraph predicate participates in protected-owner equality with the caller's ignore_formatting option. That predicate omits the run-property spans when the option is true. The protected source residual separately omits known modeled properties, so the combination can admit a changed ordinary run property in a paragraph containing an existing revision. This extends the exception from equal-format physical run partitions to a real typed formatting change.

The approved contract requires equal formatting for coalesced equality and retains refusal for edits to an F272 revision-bearing owner. See `.claude/plans/F-X279b-design.md:225` and `.claude/plans/F-X279b-design.md:229`, with the original whole-owner rule at `.claude/plans/F-X272-design.md:22`. The ordinary unprotected ignore-formatting policy does not authorize loosening that protected admission rule.

Root executed the concrete discriminator using the actual authenticated successor wheel: a paragraph contains an unchanged Old insertion9with text old, followed by an ordinary bold run with text plain. The edited document changes only that ordinary run from bold to italic. With ignore_formatting=true, all three granularities return success, no diagnostics and the same package bytes, retaining the old revision. The required outcome is location-qualified refusal without publication. Retaining output bytes is not sufficient when the operation admitted an unsupported changed protected owner.

The successor wheel is run-9517cr09, SHA4498124214c76bb979d1ea8422368ac8ee3c4c4c1034f5f70e9093ad9e7a672f, extension SHA40eeb5099122894f12b50bebc11e647a8dc90f75abeae1f66c27c5c0abf3da72. Root verified the loaded extension digest before the in-memory probe. Its results were:

| Granularity | Actual successor | Package |
|---|---|---|
| run | Accepted,0 diagnostics,1 old revision | Unchanged |
| word | Accepted,0 diagnostics,1 old revision | Unchanged |
| character | Accepted,0 diagnostics,1 old revision | Unchanged |

Root also executed the identical public native calls against the preserved accepted-provider wheel, SHA24c0a43370b9b48694a6f12915763e36f0190d18415f596a1ef993191e48454a, extension SHA92894246b6175a207d693df0f4f807fd550651abc4475ecfc84b9c843387b3ca. Each granularity refused with comparison cannot change body/paragraph[0], which holds an existing tracked revision: the edited document changes it. Each retained the original bytes. This was actual execution of the old source-qualified artifact, not a new build at the claim SHA. Its affected product-source equivalence is retained in the earlier L1_ACCEPTED_PROVIDER_BASE_GREEN.json proof.

The fixture used a generated default package with only its main XML replaced in memory by a canonical Word document containing that paragraph. Both modules were authenticated from their retained wheel payloads. The old extension was loaded from a temporary directory under the review directory, then removed. No source, test or fixture file outside the review directory was written.

Keep the protected coalescing admission format-sensitive independently of the ordinary caller policy. Preserve the positive protected-owner case with equal properties and physical partitions, the ordinary mixed-format ignore-formatting behavior, both owner directions and atomic refusal. Add this actual current-source discriminator before repair. Do not globally change caller options or the shared paired Field policies.

## Smells

Zero smells.

## Nitpicks

Zero nitpicks.

## Closed peer findings and examined scope

Peer D1 is closed at source level and by actual local regressions. The existing checked authoring replay now becomes the actual edited document before relationship remap, model and Source extraction, so ignored body and related values are not copied from the caller's changed input. Snapshot rejection removes only the owned added reply identities before restoring the old comment parts. Peer D2 now uses concrete row and cell preparation through both typed table and row control collections. The existing nesting, explicit font components, bounded table/style precedence and destination-default preservation remain.

Read the complete production diff, approved operative contract, HLD03/04/10 changes, public Rust and CLI assertions, Python fixture and lifecycle assertions, peer reports and retained local evidence. Traced coalesced equality and boundary coordinates, ordinary versus paired projection, physical content and formatting owners, one-word revision envelopes, Source interleaving and existing paired nontext publication. Traced native and Python insertion policies, empty-source and clamped destination paths, authoritative body preparation, numbering reuse and companion/relationship exclusions. No other concrete correctness, contract, panic/indexing, OOXML order/namespace, API/stub or structural finding was established in this bounded review. No crate, module, trait, generic abstraction or feature flag was added.

## Evidence authentication and limits

SOURCE_PINS.json SHAa40ca895d24e25fe5f966131dbd884bf8b327f0e91971b2b56056b40faa5b3d9 and all16 current source/configuration entries authenticate. All3 context entries also remained unchanged before and after review. The independent pass2 report is clean and remains historical as written. Its pending Python statement is not rewritten.

The later actual installed successor result passed exactly17 selected nodes, including the three peer-repair nodes and all four C1 provider/lifecycle nodes, with no skipped/error/failed JUnit case. Root checked the real wheel digest, every recorded archive member, imported extension identity, successful build/install/authentication receipts and exact JUnit membership/digest. This existing green batch does not cover D3. Its local opt2 macOS CPython3.12.13 configuration is not final release, hosted CI, full installed or typing proof. Existing public18, Native29, CLI4and9 and strict scoped observations remain their actual recorded executions, not Root reruns. Source6 reuse remains historical per-case branch applicability.

Final integrated full suites, native Word comment-candidate and comparison outputs, schema/SDK, hash/pixel, installed/typing matrices, performance, packaging, hosted CI and sprint acceptance remain due. The approved Q28 case stays unproved. Deeper optional security assessment remains deferred in its existing document. No gate was waived or publication authorized.

Only this review file was written. The pass ends before remediation.
