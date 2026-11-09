# S96 integrated repair, bounded Root pass5

Read-only follow-up to passes3and4. This phase writes only this review record.
This is a source-availability review, not final sprint review or completion.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Causal styles repair

The actual hash failure changed only the showcase styles XML. The diagnostic
proved exactly three redundant nested namespace declarations, with all other
bytes unchanged. Removing those declarations in a diagnostic copy recovers the
exact baseline digest. Production never strips XML and the baseline is unchanged.

The existing private cache now keeps two distinct facts: actual parsed written
source and an immutable actual emitted-model witness. Source allocation identity
is checked first. A matching emitted witness retains the live authored model,
so later edits preserve original lexical emission. Actual parsed models remain
authority for staged reopen. Every cache constructor supplies a genuine parse,
including fresh compatible documents. Actual parse/proof work is counted.
Copy-on-write breaks an edited model's pointer witness, full equality remains
honest and replacement source bytes invalidate both cache routes. Successful
writer proof does not relabel a failed parsed-model serialization proof.

Independent exact proposal review is0D/0S/0N. The formatted adopted source is
reviewed here. All16style controls pass, including the new intermediate versus
uninterrupted write discriminator and exact parsed/reopen assertions. All49
hashes match the unchanged baseline in deterministic font mode. Full workspace
all-target/all-feature Clippy passes in37.61s with both bindings. Full workspace
Rustdoc passes in10.56s. Every receipt records stable actual source. Prior
failed hash and successful earlier artifacts retain their own identity.

## Remaining union and checkpoint

Pass3's other product-source findings remain0, with the later test-only coercion
accounted by pass4. Actual73regression,18Native and18lower-control results are
retained with exact original source pins. Style changes receive the affected
16control and49hash successor, with final complete/hosted suites still pending.
The four affected policy methods pass after the actual sandbox environment
failure. No product policy assertion was removed or relaxed.

The current rdocx archive successor is actually measured, source authenticated,
and below10MiB. Its root README and script tuple agree. Other21tuples and
performance rows are unchanged. The affected post-adoption repackage is pending
its actual terminal receipt and must pass before push. Other three source
archive packages retain authenticated unchanged inputs. Plan and HLD03 describe
the cache's current mechanism, with no public API or budget change.

The existing three CI source-availability records retain the H5failed run and
all final acceptance debt. Only the canonical sprint ref and normal CI may be
made available under the consumed checkpoint. Exact clean SHA, actor, ref,
remote and ancestry still require observation before push/dispatch. No main,
tag, reporter comment, yank or publication follows from this review.

## Reviewed current tracked diff pins

- `.claude/plans/F-X271-design.md`: `6db805eee25578fb084bedf9e884051e9b43c3f574d28aa4477928d058641332`.
- `.claude/plans/F-X272-design.md`: `db579acb1613fb5f15ad60f06dd8e8a5194366e282749eaf9874e544c6fa4b95`.
- `.claude/plans/F-X275-design.md`: `70c15978a568d27fd734998180d7c8d0aa95b96aff32fef3c12dd5827f7b0be2`.
- `.claude/plans/F-X282-design.md`: `d843e3c6453d42a20e470c77460194a59aa1bf80406ac084a170a6db6ad42b94`.
- `.claude/plans/F-X283-design.md`: `7ab81366c8193a4ac6346ea438968935a90fb9b326cd5066811f043a0d8cdb63`.
- `README.md`: `fe351390599ea7bef4bf8a57c77bb318c4c563fb37cc78eb6735e6aa939f0592`.
- `crates/oxml-layout/README.md`: `a08936afc5dda7e4231b918a5d39861acc37b45f43093de6b9e0dc7810e38fcc`.
- `crates/oxml-layout/src/font.rs`: `ec37e321714d850852c278af262ab0ff2bfe74dca1bd8df30a140690d29a4469`.
- `crates/rdocx-layout/README.md`: `b6bed679bad71bf5e4d3f3c1a7c29e712f9659528056cb51ebf2d831e280727d`.
- `crates/rdocx-layout/src/engine.rs`: `d42eb0186cb92e05fb42ede2d56817c595aafb8d35eabad9b9fd842e18394932`.
- `crates/rdocx-layout/src/paginator.rs`: `92a42a9286231d3b2b147728dec7636f535340800e26213d22683d040d0544e2`.
- `crates/rdocx-oxml/README.md`: `a739822c5d9c76b1872025adeb1049d55a7d890afc292f45030439e725f6bab7`.
- `crates/rdocx-oxml/src/footnotes.rs`: `1d6edff7d58f19f6a2378a2d63163d36a6fd66f80423b1996077f89959931818`.
- `crates/rdocx-oxml/src/revision.rs`: `1e8d4888a6ae9ca1981f9a41b408d799de9b7f8e90a6709bd33fb7fcba632c01`.
- `crates/rdocx-oxml/src/text.rs`: `6873a90b91a5935d3d2cb746a6587097ae1d029f6277b410c57be56637d65878`.
- `crates/rdocx/src/comments.rs`: `19dff2dae6b6152c98cb2f9cbe6e37e2cdb96483404f3cc6d503c6dafd983689`.
- `crates/rdocx/src/comparison.rs`: `8633ea329b1d68ec770b6de4f206edaa7cf42a79da5355a0a70e08b9e9c4d1ec`.
- `crates/rdocx/src/document.rs`: `981fe3dd35743d5103701da70ba7e36a7801a07c1cfd77846e4b2737ae68275f`.
- `crates/rdocx/src/field.rs`: `5b1dbefd42073b81ea231a8708626d8f47f1760cf1d0ec042638acb9eb849ba1`.
- `crates/rdocx/tests/regression_test.rs`: `d3c4c07da3d95a338d04f612ce2e940fee3e59d289fda05b3d0d0c76eb5f0712`.
- `docs/hld/03-architecture.md`: `3070a46f2d980cfa58e1cc893e19af1b33a8cf17e493ff7af1bfd68e5470d8f3`.
- `docs/hld/04-opc-and-packaging.md`: `5da5527fb122dc7f64c656de0945c203ae748ab6782eb7dafb42e4433afb5ae9`.
- `scripts/readme_doctests.py`: `fcb40e4dd0e158dfe168d9863062fc91ded4073831580ff069f4a12f6f9ee0ca`.
