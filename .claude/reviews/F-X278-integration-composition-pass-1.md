# F-X278 integration composition, pass 1

**Reviewed**: staged F278 squash over canonical HEAD `91dfb0a7c6b48c054fc4a28df6bca7994f717fc9`, which already contains F276. The staged diff has27files,3711insertions and362deletions, including plan and review artifacts. F278 source commit is `512ef799d657e9e1d7efe1c37bed294eec5285b5`, with handoff tip `9ffe490409732f82280b21f8b9e9f3ae42df0b0e`. F272 carry is outside this clean staged union and remains separately preserved. This is a bounded integration audit, not another implementation peer or counter reset.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the inspected composition. Full integrated qualification and subsequent F272 carry reconciliation remain pending.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Source composition

Compared the staged tree against both canonical HEAD and the reviewed F278 source commit. No F276 PowerPoint product or test path has a staged change against HEAD. Against F278 source, the product differences are exactly the ten already-integrated F276 PowerPoint paths. No unexpected Word or shared product-source difference appears. The supplied `F278_CLEAN_INTEGRATION_BUILD_DIFFERENCES.json` names those same ten paths relative to the326 F278 build inputs. This distinction is preserved, rather than claiming that the isolated F278 extension authenticates the entire new union.

Read both approved plans and applicable local final reports through their current preparation updates. The F276 and F278 functional contracts remain intact. Plan preparation records do not mark final hash/native/installed gates passed.

Inspected shared `FontManager` selection, symbol wrappers, diagnostic publication and contribution scopes with their existing PowerPoint consumers. `crates/oxml-layout/src/font.rs:1992` retains the old symbol entry point with Word mode false. The unknown-symbol wrapper at `:2119` also passes false. The additional Wingdings2 mapping is behind the Word-mode condition. PowerPoint render and authoring still call these old methods, including `rpptx-render/src/text.rs:515` and `:912`.

PowerPoint diagnostics still consume `current_font_diagnostics`, including renderer `rpptx-render/src/lib.rs:367` and the authoring replay/capture path at `rpptx/src/lib.rs:27103`. They do not read the new Word substitution list. RF278-01's reviewed repair is present: the contribution scope snapshots the Word-list prefix, and both success and overflow rollback truncate it alongside the old channel. Final replay still records contribution facts. This preserves the shared transaction boundary without mixing Word facts into F276 report diagnostics or changing the old mapping policy.

The new Word list can exist internally during shared contribution recording, but this audit makes no claim that it changes PowerPoint output or that its allocation cost was benchmarked. The inspected PowerPoint output channels remain separate.

## HLD composition

Compared merged HLD08 and HLD10 against the reviewed F278 source and canonical HEAD. The differences against F278 are the expected F276 additions, in separate sections:

- HLD08 retains F276 whole-deck fit reporting, width-qualified transformed grown-table overflow, duplicate physical leaves, mutating refusal and read-only notes fallback.
- HLD10 retains live slide index and preserved accessor errors, validation Display parity, empty-note diagnostic suppression, fit exits/JSON reasons and assembled Markdown escaping.
- The Word TOC, indent, borders, row split, automatic colour, Word-only symbol mapping and bundled/system font-policy sections from F278 remain present.

No F276 section is lost through the automatic merge, and neither story's policy is substituted for the other's font or semantic export contract. HLD15 remains the F278 deterministic/default-font policy update. No raw XML or schema writer changes are introduced by this integration composition itself.

## Focused union evidence and limits

The supplied focused union run completed during this audit. Read its actual log and adjacent receipt at `/private/tmp/rdocx-cross-review/f278-integration-f276-interaction.log`. The receipt records child0, wrapper0 and no timeout for the canonical `cargo test -p rpptx --lib --locked --offline issue431 -- --test-threads=1` execution. The log records seven passing tests, including geometry, absent-pair width, on-page table growth, empty frames, repeated ids with mutation refusal, notes fallback and table storage/growth sharing. This is seven focused cases, not full PowerPoint or integrated acceptance.

No test, build, benchmark, network, native control, Git mutation, source fix, cleanup or delegation was executed by this reviewer. Correctness, contract, panics, OOXML, tests and structure produced no new finding within this bounded source/contract composition. Unrelated giant source files and full artifact matrices were not reaudited.

Final integrated suites, hash49/pixels7 and reviewed actual14-key deltas, native Word and pinned oracles, installed Python versions/typing/stubs, WASM, archives, wheels, hosted CI, release dry-run/size and approval remain due. Historical broader failures are not waived. The stopped F272 carry and its unresolved C1 fallback work require their own restoration and reconciliation after the integration commit. This report ends the audit and returns the canonical writer lease to Root.
