# S96, Root note and comparison repairs, support pass 1

**Reviewed**: Bounded working-source review against HEAD `a5765aa00709d88604093bfe9d46ba62d5af397a`. The examined changes are qualified note-ID decoding and its lower regression, the Run-granularity signature and unsplit Text repairs, the normal-note dependency fixture, and the x188 physical thread-marker fixtures. This is not another complete F272 or F279 peer pass.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Source identity

The sampled full-file SHA-256 values are:

| File | SHA-256 |
| --- | --- |
| `crates/rdocx-oxml/src/footnotes.rs` | `1d6edff7d58f19f6a2378a2d63163d36a6fd66f80423b1996077f89959931818` |
| `crates/rdocx/src/comparison.rs` | `a935a473da9a433e5bd44c7da5342c285d40a97efff9b4f419ae017a281d09d5` |
| `crates/rdocx/tests/regression_test.rs` | `c4baab3cbdeac810a270ee95a4a06f97fea793f8e3e2770ed81d3484b14d6dc5` |

These files also contain other working changes. The full working diffs contain 24 additions and 2 deletions in footnotes, 17 additions and 12 deletions in comparison, and 518 additions and 193 deletions in the regression module. Those totals are context identities, not a claim to have reviewed every unrelated hunk.

## Findings

### Defects

0.

### Smells

0.

### Nitpicks

0.

## Examined mechanisms

- `crates/rdocx-oxml/src/footnotes.rs:130` resolves the qualified ID attribute through the existing Word-attribute predicate and decodes its actual XML value before integer parsing. This fixes escaped minus and digit references without redefining missing or nonnumeric fallback. The neighboring type attribute already uses the same decoder. The lower regression at line 425 distinguishes negative, zero and positive decoded IDs. The intended normal-note law is retained in `.claude/plans/F-X277b-design.md:277`, which explicitly distinguishes the historical nonpositive-ID fallback from the approved F272 contract.
- `crates/rdocx/src/comparison.rs:13879` selects the whole-run signature only for Run granularity without ignored whitespace. `granular_text` at line 16316 now retains the actual unsplit `CT_Text`, including its original space flag. The signature and emitted unit therefore agree. Word, Character and whitespace-split paths retain attributed edge-space normalization. The existing bidirectional matrix at `crates/rdocx/tests/regression_test.rs:41292` still distinguishes genuine edge-space changes from non-edge producer noise, tests bookmarks, and checks that split whitespace remains preserved. Its assertions were not relaxed. F279's existing larger-run behavior remains required by `.claude/plans/F-X279b-design.md:256`.
- `crates/rdocx/tests/regression_test.rs:49252` retains the actual normal note ID returned by the creator. The binding is inserted into that qualified owner's body using `f276_word_owner_body_range`, rather than into the first physical separator. Existing dependency, remap and closure assertions remain in place.
- `crates/rdocx/tests/regression_test.rs:64800` derives removal expectations from all actual comment definitions. Each definition contributes its start, end and reference marker. The test still compares complete retained story bytes and validates reopened ownership. The shared-note fixture at line 65549 uses the same complete marker set, retaining its annotated-versus-unannotated distinction. The qualified x188 owner helper at line 64209 avoids assuming a paragraph start tag has no attributes.

Correctness, contract, panic surfaces introduced by these changes, OOXML namespace and whitespace handling, tests, and structural constraints produced no findings in this bounded scope.

## Evidence boundary

The retained MSRV70 failure inventory records actual exit 101, including the original escaped-note and edge-space failures. Root reports the new lower note-ID unit as green and retains a first comparison-fix RED. This review inspected source and fixture assertions without running Cargo or independently executing those outcomes. The repaired facade escaped-note case and combined regression batch remain pending at this report boundary.

Other working comparison changes, including the text-box scan policy, the separate comment publisher repair, unrelated fixture migrations, native Word qualification, performance observations and final union gates are excluded. This report does not certify complete provider or sprint acceptance.
