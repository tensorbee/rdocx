# F-X279b, Root final 2 C5 supplement

**Verdict**: 1 defect, 0 smells, 0 nitpicks. This subsequent contract check supersedes final2 preparation acceptance. No fourth peer, cap reset or new scope.

## Defects

### D4, Equal plain-run partitions inside held table and control owners refuse

`crates/rdocx/src/comparison.rs:10789`
`crates/rdocx/src/comparison.rs:10802`
`.claude/plans/F-X279b-design.md:238`
`.claude/plans/F-X279b-design.md:438`

The approved C5 contract says equal coalesced content inside a revision-bearing owner copies that owner byte for byte. The new eligibility comparison only coalesces BodyContent::Paragraph. A whole table or block-control owner depends on exact typed equality, so an otherwise equal physical plain-run split in a contained paragraph refuses when another paragraph of that same owner contains the existing revision. This is a supported-input overrefusal, distinct from the earlier D3 permission bypass.

After the clean final2 report, the author raised this precise contract discrepancy without editing source. Root executed the public calls against the actual fresh32 installed native SHA c3b13e62e03db0e885cc0b1fc7ee6efcc8a1f7de83cf96c5e1f47bbea8945134, verified before import use. An in-memory generated main part has one table cell containing a plain invoice paragraph and a sibling paragraph with old insertion9. The edited copy changes only the plain paragraph to two equal-property runs invoi and ce. Comparison refuses at body/table[0] at run, word and character granularity. Replacing the table with a block content control yields the same refusal at body/content-control[0] in all three modes. All six refusals preserve original bytes. No file or fixture was modified. This is actual runtime failure of C5, not a source-only counterfactual.

Preserve the strict typed whole-owner and positive source-residual requirements. Extend the same physical plain-run equality exception to contained ordinary paragraphs without allowing typed structural, formatting or text changes. Keep Field/control source facts and all nontext barriers. Add genuine predecessor regressions for these supported positives, alongside real nested changes that must refuse. No blanket comparison normalization, public API, XML validator or security assessment is requested. Original independent peers remain exhausted. Root affected final review follows bounded remediation directly.

## Smells

0.

## Nitpicks

0.

## Not found

No additional finding established in this finite check. D1/D2/D3 closures remain applicable. Local32 success does not discharge this newly exercised C5 combination, and all final union checks remain pending. Only this supplement was written. The audit ends and returns to remediation.
