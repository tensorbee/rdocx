# F-X272, all aspects, pass 9

**Reviewer**: independent Codex reviewer, not an implementer of this repair.
**Scope**: D5 physical run-parent eligibility repair, its two added existing-carrier regressions, and the current public Box/SDT successor packet. Read the current approved plan and pass 8. Root froze canonical source and closed acquisitions before authorizing this report.
**Verdict**: 0 product defects, 0 smells, 0 nitpicks in this bounded repair review. D5 is addressed for the examined invariant. The SDT public packet is still NONPASS because its graph assertion rejects an authorized deletion transformation. Its unexecuted lifecycle and compared-schema branches remain evidence debt. This is not full feature or native readiness.

## Source identity and D5 disposition

Current document.rs SHA256 is 22d551eff5f1c59e0535014ca15cec11011245ce4be3166ac47c01393e4f416c. Regression carrier SHA256 is bbf6fcf77d9c2544adf3ee5864bf32202bbf7fbc71d29f46a790cffabcb6e38e. Approved plan SHA256 is c9cea5ca6f6553c6da13469dd9deaa9ae949e5b5a511b4858102a835aca1154c. These were read from actual tracked inputs during review. The wider working diff includes unrelated F-X273 and prior F-X272 changes and is not attributed to this repair.

At crates/rdocx/src/document.rs:16715, private StoryFieldRun carries the actual enclosing run range and concrete immediate parent's source start from the existing frame stack. Both Start and Empty marker callers use it. At document.rs:16834, emission of a typed candidate requires the begin and end run parents to match. Intrinsic accepted status and parent validity propagation are computed before this eligibility gate. Template marker ranges retain their independent update path. The existing typed isolated probe remains responsible for actual same-context admission. No catch-all parse-error suppression, new public model, paragraph rescan or cache was introduced.

The source coordinate identifies physical parent identity, so sibling revisions with equal metadata do not become one context. Equal immediate parents establish a balanced sequence of sibling runs and intervening complete nodes for the existing detached probe. Same-parent direct and revision-owned cases remain eligible. Cross-parent complete fields remain source-preserved and untyped until their actual wrappers are resolved.

The added unit at document.rs:17759 covers direct fields, one revision parent, equal-metadata distinct revision parents, checked standalone snapshots and selected-author accept/reject discovery. The added public regression at crates/rdocx/tests/regression_test.rs:25383 covers actual Box and atomic SDT comparison, checked snapshots before resolution, both selected-author resolutions, counts, no remaining revisions, text, a comparison postcondition and field/marker counts. These address the observed failure without inventing exact predecessor runtime offsets, which remain unobserved.

## SDT successor assertion is a fixture mismatch

The retained packet is /private/tmp/rdocx-s96-run-refresh/fx272-d5-public-scope-successor-v1-root. Its run/probe.py:81 requires every compared SDT graph to equal the original graph, using the same opaque check as the VML pict. The saved same-atomic-sdt-compared.xml has two physical control copies. The selected deletion contains sdt/sdtContent/fldSimple/r/delText, while the insertion copy retains t. The rejection is specifically this t to delText mapping.

That mapping is the intended modeled Word deletion transformation at crates/rdocx/src/revision.rs:997 and revision.rs:1013, not mutation of unknown opaque content. The approved local inline path includes qualified SDT leaves even though the SDT remains atomic for publication framing. Atomic framing means the shell remains a single wrapped child. It does not make its modeled text opaque to deletion semantics. The foreign drawing/textbox barrier retains a separate preservation rule.

The retained primary transitional schema independently supports this shape. In ecma-transitional-schemas/wml.xsd, CT_RunTrackChange:995 includes EG_ContentRunContent, which admits CT_SdtRun at line 2086. CT_SdtRun:2151 uses CT_SdtContentRun. CT_SimpleField:1231 admits paragraph content, and CT_R uses EG_RunInnerContent, whose line 1748 explicitly admits delText. Static grammar inspection supports the examined SDT ancestry and leaf. It does not substitute for the complete saved SDT main-part schema execution that the assertion prevented.

Do not weaken unknown-payload preservation or normalize arbitrary XML to obtain a pass. A successor fixture should keep the VML opaque graph equality, require the SDT's precise selected deletion mapping, preserve its shell/instruction/other attributes, and assert exact role graphs after actual selected-author resolution. The current fixture failure is not recorded as a product defect or smell. It is an explicit uncompleted validation step, with its original NONPASS packet preserved.

## Actual evidence and limits

Independently checked all 39 public execution-seal member hashes with no mismatch. Receipt SHA256 is a18b172fa9b7625540ece60a6964af72e02254468ce0bc00db58c2a4e5864640 and the sealed outcome remains NONPASS. Read actual results and saved candidate XML, rather than treating intended qualification text as executed coverage.

| Public successor observation | Actual result |
|---|---|
| Box compare | Completed |
| Box selected-author accept/reject | Both completed, with preserved graph and no remaining revisions |
| Box complete main-part schema | Original, edited, compared, accepted and rejected all valid |
| SDT compare | Completed, previous detached-run parse failure absent |
| SDT input main-part schema | Original and edited valid |
| SDT compared schema and selected-author lifecycles | Not executed because graph assertion stopped first |
| Total actual main-part schema rows | 7 valid, not the intended 10 |

For fx272-d5-parent-context-v1, independently checked all 36 exact child stdout/stderr hashes, all 18 named starts, exits and PASS outcomes. The receipt reports source stability. Strict lint receipt reports exit 0 and source stability in 45.52 seconds. Root also supplies fresh installed Python v6 exact9 PASS and the existing D3 public successor 4 comparisons, 8 lifecycle outputs and 16 main-part XSD validations PASS. Those root-supplied observations remain finite and do not fill the missing SDT successor branches. No product execution, schema run or native operation was performed by this reviewer.

## Defects

Zero independent product defects established in this repair scope. D3 and D4 retain their examined repairs. D5 is addressed for the reviewed physical-context invariant, including actual public comparison completion in both formerly failing cases. No clean full-feature attestation follows the fixture NONPASS or the still-unexecuted SDT checks.

## Smells

Zero independent smells.

## Nitpicks

Zero.

## Handback

Correct the bounded diagnostic contract in a separate phase and acquire its missing SDT compared-schema and selected-author outcomes before claiming that packet GREEN. No additional confirmation microscope is requested solely to repeat this zero-finding review. A new source change or a newly established product failure would justify a focused review.

All current native Word, final union, CLI and remaining workspace, installed, hosted, hash and risk obligations retain their delivery-record status. This report does not complete F-X272, Issue 430, F-X273, the sprint or either release line. Only this review file was written. No source, tests, plans, HLD, progress, TEMP or other file was modified. No product build/test, network action, oracle, commit or push ran. Return control to root.
