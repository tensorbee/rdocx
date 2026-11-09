# F-X273, all, pass 7

**Reviewed**: Focused source-only working review of the newly established genuine-rich tab delta against HEAD `6506fa660372b2cfe7dd620cfaaf0959a62ba640`. The affected source is the MultilingualText handling in PendingTab::add and decimal_alignment_offset, plus actual shaping/breaker callers, the approved unchanged-default contract and existing parent test. This is a new concrete fault review, not an unchanged-scope confirmation of pass6.

**Verdict**: 1 defect,0 smells,0 nitpicks. No runtime RED, native measurement or PowerPoint output delta is claimed.

## Defects

### D1, legacy tab scanning changes genuine-rich default alignment

`crates/oxml-layout/src/line.rs:1112`
`crates/oxml-layout/src/line.rs:1175`

PendingTab::add now accumulates trailing MultilingualText whitespace and measures a mixed span's trimmed prefix without checking uses_legacy_shaping. decimal_alignment_offset likewise admits every MultilingualText base and measures its numeric prefix. At the reviewed HEAD baseline, genuine MTS fell through the non-Text fallback: trailing remained0and decimal scanning returned no offset while no earlier plain number was active. The new branches therefore change existing genuine-rich output, including generic default owners whose legacy flag is false and kerning override is None.

A source-established supported counterexample uses public FontManager::shape_multilingual_paragraph at `crates/oxml-layout/src/font.rs:2579` to shape `123.45 ` in LTR, then supplies those MTS spans after a Tab to public break_multilingual_into_lines at `crates/oxml-layout/src/line.rs:1224`. Use wrap=false, ample available width and an explicit right or decimal stop sufficiently far beyond the content so clamping cannot hide the difference. Let W be the actual nonspace shaped width, S the actual positive space advance and P the actual digit-prefix width. The old right-tab gap subtracts W+S, while the new branch subtracts W. The old decimal path reaches the complete following width without finding a decimal point, while the new path subtracts P. No fixed runtime widths are invented here.

Nonlegacy shaping segmentation does not eliminate this case. At `crates/oxml-layout/src/font.rs:2782`, trailing spaces can become a separate span. That span still reaches the new all-whitespace branch, so the right-tab difference remains S. A decimal point inside the preceding numeric span still reaches the new prefix scan. LineState::push and width_with call PendingTab::add at `crates/oxml-layout/src/line.rs:843` and line869. PendingTab is created for right/center/decimal stops at line969. Later LTR bidi movement does not recompute the original tab gap back to the prior result.

The shared default route is publicly callable. Word also has a genuine-rich producer and explicit Word tab parameters through `crates/rdocx-layout/src/engine.rs:12869` and line16761. A rich-script sibling can select that route while the tab's numeric run remains genuine MTS. K1 authorizes context-qualified Word tabs, but `.claude/plans/F-X273-design.md:325` preserves generic/PPT defaults and line818preserves existing entrypoint outputs. It does not authorize broadening all genuine-rich tab semantics as part of legacy context transport.

The parent test at `crates/oxml-layout/src/line.rs:3834` adds right/center/decimal MTS controls only with shape_paragraph_with_kerning legacy=true. It does not discriminate genuine default None or explicit-context legacy=false owners against the old fallback. Its PASS does not establish unchangedness for this default path. Earlier pass6's finite local gates remain actual historical results, but that specific unchangedness was not established by its review.

The required correction is to preserve the approved genuine-rich fallback and apply new MTS trailing/prefix scanning only to known legacy owners. Add a finite discriminator in the existing test body for both default None and explicit-context nonlegacy owners, beside the legacy scanning positives. This is a remediation requirement, not a code change performed in this pass. Actual old-current RED and repaired GREEN remain to be acquired by the separate implementing phase.

## Smells

0 findings.

## Nitpicks

0 findings.

## Examined exclusion and evidence limits

A current normal PowerPoint right/decimal-tab delta is not established. `crates/rpptx-render/src/text.rs:943` uses genuine default/caps shaping and line1062calls the shared rich breaker, but its line-break parameter construction at line1334supplies an empty explicit tab-stop list. The shared default stop at `crates/oxml-layout/src/line.rs:957` is left-aligned, which does not instantiate PendingTab. This excludes this changed PendingTab branch from that current normal PPT route. It does not prove the whole font/line increment neutral for all public consumers.

Source identities independently checked: line.rs SHA256 `457e39eeac901cf3cabcfe17147689e18f93aed59bbfd546b50a075c909f15a4`, font.rs `e421da3a39b4d004127632ff643685b4ba1721891f4e905935f78f979b3d18ef`, approved plan `f39cf22966952420d61a45b8d88e26179442b3b2dc2ffb418b4ba55ef77f130c`, and PPT text.rs `19557f9eb0c421675a8595dd96532cbe51d000961953283a082b58b53aa9f9f6`. Compared actual baseline source through git show. No builds, tests, imports, TEMP artifacts, source edits or runtime experiments occurred.

## Not found and return boundary

No additional correctness, panics, OOXML or structure finding was established outside D1's concrete source/contract/test gap. Earlier reviewed ownership, annotation, cache and R4 findings are not reopened by this narrow pass. All native/full/hash/raster/installed/WASM/hosted and sprint acceptance debt remains pending. No intentional-delta or neutrality label, baseline update or completion is approved while D1 remains.

End review and return the sole canonical writer lease to root for separate remediation.
