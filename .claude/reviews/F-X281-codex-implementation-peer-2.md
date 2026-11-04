# F-X281 implementation peer 2

**Reviewed**: Working continuation at Base6bea9093818d6b691bf88b2980e9cc17927be015 in /private/tmp/rdocx-s96-f281-chart-continuation. Seven paths,129insertions37deletions. Frozen diff SHA25649596e59222a2a71645e4e89a33c0edeefde4cfd5a96e7a5a65b73aef332d9cd. All seven path hashes match SOURCE_FREEZE_1.json. This is actual independent implementation peer2. Earlier peer1 remains used, with a total of2of3implementation peers used. No design peer or review cap reset.

**Verdict**: 1 defect, 0 smells, 0 nitpicks. Source continuation checks below are otherwise coherent, subject to the unchanged final acceptance riders.

## Defects

### D1, the current chart HLD still requires the rejected MIME match and obsolete decode cap

`docs/hld/09-charts-spec.md:736`, `docs/hld/09-charts-spec.md:737`, `docs/hld/09-charts-spec.md:744`, `.claude/plans/F-X281-design.md:368`

The canonical chart spec says a cached fallback is usable only when its bytes match the declared PNG/JPEG content type and lists content-type mismatches as placeholder cases. The operative continuation explicitly corrects that behavior, and `crates/rpptx/src/lib.rs:29153-29159` now derives the transient render type from admitted byte format so a MIME mismatch remains drawable. New tests at `crates/rpptx/tests/integration.rs:14583-14608` and :14649-14672 assert those actual positioned images, including metafile declarations. Updated HLD08:3010-3012 describes the new sniff-based behavior. HLD09 therefore contradicts the actual approved outcome and remains outside the plan's HLD impact list.

The same paragraph incorrectly gives both encoded and decoded data16MiB caps. The actual existing constants at `crates/rpptx/src/lib.rs:29435-29437` are16MiBencoded and64MiBdecoded, also correctly stated in HLD08:3019. A valid preview above16MiBdecoded and within64MiB is described as refused by one spec while allowed by source and the other spec.

Add the existing HLD09 paragraph to the plan's impact and reconcile its current intent with sniff-based PNG/JPEG eligibility and the existing16/64caps. Preserve the PNG/JPEG-only cached-chart owner policy, the odd-channel/non8bit JPEG and invisible-preview residuals, scoped relationship ownership and raw package preservation. No production quota change or general cached EMF admission follows from this documentation repair.

## Smells

None found.

## Nitpicks

None.

## Correctness and consumer trace

The exact-empty authoring guard at `crates/oxml-chart/src/lib.rs:1311-1316` is shared by authored_chart_parts:246 and replace_chart_data:920. It uses the existing authoring error before building or publishing new candidates and leaves the existing category/series/finite/scatter failure ordering intact. No whitespace trimming or broad renderer-supported-only restriction was introduced.

Imported NumericData::from_xml_with_namespaces at5211-5240 still reaches the unchanged XML-character-only format validator through data.validate. Serialization uses the same unchanged validation. The empty stored c:numFmt/cache-format cases therefore reach format_value Err plus General rendering fallback rather than being rejected on import. Nonempty unsupported formats remain authorable and drawable through fallback. Exact General and None remain positive controls.

Word's add_chart_package at `crates/rdocx/src/document.rs:27158` and presentation add_chart at `crates/rpptx/src/lib.rs:13928` use that same authored-chart provider. Presentation data replacement at13745-13768 obtains candidates before publish_chart_parts. The low-level replacement regression exercises shared admission. No alternate authoring caller was found bypassing the guard in the inspected consumer set.

The transient media map correction at `crates/rpptx/src/lib.rs:29150-29161` changes only derived ScopedMediaIds metadata for already admitted PNG/JPEG bytes. Original package declarations and MediaData bytes/type remain unchanged. This removes the second cached-chart MIME veto without rerunning a decoder, changing decoder limits or allowing skipped/failed resources into scoped media_ids. Both normal slide assembly and export-surface assembly call render_scoped_resources. Existing slide/layout/master relationship scoping remains intact. The field/method comments in rpptx-layout now accurately distinguish render-effective type.

The general shared EMF classifier admits drawable bounded EMF into native image resources. However HLD09:733-740 supplies the narrower existing unsupported-chart cached-preview policy: PNG/JPEG only. The unchanged context guard at `crates/rpptx-layout/src/context.rs:1345-1348` and :1647-1651 retains that owner restriction. A genuinely drawable EMF cached chart preview therefore remains a placeholder under the current explicit owner contract. This is not inferred as a new waiver, new cached decoder requirement or an unreported implementation finding. Ordinary EMF picture rendering and its final acceptance rows remain due.

No new public API, quota, trait, generic, module or source file was introduced by this continuation. Error/atomicity behavior is restored at checked authoring admission, not imposed on retained imported XML.

## Evidence inspected

No tests or builds were run by this review. Actual recorded receipts and their complete named logs were read, and their SHA256values were recomputed successfully:

- CHART_BASELINE and diagnostic output identify the real Some(empty) authoring failure and stale mismatched-PNG placeholder assertion. Original cohort10pass/2fail/1ignored, actual child101. The failed source/artifact history remains retained.
- The separate cached-metafile-declaration RED is actual child101, before the transient-map fix. It is distinct from the two original failures.
- CHART-FINAL-COHORT has actual child0 and12pass/0fail/1ignored. Its log hash is ae1b0fb0084796090ad314d13624abcd4b82cd9b397fb321afd6d93f1e8d5c4d. The actual image-byte assertions distinguish corrected previews from a placeholder, rather than merely observing no error. Existing malformed/corrupt/oversized/grayscale/CMYK negative cases remain in the placeholder group.
- SHARED-GREEN has actual child0 and7pass. Its new authoring/replacement negative is paired with None/General/nonempty unsupported positive formats, and the imported-empty serialization/projection/render test remains intact.
- The original LAYOUT-FALLBACK-CONTROL selected0tests and is not accepted as proof. LAYOUT-FALLBACK-CONTROL2 has actual child0 and1named test passed. Its log hash72fe8648f5f6dd6d836d7c661fc91bf84d60a9fad1d047f3341763793a03425d matches, and its assertions preserve cached-image opacity/diagnostic plus direct unsupported-WMF-map fallback.
- Actual CONTINUATION-CLIPPY, FMT, PROSE and WHITESPACE receipts now each report child0. Their retained log hashes match. Clippy targets oxml-chart, rpptx and rpptx-layout, all targets/all features with warnings denied. This is scoped continuation evidence, not a full product suite.

The caller-source and existing negative/positive paths were inspected. These receipts establish the stated bounded checks only. Installed binding, native Office, SDK/schema, oracle pixels, full hash/workspace, archives, public API qualification, hosted CI and final integrated acceptance are still pending. No source-bound receipt from another product state was promoted into those final gates.

## Not found

No further correctness, panic, raw-source mutation, namespace/XML preservation, public API, structure or test-discriminator defect was found in the inspected continuation. No source code or incidental formatting was changed by this review.

Review ends with this report. Return to Root for separate D1 remediation, preserving the existing peer count and final gate debt.
