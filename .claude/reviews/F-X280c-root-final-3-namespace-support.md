# F-X280c Root final3 namespace aspect support

Independent bounded audit supporting Root final3, not another author-model peer or fourth GLM self-review. Reviewed SOURCE_FREEZE_3.json, source diff e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007against72b1f8cbd0fcbcf1b0fefb75c22a3147bc8a2c01. All12listed current source hashes match. This pass wrote only this TEMP review file. No source/format changes, builds or tests ran.

**Verdict: zero defects, zero smells, zero nitpicks in the namespace aspect.** The earlier explicit-empty prefix discrepancy is repaired. This supports Root's separate aggregate final review, not sprint verification, complete-feature or release acceptance.

## Audited mechanics

At crates/oxml-core/src/xml.rs:55-82the strict shared helper uses plain Reader. Private LexicalNamespaces at:88-159holds one declaration-only vendor resolver, active raw bindings and flat undo/frame vectors. enter at:95-119pushes the reusable declaration resolver at baseline0to1and pops it after source-ordered raw declaration capture. Empty declarations are copied explicitly, rather than lost through bindings_of. Each logical Startadds only a usizeframe mark. leave at:122-135restores all scoped values in reverse. Empty/Endleave occurs after lexical event validation at:75-76, preserving names on the closing event. No logical depth reaches the vendor u16counter and no new depth quota/API is introduced.

Vendor quick-xml0.41.0name.rs:652-731supplies the same reserved declaration validation, raw URI semantics, attribute iterator behavior and inherited256declarations-per-element cap. Capturing raw values in the own map preserves existing distinctions between escaped or normalized URI spellings. Default namespace applies only to element names. Unprefixed attributes remain unbound at corexml.rs:140-143, and implicit xml/xmlnsbindings remain standard. Existing QNamevalidation precedes resolution. An emptied named binding's private Unboundvariant differs internally from vendor Unknown(prefix), but the unchanged validate_bound_prefix receives the same explicit prefix and returns the same observable InvalidNamespacecategory/message. No external failure follows from that private representation.

The strict helper preserves declaration/QName/reference/character/PI/comment/expanded-attribute error classes. It retains current event-order precedence: raw reserved-binding errors arise during declaration entry, element QName/binding validation comes next, then checked per-attribute decoding/declaration and duplicate expanded-name checks. No MCpolicy, root-cardinality gate or unrelated attribute admission policy was added to the format-neutral helper.

At rpptx/src/lib.rs:31871-31943the legacy field-ID scan uses its own narrow raw prefix-to-Aboolmap, declaration-only vendor validation and flat scope restore. Foreign and empty declarations insert falseinstead of inheriting an ancestor Abinding. At:31907-31911the exact :fldfix now distinguishes Some(empty)from None:the malformed explicit-empty prefix is non-A, while no prefix consults the default namespace. This preserves prior permissive cache selection without imposing a new QNamegrammar gate. The new control at:32326-32343includes :fldid=skip, valid default/prefixed fields, default reset, foreign override, sibling restoration, unrelated vendor:unknownand unbound MC:Ignorable. Expected IDs A/B/C/D/Eand a100000level shadow case with only OUTERmatch the intended old namespace selection.

Field attributes are still checked/decoded only when a DrawingMLfldis selected. Only the raw unqualified idkey is uppercased and retained. Unrelated unknown element/attribute prefixes and MCtokens are not newly rejected. A malformed declaration attribute stops the raw unchecked declaration iteration just as the vendor does, while selected fldchecked-attribute errors still return footer_malformed. The scanner does not reuse PptXmlScope::enter's stronger MC/unbound-attribute rules. No partial successful vector is fabricated after a selected-field error.

## Selected consumers and residual scope

The previous actual phase3failure is correctly a preexisting legacy opening panic: from_read_package→Parsed::of→all.xmlfield cache→part_field_identities. That last scan now uses plain Reader. Stored lexical diagnostics at lib.rs:3376consume the shared strict helper. The new stored root checker at:3390uses plain Reader, and the finite structural scanner at:3515uses plain Readerplus its existing flat StoredFrame/PptXmlScope state. The intermediate leftover NsReadersin those two new passes were removed before final deep CLIgreen. No recursive namespace resolver or recursive scan-owner drop tree is introduced there.

Existing ordinary replacement source checks such as check_ppt_xml_input at:2489still contain their older subsequent NsReaderand other modeled XML parsers retain their ordinary boundaries. This bounded change does not claim every deep XMLoperation or all repository NsReadercalls are panic-safe. Those unrelated existing paths are not new stored validation regressions or a mandate for a global reader rewrite. Root's other aspects audit nested-group/role/metadata behavior separately.

## Actual retained verification evidence

All14named successful receipt logs were read and independently SHAchecked against their JSONreceipts, each actual child0. The scoped corefilter reports8passes, including the lexical class table and new100000level namespace/256-257cap control. The field inventoryfilter reports1pass with its exact new :fldcontrol. Native Issue435reports14passes. The deep CLIchild control at crates/rpptx-cli/tests/integration.rs:572-634checks100000opaque levels inside a real slide role extension island and inside a separate unused custom part, both clean. These are actual stored CLIcalls, not metadata-only fixture construction.

Fresh Python ROOT-PY-BUILD3succeeded. ROOT-PY-DEEP3identifies installed extension949fca1055f3e4c1325f8c17d4ecc6cb8812c31eb4288989c0f02c0329534798and prints deep_parsed_issues (). ROOT-PY-VALIDATE3reports1focused validatepass,1037deselected. The installed deep control therefore now crosses from_bytesandordinaryvalidate rather than assuming opening success from stored CLIbehavior. These are scoped installed observations, not the full binding suite or final release-installed artifact qualification.

Historical source labels remain honest. Core/field receipts record aggregate diff27bb9352..., affected native/CLIreceipts e46b8fc7..., and final clippy/policy/Pythonreceipts e9fa82fa.... They were not all re-executed at the final aggregate label. Current relevant source was audited at freeze3and all12file hashes match it. Exact receipt labels are retained below for Root's source applicability join. No new redundant superset rerun is required merely to erase those historical labels.

Failures are preserved separately: ROOT-CHECK1child101was the initial unqualified Readerimport compilation error. ROOT-CLI-I435child101recorded3passes/1failure when new stored root/structural readers still used NsReader. ROOT-PY-BUILDchild1lacked VIRTUAL_ENV, and ROOT-PY-BUILD2child1failed uv-cache access during installation. These failed observations were not rewritten into successes or treated as semantic REDwhere only setup failed. The later fresh green receipts remain distinct.

## Defects

None found in this bounded namespace aspect.

## Smells

None found.

## Nitpicks

None found.

## Pending gates and exclusions

Root aggregate final3review, the other source aspects, eventual canonical semantic integration, complete final union verification, schema/native/installed platform riders, hosted CIand sprint/release approvals remain separate. This support pass does not reset author review caps, supply a new GLMpass, require an unrelated global NsReadercampaign, claim old aggregate receipt source identities changed, or qualify final packaging/performance. Review ends with this file.

Authenticated retained successful receipt identity list:

```json
[
  {
    "name": "ROOT-CORE-XML",
    "actual_child_exit": 0,
    "source_diff_sha256": "27bb9352089e13880763747d358cdc456fd4392f238b17ab809b0e5694482aba",
    "log_sha256": "0c256bb440dd9d16cf662ff68ec1548c5d18f76e72b568f24837edc3f082c572"
  },
  {
    "name": "ROOT-FIELD-INVENTORY",
    "actual_child_exit": 0,
    "source_diff_sha256": "27bb9352089e13880763747d358cdc456fd4392f238b17ab809b0e5694482aba",
    "log_sha256": "4eb95f11c55516e58c0afef4e804a8726114ba21a4644c7cfd0f808e281f3b35"
  },
  {
    "name": "ROOT-NATIVE-I435-FINAL",
    "actual_child_exit": 0,
    "source_diff_sha256": "e46b8fc707c7235bcb561ac478811a39d351b9eac674989628f64f997f1442b7",
    "log_sha256": "6a4a0a8a9e1e75c9023d44ce146150647965af225338720a50eebf8deae501f1"
  },
  {
    "name": "ROOT-CLI-DEEP2",
    "actual_child_exit": 0,
    "source_diff_sha256": "e46b8fc707c7235bcb561ac478811a39d351b9eac674989628f64f997f1442b7",
    "log_sha256": "c2a806801fec22a94b056b54a97d55fb49a70414216ed8d367613a17e5e3f61f"
  },
  {
    "name": "ROOT-CLI-AFFECTED-FINAL",
    "actual_child_exit": 0,
    "source_diff_sha256": "e46b8fc707c7235bcb561ac478811a39d351b9eac674989628f64f997f1442b7",
    "log_sha256": "0bedf8b6363f7f905e8eb3fdcef87dcafe3e5ddaa2309a2df65d53903ed4ff8f"
  },
  {
    "name": "ROOT-LEGACY-FIELDS",
    "actual_child_exit": 0,
    "source_diff_sha256": "e46b8fc707c7235bcb561ac478811a39d351b9eac674989628f64f997f1442b7",
    "log_sha256": "30a6f820a0905ecaf8c7fe7407321cb3d5b153a00d6388159257716a129fd4e8"
  },
  {
    "name": "ROOT-CLIPPY",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "4429a4ef2db9e3e8e95d0366d9c8b11dc0558a7e3e7228afe516eeea357f2486"
  },
  {
    "name": "ROOT-FMT-FINAL",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
  },
  {
    "name": "ROOT-PROSE",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "ab61ec46feb7cd020575d0b1c0e6c34f5517883168c06743458cc0d30f51befe"
  },
  {
    "name": "ROOT-ADAPTERS",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "37ea120c97672de8f7d7c358b7636a0704ab6674c7cc80d1761c15dc83afdc15"
  },
  {
    "name": "ROOT-WHITESPACE",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
  },
  {
    "name": "ROOT-PY-BUILD3",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "3037f6dc428afdbfff4240a00cb5a436368bb13e2ba3b8f7c4821ec3a493b4a9"
  },
  {
    "name": "ROOT-PY-VALIDATE3",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "5a7bc03459edfaad203664c339d91c11d79ad30ae2701f4c84c7d3b46a0d34f7"
  },
  {
    "name": "ROOT-PY-DEEP3",
    "actual_child_exit": 0,
    "source_diff_sha256": "e9fa82fa5604cf295407361daf39cb860902ee5a499364bd25507e88fdbb0007",
    "log_sha256": "ea0908431992820f97c7e5351ab0958882af3d7d1f67ae440038d47ed4c97842"
  }
]
```
