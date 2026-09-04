# F-233, Advanced mail merge

**Status**: completed
**Sprint**: S67
**Size**: L
**Depends on**: F-166

## Problem

The existing merge API accepts only flat `BTreeMap<String, String>` records
through `Document::mail_merge` and `Document::mail_merge_sections`
(`crates/rdocx/src/field.rs:577`). `FieldEvaluationContext` likewise exposes
only flat string merge values (`crates/rdocx/src/field.rs:39`). There is no
region expansion, recursive record scope, rich field replacement, or formatting
callback.

The repository already has the necessary lower-level pieces, but they do not
yet form this feature. Structured template cloning handles nested JSON controls
(`crates/rdocx/src/template.rs:681`), image authoring creates package media
relationships (`crates/rdocx/src/document.rs:2803`), document append remaps
styles and numbering (`crates/rdocx/src/document.rs:4196`), and F-166 provides
atomic staging plus body identity remapping (`crates/rdocx/src/field.rs:597`).
The advanced merge must compose those boundaries without changing flat merge
behavior.

## Spec reference

- `docs/hld/03-architecture.md`, the flat mail-merge and structured-template
  ownership paragraphs.
- `docs/hld/04-opc-and-packaging.md`, the staged template and mail-merge package
  rules.
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability".
- `docs/hld/12-testing-strategy.md`, "Test taxonomy" and the flat mail-merge
  regression gate.
- `docs/hld/14-development-backlog.md`, "F-233, Advanced mail merge".

## Approach

Keep both flat F-166 methods unchanged. Add an additive native-only rich input
boundary in the existing `field.rs`, re-exported from `lib.rs`:

```rust
#[derive(Debug, Clone, PartialEq)]
pub struct MailMergeImage {
    pub data: Vec<u8>,
    pub filename: String,
    pub width: Length,
    pub height: Length,
}

#[derive(Debug, Clone, PartialEq)]
pub enum MailMergeValue {
    Text(String),
    Image(MailMergeImage),
    Fragment(Vec<u8>),
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct MailMergeRecord {
    pub values: BTreeMap<String, MailMergeValue>,
    pub regions: BTreeMap<String, Vec<MailMergeRecord>>,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct MailMergeData {
    pub records: Vec<MailMergeRecord>,
    pub sources: BTreeMap<String, Vec<MailMergeRecord>>,
}

pub struct MailMergeFormatContext<'a> {
    pub source_name: &'a str,
    pub region_path: &'a [String],
    pub field_name: &'a str,
    pub record_number: u32,
    pub output_sequence_number: u32,
}

pub struct MailMergeFormattedText {
    pub text: String,
    pub run_properties: Option<RunProperties>,
}

pub fn mail_merge_rich(
    &self,
    data: &MailMergeData,
    formatter: Option<&mut dyn FnMut(
        &MailMergeFormatContext<'_>,
        &mut MailMergeFormattedText,
    ) -> Result<()>>,
) -> Result<Vec<Document>>;

pub fn mail_merge_sections_rich(
    &self,
    data: &MailMergeData,
    formatter: Option<&mut dyn FnMut(
        &MailMergeFormatContext<'_>,
        &mut MailMergeFormattedText,
    ) -> Result<()>>,
) -> Result<Document>;
```

Recognize whole-block `MERGEFIELD TableStart:<name>` and
`MERGEFIELD TableEnd:<name>` markers in paragraphs and table rows. Pair nested
regions with one bounded stack. Resolve region records from the current lexical
record first, then from `MailMergeData::sources`. Preserve template order and
source record order.

Expand paragraphs and adjacent table rows while retaining properties, lists,
content controls, and ordered raw sidecars. Replace scalar fields through the
existing field grammar and invoke the formatter after scalar resolution. An
image replaces one field with one inline drawing and a candidate-local media
relationship. A fragment is a complete DOCX package and is valid only when the
field owns a whole block paragraph. Import its body without final section
properties plus its reachable styles, numbering, media, hyperlinks, and other
internal relationship closure. Remap every package, relationship, numbering,
bookmark, content-control, and drawing identity.

Remove all consumed region markers and merge-field shells. Missing scalar
values retain the F-166 empty policy. Missing region sources, crossed markers,
wrong value kinds, malformed or externally linked fragments, formatter errors,
or allocation failures reject the whole operation. Reopen every candidate
before publication. Section mode then reuses F-166 section assembly. Python,
WASM, and CLI remain unchanged. Add no trait, generic parameter, dependency,
feature flag, module, file, or test binary.

## Rejected alternatives

- Generalize the flat F-166 methods. That would disturb their established
  source and behavior contract.
- Reuse structured-template tag syntax. It would conflate two existing public
  automation contracts.
- Use `Document::append` directly for fragments. It does not import the complete
  relationship closure.
- Permit arbitrary inline fragments. A DOCX body may contain blocks that are not
  legal run content.
- Add a formatter trait. There is no second implementer today. The existing
  `FnMut` abstraction already represents caller callbacks.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| regression, gate | `nested_source_built_records_generate_ordered_rich_content_without_stale_fields` | Nested and named sources emit exact paragraph, list, table, image, fragment, and formatted-run order with no stale field. |
| unit | `nested_merge_regions_resolve_lexically_before_named_sources` | Nested scopes shadow named sources deterministically and sibling scopes do not leak. |
| regression | `rich_merge_imports_images_and_fragments_without_relationship_or_identity_collisions` | Repeated insertion retains valid media, styles, numbering, hyperlinks, bookmarks, controls, and drawings. |
| regression | `formatting_hooks_change_only_the_selected_merge_field_runs` | The callback sees stable context and changes only the selected result. |
| round-trip | `rich_merge_preserves_schema_order_and_unmodelled_xml_after_reopen` | Fixed-prefix schema order and unrelated raw XML survive save and reopen. |
| regression | `invalid_rich_merge_input_leaves_the_template_and_outputs_uncommitted` | Every validation, callback, relationship, parse, and allocation failure is atomic. |
| regression | `flat_mail_merge_output_remains_unchanged_after_rich_merge_is_added` | Existing F-166 separate and sectioned output remains stable. |

The **test gate**, from the backlog, is regression. Nested source-built records
generate the expected ordered paragraphs, lists, tables, images, and formatting
without stale fields.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/04-opc-and-packaging.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- **Unit conversion**. Reuse existing truncating `Length` to EMU conversion for
  rich images, assert exact drawing extents, and declare the harness result.
- **Any parser or serialiser**. Add save and reopen coverage for fixed prefixes,
  schema child order, and byte-identical preservation of unrelated raw
  subtrees.
- **Public API of a published crate**. The owned values and two methods are
  additive and native-only. Run rustdoc with warnings denied, the exact patched
  workspace publish dry-run, and every generated archive size assertion.

## Hash harness

Expected unchanged across all 49 entries. The API is opt-in and no sample
invokes rich mail merge. Any delta is unexplained and blocks integration.

## Implementation checklist

- [x] Add and export the owned rich merge data and formatting types.
- [x] Add the two additive rich merge methods.
- [x] Parse and validate bounded nested block regions.
- [x] Resolve lexical records and named sources deterministically.
- [x] Replace scalar values and invoke formatting callbacks in field order.
- [x] Insert images with exact dimensions and collision-free relationships.
- [x] Import fragment body content and its internal relationship closure.
- [x] Remap package and document identities.
- [x] Remove consumed region and field shells.
- [x] Stage, serialize, reopen, and publish atomically.
- [x] Add source-built tests to the existing unit and regression binaries.
- [x] Run focused checks and every risk rider.
- [x] Update exactly the four listed HLD files.

## Open questions

None. Whole-block region markers, lexical lookup, field-only rich replacement,
complete DOCX fragments, exact image dimensions, and the formatting callback
contract are approved.
