# F-237, all, pass 3

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the two earlier review artifacts, 17 files and 3,433 changed lines, comprising 3,427 additions and 6 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 8 defects, 0 smells, 0 nitpicks

## Defects

### D1, changing a building-block body drops its root metadata and inherited namespaces
`crates/rdocx-oxml/src/glossary.rs:1054`

Every changed body is wrapped in a newly constructed `w:docPartBody` start tag
that retains only `xmlns:w`. For example, changing a body whose original root
declares `x:producer="keep"` deletes that attribute. If a retained raw child
uses a foreign prefix declared only on the original body root, the emitted
child also becomes namespace-unbound. This violates the contract to preserve
unmodelled attributes, children, and namespace context during selected-entry
replacement.

### D2, a legal greater-than character in a property attribute corrupts the rewritten XML
`crates/rdocx-oxml/src/glossary.rs:662`

The surgical property writer finds the end of the original start tag by taking
the first `>` byte without accounting for quoted attribute values. XML permits
a literal `>` inside an attribute value. Changing a property such as
`<w:description x:producer="a>b" w:val="old"/>` therefore appends part of the
old start tag after the rewritten one, producing malformed or duplicated XML
instead of preserving the producer attribute and changing only `w:val`.

### D3, nested glossary descendants acquire typed schema meaning
`crates/rdocx-oxml/src/glossary.rs:965`

Container parsing accepts every matching Word element below the container,
regardless of depth. An ignorable producer wrapper inside `w:types` that owns a
retained `w:type w:val="autoExp"` descendant is consequently reported as a
direct type and classifies the entry as AutoText. Expanded name and direct
schema position must both decide typed meaning. The wrapped subtree should
remain unmodelled rather than changing the public building-block projection.

### D4, a property-local prefix shadow prevents adding a repeated value
`crates/rdocx-oxml/src/glossary.rs:755`

New children of an existing container use one output prefix selected at the
`w:docPartPr` root, without checking the container's local namespace scope. A
valid `r:types` container can bind `r` to WordprocessingML while shadowing the
root's Word alias `q` with a producer namespace. Changing that entry from a
non-AutoText type to AutoText emits the added `q:type` under the foreign local
binding. Reopen does not observe `autoExp`, so the supported replacement fails.

### D5, inline form mutation discards run-owned preservation context
`crates/rdocx-oxml/src/content_control.rs:483`
`crates/rdocx-oxml/src/text.rs:1056`

The inline-control adapter serializes every direct run through `CT_R`, whose
writer constructs a new `w:r` root with no retained producer attributes or
run-local namespace declarations. If a form run declares `xmlns:x` for an
unknown `x:` child inside its retained `w:ffData`, mutating the form deletes the
declaration while retaining the child. The committed document then contains
namespace-invalid preserved XML instead of changing only the typed form value
and cached display.

### D6, inline forms in package-owned stories cannot be persisted
`crates/rdocx-oxml/src/content_control.rs:448`
`crates/rdocx/src/field.rs:7786`

After an inline-control form is changed, every control-owned run is replaced by
`SdtContent::RawXml`. Header, footer, footnote, and endnote mutation then asks
`patch_story_field_sources` to discover changed fields through
`paragraph.runs()`, which no longer returns those raw replacements. The
original package story bytes are left unchanged, and staged validation rejects
the valid inventoried identity because its requested value did not persist.

### D7, adding an attribute to a valueless default-namespace form element can use an unbound prefix
`crates/rdocx-oxml/src/text.rs:3721`

The pass-2 fix declares a local prefix when inserting a wholly missing value
element, but the existing-element path still substitutes `w` for an unprefixed
element without declaring it. A header using only the default Word namespace
can validly contain `<checked/>`, whose omitted on-off value means true.
Setting it to false emits `<checked w:val="0"/>` into the original header with
no `xmlns:w`. Reopen still sees the old semantic value, so mutation fails.

### D8, nested legacy forms inside inline controls are omitted or assigned unusable ordinals
`crates/rdocx-oxml/src/content_control.rs:385`
`crates/rdocx/src/field.rs:304`
`crates/rdocx/src/field.rs:478`

The inline helper returns only top-level fields whose own `legacy_form` is
present. A form nested in a non-form field instruction is therefore never
returned, even though the ordinary paragraph path recursively inventories it.
If the outer field is also a form, facade inventory recursively counts both
forms, but mutation computes `inline_count` from only the single top-level
field and cannot address the nested ordinal. Supported field nesting therefore
either disappears from inventory or produces an identity that mutation rejects.

## Smells

None.

## Nitpicks

None.

## Not found

All thirteen exact pass-1 and pass-2 reproduction cases have concrete code and
named regression closure. No additional findings were found in panic safety,
dependency direction, public API structure, schema order for the directly
modelled children, staged commit atomicity, HLD file scope, or repository
structure. All 386 `rdocx-oxml` unit tests and its doc test pass, all 12 focused
F-237 integration tests pass, and `git diff --check` passes.
