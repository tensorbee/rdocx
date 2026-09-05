# F-237, all, pass 8

**Reviewed**: Full uncommitted implementation and HLD working tree against `4ba8b6b`, excluding the seven earlier review artifacts, 17 files and 5,652 changed lines, comprising 5,577 additions and 75 deletions, including untracked `crates/rdocx-oxml/src/glossary.rs` and `crates/rdocx/src/building_block.rs`
**Verdict**: 6 defects, 0 smells, 0 nitpicks

## Defects

### D1, out-of-order glossary properties remain selectable and can survive replacement
`crates/rdocx-oxml/src/glossary.rs:582`
`crates/rdocx-oxml/src/glossary.rs:586`

The properties parser recognizes a modeled child that appears after a later
schema slot, but downgrades it to retained extra XML instead of rejecting the
invalid sequence. For example, `w:name`, `w:category`, `w:style`, and
`w:description` exposes the name, category, and description while retaining
the misplaced style. Changing the description replays that style after the
category, reopens to the same typed projection, and commits the still-invalid
`w:docPartPr`. The schema-order gate therefore covers insertion but not an
existing modeled property sequence that selected replacement preserves.

### D2, valueless modeled glossary properties are treated as absent
`crates/rdocx-oxml/src/glossary.rs:1157`
`crates/rdocx-oxml/src/glossary.rs:1188`

The scalar and container value readers convert a modeled child without its
required `w:val` attribute into `None` or omit it from the typed vector. An
entry containing `<w:description/>` or a valueless direct `w:type` can
therefore be inventoried. Replacing another supported property retains the
invalid element as an unchanged slot, and staged reopen sees the same absent
typed value, so the facade equality check accepts and commits it. This is the
same ambiguity that duplicate modeled properties now reject, and it still
allows malformed bounded metadata to survive a successful typed mutation.

### D3, legacy form child order is not validated
`crates/rdocx-oxml/src/text.rs:1696`
`crates/rdocx-oxml/src/text.rs:1722`
`crates/rdocx-oxml/src/text.rs:1739`

The form projector tracks singleton cardinality but has no schema-slot state
for either `w:ffData` or its kind-specific container. It accepts common fields
after the final form-kind choice, such as `w:checkBox` followed by `w:name`,
and accepts kind children in the wrong sequence, such as `w:checked` before
`w:default`. Value mutation patches the existing element in place, staged
reopen accepts the same ordering, and the invalid sequence commits. This
violates the repository child-order rule and the story's schema-order test
contract.

### D4, malformed form kinds and selections disappear instead of failing closed
`crates/rdocx-oxml/src/text.rs:1641`
`crates/rdocx-oxml/src/text.rs:1650`
`crates/rdocx-oxml/src/text.rs:1670`

A present `w:ffData` with no required kind choice, a kind that contradicts the
field instruction, or a drop-down result outside its entry list all return
`Ok(None)`. Only an error sets `legacy_form_parse_error`, so these malformed
owners are silently omitted while later forms acquire lower ordinals. The
planned malformed-owner and bounds policy requires inventory and mutation to
fail closed, not reinterpret a malformed form as an ordinary untyped field.

### D5, normal note ownership is incorrectly restricted by the numeric id
`crates/rdocx/src/field.rs:8337`
`crates/rdocx-oxml/src/footnotes.rs:13`

The package-story scanner requires a note id greater than zero even when the
namespaced `w:type` says the note is normal. The existing note model explicitly
documents that separator identity comes from `w:type` and that conventional
numeric ids are not a guarantee. A normal footnote or endnote with id zero or
a negative id is therefore valid in the typed story model but its forms are
omitted from the new inventory and cannot be mutated.

### D6, multiple footnote or endnote relationships are accepted as separate stories
`crates/rdocx/src/field.rs:151`
`crates/rdocx/src/field.rs:157`
`crates/rdocx/src/field.rs:190`

Story discovery iterates every matching relationship and deduplicates only an
identical resolved part. It does not enforce the singleton relationship
constraint for the main document's footnotes and endnotes roles. A malformed
package with two different footnotes targets therefore exposes forms from both
parts, and mutation of either identity survives staged reopen because the same
collector again accepts both. Relationship-role validation must reject this
ambiguous graph before publishing part-scoped identities.

## Smells

None.

## Nitpicks

None.

## Not found

All forty-one exact findings from passes 1 through 7 have concrete closure. In
particular, character references outside glossary and package-story roots now
reject, direct entry child order rejects, duplicate modeled singleton glossary
properties reject, legacy form singleton and lexical token validation is
strict, note attributes resolve by namespace, and conflicting relationship
roles reject before part deduplication. No additional findings were found in
selected-entry raw-subtree preservation, inline and nested source-order
identity, cached-display mutation, staged failure atomicity, panic safety,
dependency direction, public API structure, HLD file scope, or repository
structure. All 22 focused glossary tests, the 21-case malformed-form matrix,
and all 25 focused F-237 integration tests pass. `git diff --check` passes for
the implementation and HLD diff.
