# F-X280c, GLM self-review, all aspects, pass 1

**Reviewer**: GLM-5.3-Flash-EXL3 (dgx/GLM-5.3-Flash-EXL3#high), implementing
author self-review. This is explicitly not the independent review. Root
supplies that separately.

**Reviewed**: uncommitted working tree versus Base 72b1f8cb, 10 files, 1946
insertions and 12 deletions (git diff --stat). Product focus
crates/rpptx/src/lib.rs stored-validation section, CLI route, both carriers,
READMEs, HLD04/HLD10 stored-validation sections.

**Verdict**: 2 defects, 1 smell, 2 nitpicks

## Defects

### D1, untrusted gridSpan sum can overflow and panic in debug builds
`crates/rpptx/src/lib.rs:3722`

`coverage += span` sums attacker-controlled `gridSpan` attribute values
parsed as `u64`. A part with two cells carrying `gridSpan="9223372036854775808"`
or larger values overflows the sum, which panics under debug assertions.
`validate_presentation_file` is a public read-only validator, so untrusted
input reaching a panic violates the non-panicking validator expectation.
The sum should saturate (`saturating_add`), which cannot change any
legitimate verdict because the comparison target is the grid column count.

### D2, the tree-root spTree's own nvGrpSpPr cNvPr id and name are never checked
`crates/rpptx/src/lib.rs:3592`

The approved scope table requires unqualified `id` and `name` presence for
`cNvPr` under physical owners including "tree-root P:spTree/P:nvGrpSpPr
paths reached from the role tree". The stored walk starts at the spTree's
children, so the spTree's own `p:nvGrpSpPr` falls into the
membership-allowed branch and its `cNvPr` is never inspected. A deck whose
spTree root `cNvPr` lacks `name` passes the stored route. The plan calls
this exact path a stored-only additional context, so the check belongs in
this story.

## Smells

### S1, tautological byte comparison in the CLI fault loop
`crates/rpptx-cli/tests/integration.rs:4518`

`after` and `before` are both read after the CLI ran, so
`assert_eq!(after, before)` compares the file with itself and proves
nothing about preservation. The clean-deck half of the same test and the
native tests do hold real before/after comparisons, but this assertion
reads as preservation evidence without being any. Compare against bytes
captured before the CLI invocation, or drop it in favor of the native
preservation assertions.

## Nitpicks

- `crates/rpptx/src/lib.rs:3625`, a DOCTYPE is not reported because the
  shared `validate_strict_xml_1_0` frontier does not classify it as a
  lexical fault. This follows the declared shared frontier, but a future
  story may want the explicit refusal the mutation checker has.
- `crates/rpptx/src/lib.rs:3580`, the text-outside-root fault is checked
  before the root-count fault, so a part with both problems reports the
  text one. Arbitrary but deterministic.

## Not found

- correctness: the remaining stored logic checked against the pinned SDK
  3.0.1 particle, the six original faults, role selection by qualified root,
  MC Choice/Fallback member inheritance, opaque extLst/contentPart/ignorable
  islands, alias and default namespaces, model-open failure precedence,
  InvalidXmlCharacter mapping and dedup by existing Eq. No further defects.
- contract: the public surface is exactly the one additive file function.
  no new module, file, crate, trait, generic, dependency, feature flag or
  OPC policy change. The five-kind helper extraction preserves the frontier
  caller's conditions. HLD edits are confined to the approved
  stored-validation sections. Python stays parsed-model only. The gridSpan
  deviation is recorded in the progress file and is deliberate, not silent.
- panics: no other new unwrap, expect, slicing or indexing on untrusted
  input. `depth` arithmetic is guarded. Archive and entry reads are
  ok-guarded. Only D1.
- ooxml: namespace resolution reuses PptXmlScope and PptXmlNames unchanged.
  attribute checks require unqualified names. Original bytes are never
  reserialized for scanning and the input file hash is asserted unchanged.
- structure: no wrapper that only forwards, no second namespace resolver,
  concrete types only.
- tests: the A-ID rows map to actual assertions, the two named controls
  exist, RED at Base is recorded through the real CLI producer, and the
  pinned corpus validates clean through the new route. Only S1.
