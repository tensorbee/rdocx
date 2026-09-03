# F-231, Extended field evaluation

**Status**: approved
**Sprint**: S66
**Size**: L
**Depends on**: F-161, F-162

## Problem

The native field evaluator currently dispatches only the field families from
F-161 and treats every other normalized instruction as unsupported in
`crates/rdocx/src/field.rs:1499`. That leaves formula, TOC, TC, mail-merge
control, and barcode instructions at the generic stored-display fallback even
though the recursive field grammar, ordered story traversal, and atomic cache
update path already exist.

The current public result boundary has only resolved text, pagination
deferral, and stored-display fallback in `crates/rdocx/src/field.rs:54`. Some
new families are control or generated-content decisions rather than plain
text, so the design must define their result representation before extending
the evaluator.

## Spec reference

- `docs/hld/03-architecture.md`, the field grammar, pure field evaluation, and
  explicit cache update paragraphs beginning with "The Word text model also
  projects bookmark starts".
- `docs/hld/04-opc-and-packaging.md`, the package mutation paragraphs beginning
  with "Mail merge uses the same fail-closed package boundary".
- `docs/hld/08-rendering-spec.md`, "Word bookmark field pagination".
- `docs/hld/10-bindings-spec.md`, "Native Word facade stability", specifically
  the `Document::evaluate_fields` and explicit update method paragraphs.
- `docs/hld/12-testing-strategy.md`, the Word field regression matrix paragraph
  beginning "The Word field regression matrix records Microsoft Word 16.104".
- `docs/hld/14-development-backlog.md`, "F-231, Extended field evaluation".

## Approach

Extend the existing evaluator in `crates/rdocx/src/field.rs` without adding a
module, trait, generic parameter, or dependency. Keep the recursive traversal,
document-order indexing, story-isolated state, formatting switch application,
and staged cache update path as the single execution boundary.

Add instruction validation and evaluation for `TOC`, `TC`, formula,
`NEXT`, `NEXTIF`, `SKIPIF`, `MERGEREC`, `MERGESEQ`, `DISPLAYBARCODE`, and
`MERGEBARCODE`. Reuse the current recursive argument resolver for nested
operands. Use bounded local parsers for formula and barcode syntax. No runtime
oracle, filesystem access, ambient clock, or renderer dependency enters the
evaluator.

Add structured native outcomes for TOC or TC decisions, mail-merge controls,
and barcode specifications. Formula fields continue to return resolved text.
Cache materialization will replace a cache only when the result is an exact
supported textual value. Deferred, control, generated-content,
unavailable-context, and unsupported cases will retain the original
instruction and cached display with a stable specific diagnostic and will set
only the field-local dirty flag when the existing update policy calls for a
Word retry.

Keep Python, WASM, and CLI surfaces unchanged. Keep the oracle version and
locale in test metadata only.

## Rejected alternatives

- Add a second evaluator for the new families. This would split field order,
  nested evaluation, and fallback policy across two code paths.
- Invoke Word or another office application at runtime. Field evaluation must
  remain pure and deterministic.
- Flatten every result to cached text. Control and generated-content fields do
  not have a truthful text result at the evaluator boundary.
- Add a new source module. The existing `field.rs` is the single owner and the
  repository requires an explicit ask before a new module or file.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| differential | `extended_field_families_match_the_pinned_word_result` | The approved source-built field matrix matches Microsoft Word 16.104 build 16.104.25121423 under the pinned locale and UTC context, including exact ordered outcomes and diagnostics. |
| unit | `formula_fields_use_bounded_precedence_and_stable_failures` | Formula precedence, nesting, pictures, bounds, malformed input, division by zero, and unsupported functions resolve or retain the stored display exactly. |
| unit | `mail_merge_control_state_is_story_and_record_scoped` | Approved control instructions do not leak state between stories or records and unavailable context retains the cache. |
| unit | `toc_tc_and_barcode_fields_preserve_non_text_results` | Supported non-text decisions use the approved outcome form while unsupported syntax keeps instruction text, cache, and a stable diagnostic. |
| regression | `extended_field_updates_preserve_instruction_and_result_scaffolding` | Simple and complex fields retain run formatting, dirty spelling policy, nested order, unmodelled XML, and surrounding package content after save and reopen. |

The **test gate** is differential. Supported field results match the pinned
Word values, and unsupported instructions remain intact with diagnostics.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/12-testing-strategy.md`

## Risk routing

- Public API of a published crate. Read `docs/hld/10-bindings-spec.md` and the
  `CLAUDE.md` structural rules. State the pre-1.0 additive or breaking impact,
  run `cargo publish --dry-run -p rdocx`, and assert the generated `.crate`
  remains below 10 MiB.
- External oracle comparison. Read
  `.claude/skills/differential-testing.md`. Pin Microsoft Word 16.104 build
  16.104.25121423 and the locale in the source test metadata. Keep the oracle
  out of published crate dependencies and use only source-built fixtures.

## Hash harness

Expected to be unchanged. The story adds source-built test inputs and does not
change any harness sample. Any output delta is unexplained and blocks
integration.

## Implementation checklist

- [ ] Add the approved structured outcomes and supported instruction subset.
- [ ] Add bounded syntax validation for each approved field family.
- [ ] Evaluate supported text, control, deferred, and generated-content cases through the existing ordered evaluator.
- [ ] Preserve instruction source, cached display, formatting, dirty policy, and unmodelled XML for every fallback.
- [ ] Add the pinned Word differential matrix and focused bounded regressions to the existing test binaries.
- [ ] Update the listed HLD sections and native facade contract.
- [ ] Run focused `rdocx` checks plus every risk rider.

## Open questions

None. Structured native outcomes are approved for TOC or TC decisions,
mail-merge controls, and barcode specifications. Formula fields return
resolved text. The supported new instruction set is `TOC`, `TC`, `NEXT`,
`NEXTIF`, `SKIPIF`, `MERGEREC`, `MERGESEQ`, `DISPLAYBARCODE`, and
`MERGEBARCODE`.
