# F-008, Non-consuming setter twins

**Status**: completed
**Sprint**: S02
**Size**: M
**Depends on**: none

## Problem

The mutable facade types in `crates/rdocx/src/paragraph.rs`, `run.rs`, and
`table.rs` expose formatting primarily as `mut self -> Self` builders. A caller
holding `doc.paragraph_mut(0)` cannot mutate a property in place without
rebinding the wrapper, and these consuming methods cannot directly back Python
property setters.

## Spec reference

- `docs/hld/10-bindings-spec.md`, "Two supporting decisions".
- `docs/hld/03-architecture.md`, "The two APIs have different memory models".

## Approach

For every public consuming builder in the three named files, add a
non-consuming `set_*(&mut self, ...)` sibling that performs the mutation. Keep
the existing builder names and return types for compatibility, but make each
builder delegate to its setter and return `self`. This covers `Paragraph`,
`Run`, `Table`, `Row`, and `Cell`, including all 61 current builders: 24 on
`Paragraph`, 19 on `Run`, and 18 across the table facade types. No-argument
builders such as `superscript`, `layout_fixed`, and `v_merge_restart` receive
the same mechanical `set_*` twin.

No trait, generic helper, wrapper, module, or feature flag is introduced. The
surface is additive and directly required by F-008 and the future Python
bindings.

## Rejected alternatives

- Replace the builders. That would be a breaking API change with no benefit to
  existing Rust callers.
- Generate setters with a macro. The methods have varied arguments and bodies,
  and a macro would make readers follow indirection across three small files.

## Test plan

| Category | Test | Asserts |
|---|---|---|
| integration | `non_consuming_setters_match_consuming_builders` | Representative paragraph, run, table, row, and cell setters produce the same serialized XML as their builder twins |
| compile-time regression | `non_consuming_setters_mutate_borrowed_wrappers` | A run obtained through `doc.paragraph_mut(0).unwrap().add_run(...)` accepts `set_bold(true)`, and paragraph, table, row, and cell calls compile without rebinding |

The backlog's borrowed paragraph call to `set_bold(true)` names the wrong
facade because bold is a run property. Correct the HLD gate to obtain a `Run`
through the borrowed paragraph and call `set_bold(true)` there, with serialized
equivalence to the existing builder form.

## HLD impact

- `docs/hld/03-architecture.md`
- `docs/hld/10-bindings-spec.md`
- `docs/hld/14-development-backlog.md`

## Risk routing

- Public API of a published crate. State additive semver impact, run
  `cargo publish --workspace --dry-run`, and assert every archive is below
  10 MiB. Confirm every new method has a named builder twin and a future Python
  property consumer.

## Hash harness

Expected to remain unchanged. Builders delegate to setters without changing
their serialized output.

## Implementation checklist

- [x] Inventory every public `mut self -> Self` builder in the three files.
- [x] Add one in-place setter per builder with the same mutation semantics.
- [x] Delegate every existing builder to its setter.
- [x] Add borrowed-wrapper compilation and serialized-equivalence coverage to
      the existing rdocx integration test binary.
- [x] Run focused rdocx tests, rustdoc, and the packaging rider.

## Open questions

None. "Every consuming builder" is interpreted literally across all five
mutable facade types in the three files named by the story. The invalid
paragraph-level bold example is corrected to the equivalent borrowed-run call.
