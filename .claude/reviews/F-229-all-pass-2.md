# F-229, all aspects, pass 2

**Reviewed**: remediated uncommitted worker diff, 15 files, 2,429 additions and 15 deletions
**Verdict**: 2 defects, 1 smell, 0 nitpicks

## Defects

### D1, matrix base justification is ignored
`crates/rdocx-layout/src/math.rs:928`

Every matrix receives a midpoint baseline even when the typed F-228
`base_justification` requests the top or bottom row. A top- or bottom-justified
matrix therefore reports the wrong ascent and descent and moves adjacent text.

### D2, delimiter no-grow is ignored
`crates/rdocx-layout/src/math.rs:1145`

The delimiter target height always follows its contents. When the typed
`grow` property is `Some(false)`, the begin and end glyphs are still stretched,
which contradicts the modeled property and changes the approved expression
geometry.

## Smells

### S1, argument and equation expression loops duplicate traversal
`crates/rdocx-layout/src/math.rs:241`

`layout_expressions` and `layout_equation_expressions` repeat the same recursive
loop and differ only in final spacing. A construct added to one path can drift
from the other. One spacing parameter on the private helper removes the second
place a reader must inspect.

## Nitpicks

None.

## Not found

Pass 1 cache invalidation, external-oracle connection, raster mutation, and
exact-token findings are fixed. No additional findings in panics, OOXML
preservation, or public structure.
