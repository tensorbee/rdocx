# S65 sprint review, pass 1

**Reviewed**: `sprint/s65` through `4fba87e13d52f0cf784b2f39985486c1bb6d8585`
against `e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 30 files, 5,080 lines,
crates: `rdocx-oxml`, `rdocx`
**Verdict**: 3 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, the completed model cannot report unsupported content to its layout consumer

`crates/rdocx-oxml/src/math.rs:13`

`.claude/plans/F-229-design.md:76`

Every retained raw child and property child is reachable only through the
private `Preservation` value, and the public equation values expose no bounded
read-only query for that state. The separate `rdocx-layout` crate therefore
cannot distinguish a fully supported expression from one carrying opaque root,
property, or argument content. That makes the approved F-229 requirement to
emit one stable source-path diagnostic for visible unsupported content
unimplementable without reparsing serialized XML or duplicating the grammar.
The dependency prefix is not ready until F-228 exposes sufficient
consumer-readable unsupported-content metadata, or the downstream contract is
explicitly redesigned and reapproved.

### B2, absent optional property nodes move retained raw siblings to a different slot

`crates/rdocx-oxml/src/math.rs:550`

`crates/rdocx-oxml/src/math.rs:589`

The fraction parser numbers raw slots only around modeled children that exist in
the source. For a source fraction with no `m:fPr`, a raw child between `m:num`
and `m:den` is recorded at slot 1. The writer always inserts `m:fPr`, then emits
slot 1 before `m:num`, so the raw child changes owner-relative position even
without a fraction-property mutation. The same unconditional optional-property
insertion occurs for radicals at `crates/rdocx-oxml/src/math.rs:950`, n-ary
operators at `crates/rdocx-oxml/src/math.rs:1373`, delimiters at
`crates/rdocx-oxml/src/math.rs:1470`, and accents at
`crates/rdocx-oxml/src/math.rs:1553`. This violates the sprint contract that
unsupported siblings retain their owning schema slot. The fix must preserve
optional-property presence or use slot coordinates that remain stable when a
default property container is inserted, with regressions for raw children
before, between, and after required arguments.

### B3, the named round-trip gate does not verify raw preservation after reopen

`crates/rdocx-oxml/src/math.rs:3000`

`crates/rdocx-oxml/src/math.rs:3029`

The mandatory corpus test asserts the four raw fragments only in the first
serialized byte buffer. It then reparses that buffer but checks only expression
variants and the changed run text. A parser that drops every opaque sibling on
reopen would still pass the named gate. The F-228 contract at
`docs/hld/14-development-backlog.md:2092` requires opaque root, property, and
argument siblings to survive mutation and reopen. The gate must serialize the
reopened value and assert the retained bytes and logical slots again.

## Should-fix

None.

## Nice-to-have

None.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`). It does not hold at this
dependency-prefix boundary. F-229 and F-230 remain pending at
`docs/sprints/CURRENT_SPRINT.md:31`, so rendering and conversion evidence does
not exist yet. The prefix also cannot advance on F-228 evidence alone because
B2 violates raw-slot preservation and B3 leaves the story's named round-trip
gate incomplete.

## Not found

- `duplication`: no duplicate OfficeMath model or competing helper family was
  added.
- `layering`: no `oxml-*` dependency on an `rdocx-*` or `rpptx-*` crate was
  introduced.
- `harness`: the baseline file is unchanged, and the delivery record reports 49
  of 49 hashes unchanged at `docs/sprints/AS_BUILT.md:11271`.
- `docs`: all six HLD files listed by the approved F-228 design were updated.
- `deps`: no manifest or lockfile changed, and no unnamed dependency was added.
- `surface`: the native equation, paragraph, and settings additions match the
  story-owned pre-1.0 surface documented at
  `docs/hld/10-bindings-spec.md:227`. Python, WASM, and CLI exposure did not
  change.
- `delivery`: `CURRENT_SPRINT`, `BACKLOG`, `SPRINT_TRACKER`, and `AS_BUILT`
  consistently record F-228 as completed. Six feature-review passes culminate
  in zero defects, smells, and nitpicks at
  `.claude/reviews/F-228-all-pass-6.md:5`.
