# S65 sprint review, pass 5

**Reviewed**: `sprint/s65` at
`e50d5a87726d7736a6e71eb6b76519fc71440cfd` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 34 files, 6,208 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, shortening a public expression vector can discard retained raw siblings

`crates/rdocx-oxml/src/math.rs:38`

`crates/rdocx-oxml/src/math.rs:64`

`crates/rdocx-oxml/src/math.rs:98`

`crates/rdocx-oxml/src/math.rs:415`

The parser records each opaque child at the number of typed expressions that
precede it, while the writer emits only slot zero and the slots reached by the
current public `expressions` vector. If a caller removes the final parsed
expression, any raw sibling after that expression retains its old higher slot,
which the writer never visits. Clearing the vector can therefore discard every
raw sibling outside slot zero. The same loop shape exists for nested math
arguments at `crates/rdocx-oxml/src/math.rs:440` and
`crates/rdocx-oxml/src/math.rs:459`. Public multi-child collections for display
equations, matrix rows and cells, and delimiter arguments expose the same
structural mutation risk.

This violates the sprint requirement that supported OfficeMath remain editable
while unsupported sibling XML is preserved verbatim at
`docs/sprints/CURRENT_SPRINT.md:51`. The mandatory corpus gate mutates only one
run's text at `crates/rdocx-oxml/src/math.rs:3634`, so it cannot detect
collection shortening. The fix must retain every raw child when a parsed public
collection becomes shorter and add first-write and reopen regressions for both
root and nested collections. It must also define and test the retained slot
policy for insertion, removal, and reordering through the public vectors.

## Should-fix

None.

## Nice-to-have

None.

## Pass-4 remediation status

- Pass-4 B1 is resolved. Run and display parsers retain the existing leading
  property container at `crates/rdocx-oxml/src/math.rs:145` and
  `crates/rdocx-oxml/src/math.rs:574`. Their writers retain empty containers
  and both parent raw slots at `crates/rdocx-oxml/src/math.rs:195` and
  `crates/rdocx-oxml/src/math.rs:599`, with first-write and reopen coverage at
  `crates/rdocx-oxml/src/math.rs:3456`.
- Pass-4 B2 is resolved. A run with a comment or processing instruction inside
  `m:t` fails closed to its opaque owner at
  `crates/rdocx-oxml/src/math.rs:2702`, with byte-retention and diagnostic
  coverage at `crates/rdocx-oxml/src/math.rs:3513`.
- Pass-4 B3 is resolved. Modeled property attributes are selected per leaf at
  `crates/rdocx-oxml/src/math.rs:2128`, with both `m:val` and `m:alnAt`
  false-negative cases covered at `crates/rdocx-oxml/src/math.rs:3533`.
- Pass-1 through pass-3 findings remain resolved. The raw-slot gap in B1 is a
  separate structural mutation path through the public vectors.

## Review-bound extension

The user approved as many additional review and remediation passes as required
to reach a clean verdict on 2026-09-03. Pass 5 and later passes are authorized
under that explicit extension.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`). It does not hold at this
dependency-prefix boundary. F-229 and F-230 remain pending at
`docs/sprints/CURRENT_SPRINT.md:32`, so rendering and conversion evidence does
not exist. F-228 also cannot advance as the reviewed dependency prefix because
B1 permits unsupported XML loss during an exposed typed mutation.

## Not found

- `interaction`: no implemented consumer interacts with F-228 at this boundary.
  The remaining consumer-readiness failure is the preservation blocker above.
- `duplication`: no second OfficeMath model or competing preservation helper
  family was added.
- `layering`: no manifest or lockfile changed, and no forbidden dependency
  direction was introduced.
- `harness`: the baseline file is unchanged. The hash harness passed with 49 of
  49 entries at the reviewed SHA, matching
  `docs/sprints/AS_BUILT.md:11271`.
- `gate`: all 25 OfficeMath-focused unit tests, both named facade integration
  tests, and the legacy Equation Editor regression passed. B1 identifies a
  mutation path that none of those tests exercises.
- `docs`: all six HLD files listed by the approved F-228 design were updated.
  The implementation mismatch is reported in B1 rather than hidden by a doc
  change.
- `deps`: no dependency, feature flag, crate, trait, generic parameter, or new
  integration binary was added.
- `surface`: the additive native equation, paragraph, settings, and diagnostic
  query APIs match the approved surface. B1 concerns preservation under that
  requested public surface, not an unrequested API.
- `delivery`: `CURRENT_SPRINT`, `BACKLOG`, `SPRINT_TRACKER`, and `AS_BUILT`
  consistently record F-228 as completed and its consumers as pending.
- `differential`: F-228 declares no external oracle comparison. The pinned Word
  rendering and Pandoc conversion oracles remain obligations of F-229 and
  F-230.
