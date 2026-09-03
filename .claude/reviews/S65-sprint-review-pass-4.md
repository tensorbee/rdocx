# S65 sprint review, pass 4

**Reviewed**: `sprint/s65` at
`1f836b1bb7e2f1b2ee0ea536570c3fcb12c40cb4` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 33 files, 5,923 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 3 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, math-run and display property containers lose their parent raw slot

`crates/rdocx-oxml/src/math.rs:143`

`crates/rdocx-oxml/src/math.rs:187`

`crates/rdocx-oxml/src/math.rs:560`

`crates/rdocx-oxml/src/math.rs:582`

The fraction, script, radical, matrix, n-ary, delimiter, and accent parsers
record an existing leading property container in their parent preservation
sidecar. The math-run and display parsers do not. They increment the modeled
slot after `m:rPr` or `m:oMathParaPr`, but the writers then conclude that the
property container did not exist. A parsed run such as `raw, rPr, raw, t` is
therefore written as `rPr, raw, t, raw`, and the corresponding display case
moves raw content across `oMathParaPr`. An empty existing property container is
also omitted. This violates the owner and schema-slot contract at
`docs/hld/04-opc-and-packaging.md:206`. Both parsers must preserve the leading
container in the parent sidecar, both writers must retain an existing empty
container, and focused first-write and reopen tests must cover raw children on
both sides.

### B2, non-text nodes retained inside math text are silently discarded

`crates/rdocx-oxml/src/math.rs:2066`

`crates/rdocx-oxml/src/math.rs:2148`

`crates/rdocx-oxml/src/math.rs:2894`

`element_text` accepts comments and processing instructions inside `m:t` by
falling through its event match. The run records the complete original `m:t`
as a modeled child, but `write_math_text` reparses only its root attributes and
writes a new text-only element. The comment or processing instruction is lost
on the first save. The public query also examines only text attributes, so it
reports this value as fully supported before the loss. This contradicts the
unsupported-descendant preservation contract at
`docs/hld/03-architecture.md:372` and the query contract at
`docs/hld/10-bindings-spec.md:237`. The reader must either keep such a run
opaque or preserve the non-text nodes through mutation and reopen, and the
query must report whichever retained content remains outside the typed model.

### B3, the unsupported-content query treats the wrong math attribute as modeled

`crates/rdocx-oxml/src/math.rs:2110`

`crates/rdocx-oxml/src/math.rs:2350`

`crates/rdocx-oxml/src/math.rs:2812`

The property query permits both `m:val` and `m:alnAt` on every modeled leaf,
although the writer models exactly one attribute according to the leaf type.
For example, `m:type` with valid `m:val="bar"` plus retained `m:alnAt="4"`
remains a typed fraction. The writer retains the extra attribute, but
`has_unsupported_content` returns false because `m:alnAt` is globally
allow-listed. The inverse false negative exists for `m:brk` with an extra
`m:val`. F-229 and F-230 therefore cannot rely on the public query promised at
`docs/hld/10-bindings-spec.md:237`. The query must compare each property leaf
with its actual modeled attribute and focused tests must cover both attribute
families.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding status

- Pass-1 B1 and pass-2 B1 remain resolved for the retained raw content already
  covered by the recursive query. B2 and B3 above identify separate untested
  false-negative paths.
- Pass-1 B2 remains resolved for fraction, radical, n-ary, delimiter, and
  accent property insertion. B1 above identifies the two leading-property
  writers that do not use the repaired parent-preservation path.
- Pass-1 B3, pass-2 B3, and pass-3 B1 are resolved. The named corpus gate now
  proves the reopened root extension after the final accent, the property
  extension before `m:type`, and the argument extension before its first run at
  `crates/rdocx-oxml/src/math.rs:3554` and
  `crates/rdocx-oxml/src/math.rs:3571`.
- Pass-2 B4 remains resolved. Radical degree and n-ary limit arguments are
  optional in the validator at `crates/rdocx-oxml/src/math.rs:2711`, with all
  absence combinations exercised at `crates/rdocx-oxml/src/math.rs:3407`.

## Review-bound extension

The user-approved extension is recorded at
`.claude/reviews/S65-sprint-review-pass-3.md:35`. Pass 4 and any remediation
passes required for a clean verdict are authorized.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`). It does not hold at this
dependency-prefix boundary. F-229 and F-230 remain pending at
`docs/sprints/CURRENT_SPRINT.md:32`, so rendering and conversion evidence does
not exist. F-228 also cannot advance as the reviewed dependency prefix because
B1 and B2 violate raw preservation, while B3 leaves its consumer diagnostic
surface incomplete.

## Not found

- `interaction`: no implemented consumer interacts with F-228 at this boundary.
  The remaining consumer-readiness failures are reported in B3.
- `duplication`: no second OfficeMath model or competing helper family was
  added.
- `layering`: no manifest or lockfile changed, and no forbidden dependency
  direction was introduced.
- `harness`: the baseline file is unchanged. The hash harness passed with 49 of
  49 entries at the reviewed SHA, matching `docs/sprints/AS_BUILT.md:11271`.
- `docs`: all six HLD files listed by the approved F-228 design were updated.
  The implementation mismatches are reported above rather than papered over.
- `deps`: no dependency, feature flag, crate, trait, generic parameter, or new
  integration binary was added.
- `surface`: the additive native equation, paragraph, settings, and diagnostic
  query APIs match the approved surface. B3 concerns their behavior, not an
  unrequested API.
- `delivery`: `CURRENT_SPRINT`, `BACKLOG`, `SPRINT_TRACKER`, and `AS_BUILT`
  consistently record F-228 as completed and its consumers as pending.
- `focused checks`: all 22 OfficeMath-focused unit tests, both named facade
  integration tests, the legacy Equation Editor regression, and the 49-entry
  hash harness passed at the reviewed SHA. None exercises B1 through B3.
- `differential`: F-228 declares no external oracle comparison. The pinned Word
  rendering and Pandoc conversion oracles remain obligations of F-229 and
  F-230.
