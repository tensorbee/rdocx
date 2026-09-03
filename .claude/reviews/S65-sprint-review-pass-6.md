# S65 sprint review, pass 6

**Reviewed**: `sprint/s65` at
`ab4800038c6bfde423e87de9767a1229c80eda5c` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 35 files, 6,514 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, a nested extension in math text aborts document parsing instead of remaining raw

`crates/rdocx-oxml/src/math.rs:2704`

`crates/rdocx-oxml/src/math.rs:2939`

`crates/rdocx-oxml/src/text.rs:2466`

The shape check for `m:t` treats only comments and processing instructions as
non-text nodes. A nested start or empty element therefore passes the typed-run
shape check, reaches `element_text`, and returns `UnexpectedElement`. Paragraph
equation projection propagates that error, so opening a document containing a
producer extension inside `m:t` fails instead of leaving the run or equation as
unmodelled raw XML. This contradicts the owner-preservation contract at
`docs/hld/03-architecture.md:371` and the explicit malformed-grammar fallback
at `docs/hld/04-opc-and-packaging.md:209`.

The existing nested-text test asserts only that standalone `CT_OMath` parsing
fails at `crates/rdocx-oxml/src/math.rs:4129`. It does not exercise the direct
paragraph boundary or prove that the original equation bytes survive save and
reopen. The fix must make paragraph projection decline the typed equation when
its supported descendant cannot be parsed, or classify every nested element in
`m:t` as unsupported before text extraction. A paragraph-level regression must
prove first-write and reopen retention of the original raw equation.

## Should-fix

None.

## Nice-to-have

None.

## Pass-5 remediation status

- Pass-5 B1 is resolved for repeated-child collection edits. Root equations
  and nested arguments emit every unreached higher raw slot at
  `crates/rdocx-oxml/src/math.rs:103` and
  `crates/rdocx-oxml/src/math.rs:468`. Display equations, matrix rows and cells,
  and delimiter arguments use the same tail rule at
  `crates/rdocx-oxml/src/math.rs:213`,
  `crates/rdocx-oxml/src/math.rs:1313`,
  `crates/rdocx-oxml/src/math.rs:1318`, and
  `crates/rdocx-oxml/src/math.rs:1669`.
- The shared tail writer preserves source order for every raw child above the
  last reached ordinal boundary at `crates/rdocx-oxml/src/math.rs:2401`.
  Shortening is covered across all five public repeated-child shapes through
  first write and reopen at `crates/rdocx-oxml/src/math.rs:3734`. Insertion and
  reordering use the documented ordinal-boundary policy and are covered through
  reopen at `crates/rdocx-oxml/src/math.rs:3888`.
- Pass-1 through pass-4 findings remain resolved. The blocker above is a
  separate fail-closed path at the paragraph owner.

## Review-bound extension

The user approved as many additional review and remediation passes as required
to reach a clean verdict on 2026-09-03. Pass 6 and later passes are authorized
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
B1 turns supported-document parsing into an error for an unsupported math-text
descendant instead of preserving the original XML.

## Not found

- `interaction`: no implemented consumer interacts with F-228 at this boundary.
  The recursive unsupported-content query remains available to both approved
  consumers. B1 concerns owner fallback before a typed value exists.
- `duplication`: no second OfficeMath model or competing preservation helper
  family was added.
- `layering`: no manifest or lockfile changed, and no forbidden dependency
  direction was introduced.
- `harness`: the baseline file is unchanged. The hash harness passed at the
  reviewed SHA with 49 of 49 entries, matching
  `docs/sprints/AS_BUILT.md:11271`.
- `gate`: all 25 OfficeMath-focused unit tests, both named facade integration
  tests, and the legacy Equation Editor regression passed at the reviewed SHA.
  B1 identifies a direct paragraph malformed-descendant path not covered by
  those checks.
- `preservation`: collection shortening, insertion, removal, and reordering
  retain raw bytes under the documented ordinal policy. The remaining
  preservation failure is isolated in B1.
- `diagnostics`: the recursive `has_unsupported_content` surface includes root,
  property, argument, matrix-row, and math-text retained content. Prior
  false-negative findings remain resolved.
- `grammar`: schema ordering, optional radical and n-ary arguments, property
  insertion, and malformed supported constructs outside B1 remain covered by
  the focused tests. B1 is the only grammar fallback defect found.
- `docs`: all six HLD files listed by the approved F-228 design were updated.
  The implementation contradiction is reported in B1 rather than hidden by a
  documentation change.
- `deps`: no dependency, feature flag, crate, trait, generic parameter, or new
  integration binary was added.
- `surface`: the additive native equation, paragraph, settings, and diagnostic
  query APIs match the approved F-228 contract. Python, WASM, and CLI surfaces
  remain unchanged.
- `delivery`: `CURRENT_SPRINT`, `BACKLOG`, `SPRINT_TRACKER`, and `AS_BUILT`
  consistently record F-228 as completed and its two consumers as pending.
- `differential`: F-228 declares no external oracle comparison. The pinned Word
  rendering and Pandoc conversion oracles remain obligations of F-229 and
  F-230.
