# S65 sprint review, pass 3

**Reviewed**: `sprint/s65` at
`cc7b952f5e370b083b67facfe050b8aa5bd6e40d` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 32 files, 5,825 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 1 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, the named reopen gate still does not prove the original raw-child slots

`crates/rdocx-oxml/src/math.rs:3562`

`crates/rdocx-oxml/src/math.rs:3577`

`docs/hld/12-testing-strategy.md:655`

The source inserts the property extension before `m:type`, the argument
extension before the numerator expression, and the trailing root extension
after the final accent. The reopened assertions only prove that the first two
extensions remain somewhere inside `m:fPr` and `m:num`, and that the trailing
root extension remains somewhere after the fraction. A regression that moves
the property extension after `m:type`, moves the argument extension after its
expression, or moves the trailing root extension from the root tail to directly
after the fraction would still pass. This does not establish the logical-slot
preservation claimed by the mandatory corpus gate. The gate must place the
modeled neighbors in the reopened order assertions so all four raw siblings are
proved in their original slots.

This is the third bounded pass. The remaining actionable finding is blocking,
and a fourth pass is not permitted without an explicit decision to extend the
review bound. The dependency prefix is not ready to advance.

## Review-bound extension

On 2026-09-03 the user explicitly approved as many additional review passes as
required to reach a clean verdict. Pass 4 and later remediation passes are
therefore authorized under the same sprint state.

## Should-fix

None.

## Nice-to-have

None.

## Prior finding status

- Pass-1 B1 and pass-2 B1 are resolved. The public recursive query includes
  retained `m:t` extensions at `crates/rdocx-oxml/src/math.rs:283`, with the
  foreign-attribute and modeled `xml:space` cases at
  `crates/rdocx-oxml/src/math.rs:3393`.
- Pass-1 B2 remains resolved. Leading property insertion preserves source raw
  coordinates through the shared writer at
  `crates/rdocx-oxml/src/math.rs:2213`.
- Pass-1 B3 and pass-2 B3 remain only partially resolved, as described in B1.
- Pass-2 B4 is resolved. Radical degree and n-ary limits are optional in the
  shape validator at `crates/rdocx-oxml/src/math.rs:2711`, and all absence
  combinations remain typed after reopen at
  `crates/rdocx-oxml/src/math.rs:3406`.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`). It does not hold at this
dependency-prefix boundary. F-229 and F-230 remain pending at
`docs/sprints/CURRENT_SPRINT.md:32`, so rendering and conversion evidence does
not exist yet. F-228 also cannot advance as the reviewed dependency prefix
because B1 leaves its mandatory raw-slot gate incomplete.

## Not found

- `interaction`: no implemented consumer interacts with F-228 at this boundary.
  The unsupported-content query now gives the planned layout and conversion
  consumers a recursive diagnostic signal.
- `duplication`: no second OfficeMath model or competing preservation helper
  family was added.
- `layering`: no manifest or lockfile changed, and no forbidden dependency
  direction was introduced.
- `harness`: the baseline file is unchanged. The hash harness passed with 49 of
  49 entries at the reviewed SHA, matching `docs/sprints/AS_BUILT.md:11271`.
- `docs`: all six HLD files listed by the approved F-228 plan were updated. The
  remaining evidence mismatch is reported in B1.
- `deps`: no dependency, feature flag, crate, trait, generic parameter, or new
  integration binary was added.
- `surface`: the additive native equation, paragraph, settings, and
  unsupported-content query surfaces match the approved F-228 contract at
  `docs/hld/10-bindings-spec.md:230`. Python, WASM, and CLI surfaces remain
  unchanged.
- `delivery`: `CURRENT_SPRINT`, `BACKLOG`, `SPRINT_TRACKER`, and `AS_BUILT`
  consistently record F-228 as completed and its two consumers as pending.
- `focused checks`: all 20 OfficeMath unit tests, both named facade integration
  tests, the raw-boundary rebase test, the legacy Equation Editor regression,
  and the 49-entry hash harness passed at the reviewed SHA. These green checks
  do not close B1 because none asserts all original reopened raw slots.
- `differential`: F-228 declares no external oracle comparison. The pinned Word
  render and conversion oracles remain obligations of F-229 and F-230.
