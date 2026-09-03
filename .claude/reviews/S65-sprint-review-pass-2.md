# S65 sprint review, pass 2

**Reviewed**: `sprint/s65` at
`fd56379f7d8e0d5b5076ce27815e79962ae180e1` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 31 files, 5,508 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 3 blocking, 0 should-fix, 0 nice-to-have

## Blocking

### B1, the unsupported-content query still misses retained math-text extensions

`crates/rdocx-oxml/src/math.rs:283`

`crates/rdocx-oxml/src/math.rs:565`

`crates/rdocx-oxml/src/math.rs:2030`

The remediation added a recursive `has_unsupported_content` query, but its run
branch checks only the run container and run-property preservation. Parsing
stores the modeled `m:t` source separately, and writing reparses that source so
attributes other than the modeled `xml:space` attribute survive. The existing
case with `m:t x:keep="yes"` proves this retained extension at
`crates/rdocx-oxml/src/math.rs:3476`, yet the public query returns false for that
run. This contradicts the consumer contract at
`docs/hld/10-bindings-spec.md:236` and still lets F-229 or F-230 silently omit
retained run-text metadata. The query must inspect the preserved `m:t` source,
with a regression that reports a foreign attribute while treating modeled
`xml:space` as supported.

### B3, the named reopen gate still does not prove the logical raw slots

`crates/rdocx-oxml/src/math.rs:3313`

`crates/rdocx-oxml/src/math.rs:3383`

`crates/rdocx-oxml/src/math.rs:3394`

The named gate now serializes the reopened value, but its second check only
asserts that each raw byte fragment occurs somewhere. It does not call the
available ordering helper or prove that the property and argument fragments
remain owned by `m:fPr` and `m:num`. Moving all four fragments to a root tail
would still pass. The contract at `docs/hld/14-development-backlog.md:2092`
requires opaque root, property, and argument siblings through reopen. The named
gate must assert owner-relative placement and order in the second serialized
value.

### B4, schema-valid optional radical and n-ary arguments are rejected as opaque

`crates/rdocx-oxml/src/math.rs:2646`

`crates/rdocx-oxml/src/math.rs:2649`

`crates/rdocx-oxml/src/math.rs:1002`

The `CT_Rad` shape check requires `m:deg`, although the OfficeMath degree is
optional, and the `CT_Nary` check requires both `m:sub` and `m:sup`, although
both limits are optional. A valid square root containing only `m:e`, or a valid
n-ary operator with one or no limits, therefore remains opaque instead of
using the empty arguments already represented by the public constructors. This
contradicts the claimed supported radical and n-ary coverage at
`docs/hld/12-testing-strategy.md:655` and blocks F-230's optional-root-degree and
n-ary conversion contract at `.claude/plans/F-230-design.md:76`. The validator
must accept the optional schema positions and focused tests must cover every
absence combination.

## Should-fix

None.

## Nice-to-have

None.

## Pass-1 remediation status

- Pass-1 B1 is only partially resolved. The public recursive query exists, but
  the retained `m:t` case above remains invisible.
- Pass-1 B2 is resolved. The shared leading-property writer preserves slot zero
  before accounting for a newly inserted property container at
  `crates/rdocx-oxml/src/math.rs:2147`, and the regression checks first write
  and reopen order at `crates/rdocx-oxml/src/math.rs:3213`.
- Pass-1 B3 is only partially resolved. Reopened bytes are serialized and
  checked, but their logical owners and slots are not checked again.

## Milestone gate

The M22 gate at `docs/hld/14-development-backlog.md:2079` does not hold at this
dependency-prefix boundary. F-229 and F-230 remain pending at
`docs/sprints/CURRENT_SPRINT.md:31`, so equation rendering and conversion
evidence does not exist. The prefix also cannot advance because B1 leaves the
consumer diagnostic surface incomplete, B3 leaves the mandatory reopen gate
incomplete, and B4 leaves valid members of the declared supported grammar
opaque.

## Not found

- `duplication`: no second OfficeMath model or competing preservation helper
  family was added.
- `layering`: no forbidden dependency direction was introduced. The sprint
  diff changes no manifest or lockfile.
- `harness`: the reviewed baseline is unchanged. The focused hash check passed
  with 49 of 49 entries, matching `docs/sprints/AS_BUILT.md:11271`.
- `docs`: all six HLD files listed by the F-228 plan were updated. The remaining
  contract mismatch is reported in B1 and B4.
- `deps`: no dependency, feature flag, crate, trait, generic parameter, or new
  integration binary was added.
- `surface`: apart from B1, the additive native equation, paragraph, and
  settings surface matches `docs/hld/10-bindings-spec.md:230`. Python, WASM,
  and CLI surfaces remain unchanged.
- `delivery`: `CURRENT_SPRINT`, `BACKLOG`, and `AS_BUILT` consistently record
  F-228 as completed, with F-229 and F-230 pending. The prior microscope ended
  at zero findings in `.claude/reviews/F-228-all-pass-6.md:5`.
- `focused checks`: 20 OfficeMath unit tests, both named facade integration
  tests, the legacy Equation Editor regression, and the 49-entry hash harness
  passed at the reviewed SHA. These green checks do not exercise the three
  blockers above.
- `differential`: F-228 declares no external oracle comparison. The pinned Word
  render and conversion oracles remain obligations of the pending consumer
  stories.
