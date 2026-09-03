# S65 sprint review, pass 7

**Reviewed**: `sprint/s65` at
`1b01ee7c596c4c5afa88953ca06fa42d4cbc76b6` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 36 files, 6,677 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 2 blocking, 1 should-fix, 0 nice-to-have

## Blocking

### B1, omitted property values are rejected or collapsed without their schema defaults

`crates/rdocx-oxml/src/math.rs:1603`

`crates/rdocx-oxml/src/math.rs:1650`

`crates/rdocx-oxml/src/math.rs:2833`

`crates/rdocx-oxml/src/math.rs:3192`

The property validator requires `m:val` on several schema-defaulted leaves. A
run script with `<m:scr/>`, an n-ary limit location with `<m:limLoc/>`, or a
display justification with `<m:jc/>` is therefore rejected from the typed
model instead of receiving its defined default. This leaves valid supported
OfficeMath opaque to F-229 and F-230.

Delimiter characters take the opposite path. The validator accepts a missing
value, but `child_property_value` maps both an absent child and a present child
without `m:val` to `None`. The delimiter parser then substitutes the
constructor character, and its writer always emits an explicit value. A
present leaf whose value is missing therefore neither receives
schema-specific treatment nor fails closed as retained raw XML.

This contradicts the claimed focused coverage of property defaults and domains
at `docs/hld/12-testing-strategy.md:660` and the F-228 typed-default contract at
`docs/hld/14-development-backlog.md:2084`. The fix must handle omitted values
per property schema, keep valid defaulted leaves typed, preserve malformed
required-value leaves raw, and retain the distinction between an absent
property and a present valueless property. Add first-write and reopen
regressions for run, display, n-ary, and delimiter properties.

### B2, an invalid xml:space value on math text is silently discarded

`crates/rdocx-oxml/src/math.rs:2095`

`crates/rdocx-oxml/src/math.rs:2175`

`crates/rdocx-oxml/src/math.rs:2188`

The unsupported-content query excludes every attribute spelled `xml:space`
without validating its value. The writer then removes that attribute
unconditionally and recreates only `xml:space="preserve"` from leading or
trailing whitespace. A run containing `xml:space="invalid"` remains typed,
reports no unsupported text attribute, and loses the original attribute on
save.

This violates the owner-preservation contract at
`docs/hld/03-architecture.md:370` and the sprint requirement that unsupported
XML remain verbatim at `docs/sprints/CURRENT_SPRINT.md:51`. The fix must accept
only the defined `default` and `preserve` values as modeled. Any other value
must make the run fail closed to raw XML, or remain retained and observable as
unsupported. A regression must prove diagnostic state and exact attribute
retention through first write and reopen.

## Should-fix

### S1, the two newest sprint commits use the wrong required trailer

`.claude/WORKFLOW.md:259`

`CLAUDE.md:100`

Commits `39fe2392f086ba210da39d64e9cd074ba3d3b170` and
`1b01ee7c596c4c5afa88953ca06fa42d4cbc76b6` end with `Migration, none` instead
of the required `Harness,` trailer. Their bodies already state the unchanged
49-entry result. Rewrite only those messages so the required trailer records
that result consistently.

## Nice-to-have

None.

## Pass-6 remediation status

- Pass-6 B1 is resolved. The run shape check asks the classifier before text
  extraction at `crates/rdocx-oxml/src/math.rs:2704`, and nested start or empty
  elements are classified as unsupported at
  `crates/rdocx-oxml/src/math.rs:2969`.
- The direct `CT_OMath` regression proves the nested run stays opaque and
  byte-identical through reopen at `crates/rdocx-oxml/src/math.rs:4122`. The
  paragraph regression proves surrounding runs still parse while the equation
  remains present and unsupported through reopen at
  `crates/rdocx-oxml/src/text.rs:7242`.
- Pass-1 through pass-5 findings remain resolved. B1 and B2 above are separate
  property-value and text-attribute paths.

## Review-bound extension

The user approved as many additional review and remediation passes as required
to reach a clean verdict on 2026-09-03. Pass 7 and later passes are authorized
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
B1 leaves valid supported properties opaque or rewrites valueless properties,
and B2 permits an unsupported math-text attribute to disappear.

## Focused evidence

- `cargo test -p rdocx-oxml` passed 363 unit tests and 1 doctest. The narrower
  OfficeMath filter passed all 28 matching tests.
- The named `rdocx` public authoring integration, paragraph source-order
  integration, and legacy Equation Editor regression each passed.
- `python3 scripts/hash_harness.py --check` passed 49 of 49 entries with the
  baseline unchanged, matching `docs/sprints/AS_BUILT.md:11271`.
- `cargo fmt --all --check`, `python3 scripts/prose_check.py`,
  `python3 scripts/sync_agent_skills.py --check`, and the merge-base diff check
  all passed.

## Not found

- `interaction`, 0 additional findings. No F-229 or F-230 implementation is
  present to conflict with F-228. Their dependency on the defective normalized
  inputs is already accounted for in B1 and B2.
- `duplication`, 0 findings. No second OfficeMath model or competing
  preservation helper family was added.
- `layering`, 0 findings. No manifest or lockfile changed, and no forbidden
  dependency direction was introduced.
- `harness`, 0 findings. The baseline file is unchanged and the reviewed SHA
  passed 49 of 49 entries.
- `gate`, 0 additional findings. The focused gates are green. Their uncovered
  property and attribute cases are B1 and B2.
- `preservation`, 0 additional findings. Prior raw-slot and nested-text defects
  remain resolved. The remaining loss path is B2, with the related valueless
  property canonicalization in B1.
- `diagnostics`, 0 additional findings. The recursive unsupported query covers
  retained root, property, argument, matrix-row, and nested-text content. Its
  remaining false negative is B2.
- `grammar`, 0 additional findings. Schema child order, optional radical and
  n-ary arguments, malformed sequences, and repeated-child edits remain
  covered. The remaining property-default defect is B1.
- `docs`, 0 findings. All six HLD files listed by the approved F-228 design were
  updated. B1 and B2 are implementation contradictions rather than missing HLD
  edits.
- `dependencies`, 0 findings. No dependency, feature flag, crate, trait,
  generic parameter, or new integration binary was added.
- `public surface`, 0 findings. The additive native equation, paragraph,
  settings, and diagnostic APIs match the approved F-228 contract. Python,
  WASM, and CLI surfaces remain unchanged.
- `delivery records`, 0 additional findings. `CURRENT_SPRINT`, `BACKLOG`,
  `SPRINT_TRACKER`, and `AS_BUILT` agree that F-228 is complete and F-229 and
  F-230 are pending. Commit-message hygiene is reported separately as S1.
- `differential`, 0 findings. F-228 declares no external oracle comparison. The
  pinned Word rendering and Pandoc conversion oracles remain obligations of
  F-229 and F-230.
