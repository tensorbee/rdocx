# S65 sprint review, pass 8

**Reviewed**: `sprint/s65` at
`348ec31561b5950a006ada6f40c09ef2b7b2dcc0` against merge base
`e697af64bb172f6ae2df7a0a29e774e88778b5ab`, 37 files, 6,999 changed
lines, crates: `rdocx-oxml`, `rdocx`
**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have

## Blocking

None.

## Should-fix

None.

## Nice-to-have

None.

## Pass-7 remediation status

- Pass-7 B1 is resolved. The property-default table covers the supported
  schema-defaulted leaves at `crates/rdocx-oxml/src/math.rs:2454`, and the
  validator applies those defaults only after rejecting wrong-namespace or
  duplicate modeled attributes at `crates/rdocx-oxml/src/math.rs:2867`.
  Child-property reads retain the distinction between an absent leaf and a
  present valueless leaf at `crates/rdocx-oxml/src/math.rs:3282`.
- The first-write and reopen regression covers run, fraction, matrix, n-ary,
  delimiter, accent, and display defaults at
  `crates/rdocx-oxml/src/math.rs:4160`. The document-wide regression keeps
  valueless required properties raw while typing the supported global
  defaults at `crates/rdocx-oxml/src/math.rs:4271`.
- Pass-7 B2 is resolved. The run-shape classifier rejects an invalid
  `xml:space` value before text extraction at
  `crates/rdocx-oxml/src/math.rs:3008`. The regression proves that the run
  stays opaque and byte-identical through first write and reopen at
  `crates/rdocx-oxml/src/math.rs:4316`.
- Pass-7 S1 is resolved. The replacement pass-6 review and remediation commits
  use the required `Harness,` trailer defined at `.claude/WORKFLOW.md:259`, and
  the pass-7 review and current remediation commits use the same trailer.
- Pass-1 through pass-6 findings remain resolved. The full grammar crate and
  named facade regressions pass at the reviewed SHA.

## Review-bound extension

The user approved as many additional review and remediation passes as required
to reach a clean verdict on 2026-09-03. Pass 8 is authorized by the extension
recorded at `.claude/reviews/S65-sprint-review-pass-7.md:97`.

## Milestone gate

The M22 gate is: "a representative modern document authors and renders
equations, rebuilds fields and a table of contents, performs advanced merge and
comparison, inventories embedded content, and round-trips its modern package
variant without losing unsupported XML or executable payloads"
(`docs/hld/14-development-backlog.md:2079`). It does not hold at this
dependency-prefix boundary. F-229 and F-230 remain pending at
`docs/sprints/CURRENT_SPRINT.md:32`, so rendering and conversion evidence does
not exist yet. F-228 is clean to advance as their reviewed dependency.

## Focused evidence

- `cargo test -p rdocx-oxml` passed 366 unit tests and 1 doctest. This includes
  the named OfficeMath round-trip gate and all pass-7 remediation regressions.
- The named `rdocx` public authoring integration, paragraph source-order
  integration, and legacy Equation Editor regression each passed.
- `cargo clippy -p rdocx-oxml --all-targets --all-features -- -D warnings`
  passed.
- `python3 scripts/hash_harness.py --check` passed 49 of 49 entries with the
  baseline unchanged, matching `docs/sprints/AS_BUILT.md:11271`.
- `cargo fmt --all --check`, `python3 scripts/prose_check.py`,
  `python3 scripts/sync_agent_skills.py --check`, and the merge-base diff check
  all passed.

## Not found

- `interaction`, 0 findings. No F-229 or F-230 implementation is present, and
  F-228 now supplies typed schema defaults without weakening fail-closed raw
  preservation.
- `duplication`, 0 findings. The remediation uses one property-default table
  and one modeled-attribute classifier rather than parallel property-specific
  paths.
- `layering`, 0 findings. No manifest or lockfile changed, and no forbidden
  dependency direction was introduced.
- `harness`, 0 findings. The baseline file is unchanged and the reviewed SHA
  passed 49 of 49 entries.
- `gate`, 0 findings in the F-228 prefix. The named round-trip gate and the
  pass-7 regressions pass. The full M22 gate remains pending on F-229 and F-230.
- `preservation`, 0 findings. Valueless modeled leaves retain typed meaning,
  malformed same-local attributes remain raw, and invalid math-text spacing
  remains byte-identical as opaque content.
- `diagnostics`, 0 findings. Unsupported property and math-text content remains
  observable through the recursive query.
- `grammar`, 0 findings. Supported property defaults, value domains, schema
  order, optional arguments, malformed sequences, and reopened output are
  covered by passing regressions.
- `docs`, 0 findings. All six HLD files listed by the approved F-228 design
  remain aligned with the implemented behavior.
- `dependencies`, 0 findings. No dependency, feature flag, crate, trait,
  generic parameter, or new integration binary was added.
- `public surface`, 0 findings. The additive native OfficeMath surface matches
  the approved F-228 contract. Python, WASM, and CLI surfaces remain unchanged.
- `delivery records`, 0 findings. `CURRENT_SPRINT`, `BACKLOG`,
  `SPRINT_TRACKER`, and `AS_BUILT` agree that F-228 is complete and F-229 and
  F-230 are pending. The reviewed commit messages use the required trailers.
- `differential`, 0 findings. F-228 declares no external oracle comparison. The
  pinned Word rendering and Pandoc conversion oracles remain obligations of
  F-229 and F-230.
