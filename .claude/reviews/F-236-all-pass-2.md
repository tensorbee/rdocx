# F-236, all, pass 2

**Reviewed**: Remediated uncommitted implementation diff against `dbb5ab1`, 7 files and 2,521 changed lines, comprising 2,515 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus pass 1 closure evidence
**Verdict**: 2 defects, 0 smells, 0 nitpicks

## Defects

### D1, same-namespace invalid paths are accepted as schema-positioned owners
`crates/rdocx/src/embedded.rs:1056`

The owner check accepts every ancestor between a paragraph and a run when it
is in the Word namespace. It does not verify that those ancestors form a valid
run-owning schema path. A crafted `w:r` below an unmodelled same-namespace
child such as `w:pPr` can therefore expose a nested `w:object` or `w:control`
to inventory and removal. Removing that identity patches content inside an
unsupported subtree that the preservation contract requires to remain opaque.

### D2, unknown relationship target modes are treated as internal
`crates/rdocx/src/embedded.rs:815`

`is_external` returns false for every value except a case-insensitive exact
`External`. `safe_internal_target` consequently accepts relationships with an
invalid or padded target mode such as `External ` or `ProducerDefined` and
resolves their target as an internal part. The approved contract requires a
proven internal relationship, so an unknown mode must fail closed instead of
being inventoried, extracted, replaced, or used as reachability evidence.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 1 defects D1 through D5 are closed. ActiveX binary multiplicity, shared
ActiveX properties, unrelated package-signature incoming edges, synchronized
signature invalidation, and compatibility-wrapper preservation now have the
required implementation and regression coverage.

No additional findings in panic safety, signature-policy cleanup, test-gate
sensitivity, public API shape, dependency direction, or repository structure.
