# F-X272, Root glossary omission functional review, pass 1

**Reviewed**: The sealed three-path Native increment in `ROOT_TAKEOVER/root-glossary-omission-1/increment.patch`, 428 additions and 35 deletions. The separate progress checkpoint is not product source. All final source and patch hashes authenticate against `SOURCE_BOUNDARY.json`. This is bounded independent Root review, not whole-provider acceptance or a reset of a previous peer cycle.

**Verdict**: 0 defects, 1 smell, 0 nitpicks.

## Defects

0 defects in the reviewed semantic handoff.

## Smells

### S1, body selection first copies the entire main part unnecessarily

`crates/rdocx/src/document.rs:36207` copies the complete parent Source with repeat_source solely to pass ownership into package_authoritative_body_source. That private selector at `crates/rdocx/src/document.rs:14691` only borrows its input while creating its selected outputs. No subsequent operation needs that intermediate complete-part buffer. The added copy allocates and charges every positively classified Field in the parent before selection, including material outside the selected body. This is avoidable work on the resource-sensitive path being repaired. Make the existing private selector borrow its Source and pass the actual retained parent directly, adjusting its existing callers. Keep actual selected fragment and final publication copies charged normally. No new public carrier or transport exemption is needed.

## Nitpicks

0 nitpicks.

## Functional and evidence assessment

The qualified marker planner preserves reverse marker validation order and opaque-content refusal. The byte consumer and checked Source consumer share its completed deletion plan. Namespace closure uses the actual selected body and optional final section ranges. Recursive note discovery preserves the existing pending and seen policy. Only selected note owners are replaced, with actual paid bytes supplied to the non-main setter. Main publication retains the actual setter reader handback and installed package Arc. The initial fragment handoff records its actual supplied main Source with its package storage, without equality-based origin recovery.

The new controls inspect ordinary/scoped bytes and package parity for both note families and both final-section modes. They retain an unused marked note, follow a recursively referenced note, check the section-only note decision and reach charged body and note Text consumers. Opaque marker refusal retains held package/input state. The reviewer inspected the actual assertions rather than treating pass counts as complete acceptance.

The seal records two final-source new test passes and strict all-targets/all-features checking for lower and Native crates at exit 0. The 21 scoped template, 18 versioned, four forward, two reverse and three transfer observations precede the final last-use Option cleanup and stay attached to their preserved source and binaries. Their results are not relabeled as new final-source executions. Compile/setup errors, the introduced empty-note publication error and the locator fixture failure remain separately classified. No historical Base RED is claimed. No reviewer compilation, native application or external oracle ran.

## Explicit boundaries

Comments dirty emission, rebuilt revision preparation and installation, complete non-emitting expression evaluation, callable closure, HLD and canonical source reconciliation, whole C1/provider readiness and final integrated checks remain pending. This review does not certify them. The existing 42-carrier and seven-accessor contract is unchanged by this increment.

## Not found

No additional correctness, panic, namespace/schema-order, unmodelled XML preservation, recursive selection, operation ownership or error-order defect was found in this exact bounded increment. No new trait, generic, source file, module, quota, fresh ingress or public callable is introduced.
