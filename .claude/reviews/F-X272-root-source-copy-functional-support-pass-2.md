# F-X272, Root source-copy functional support, pass 2

**Reviewed**: Only the immutable follow-on from `ROOT_TAKEOVER/root-source-copy-review/source` to `ROOT_TAKEOVER/root-source-closure-review/source`. The two pinned files contain one changed document.rs and unchanged text.rs. Scope is the actual namespace-insertion copy and imported-rewrite reservation, plus the new closure regression. This is not a broader provider review or a reset of the prior review cycle.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this bounded follow-on.

## Defects

0 defects.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Assessment

`crates/rdocx-oxml/src/document.rs:6812` now reserves the real namespace-insertion copy through reserve_source_publication_copy using the pre-insertion source map and the previous positive Field windows. These coordinates belong to the same original buffer, before generated declarations are inserted. Only surviving original pieces in those actual copied windows can select raw input. Declaration bytes gain no input credit and retain their existing closure policy. Constructor admission and map/length preparation still precede reservation and final buffer growth. The genuine no-insertion move path is unchanged.

`crates/rdocx-oxml/src/document.rs:7600` now supplies the actual assembled positive Field windows to that same helper for imported rewrite. Its demand remains the prior demand calculation: retained Field prefix/suffix bytes are charged, removed Field bytes are excluded and the actual prepaid scalar payload is excluded from a second debit. The reservation occurs once before materialize_source. It does not route through compose_sources and reserve again.

The imported scalar producer at `crates/rdocx-oxml/src/document.rs:7484` assigns generated source and Word maps to its output. Its one-shot payload receipt remains bound to the parent occurrence and exact produced length. rewrite_imported_source still validates anchors, parent identity, edit geometry, actual owner role and prepayment before this reservation. Consequently including that scalar's positive Field output window in the assembled selection cannot add original input: it has no original pieces. An ordinary payload likewise does not acquire an Original or ordinary-XML input role from replacement text. No raw complement, value equality, new baseline or semantic-current membership is used.

The shared helper itself is byte-identical to pass 1. It stages the original-range union and checked allowance/demand before committing, preserves sticky refusal and cannot renew an already selected physical range. This follow-on does not change its numerical units or introduce another quota.

## Controls and evidence

The new namespace control reaches the actual private paragraph factory while leaving input state zero. A real declaration insertion then performs the first copied Field selection and asserts exactly field.len() raw input, zero semantic-current membership and the same physical union. Repeating that result refuses without renewing the input union. This distinguishes an actual insertion copy from the unchanged no-insertion move. The retained RED2 log is the semantic copied-reservation failure. Root's earlier failed fixture precondition is not relabeled as that RED.

Both source hashes authenticate against SOURCE.json:

| Path | SHA256 |
|---|---|
| crates/rdocx-oxml/src/document.rs | 37db91b2972b02442bda0670c8b27323af9e484f3188599fa7ab763e6ffddcdc |
| crates/rdocx-oxml/src/text.rs | ba8df7efeb23229bdfbec2e359c666f6b976164b5de50e16508ede91a5e42f5f |

The archived lower binary authenticates as `26d0cae70752eb4a996331647fe87f55bdb552b5ed29ccfeaeae9d063b52c7eb`. The scoped log records 198 passes with zero failures. The archived strict log records completed checking in17.41 seconds. This is the requested all-targets, warnings-denied check, not an inferred all-features result. No reviewer compilation, test or native application execution was run.

## Boundaries

This closes the two equivalent reservation handoffs named as pending in pass 1, within their actual positive Field/raw-copy role. It does not certify all namespace behavior, complete imported-owner classification or every import caller. The scalar producer, prepaid receipt and existing validation semantics are preserved, not given a broader exemption.

Other returned contexts, remaining comments/glossary work, complete non-emitting expression evaluation, Native handoff closure, final integration, callable inventory, SemVer, full local/integrated checks and hash/native/oracle riders remain pending. No whole F-X272 or provider acceptance follows.

## Not found

No additional concrete coordinate, prepaid-demand, duplicate-reservation, sticky-refusal, atomic-commit, input-renewal or structural-rule defect was found in the exact follow-on examined.
