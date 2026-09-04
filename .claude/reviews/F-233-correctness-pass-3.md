# F-233 microscope, correctness pass 3

## Verdict

0 defects, 0 smells.

## Defects

None.

## Smells

None.

## Reviewed categories

- The owned native types and two rich methods retain the approved public
  signatures at `crates/rdocx/src/field.rs:58` and
  `crates/rdocx/src/field.rs:668`. Their narrow type-complexity allowances keep
  the approved callback signature without adding a public alias.
- Candidate staging, reopen, and shared flat and rich section assembly remain
  atomic at `crates/rdocx/src/field.rs:668` and
  `crates/rdocx/src/field.rs:754`.
- Whole-block marker validation and bounded body and row expansion reject
  semantic marker payloads, crossed containers, and stale marker fields at
  `crates/rdocx/src/field.rs:816` and `crates/rdocx/src/field.rs:1665`.
- Fragment import copies recursively reachable internal relationship parts and
  remaps style, numbering, bookmark, hyperlink-anchor, content-control, and
  drawing identities per occurrence at `crates/rdocx/src/field.rs:1001` and
  `crates/rdocx/src/field.rs:1106`.
- Flat and rich text evaluation share instruction validation plus merge prefix
  and suffix resolution before general, numeric, and date formatting at
  `crates/rdocx/src/field.rs:1375` and `crates/rdocx/src/field.rs:1502`.
- The seven approved tests independently cover ordered rich output, lexical
  scope, exact EMU images, recursive fragment relationships and identities,
  callback context, round-trip preservation, the invalid-input matrix, and
  byte-exact flat compatibility at
  `crates/rdocx/tests/regression_test.rs:11293`.
