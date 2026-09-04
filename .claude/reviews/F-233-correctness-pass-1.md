# F-233 microscope, correctness pass 1

## Verdict

3 defects, 1 smell.

## Defects

1. `crates/rdocx/src/field.rs:996` imports fragment relationships and body
   content without remapping fragment bookmark, content-control, or drawing
   identities against the destination. Repeated fragments can therefore emit
   duplicate document-scoped identities.
2. `crates/rdocx/src/field.rs:914` reuses the outer record number and output
   sequence for every nested source record. The callback context cannot
   distinguish records within a region or successive emitted fields.
3. `crates/rdocx/tests/regression_test.rs:11326` asserts only two image
   relationships and fragment text. It does not exercise repeated fragment
   insertion or the approved style, numbering, hyperlink, bookmark, control,
   drawing, and exact identity collision contract.

## Smells

1. `crates/rdocx/src/field.rs:782` duplicates the flat section assembly at
   `crates/rdocx/src/field.rs:648`, leaving two implementations of the same
   section boundary behavior.

## Clean categories

- Public type shape and exports: no finding.
- Atomic candidate publication: no finding.
- Region marker removal and scalar shell removal: no finding.
- Image relationship staging and exact EMU use: no finding.
- Flat API source compatibility: no finding.
- Method-level type-complexity allowances: justified by the approved public
  callback signature, no finding.
