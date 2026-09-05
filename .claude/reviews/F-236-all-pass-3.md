# F-236, all, pass 3

**Reviewed**: Final-pass remediated uncommitted implementation diff against `dbb5ab1`, 7 files and 2,681 changed lines, comprising 2,675 additions and 6 deletions, including untracked `crates/rdocx/src/embedded.rs`, plus pass 1 and pass 2 closure evidence
**Verdict**: 1 defect, 0 smells, 0 nitpicks

## Defects

### D1, schema positioning is not validated above the paragraph
`crates/rdocx/src/embedded.rs:1069`

The owner check validates the ancestors between a paragraph and its run, but
for the path above that paragraph it checks only that the first open node is a
recognized Word story root. Every intervening node is ignored. A crafted path
such as `w:hdr/w:pPr/w:p/w:r/w:object` therefore passes even though `w:pPr`
cannot own a paragraph. Inventory and removal can consequently act on content
inside an invalid or unsupported same-namespace subtree instead of failing
closed and preserving it as opaque XML. The schema-position contract requires
the complete owner path, not only the paragraph-to-run suffix, to be proven.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 1 defects D1 through D5 are closed. ActiveX binary multiplicity, shared
ActiveX properties, unrelated package-signature incoming edges, synchronized
signature invalidation, and compatibility-wrapper preservation now have the
required implementation and regression coverage.

Pass 2's unknown target-mode defect is closed, and its cited invalid owner path
between a paragraph and run is rejected. The remaining defect is the distinct
unchecked prefix between the story root and paragraph.

No additional findings in signature-policy cleanup, target resolution,
mutation atomicity, panic safety, public API shape, dependency direction,
test-gate sensitivity, or repository structure. The focused embedded
regression selection passes all 6 tests.
