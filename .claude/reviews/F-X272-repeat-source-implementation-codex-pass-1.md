# F-X272 repeat source, Codex implementation pass 1

**Scope**: The sealed repeat_source increment only. Source citations refer to
`final/` beneath
`/private/tmp/rdocx-s96-root-c1-implementation/word-grammar-carriers/repeat-source-increment`.
The approved contract is `.claude/plans/F-X272-design.md:4322` and its partial
publication rule at `.claude/plans/F-X272-design.md:5339`.

**Verdict**: 0 defects, 0 smells, 0 nitpicks. No whole C1, Native migration or
provider completion is established by this bounded review.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Checked boundaries

- Operation anchoring and refusal occur before copying at
  `crates/rdocx-oxml/src/document.rs:6295`. The generated occurrence has
  initial_delivery false at `crates/rdocx-oxml/src/document.rs:6333`, while
  actual operation identity, constructor state and both correspondence maps
  survive at `crates/rdocx-oxml/src/document.rs:6321`. There is no new Original,
  frozen-model, Word or template-input registration in this producer.
- Positive interval coverage is materialized before copying at
  `crates/rdocx-oxml/src/document.rs:6301`. Existing materialization resolves
  live weak parent collectors at `crates/rdocx-oxml/src/document.rs:618`, and
  add_ranges sorts and unions overlap at
  `crates/rdocx-oxml/src/document.rs:101`. The result is validated and measured
  once in current coordinates at `crates/rdocx-oxml/src/document.rs:6311`.
  Numeric publication_owned ranges are installed at
  `crates/rdocx-oxml/src/document.rs:6322`, so predecessor drop does not erase
  the already observed demand. A new independent collector retains that same
  positive coverage. Its returned_fields_joined value stays unknown, with no
  claim that unobserved regions are ordinary or fieldless.
- Word event metadata, source maps and constructor correspondence are cloned
  intact before demand reservation at
  `crates/rdocx-oxml/src/document.rs:6323`. Nothing parses, normalizes or
  sanitizes the copied bytes. Events are transported without activating their
  decoded operands. Metadata lying within charged Field intervals remains
  metadata, rather than an independent input or a second debit.
- The reservation uses the existing normalized-output ledger at
  `crates/rdocx-oxml/src/text.rs:4530`. Setting the local emission flag enables
  output demand only. It neither alters the ledger allowance nor registers
  inputs. The existing checked increment refuses without committing emitted
  demand at `crates/rdocx-oxml/src/text.rs:5737`. All recoverable fallible
  preparation precedes that reservation, and the single byte clone follows it
  at `crates/rdocx-oxml/src/document.rs:6331`. Source bytes and publication
  metadata are not mutated on failure. The existing resource-refused latch is
  preserved rather than cleared.
- The actual four named tests at
  `crates/rdocx-oxml/src/document.rs:26512` exercise exact observed Field demand,
  repeated occurrences after predecessor drop, unchanged input state, uncapped
  unknown bytes, inert Word events, generated document coverage without an
  input grant, refusal persistence and foreign-operation rejection. Coverage
  union is also directly source-traced through the existing add_ranges
  implementation above. These tests do not establish whole-source ordinary
  proof or a completed Native consumer. The unchanged stored ordered accessor
  is authenticated BEFORE input, not part of this increment.
- The HLD descriptions at `docs/hld/03-architecture.md:3310` and
  `docs/hld/10-bindings-spec.md:5155` retain positive partial demand, unknown
  coverage and pending Native consumers. They do not create a universal XML
  cap or grant another allowance.

## Evidence and limits

Authenticated SOURCE_BOUNDARY.json SHA256
572743c26eb83055373434f309ed7fce5be9151a4b90ae13d928019089165375,
increment.patch SHA256
87a59c83904f3a8bb8ea9c5bcd08ba4cd78e5deb2ac152aa4cbdb35b5b1bea73,
all five BEFORE and final path hashes and the retained GREEN binary hash.
The retained final-green.log records4PASS. Its actual test execution was by
Root, not this review. The earlier paragraph-consumer controlled omission
remains exit0. The corrected document-consumer omission remains exit101.
Neither is relabeled historical Base RED.

Root reports strict lower and Native all-target/all-feature Clippy, formatting,
prose, drift and diff checks passed after the frozen source. The actual
/private/tmp/f272-repeat-clippy.log records successful completion. Those gates
remain separate Root receipts, not a fresh before/after53claim or execution by
this reviewer. No new builds or tests were run, and no source, private artifact
or plan was edited. All remaining C1, Native and final union riders remain due.
