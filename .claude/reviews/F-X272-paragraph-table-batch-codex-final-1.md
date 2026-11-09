# F-X272 paragraph/table producer batch, Root final1

Root reviewed the complete frozen patch
9137f6b19b5a146d16ee14d73ef1bf4b8db942c412949df135f70206bdc7fe53
and its immutable final-snapshot. All manifest source/progress, patch and test
binary hashes match. Actual final output reports610passed, zero failed and
zero ignored. The compiled controlled CT_P-to-PPr omission exits101 without
timeout, observing one large materialization instead of zero. Scoped lower and
native checks, lower all-feature Clippy, formatting, whitespace and prose
record0. The styles base is an authenticated earlier frozen carry, explicitly
not a falsely reconstructed pre-edit observation.

Verdict: one defect, one allocation smell and zero nitpicks. This bounded batch
needs remediation before acceptance. These are source-proven findings, not
new runtime RED results produced by Root. The sole author retains the lease.

## D1, empty cell properties reserve an unproduced event

At crates/rdocx-oxml/src/table.rs:2211 the checked tcPr writer creates and reserves
its root before the empty-property return at2218. FieldXmlStart::new reserves
the QName through the inherited normalization ledger. An empty CT_TcPr returns
without ever writing that event, so it consumes work and can refuse despite
producing no XML. The ordinary writer still returns empty bytes. This violates
the actual-production boundary and can make an otherwise accepted containing
Field refuse solely because a cell holds Some(empty properties).

Move the existing empty check ahead of guarded event construction and Source
qualification. Add an affected control for an actual empty cell-property owner
in both roles, with exhausted remaining normalization work. It must emit no
bytes, consume no additional work and agree with the ordinary route. Retain
the existing nonempty before-allocation refusal test. Batch these controls with
the next earned property checks rather than rebuilding separately per case.

## S1, new owned canonical prefix for ordinary border output

At crates/rdocx-oxml/src/borders.rs:355 the new shared border writer allocates
"w".to_owned on its Publication branch. The preceding public border writer used
the literal prefix directly through write_xml_attrs and did not allocate this
standalone prefix. Source qualification needs its actual returned owned prefix,
but ordinary Publication does not. The same file already uses a borrowed Cow
for this distinction in paragraph borders and tabs. Preserve that borrowing
here as well. This is an avoidable added allocation, not a measured slowdown.
It repeats the owned-static producer class identified in the prior event review.

## Qualified implementation and explicit remainder

The actual paragraph, table, row, cell and grid callers forward the inherited
context. Shared frame, shading, border, width, margin, look and toggle producers
guard their real event payloads. Positioned PPr avoids the complete property
clone and replays borrowed generated spans. Its modeled tabs/border suppression
retains the previous selection. The styles scalar merger validates a borrowed
lexical span plan, sizes the actual output, reserves before the result buffer
and preserves lexical source spans and existing changed-attribute ordering.

The focused controls exercise actual typed owners, both roles, lexical parity,
sticky refusal, unchanged inputs and retained-border intermediate production.
They do not qualify section/revision metadata, numbering retained leaves,
positioned border/tabs container mergers, retained-toggle decoder allocation,
drawing, remaining constructor outcomes, deferred input roles or Native bridges.
Those remain the author's explicit finite work list. Whole C1, dependency
providers, Issue430 and final integrated riders remain open. No source was
changed during this review. This ends the review and returns to orchestration.
