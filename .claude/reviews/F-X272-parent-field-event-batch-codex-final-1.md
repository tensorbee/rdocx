# F-X272 parent Field event batch, Root final1

Root reviewed the complete frozen incremental patch27d424d7a0acac2684b29a7aab332b4e78768cbf1ace4da09605ff569b6f54a3
and its source snapshot. Every manifest source/progress, patch and final binary
hash matches. Actual binaryfbbbd9f65889494997918a3fea93cc1a229c3d9adfaa9f9c3bfff7f508d371ed
reports539passed, zero failed and zero ignored. The controlled missing-emission
reproduction compiled and exited101, without timeout. Scoped lower/native
checks, lower all-feature Clippy, formatting, whitespace and prose record0.
This is bounded Root review, not another peer or whole C1/provider qualification.

Verdict: zero new correctness defects, one allocation smell, zero nitpicks.
Prior mapped-kernel D1 is closed in this frozen increment. S1 needs repair before
a clean emitter-batch verdict. The author keeps its sole worker lease for the
next complete guard batch and may fold S1 into that remediation.

## D1 disposition

Actual optional parent emission now reaches run removal. Initial collapsed
hyperlink creation at text.rs:16865 and later boundary reconstruction at24673
share the same substantive record writer. The concrete new-record output at
text.rs:3605 isolates an outer prepaid receipt, counts actual retained raw and
revision chunks before the new Vec grows, preserves immediate event prepayments
and marks generated source correspondence. Only the resulting exact record is
prepaid for its later synchronous transfer. Source-context None cannot swallow
a sticky resource refusal. Existing local upper estimates remain separate
admission checks and are not debited as exact output demand.

The actual Field-owned kernel allocation discriminator at text.rs:27893 passes,
while restoring only the reviewed omitted arguments produces two large header
materializations against expected zero. The cumulative retained-revision control
at28012 pays exactly the first record length and refuses the next under its
remaining work. These observations close the reported class at both equivalent
publishers. They do not repair the separately preserved missing original/current
role or unqualified descendants.

## S1, unnecessary owned static QName in ordinary writer paths

FieldXmlStart::new at crates/rdocx-oxml/src/text.rs:3370 accepts a borrowed name
but unconditionally calls name.to_owned at3376. All current production calls
supply static literal names. This adds an owned QName allocation to ordinary
Publication/None paths whose old BytesStart borrowed that literal, including
CT_P at text.rs:18414 and table cells at table.rs:2672. The template/performance
sprint should not add that avoidable per-owner allocation to unrelated output.
This is a source-proven added allocation, not a measured release slowdown.

The same avoidable-copy class also appears in word at text.rs:3383. It calls
planned_word at3388, which renders an owned copy of the already supplied prefix
at3407 and returns that String only for word to discard it. Existing callers
already own or borrow the selected prefix. Formatting the QName directly from
that existing prefix avoids the extra intermediate allocation. This affects
ordinary Ruby, SDT and retained-font paths as well as inherited Field emission.
It is not a request to remove owned prefixes where a caller actually needs one.

Preserve the borrowed static name until actual attributes or qualification need
ownership, retaining reservation before those real growth points. The existing
dynamic planned-word constructor covers nonstatic QNames. Do not add a second
builder abstraction or alter event debit/prepaid semantics. Existing affected
byte/error/debit controls and scoped compile are sufficient for this small
repair, together with the already earned next-batch checks. No artificial
historical regression or new standalone test binary is requested.

## Qualified scope and remaining work

The event producer increment preserves raw lexical RPr attributes, property
order, original Source namespace decoding and generated prefix policy. It guards
actual QName/attribute growth before allocation, with exact final payload
prepayment. Ruby properties, SDT/dataBinding, retained font/color/language,
paragraph/table roots and hyperlink roots/ends use it. Fixed-name and namespace
Source/Publication controls pass. Vec and NormalizedWordOutput are the two real
current sinks of the shared record-writer generic.

The manifest explicitly leaves type-source reconstruction, positioned replay,
run/property serializers and remaining raw/parser producers unqualified. Those
remain actual work in the next batch, not new findings manufactured from an
excluded scope. Original/frozen-model deferred result roles, Native grammar
carrier/format/import/repeat, Issue430 and all final riders remain open. No
public API or new quota was added by this increment. This report ends the audit
and returns to orchestration. No source repair occurred in this pass.
