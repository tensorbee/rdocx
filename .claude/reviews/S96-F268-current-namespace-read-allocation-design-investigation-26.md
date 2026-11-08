# S96 current namespace read allocation, investigation26

## Result and boundary

One small technically safe temporary-copy candidate exists in RootAttributePlan::read: preserve the actual attribute decoder at its current point, but defer into_owned until after the existing namespace-declaration discard. This avoids constructing a discarded String only when that decoder actually returned Borrowed. It adds no parser, cache, receipt, changed namespace policy or retained state. Current workload frequency and usefulness remain unmeasured, so no candidate is adopted or elapsed benefit predicted.

A second borrowed foreign-namespace predicate is semantically feasible, but its actual exposure may be sparse. Keep it unselected without current positive observations. No larger envelope, capture, namespace-map or reader rewrite is recommended by this investigation.

Zero existing product defects, zero smells and zero nitpicks established. This pass read source and retained records and wrote only this report. No code/plan mutation, Cargo, product import, test, runtime, timing or network operation occurred. Root remains sole product/Cargo writer. Existing implemented11/12/13/16/18/19/20/21/24A contracts remain intact.

I independently matched all4309 current inputs to the ordinary Root31 producer manifest in f268-guard-iterator-release-qualification-20261108-1. OXML text.rs is `dd52458b751a823c91532bd35f69a3570da6bd809e315beb1e5a923c45259765`, and OXML document.rs is `e9ed3cbf18daadbcf5abad95440d9e083cc607203251b21a9936022c7ea70998`. Investigation22 was read with its actual Root29 provenance retained. Its sample categories overlap and do not establish current elapsed or allocation totals. The current Root31 ordinary result remains76/79flat, with reply, Document picture and Cell picture failing, all five linear rows and both407protocols passing. No quiet window or favorable artifact retry is selected here.

## Primary candidate: discard decoded namespace Cow before owning it

RootAttributePlan::read iterates each actual attribute, propagates its iterator error, decodes and normalizes its value, then immediately calls into_owned. Only afterward does namespace_declaration(name) discard xmlns or xmlns-prefixed declarations. For a decoded borrowed namespace URI, the String allocation/copy therefore has no consumer. For a decoded owned value, into_owned merely moves the decoder's existing allocation, so no extra copy should be claimed. References: `crates/rdocx-oxml/src/text.rs:743`, `crates/rdocx-oxml/src/text.rs:804`, `crates/rdocx-oxml/src/text.rs:811`.

The finite arrangement is:

```rust
let value = attribute
    .decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?;
if namespace_declaration(name) {
    continue;
}
let value = value.into_owned();
// Existing non-namespace attribute logic stays unchanged.
```

Keep the decoder invocation before the namespace test. Do not skip decoding declarations, change the decoder/entity policy, use raw attribute bytes, or move non-namespace ownership past root_attribute_namespace, expanded-name checks or retained attribute construction. The existing exact namespace_declaration predicate must remain, rather than a broader starts_with-xmlns test. Every actual attribute is still yielded and decoded in the same order, and any error still precedes that attribute's discard or later attributes.

The pinned dependency is quick-xml0.41.0. Its decoded_and_normalized_value delegates to its existing normalizer and returns a Cow. Its implementation preserves borrowed decoded input when normalization performs no replacement. This establishes that the discarded-copy case is possible, not its frequency in a workload. References: `Cargo.lock:1452`, `/Users/atulsharma/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/quick-xml-0.41.0/src/events/attributes.rs:205`, `/Users/atulsharma/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/quick-xml-0.41.0/src/events/attributes.rs:256`.

On the non-declaration branch, the same String is owned before the same namespace resolution, namespace-valued prefix dependency extraction, duplicate-expanded-name check, name UTF8 conversion and retained attributes push. Attribute and used-prefix order stay exact. On the declaration branch, the decoded Cow drops at the existing continue and cannot escape the event lifetime. A decoder-owned value still drops its owned payload there. There is no new lifetime parameter, helper, type or generic. References: `crates/rdocx-oxml/src/text.rs:822`, `crates/rdocx-oxml/src/text.rs:834`, `crates/rdocx-oxml/src/text.rs:849`.

## Real consumers and Source/error preservation

This is an existing shared reader primitive, not a toy helper. capture_root_attribute_record builds the current scope, then invokes this plan and its existing emitter. Actual consumers include empty and nonempty run/paragraph readers, body readers, table/row/cell readers and placeholder metadata. References: `crates/rdocx-oxml/src/text.rs:907`, `crates/rdocx-oxml/src/text.rs:9848`, `crates/rdocx-oxml/src/text.rs:10459`, `crates/rdocx-oxml/src/text.rs:18250`, `crates/rdocx-oxml/src/text.rs:22045`, `crates/rdocx-oxml/src/document.rs:10573`, `crates/rdocx-oxml/src/document.rs:11156`, `crates/rdocx-oxml/src/table.rs:2451`, `crates/rdocx-oxml/src/table.rs:2943`, `crates/rdocx-oxml/src/table.rs:3181`, `crates/rdocx-oxml/src/table.rs:3434`, `crates/rdocx-oxml/src/placeholder.rs:2921`.

The plan also participates in actual constructor qualification through root_attribute_plan_failed, which intentionally maps its result to a failure Boolean, and in retained run-root reconstruction/merge. Keep both consumers and their precise failure handling. No changed disposition of malformed attributes, reserved bindings, raw fallback or failed constructors is permitted. References: `crates/rdocx-oxml/src/document.rs:2424`, `crates/rdocx-oxml/src/document.rs:4608`, `crates/rdocx-oxml/src/document.rs:4644`, `crates/rdocx-oxml/src/text.rs:11139`.

emit_in and write_record keep the same retained attributes, prefix dependencies, output ordering and FieldXmlEmission behavior. Declaration pooling, capture records, Source maps, actual XML reads, constructor/attribute validation, normalization work, allowances and private Field reads stay unchanged. No skipped decoder can be justified by a declaration being unused. References: `crates/rdocx-oxml/src/text.rs:855`, `crates/rdocx-oxml/src/text.rs:875`, `crates/rdocx-oxml/src/document.rs:8508`, `crates/rdocx-oxml/src/document.rs:8547`.

Investigation22's actual reply and Cell picture profiles reach capture_root_attribute_record and RootAttributePlan::read. That supplies a real consumer connection, but neither isolates this discarded into_owned operation nor measures it on Root31. Examples: `.claude/scratch/S96-final-evidence/f268-five-remaining-symbol-diagnostic-20261108-1/reply-k2-attempt-1/profile.sample.txt:1053`, `cell-picture-k1-attempt-1/profile.sample.txt:775` in that same packet. Fresh current profiles remain separate evidence when available.

## Required observations before selection

Root should first observe actual current original warm states and both sizes, without added setup getters. Count yielded attributes, decoder attempts/successes, namespace versus retained attributes, and decoded Borrowed/Owned forms at the actual conversion site. Separately count old discarded Borrowed-to-owned copies and their exact decoded UTF8 payload lengths. In a successor, count those same decoded borrowed declarations being dropped without owning. Decoder-owned declarations are an unchanged allocation class and must not be credited as removed copies. Retained attribute ownership is separate and must remain.

The key discriminator is positive removed work in current relevant actions, with the same input/reader/constructor/output behavior. If the values are already decoder-owned, the site is rare, or counts remain constant and too small to justify another patch, retain that negative observation and leave the candidate unselected. Source-level counts are not optimized malloc calls, heap/RSS or elapsed savings.

Compare complete RootAttributePlan retained attributes/used-prefix records, exact emitted bytes, qualification/fallback outcomes and typed first errors against the actual predecessor. Include plain, default and prefixed declarations, repeated/conflicting bindings, encoded entities and whitespace normalization, unused declarations, invalid UTF8/entities, malformed/duplicate attributes and earlier declaration-decode failure before a later ordinary attribute failure. Non-namespace MC/XSI dependencies, unbound prefixes, duplicate expanded names, xml attributes, Source versus Publication roles and reserved bindings must retain the old outcome. Existing resolver-limit, spelling and emitted-root controls are relevant and should be consumed rather than replaced by success-only tests. References: `crates/rdocx-oxml/src/text.rs:38903`, `crates/rdocx-oxml/src/text.rs:38980`, `crates/rdocx-oxml/src/text.rs:39314`.

Complete actual reader/admission/capture/resource and first-cause observations, exact original public semantics/held/fresh/refusal checks, same-input full ZIP/member identity and reopen remain required. This candidate has no new retained payload or omission endpoint, but its final-source representation still belongs to the existing memory/equivalence/calibration proof. It promises neither lower peak/RSS nor threshold sufficiency. Explicit plan/HLD adoption and implementation review must precede a fresh ordinary qualification.

## Secondary option remains unselected

admit_word_reader still constructs an owned namespace_bindings Vec on its foreign-namespace close branch solely for an any predicate. namespace_bindings is exactly split_namespace_binding followed by owned pair copies. A direct borrowed filter_map/any can preserve that predicate's existing pair order, duplicates, conflicts, encoded-only policy and short circuit. The Word branch, unqualified-name refusal, actual Reader clone/close validation and subsequent full admission remain unchanged. It must not become a last-binding map, default-namespace acceptance or bare-prefix fallback. References: `crates/rdocx-oxml/src/document.rs:9795`, `crates/rdocx-oxml/src/document.rs:9824`, `crates/rdocx-oxml/src/numbering.rs:632`.

However, public inline/anchor entrypoints use this foreign branch, while the actual containing drawing reader directly invokes their admitted routines after capture and its own existing preparation. Thus prominence of drawing parsing does not prove this particular Vec executes in the remaining warm operations. Do not select it without positive actual branch/pair-copy counts. References: `crates/rdocx-oxml/src/drawing.rs:1005`, `crates/rdocx-oxml/src/drawing.rs:1479`, `crates/rdocx-oxml/src/drawing.rs:2450`, `crates/rdocx-oxml/src/drawing.rs:2503`.

No whole-reader replacement, parser batching, retained cache, envelope-policy rewrite, Source/frame equality shortcut, inline-language expansion or previously rejected private model clone omission follows. The primary candidate is a narrowly bounded proposed observation, and the secondary option remains unselected. Current three elapsed failures and all final memory/Office/typing/hosted/archive/full-union riders remain pending. This pass ends here and returns control to Root.
