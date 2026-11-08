# S96 foreign prefix temporary storage, investigation28

## Disposition and source boundary

Borrowing the validated namespace prefix until its existing consumers finish is technically feasible in numbering::word_prefixes_at. Own a prefix String only at the actual Word-alias push. Keep the UTF8 check, URI decoder, Cow materialization, retain, encoded binding push and Word-alias push in their exact existing order. No API, lifetime parameter, namespace language, Source policy, helper, generic or cache change is needed.

This report does not adopt the candidate or predict a speedup. Root still needs actual current successful foreign-branch conversion/copy counts before selection. Investigation26's observed discarded namespace-value copies measure a different String site and do not prove this prefix cost. Its secondary admit_word_reader predicate remains unselected.

Zero existing product defects, zero smells and zero nitpicks established. This pass wrote only this report and executed no product, Cargo, compiler, test, runtime, timing or network operation. Root remains sole product/Cargo writer. Pass numbering and all current ordinary/final-rider debt remain unchanged.

The inspected numbering.rs SHA-256 is `925cb0e082696835e7468b38f1844521df7bcf955b96e273e83659499c4b37ef`, matching its actual frozen Root31 ordinary producer entry in f268-guard-iterator-release-qualification-20261108-1/source-inputs.json. Root's separate26 work does not define this helper's policy.

## Exact current behavior and finite arrangement

The helper visits actual attributes with their existing iterator checks. A declaration selects an empty default prefix or the exact bytes after xmlns:. It validates those bytes as UTF8 and immediately makes a String before decoding the URI. After decoding succeeds it materializes the inherited Cow, removes old bare/encoded entries for that exact prefix, appends namespace_binding(prefix, decoded URI), and appends the bare prefix String only for the Word URI. For a successful foreign declaration the temporary prefix String has no retained consumer. References: `crates/rdocx-oxml/src/numbering.rs:644`, `crates/rdocx-oxml/src/numbering.rs:659`, `crates/rdocx-oxml/src/numbering.rs:663`.

The smallest arrangement is:

```rust
let prefix = std::str::from_utf8(prefix)?;
let value =
    attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, start.decoder())?;
let prefixes = prefixes.to_mut();
prefixes.retain(|candidate| {
    candidate.as_str() != prefix
        && split_namespace_binding(candidate)
            .is_none_or(|(candidate, _)| candidate != prefix)
});
prefixes.push(namespace_binding(prefix, &value));
if value.as_bytes() == W_NS.as_bytes() {
    prefixes.push(prefix.to_owned());
}
```

The namespace selection code, loop and signature stay unchanged. String versus str comparisons here are exact value comparisons, with no normalization, case folding, QName reconstruction, new namespace lookup or hash. The encoded namespace_binding remains the same mandatory owned formatted String. The bare Word alias remains an owned String. The foreign branch removes only the discarded standalone prefix copy. References: `crates/rdocx-oxml/src/numbering.rs:627`, `crates/rdocx-oxml/src/numbering.rs:632`.

Do not delay UTF8 validation until after URI decoding, materialize Cow before either validation, skip redundant declarations, omit encoded foreign bindings, alter the Word URI comparison or move pushes. In particular, an invalid UTF8 prefix must still beat a malformed URI, even when that URI would be foreign. The attribute iterator must still reject duplicate/malformed names in its existing order. No default, reserved, Unicode or unusual prefix receives a new validation/acceptance rule.

## Lifetime, storage and error proof

The borrowed prefix refers to the immutable BytesStart attribute-name bytes and is needed only within that attribute iteration. URI decoding borrows the attribute/start and does not mutate those bytes. prefixes.to_mut materializes a separate owned Vec from inherited storage when necessary. It cannot invalidate the prefix reference, even if the caller's start bytes originally borrowed some caller-owned scope storage, because that inherited input is not mutated.

No prefix borrow enters the returned scope. namespace_binding copies its inputs into its existing owned encoded String. The Word branch copies the prefix into its owned alias String. A declaration-free successful call remains Cow::Borrowed with the same inherited pointer. Any successful declaration still causes the original Cow mutation/invalidation behavior and output ordering. Existing caller into_owned conversions and cross-event scope ownership remain required. References: `crates/rdocx-oxml/src/numbering.rs:648`, `crates/rdocx-oxml/src/numbering.rs:663`, `crates/rdocx-oxml/src/numbering.rs:1277`, `crates/rdocx-oxml/src/numbering.rs:1290`.

For a Word declaration, the owned prefix allocation moves from before decoding/materialization to the final alias push. It is still performed before the same alias enters the Vec. Namespace binding and alias values, prefix removals and push order stay exact. The intervening changed allocation/drop interleaving contains concrete Strings and Vec operations, not user callbacks, parsers or recoverable Source errors. An allocator abort is not a new recoverable error policy. Keep existing typed errors and checks unchanged.

Retained String lengths and capacities should be observed, not presumed to improve. The old to_string and proposed to_owned both copy the same validated prefix bytes into the Word alias, and output Vec length/push policy is unchanged. The proposal adds no spare-capacity owner. Default empty prefixes need special reporting: an old empty String conversion has zero copied payload and normally no heap buffer. Removing its conversion event does not establish a removed heap allocation.

## Consumers and deliberately excluded changes

This numbering helper is used beyond numbering. Text imports it directly for paragraph/run scope, and table, controls, headers/footers, comments, notes, settings, styles, borders, ruby, glossary, placeholder and property-projection code use its current scope vocabulary. Namespace closure and persistent sinks continue to receive the same owned encoded pairs and Word aliases. Representative real consumers: `crates/rdocx-oxml/src/text.rs:19`, `crates/rdocx-oxml/src/table.rs:13`, `crates/rdocx-oxml/src/content_control.rs:14`, `crates/rdocx-oxml/src/header_footer.rs:10`, `crates/rdocx-oxml/src/comments.rs:10`, `crates/rdocx-oxml/src/footnotes.rs:8`, `crates/rdocx-oxml/src/settings.rs:11`, `crates/rdocx-oxml/src/styles.rs:8`, `crates/rdocx-oxml/src/numbering.rs:728`.

Properties::word_prefixes_at is a distinct policy: it materializes/removes prefixes before URI decode and carries only Word aliases, not this encoded binding inventory. Do not alter or mechanically rewrite it under this contract. namespace_bindings, merged_owner_bindings, namespace decoder policy, Source constructors and capture/attribute scans likewise remain unchanged. Reference: `crates/rdocx-oxml/src/properties.rs:429`.

No actual reader, decoder, attribute, capture, namespace resolution or resource check is removed. No Source receipt or allowance is reused. Private Field/source identity, error order, finite language, current model/vector/standalone certificates, eager authority and publication are outside the mutation. This is separate from26's value-Cow delay and requires its own old-work observation even if composed before a later ordinary build.

## Required observations and controls

Before adoption, observe the actual old prefix to_string construction at its real source site on the declared current warm states and both sizes. Separate successful foreign, successful Word, default-empty and decoder-failed/unclassified cases. A counter added only near a later branch is not an independent copy-site observation. Record exact prefix UTF8 payload lengths and actual owned String capacities, distinguishing those bytes from URI payload, inherited-scope copies, encoded namespace_binding allocation and retained output.

The successor must show zero temporary owned prefix construction on foreign successes, unchanged owned Word-alias creation and unchanged URI decoder attempts/results. Borrowed inspected bytes are not copied bytes. Existing inherited-Cow clone and encoded binding counts must remain equal. Observe complete returned ordered scopes, bare aliases, decoded URI spellings, Cow variant/pointer behavior and retained capacities. If the foreign branch is sparse or zero on the relevant states, keep that result and leave the candidate unselected. Profile exposure alone does not establish allocation or elapsed cost.

The existing real predecessor matrix compares complete scope values and typed Debug errors, verifies Borrowed/Owned behavior, and explicitly checks invalid UTF8 prefixes before later bad entities. Consume it and its actual callers. Add only missing discriminators, including foreign/Word/default transitions, two declarations for the same prefix, redundant equal bindings, encoded and bare inherited duplicates, long/Unicode prefixes, escaped URI spellings that decode to Word, malformed URI after a valid prefix, malformed prefix before URI error, and invalid attributes before/after a declaration. References: `crates/rdocx-oxml/src/numbering.rs:7301`, `crates/rdocx-oxml/src/numbering.rs:7343`, `crates/rdocx-oxml/src/numbering.rs:7360`.

Source-bound actual old/new component outcomes must retain namespace/projection/reader first causes and complete scopes. Actual public diagnostics must preserve original input ZIPs and setup/prior actions, then compare full results, owner/staging/source observations, complete output ZIPs/every member and reopen before copy assertions. Retained scopes and actual decoder/read/capture/resource work remain mandatory, including private Field and malformed unrelated-source controls.

This introduces no new retained omission endpoint, AST, cache or history. Final memory/equivalence/calibration still binds actual representation and transient ownership, with no RSS or universal peak claim. Explicit plan/HLD adoption, scoped affected checks, ordinary installed proof and independent implementation review precede a fresh ordinary final-source qualified run. Current elapsed failures and all final memory/Office/typing/hosted/archive/full-union riders remain pending. The pass ends here and returns control to Root.
