# S96 final union repair review, Root pass28

## Verdict and boundary

Zero defects, zero smells and zero nitpicks. This pass reviews source and scoped correctness. Performance acceptance and final riders remain pending. No product, test, plan or HLD file was changed in this review.

The composed increment against Root27/support9 is five paths, +1142/-137. Native document.rs contributes +1086/-134. The other paths are the existing plan and HLD03/04/10. The final manifest has4309 entries, with4304 unchanged from that base. Native SHA-256 is `cc2bb7754f79dbb6081787493c99bd8d833e88c5df7343f58cdf0fc5f00f93c0`. Boundary SHA is `1d02cbd62a859693dd2a09100f866cb329a96b8b45beba9d71c4a4a42395a00c`. Increment SHA is `989aa1155800c9dd59de7eb45a13a71f14231f779eb31c5201ec117d3cf41208`.

The actual independent [support10 report](S96-table-values-and-native-frame-implementation-support-10.md) is complete and consumed. It reports zero findings and independently authenticates the boundary, current files, scoped artifacts and evidence. This Root review does not impersonate that reviewer.

## Correctness and contract

Zero findings. The selected storage is advisory value storage only. Physical cell selectors use checked row/cell/content access and require the Paragraph variant. The factory checks the adopted simple model domain once and retains one block Arc per table. Source-owner-local paragraph ordinals select preenumerated physical indices, never source index_path values. Missing, shifted and invalid hints cannot become physical affiliation. The explicit model-domain-only incidental equal-value contract is implemented. References: `crates/rdocx/src/document.rs:9202`, `crates/rdocx/src/document.rs:9253`, `crates/rdocx/src/document.rs:9318`.

Every hit requires a valid selected leaf, exact full XML key and successful Weak upgrade with Arc::ptr_eq. An immutable leaf clone occurs outside the metadata lock. A miss executes the actual fixed CT_P reader once, then the existing lexical Field exclusion, both finite qualifications and complete leaf equality. No whole-table/private-sibling equality, wider finite grammar, Source receipt or writer-as-reader permission was introduced. References: `crates/rdocx/src/document.rs:9353`, `crates/rdocx/src/document.rs:9374`, `crates/rdocx/src/document.rs:9390`.

Source preparation remains authoritative. Ownership and unknown-parent checks precede complete source paragraph traversal. Actual marker extraction precedes every value lookup. The source location, flow, offset, namespace closure and marker data are not obtained from model selectors. References: `crates/rdocx/src/comments.rs:3085`, `crates/rdocx/src/comments.rs:3095`, `crates/rdocx/src/comments.rs:3100`.

## Lifetime, concurrency and panics

Zero findings. The retained state owns one Weak root and one direct key or a selector map, with no AST or old-root list. Root mismatch drops the complete old map and capacity. Same-root replacement drops the selected old Vec. A successful nonmatch removes only its selector and drops an empty state. Reader errors remain uncached. Competing roots may evict optimization without mixing roots. Poison recovery retains full root/key/selector checks. All parsing, finite checks, equality and returned cloning remain outside locks. References: `crates/rdocx/src/document.rs:9358`, `crates/rdocx/src/document.rs:9412`, `crates/rdocx/src/document.rs:9446`.

Actual source-child root replacement detaches the current block's cell before staged installation. Source-only publication also clears the cell before clearing staged storage. New constructors start empty. Same-root clone/fresh-context carry retains only the existing root and metadata cell. The unit controls include actual replacement/removal, enclosing COW, independent equal roots, live held-root retention followed by expiry, competing selectors/roots and poison recovery. References: `crates/rdocx/src/document.rs:9855`, `crates/rdocx/src/document.rs:12407`, `crates/rdocx/src/document.rs:28890`, `crates/rdocx/src/document.rs:86492`, `crates/rdocx/src/document.rs:86571`.

New selector and collection lookups use checked access. Internal ordinal increments are bounded by the actual factory's finite source paragraph iteration. No new unchecked input-derived indexing, public panic, trait object or callback under lock was found.

## Native frame storage and OOXML

Zero findings. The private Cow field and substantive helper use only four exact static byte slices. Every other selected value retains its owned bytes. The helper has ten actual consumers. Three retain resolver-selected names and push-only conversion, five retain eager Start conversion, and two retain eager Start/Empty conversion. Root QName validation remains before the replacement consumer's conversion. Byte-slice reads are mechanical. Namespace, owner/item/SDT/Field decisions and actual parser/attribute/capture/resource work retain their original owners and order. References: `crates/rdocx/src/document.rs:13311`, `crates/rdocx/src/document.rs:13325`, `crates/rdocx/src/document.rs:17963`, `crates/rdocx/src/document.rs:20247`, `crates/rdocx/src/document.rs:21924`.

Neither parser optimization changes a writer, schema sequence, font mode or opaque XML preservation policy. Actual unknown/Unicode/case-different names and foreign namespace names remain distinct. Static byte storage does not classify them as Word names or grant Source authority.

## Evidence and structure

Zero findings. The genuine predecessor completed owned-output comparisons before the mechanism assertion failed with six reads/zero hits against three reads/three hits. The successor observes three first reads, three hits and three selector keys. Additional tests cover exact replacement/removal/error behavior, wrong and incidental hints, all model-domain exclusions and selected/private sibling behavior. These are discriminating controls rather than helper-constant mirrors.

The actual final binary has82 disjoint successful Native cases and zero final ignores. Its70 risk cases comprise53 preserved successes and17 resumed cases after the runner's omitted required stack setting caused an actual abort. That failure remains recorded. The same unchanged binary passes the resumed cases with8MB. The other12 cases include the complete actual scanner matrix and value-cache controls.

Nineteen complete before/after scanner outcomes, typed errors, byte outputs and Native A9 work match exactly. All ten consumers actually run. Both actual frame sizes are48 bytes. The comparison is internal predecessor preservation, not external Office parity. The joined outcome SHA is `19876d11b8888fd6e8903b2ed2387b2588b7b906e2dcac040e26506fd48914fb`.

Fresh ordinary debug wheel `80e1ff5d577bbe798ff96b7df688360e59a60088583c049235bd47ed3cf07820` and Native `0c195c4113b2963c6197528a48c27e2e7a08602efdc8f6b47cba0604aa99b98d` authenticate all ten actual installed payloads. The two public groups have64 and10 disjoint passes with zero failures/errors/skips. Diagnostic exports and environment hooks are absent from that ordinary artifact.

All sixteen actual diagnostic public states preserve complete ZIPs and every member. Common owned-name values become corresponding actual static views. Other owned values/copied bytes, projections, pushes and the scoped independent reader/attribute/admission/capture/value/retained metrics match. Third-reply table hits are20/52. Point1's event/byte meter is explicitly limited to its active paired range, and all sixteen public point1 rows are zero. No complete workspace-read, total allocation, RSS or optimized-time claim is inferred. Mechanism SHA is `41988a7f479656523d088ae8cc5574d2111b79fe18a6789642c84170f8da93ea`.

All-target check, strict all-target/all-feature Clippy and final cheap gates pass. The first lint failure was an equivalent cfgtest default-tuple expression. Its correction leaves product source unchanged. Original test assumptions, stack abort, prose failures and metadata assembly failure remain preserved. Only disclosed plan punctuation/code-format changes followed the ordinary debug producer. No unperformed check is counted as a pass.

The concrete selector/key enum has two current consumers. The mapping helper has ten current consumers. No new trait, generic parameter, module, source file, feature flag, dependency or stable API was added. The existing plan and HLD03/04/10 describe the resulting behavior and resource costs.

## Remaining acceptance

The previous ordinary14e78 reply and Document picture misses remain failed. A new ordinary release artifact, actual installed functional controls and a fresh compiler-free unchanged79-flat/5-linear/both407 batch must decide performance. No favorable artifact retry, threshold change or noise waiver follows from this review.

Final memory omission/equivalence/calibration/RSS/history work must cover the complete selector map, keys, capacity, Weak and inline overhead, as well as the actual final temporary frame representation. Native Office, whole installed typing, hosted/archive and full-union gates remain due. No state/ledger update, commit, push, tag or publication is authorized by this review. This pass ends here.
