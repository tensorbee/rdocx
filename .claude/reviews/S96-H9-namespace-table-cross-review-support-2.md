# S96 H9 namespace and table support, S1 successor

Verdict: 0 defects, 0 smells, 0 nitpicks. S1 from support pass1 is closed by the integrated nearest-binding repair. This is the same bounded support finding, not another full feature peer pass or a reset of any review cap. No source changes, Cargo commands or runtime probes were performed.

## Source authentication

The canonical base remains H8 `924b348c7d019a41b80c2a955284b2d83c519d10`. The six-file working diff SHA256 is `a85c02d128fcf88f83d97549b984eeb713472292df3193870d644f6b33ca6a5e`. Only three files changed since pass1, matching the immutable private `H9-property-nearest-binding-remediation-1/PROPOSED-v2.patch`, SHA256 `dcf87562db1d5da1b8b707251b276bc133cb019d6aceafdd889f5bb8fada684f`.

| Path | Pass1 SHA256 | Reviewed successor SHA256 |
|---|---|---|
| crates/rdocx-oxml/src/text.rs | cc4fa5bbd3ab3332dead67b12c120d827ac8e6437b24e6bd6002ff71189ebff6 | 8ccfde7e3e6851d6c52caa3b23f2867f3d7e969959616b8da483027ea02ef73b |
| crates/rdocx-oxml/src/paragraph_properties.rs | acdff79eab58f183efe5856875d407c12cc9fff2c45e7ab3f5edd853ad8ea251 | 3de57a04c4d60471485ad960598528af020ff9dd46e9170548cffe69f86f2f08 |
| crates/rdocx-oxml/src/run_properties.rs | bde9f5ace7a614cd79e873a0e9d283ecdc1186443ddaa2ef3ad98c281a93d1fe | 325900dc44a358e4958da94f3c9d5cfcfa434f2971d9f8b3cefdab05413a1273 |

The immutable before snapshots match pass1 and the after snapshots match the integrated files. Numbering, table and PPT html source hashes remain exactly as recorded in pass1. Source and diff hashes were rechecked before handback.

## S1 closure and aspect results

At `crates/rdocx-oxml/src/text.rs:509`, the shared query scans actual nearest frames and reversed local declarations, returning at the first `w`. It allocates no namespace catalog and does not deduplicate unrelated prefixes. Returning false when no declaration exists preserves the prior full-snapshot canonical Word default. A nearer canonical declaration hides an outer foreign declaration, and a nearer foreign declaration overrides a canonical ancestor. This is exactly the precedence used by the previous snapshot implementation.

The four Publication predicates now call this query at `crates/rdocx-oxml/src/paragraph_properties.rs:1400`, `crates/rdocx-oxml/src/paragraph_properties.rs:1912`, `crates/rdocx-oxml/src/run_properties.rs:1552` and `crates/rdocx-oxml/src/run_properties.rs:2107`. The predicates no longer construct the complete catalog. The generic restoration helper at `crates/rdocx-oxml/src/text.rs:737` also uses the same query after checking actual local declarations. Its old condition was absence of local `w` and absence of canonical effective `w`. The new condition is equivalent, including unbound, canonical, foreign and locally declared cases. Local declaration decoding and its error order remain unchanged.

Correctness: 0 findings. Contract: 0 findings. Panics: 0 findings. OOXML: 0 findings. Tests: 0 findings. Structure: 0 findings. Nitpicks: 0.

The repair changes only binding lookup. It does not change parser capture, Source qualification, positioned raw replay, generated names, scope lifetime, raw child ordering or metered declaration emission. Generic restoration still uses the actual local declaration and destination before adding a declaration. The existing four generated property paths and their raw/source distinction therefore retain pass1's reviewed semantics.

The existing 24-case expanded-name matrix remains intact. The added nested query checks at `crates/rdocx-oxml/src/text.rs:45189` compare the result against the old full-snapshot predicate, exercise unbound and inherited foreign state, an intermediate frame without `w`, a canonical override, a nested foreign override and guard-drop restoration at each boundary. These are meaningful precedence and equivalence controls. Source inspection establishes removal of the query's catalog construction, without claiming a measured performance improvement.

The private source boundary authenticates actual full lower library 1,289 PASS, 0 FAIL, 0 ignored, and strict all-target/all-feature lower Clippy exit0. Their raw stdout and stderr hashes were verified against the receipt. These are worker executions consumed as evidence, not reviewer reruns. The initial pass1 report and its source identities remain preserved. Final integrated gates, native application riders, diagnostics, timings and memory acceptance remain separate and pending where not independently completed.
