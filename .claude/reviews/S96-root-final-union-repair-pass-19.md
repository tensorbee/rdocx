# S96, Root final replacement preparation, pass 19

## Scope and verdict

Root reviewed the finite six-file F268 replacement-preparation increment, the adopted contract and the actual independent implementation support report. The source boundary is .claude/scratch/S96-final-evidence/f268-retired-successor-implementation-20261107-1/SOURCE_BOUNDARY.json, SHA-256 ed46bdccd5d8dfbd469cad95e86fa4b3c190aae037cea26f133f30fd96e091da. The increment SHA-256 is 6351cc1ab0f32257c2bfa222f831771e678c97e2139a49e318ba556c9cc2ac41.

Defects: 0. Smells: 0. Nitpicks: 0.

## Assessment

The substantive allocator at crates/oxml-py-support/src/lib.rs:474 implements the three actual policies without adding a forwarding-only entrypoint. Existing global/local entrypoints and numeric bump retain their behavior. This closes the API shape issue recorded as S1 in the first design assessment, with the second design assessment and explicit Root adoption retained.

The complete virtual generation budget is checked at 483. All existing scope validation and the ordered retirement inventory remain. Every generation advances at 529 before Replacement skips a nonzero-index owner construction. The document successor remains index zero, every birth keeps its original generation position and the next operation sees the same reserved generation budget. Commit retains its original pre-retirement checks, serial and first-cause writes. No lazy allocation or publication bypass is introduced.

Only the two intended Word Conservative branches select Replacement, at crates/rdocx-py/src/document.rs:6879 and 11479. AppendBody, local publication and other existing policies retain their previous paths. Future handles still capture the allocated document and births before checked publication. PowerPoint retains its actual required owner successors. No native Word transaction clone, parser, XML, Source, namespace or resource-allowance rule changes.

Debug diagnostics count actual ScopeGeneration constructions at both allocation sites, with exact value payload bytes and a separately counted owner-successor subset. The subset excludes the implicit document by retired index zero, including a supplied Document-kind scope as an ordinary requested owner. It does not double-count the total. Shipping builds omit these observations, and existing staging/splice counter units remain unchanged.

## Evidence

The actual predecessor unit completes generation, future capture, commit and retirement semantics before failing on six creations against three. The freshly built installed predecessor likewise completes the five public cold/warm splice and reply semantic checks, then fails only on positive unused owner-successor creation against zero. These are executed mechanism failures, not absent-API compilation failures.

The successor passes four new support cases and seven affected existing cases. They cover zero/one/many states and births, exact default-equivalent and subsequent generations, document capture, first causes, dropped preparation, foreign/duplicate/retired/stale refusal, serial overflow and overflow caused only by skipped virtual positions. The freshly built successor wheel passes all twenty-six selected installed controls without skips. Five new public cases prove actual zero-subset dispatch, alongside append/local retention, conservative stale causes, failure atomicity, saved/reopened values and prior Cell picture controls.

Both retained binaries, both wheel/native pairs, complete installed payloads, actual logs and final source pins were independently authenticated. Cross-family all-target checking, strict changed-crate all-target/all-feature lint, formatting, prose, adapter drift and diff checks passed. The other 4303 tracked inputs remain unchanged. Root consumed these observations and the clean independent implementation report without a confirmation rerun.

## Pending acceptance

This accepts reduced unused state creation with preserved behavior. It does not qualify release speed, memory or final sprint acceptance. The previous ordinary release artifact still has its three measured misses: reply and secondary407 cold insertion/removal. A fresh ordinary artifact must measure this changed source under the unchanged protocols. Final installed/typing, native, memory, hosted, affected archive applicability and integrated gates remain pending. No status, ledger, baseline, review-cap reset, commit, push or publication approval changes.

This review pass changed only this review file and returns to sprint orchestration.
