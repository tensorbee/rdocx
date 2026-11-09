# S96, opened plain owner batch, bounded support pass 1

**Reviewed**: The frozen four-path working diff against H10 `39287e8cb9c4a014adf0b3191cbe3cdf99c20569`, 415 additions and 3 deletions. This is bounded implementation support under the existing F268 remediation, with no review-cap reset.

**Verdict**: 0 defects, 0 smells, 0 nitpicks.

## Source and evidence identity

The current four files match every changed-path SHA in `IMPLEMENTATION_HANDOFF.json`, SHA `061ac73d96b332d8d931471a454da9b3ee98561025c6642e2ae8647ea2bdd27a`, under `/private/tmp/rdocx-s96-prepared-integration-readiness/FX237-H9-open-source-triage-1/opened-plain-batch-design-proposal-1`. The complete working diff matches SHA `7496e0c2c7997920bac1f2bf62feecb4fb6e044ce2a4e4972d8d53060251a411`. All 28 listed stdout/stderr hashes and the retained final binding binary SHA `a21711139dc2c22d46d9d89613a67cba25247a4d37824995bbaafbb81ae0d25e` authenticate. The same source pins and diff were checked again before handback.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Examined behavior

- Opening membership, current package bytes and main-part identity remain checked by `crates/rdocx/src/document.rs:29334`. The new query at `crates/rdocx/src/document.rs:34357` additionally requires the actual successful opening writer record, complete current model equality, matching namespace state and existing local namespace eligibility. Existing authoritative model resource validation runs before facts are returned. This is the approved finite reader-invariance proof, not an inference that arbitrary successful serialization proves arbitrary read equivalence.
- The body traversal is all-or-nothing. Nonparagraph children, PPr, captured metadata, controls, revisions, Fields and nonfinite payload grammar decline at `crates/rdocx/src/document.rs:34380`. Existing `paragraph_is_plain` at `crates/rdocx/src/document.rs:8495`, plus the new run-position and Text-only checks, exclude every current paragraph/run capture channel. Text is restricted to ASCII alphanumeric characters and space.
- Formatting reconstruction at `crates/rdocx/src/document.rs:34395` copies only the approved RPr fields, reconstructs shading without extra attributes and requires complete equality. Unlisted modeled, raw, namespace and revision properties cannot enter through that copy. For the admitted grammar, the actual run writer at `crates/rdocx-oxml/src/text.rs:10752` emits paired Text elements, and the actual parser at `crates/rdocx-oxml/src/text.rs:10300` preserves the corresponding Text and run sequence. Accepted readers at `crates/rdocx-oxml/src/text.rs:18303` and `crates/rdocx-oxml/src/text.rs:18580` do not depend on these formatting values.
- The binding consumes the batch once only for an empty previous block list at `crates/rdocx-py/src/document.rs:5573`. Physical vector position supplies correspondence, followed by existing route rebasing. It does not search equal values. Optional failure or decline leaves the previous per-block/full-authority route available at `crates/rdocx-py/src/document.rs:5636`. Related sources, duplicate physical part checks, link validation and authoritative fallback remain in the existing assembly/inventory path at `crates/rdocx-py/src/document.rs:5555`, `crates/rdocx-py/src/document.rs:5743` and `crates/rdocx-py/src/document.rs:8339`.
- No Source, operation allowance, publication authority, retained dormant cache, model mutation or new parser is introduced. Warm existing pieces do not acquire this whole-opening query. This matches the explicit amendment in `.claude/plans/F-X268-design.md:723` and the current binding description in `docs/hld/10-bindings-spec.md:4992`.

## Acceptance evidence and limits

The authenticated predecessor log records one actual failing test after owner-address equivalence, with direct-fragment work 8463 rather than zero. This is a mechanism RED, not an absent-API compile failure or timing result. The final binding log records four passing affected tests. Their assertions cover actual rich407 eligibility, Native checked-producer fact equivalence, exact physical addresses, zero cold fragment work, warm query absence, optional grammar and changed-context/duplicate-part fallback. The Native guard log records one passing opening-proof/package/namespace/capture test. The existing missing-link activation test and strict lint logs also authenticate as successful.

The self-closing run/Text, unused source alias and caller NUL premise corrections remain explicit in the retained progress and failed logs. The final tests distinguish actual captured self-closing carriers from paired empty carriers. An unused declaration discarded by the existing parser is admitted only for these read facts, while retained foreign metadata declines. The NUL control preserves the actual full-reader outcome rather than inventing a new refusal. No authenticated Word oracle or other external result is supplied by these local tests.

Correctness, contract, panic/index boundaries, OOXML preservation, tests and structural rules produced zero findings within this increment. No runtime, build, product import or source edit was performed during this review. Actual ordinary release cold407 measurements, all five flat-cost gates, Root final review and final integrated/native/release riders remain pending. This report does not establish a performance bound or full sprint acceptance.
