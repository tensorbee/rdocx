# S96, final union repair, Root pass10

**Reviewed**: The four-path finite opened-paragraph increment against H10 39287e8cb9c4a014adf0b3191cbe3cdf99c20569, plus its actual independent support report. Source diff7496e0c2c7997920bac1f2bf62feecb4fb6e044ce2a4e4972d8d53060251a411.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in this increment. This is a bounded repair review, not sprint acceptance or a reset of prior review caps.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Root analysis

The Native query at crates/rdocx/src/document.rs:34357 checks actual retained main-part membership, current package/source identity, successful opening serialization, complete current model equality and namespace state before model resource validation. It uses a separately argued finite reader-invariance grammar. Model equality supplies neither Source origin nor an operation allowance.

Admission requires physical paragraph slots, absent paragraph properties and all existing plain-owner exclusions. Run content is only ASCII alphanumeric/space Text. Run capture positions and unknown metadata stay excluded. RPr reconstruction admits only the approved fields, with independent shading reconstruction and complete equality against the original. The actual writer/parser preserve Text and direct run order in that grammar. Other spellings and unsupported properties retain full fallback. The existing arbitrary-paragraph writer/parser API is unchanged.

The binding at crates/rdocx-py/src/document.rs:5573 attempts this batch only when the previous piece list is empty. It maps facts by actual physical slot and retains ordinary route rebasing. Populated warm pieces and changed-context rebuilds keep their previous policy. Optional errors and misses retain the authoritative fallback. Related parts, duplicate identities and link qualification still run. No cache, authority, new parser, generic parameter, trait or file is introduced.

The genuine predecessor fails the existing direct-fragment counter at8463bytes after full owner-address equivalence. The successor passes the rich formatted fixture and zero-fragment assertion, empty/alias/raw/control/field/property and changed-state fallback cases, existing shared/case-equivalent story checks and missing-link activation. Four affected binding tests and one Native proof-guard test pass. Scoped compilation, strict all-target/all-feature lint, formatting, prose and diff checks pass. Root examined the actual code, handoff and independent report, rather than inferring correctness from counts alone.

The failed premise controls remain preserved. Empty self-closing carriers retain their actual raw-fallback distinction. Existing parser treatment of unused declarations is a positive read-fact control, while retained foreign metadata declines. The caller NUL case follows the actual legacy full-reader outcome and does not introduce a new refusal. These are explicit baseline facts, not product relaxations.

## Remaining work

H10 actual open timing passes at9.951792ms against10ms, but cold407 remains failed at about121ms. The new source has not been release-measured. All five flat-cost failures remain pending. The profiling-only symbol build is distinct from ordinary release artifacts. Its OS sampler refused its owned child before any target action, so it supplies no causal stack or successful benchmark evidence.

The current CI archive mismatch, remaining native qualification, memory and Markdown runs, current artifacts, final full verification, sprint review and close remain due. No feature delivery status, main, tag, publication or reporter comment is changed by this review.
