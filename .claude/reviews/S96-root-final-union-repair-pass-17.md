# S96, Root final union repair, pass 17

## Scope and result

Root reviewed the finite F268 reply shared-source increment over the already reviewed Native Cell and frame work. There are zero defects, zero smells and zero nitpicks. This is correctness acceptance of this increment, not sprint completion or a performance pass.

The five-path before-to-after patch is retained at /private/tmp/rdocx-s96-reply-shared-source-implementation-1/increment.patch with SHA-256 743dcaf623a0e25c4b2918f2d85d6684a65820698a048fd3541786e3027c7eab. SOURCE_BOUNDARY.json has SHA-256 6138cf32a05f62d3897acdeb2b20065b1c90290e47d428d7f3d16c98fce738be. Root independently authenticated every final source, protected input, actual log, retained binary and patch pin. The independent support review is S96-reply-shared-source-support-1.md. This review changes no product source.

## Contract, correctness and structure

At crates/rdocx/src/document.rs:30720 the first endpoint still performs complete current source and item qualification. Only complete ContentLocation equality permits reusing that actual span. Different locations retain the second qualification. At 30757 the same parsed paragraph supplies both run-bound checks, and the distinct-span branch retains its separately parsed paragraph. The new expect follows the same span inequality that constructed its Some value. Namespace closure, reverse edit order, marker insertion and publication remain in the existing implementation.

At crates/rdocx/src/document.rs:31516 the crate-private refreshed owner helper obtains actual spanned sources, discovers all stories using the same public discovery implementation, selects the same complete kind/part/index selector and qualifies the refreshed fingerprint on those same buffers. The main source is guaranteed by spanned_story_sources at 29956. It preserves complete related discovery and the existing layout fallback. The consumer at crates/rdocx/src/comments.rs:2050 retains parent preparation, ID reservation, anchoring, marker selection and adjacency correction. No public API, new input allowance, cached owner shortcut, transaction removal, parser grammar, XML sequence or unmodelled-payload policy changes.

## Tests and evidence

The genuine retained predecessor finishes successful package, anchor and reopen checks before observing two endpoint and two refresh preparations. The successor requires one of each at comments.rs:4925. The distinct-endpoint discriminator retains two endpoint reads. The complete-discovery error test at 5012 uses an actual malformed referenced header. The four-mode legacy differential retains full package and graph comparisons for body, header, aliases, unsaved content, simple and complex Fields and raw tracked content. Existing atomic refusal and cross-owner controls remain applicable.

The initial complete comments invocation remains failed with 41 passing tests, one new fixture failure and one native ignore. That fixture expected refusal for a truncated EOF form which the old reader also accepted. It was replaced with an actual unterminated attribute. The seven selected successor reply tests pass, as do the three affected filters and scoped check, strict Native and binding lint, formatting, prose, adapter drift and diff checks. This does not rename the earlier failed invocation as a passing full suite. No independent test rerun was needed to inspect these unchanged, pinned executions.

## Pending acceptance

Actual release-profile reply timing remains pending, as do installed binding, final integrated, native, memory, affected archive and hosted checks for the later repaired source. Both existing staging clones remain. The prior H11 timing failures stay failed. The separately verified H11 hosted run applies to cec87f9501eb13d0d5d60572472b227483262620 and does not qualify this later increment. No feature status, ledger, baseline, review cap, release approval or publication changes.
