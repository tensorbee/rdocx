# S96 bounded STYLEREF width-accounting source review

Verdict: 0 defects, 0 smells, 0 nitpicks in the assigned deletion branch. No source fix, Cargo or timing was performed by this reviewer.

`crates/rdocx-layout/src/engine.rs:10013-10029` now records the actual removed tab-aligned Text leaf's negative advance width before deleting it. Together with the first replacement's new-minus-old contribution at `:10032-10046`, a multi-leaf cache sums to new width minus the entire former width. The existing persistent tab group, baseline, font size and actual old x are retained, so `realigned_offset` at `:10145` aggregates the same physical alignment group and accounts for later leaves in their original order. It does not change glyph shaping, text, geometry expectations or source ownership.

The distinct ordinary substitution publisher at `:10471-10488` already computes replacement width minus old width through `field_replacement`, including an empty value. The new branch closes STYLEREF's separate direct-splice path rather than altering that consumer. Existing annotation/rich carrier branches are unchanged and no broader width-coverage claim is made.

Actual regression44 and affected field/tab controls remain pending Root's next frozen runtime batch. The historical361-versus451 outcome remains failed evidence, not a passed native or geometry baseline. No full sprint, performance or review-cap acceptance follows.

Source SHA256: `ca81dcaa805750af30a5c1a539ed36226ed12ed6e9679f0a2d8705907ce97efd`.
