# F-X272, canonical callback integration, pass 1

**Reviewed**: The actual frozen joined R2 source and failed `provider-gate-r2/native-versioned.log`. The preceding 25 template and five loader tests pass. Versioned selection has 15 passes and three failures. Two are exact authored-note spacing fixture mismatches. The remaining failure below is a real integration defect.

**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, initial program capture observes published numbering instead of the original callback owner

`crates/rdocx/src/document.rs:48089`, `crates/rdocx/src/template.rs:2544`, `crates/rdocx/src/document.rs:74991`, `crates/rdocx/src/document.rs:75019`

The existing admitted test creates two live numbering instances, uses the second on a paragraph containing a scalar, and records its original public body ContentLocation before rendering. The actual callback then receives the same physical story/index but a different fingerprint. The observed identifiers are 12212761742953017915 instead of 11004995489472932852.

The reconciled first scoped program inventory flushes through the publication-only numbering remap before capturing the callback originals. Restoring the live numbering model afterward preserves the numbering API but cannot restore the already captured published fingerprint. Preserve both contracts: original callback locations from the actual pre-publication owner, and F277b's temporary package numbering projection with live state restored. Do not weaken the callback assertion, use ordinal/value equality as provenance, or restore eager live numbering mutation. Existing private scoped producers should supply the original-location handoff without a public API or renewed resource allowance.

## Smells

Zero smells.

## Nitpicks

Zero nitpicks.

## Not found and retained evidence

The merged source compiles, external46/7client compiles after its separate fixture repair, and lower209/footnote16 cases pass. The failed Native binary and complete raw log remain associated with R2 source. No Rust regression, CLI, fresh Python or final strict result follows because the dependent batch stopped at the actual failure. Earlier independent semantic-join support is bounded source evidence and does not override this runtime counterexample. This review modifies only this report and ends before remediation.
