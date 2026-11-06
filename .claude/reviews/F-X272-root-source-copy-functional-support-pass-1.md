# F-X272, Root source-copy functional support, pass 1

**Reviewed**: The immutable two-file increment at `ROOT_TAKEOVER/root-source-copy-review/source`, compared with the corresponding `root-returned-private-repair/source` files. Scope is the factored original raw-range union, its actual compose/repeat copy consumer and the reconciled copy controls. This is bounded Root support, not a whole F-X272 peer cycle, cap reset or provider verdict.

**Verdict**: 0 defects, 0 smells, 0 nitpicks in the bounded implementation examined. Equivalent unmodified boundaries and whole C1 remain pending.

## Defects

0 defects.

## Smells

0 smells.

## Nitpicks

0 nitpicks.

## Mechanism assessed

`crates/rdocx-oxml/src/text.rs:3977` factors the existing physical range-union algorithm into original_source_union. Both the original parsed-Field activation and the actual retained-Source copy use that same staged delta. It sorts and coalesces selected physical input ranges by part, intersects the existing union and returns additions plus pending removals/replacements. It does not mutate the ledger while preparing that result.

`crates/rdocx-oxml/src/text.rs:4049` selects only original_pieces from the actual positive Field output windows. Generated images and unknown complements contribute no raw input. It checks the prospective constructor allowance, output allowance and cumulative emitted demand before committing the union. An over-budget copy sets the shared sticky refusal without committing new input membership or emitted demand. A successful copy adds each previously unselected original byte once and spends each actual copied output byte. Duplicate output occurrences still count as separate demand, while duplicate physical input images cannot renew the union.

The helper does not add original_current_inputs, template_current_inputs, model membership, Word input, callback events or JSON inputs. Its constructor allowance increment is the same raw A contribution selected through the existing union, with current contribution zero. A derived image can forward original correspondence, but the shared union prevents it from registering that original byte a second time.

`crates/rdocx-oxml/src/document.rs:6573` applies the reservation after source assembly, checked lengths, map preparation and actual constructor admission, before materializing the buffer. `crates/rdocx-oxml/src/document.rs:7649` applies it to the actual repeat map and materialized positive coverage before cloning bytes. Both use the same operation and sticky state. Neither uses equality to find input provenance. Imported rewrite keeps its distinct existing reservation, avoiding a second debit through this helper.

## Contract and regression assessment

The approved raw-range-union scope permits first selection of original physical bytes reached by an actual positively classified Field copy. The no-match capture remains inert. The later real copy is a distinct selected operation, not allowance renewal or another stored P baseline. This follows the canonical plan's metric-specific reached-range union, stored-result P union, composition non-renewal and checked-copy reconciliation sections.

The updated `scoped_publication_postcapture_record_and_original_selected_field_have_exact_windows` preserves the no-match input-state assertion and original versus postcapture exact windows in paragraph, table and control contexts. It then drops the original Source, performs the actual first copied-field composition and checks only the reached raw union with no semantic-current membership. The second actual repeat refuses without changing that union. The equal generated-input control refuses and retains zero input state. This replaces only the overstrict expectation that the first real raw copy must also refuse.

The formerly failing textbox lifetime case now passes in the recorded full lower scoped cohort. The test correction does not treat omitted end-tag whitespace as output or add whole wrappers, namespace insertions or unknown ranges to the input union.

## Authentication and evidence

Both frozen source hashes authenticate against SOURCE.json:

| Path | SHA256 |
|---|---|
| crates/rdocx-oxml/src/document.rs | b3156fde1ba0804ae7fa65373dc879dffd12c2c50ae92f24074caac3d62494e1 |
| crates/rdocx-oxml/src/text.rs | ba8df7efeb23229bdfbec2e359c666f6b976164b5de50e16508ede91a5e42f5f |

The archived lower binary authenticates as `ab3931619cb5699e64340f6a77004fbd431eb0b194908c08ba117e965e305769`. The lower log records 197 scoped passes with zero failures. Separate Native logs record 21 scoped-template, 18 versioned-template and 10 fragment passes. Their aggregate is not asserted to be a unique-test total. RESULTS.json records all four Native/strict command exits as0. Strict all-targets with warnings denied completed in16.19 seconds. No all-features, integrated or native-application result is inferred. No reviewer build or test was run.

## Explicit boundaries

The unchanged namespace-insertion copy in `close_source_namespaces` at `crates/rdocx-oxml/src/document.rs:6812` still calls reserve_publication_copy rather than this first-selection helper. This report does not certify its first-unselected-original-copy behavior. That existing equivalent boundary remains a concrete follow-on to reconcile. A no-insertion move remains a separate path.

Generated and imported payload boundaries, unknown-owner classification, remaining returned contexts, non-emitting expression evaluation, comments/glossary and final Native joins are not certified by this review. Final integration, callable inventory, SemVer, complete local and integrated checks, hash/native/oracle riders and provider completion remain pending. No new API, quota, semantic-current input role or generic transport exemption is approved.

## Not found

No concrete additional arithmetic, union-overlap, commit-order, sticky-refusal, renewal, origin-selection or structural-rule defect was found in the bounded compose/repeat repair. The zero categories do not imply closure of the explicit equivalent boundary or other pending consumers.
