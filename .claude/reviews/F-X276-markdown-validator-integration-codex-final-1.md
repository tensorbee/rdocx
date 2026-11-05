# F-X276 Markdown validator integration, Codex final pass 1

**Reviewed**: Existing TEMP benchmark semantic collector against the newly
qualified cmark-gfm0.29.0.gfm.13 XML grammar. This is a preparation integration
review prompted by new actual oracle output, not product benchmark execution
or confirmation of unchanged product source.
**Verdict**: 1 defect, 0 smells, 0 nitpicks.

## Defects

### D1, table header container is counted as an extra cell

`/private/tmp/rdocx-cross-review/S96_MARKDOWN_BENCHMARK_PREPARATION/bench_common.py:96`
counts both table_cell and table_header. The qualified actual GFM XML output
contains a table_header container with its own table_cell children. Its two
columns and one body row produce four real cells but five collector entries.
The actual output and extracted values are retained in
cmark-gfm-validator-preparation/VALIDATOR_QUALIFICATION.json under gfm-positive.
This is an observed tool/collector grammar disagreement, not a product failure.

The later W2 threshold at
`/private/tmp/rdocx-cross-review/S96_MARKDOWN_BENCHMARK_PREPARATION/bench_common.py:108`
uses that inflated count to require16cells for each selected slide. A malformed
export with a three-column, five-row table has only15real cells but can satisfy
that16-entry threshold. The required fixture actually contains four by four
cells, authored at
`/private/tmp/rdocx-cross-review/S96_MARKDOWN_BENCHMARK_PREPARATION/generate_fixtures.py:75`.
Count the actual leaf cells, retaining the qualified header grammar and positive
controls. Preserve the original preparation source and manifest. Reproduce the
collector distinction using tiny independent Markdown, without running product
benchmarks or claiming a historical product RED.

## Smells

None.

## Nitpicks

None.

## Not found

No additional finding in the bounded new XML grammar join. Acquisition, source
and executable identity, static library linkage and actual eight qualification
outcomes have been authenticated separately. No product source, fixture,
benchmark, runtime performance or final acceptance is reviewed here.

Return to orchestration for a separate private preparation repair. Do not patch
inside this review pass. The original Root static inspection remains historical,
and this new observation does not rewrite it as a failed execution.
