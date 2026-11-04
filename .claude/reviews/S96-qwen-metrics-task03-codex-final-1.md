# S96 Qwen metrics Task03, Codex final 1

ACCEPTED FIXED-CORPUS SCALARS. TASK PROTOCOL NOT MET. This is actual Root final after DeepSeek's first scoped review, not a model ranking or sprint acceptance. Accurate output and adherence to the assigned constraints are separate observations.

## Identity and independent checks

The frozen four-file directory is /private/tmp/rdocx-qwen-harness/metrics-task03. All twelve INPUTS entries retain their actual SHA-256 and byte counts. Root independently parsed all 26761 JSONL events with zero parse errors and matched per-segment event types, separate call/update/use counts, unique call identifiers, timestamp windows, final end fields and controller qualifications against metrics.json. Controller exits are nine 0, two 1 and one 130. No usage blocks were summed.

Actual DeepSeek peer1 is /private/tmp/rdocx-cross-review/qwen-metrics-task03-deepseek-peer-pass1.md, SHA-256 52b55a2022c38289405f7fe268392ee2908df5e2ce02aa72488e7970504f4872. Its scalar acceptance agrees with Root's independent recomputation. Its explicit author-history limitation was honest.

## M1, medium, historical task protocol noncompliance

/private/tmp/rdocx-cross-review/QWEN_METRICS_TASK03.md:2 requires ONLY shell and at most three successful shell calls, with output at most 2000 characters. The actual author trace /private/tmp/rdocx-qwen-harness/metrics-task03.stdout.jsonl:3, :8, :11 and :30 contains four completed shell calls, each with child exit 0 under state.metadata.metadata.exit. The fourth is not a rejected or failed attempt.

The same trace has read uses at :15, :25 and :34 and grep uses at :19, :20, :21 and :26. Its :21 and :25 outputs are 4223 and 5800 characters. Read/grep were available in the harness despite the prompt prohibition, so this also exposes a controller configuration weakness. Availability does not make the model's uses compliant.

Exact scalar-only call facts are preserved in /private/tmp/rdocx-cross-review/QWEN_METRICS_TASK03_AUTHOR_PROTOCOL_FACTS.json. No task-private credentials or personal data incident is inferred. The finding is task instruction adherence and evidence discipline. Historical extra calls cannot be undone. Do not describe a corrected handoff or a future successful run as proof that this run met its original bound.

DeepSeek's assigned audit scope was the frozen outputs and twelve input segments, and its report explicitly did not certify the author call budget. Root extended the final audit to the actual author trace. Do not portray this as a peer finding it had already claimed to disprove.

## M2, low, handoff overstates completion details

/private/tmp/rdocx-cross-review/qwen-metrics-task03-handoff.md:16 generalizes Grok stopReason to end_turn and suggests controller_exit is null where absent. The frozen corpus actually includes three Grok end_turn segments and two max_turn_requests segments. All controller exits are populated, with the nine/two/one distribution above. Its final spent-budget wording does not accurately establish compliance with the original three-successful-shell-call limit.

Have a delegated model write a new factual addendum. Preserve the immutable old handoff, four frozen outputs and author trace. State the correct per-segment qualifications and the historical protocol miss. No need to recalculate correct data solely for that wording repair.

## M3, low, reviewer count needs correction

The S2 paragraph in /private/tmp/rdocx-cross-review/qwen-metrics-task03-deepseek-peer-pass1.md:51 says eight exit0 segments. Actual INPUTS.json has nine. The review's own claim of exact scalar agreement is supported, but that prose count is wrong. Correct it in the new addendum while retaining the original peer report.

## Qualified nits and limits

The peer's truncated tool-name table, dense 150-line script, timestamp-magnitude heuristic and broad id-key heuristic do not change this corpus's values. They prevent promotion to a generic collector certification. Do not invent microsecond or top-level-event-id failures on these inputs.

The Root check establishes fixed-data equality, not that the author script generated the file during a particular command. An exit 0 is a process outcome, not feature acceptance. Grok timestamp absence yields unknown event windows. These unlike tasks and hardware support no token rate, throughput or model-quality ranking. The report contains no raw thought or source payload in its scalar artifacts. Native tool outputs in the historical author trace are a separate protocol observation above.

No project source, workflow commands or skills were changed. No test, build, Git mutation, native app or network fetch was performed for this review. Write the factual addendum through delegation and retain this qualified outcome in the eventual single canonical model log. No overall DeepSeek, Grok or Qwen verdict before sprint end.
