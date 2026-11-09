# S96 additional text-history consumer, Root final pass 3

Root reviewed the finite v2-to-v3 consumer and control deltas against the
unchanged prepared source. Verdict: zero defects, zero smells and zero
nitpicks in this residual D2 repair. D2 is closed for consumer preparation.
This bounded final is not a fourth history peer or product acceptance.

The consumer now checks the three actual unchanged-watcher drop edges
componentwise. The source reads clone_live immediately before drop(clone),
reopen_live immediately before drop(reopened), and only_control_live after
drop(document), without adding watcher entries between those reads. Existing
expired watched allocations cannot regain strong ownership through Weak.
The checks therefore reject impossible increases without requiring equality
or imposing monotonicity across arbitrary publications and queries.

The former positive expectation was incorrect and is explicitly preserved as
a historical correction. The replacement permits strict decreases. A separate
positive permits earlier publication liveness to decrease and then increase
within the accumulated totals. All three activation categories remain bounded
individually, rather than incorrectly summed against live node count.

The exact consumer delta is
`0aadb8dc7d6a1af0e3a08e82549a7ff1dcea27f4b155b75d3849d23dade1ff11`.
Final consume.py is
`ce1098f155000b720c837a0cf4ddcbc191dd1ab0d975a95a46b18d4262cf1d68`.
The packet records293synthetic controls and13offline before/after function
discriminators. These exercise protocol consistency, not Native retention or
product RED/GREEN. Independent mechanical replay is separately pending.

Correctness, contract, panic, tests, structure and evidence boundaries yield
no additional findings in the repair. Source candidate, font helper, receipt
contract and protocol are unchanged. Actual final-H compilation and execution,
positive memo-family creation, observer/COW calibration, moved/transitive
payload attribution and fixed-font/quiet-RSS riders remain unproved. Root
changed no product or TEMP source during this review.
