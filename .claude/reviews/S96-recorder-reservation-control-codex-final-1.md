# S96 recorder reservation control, Root final pass 1

Reviewed the complete129line private diagnostic patch against the finite
F-X282 reservation-failure contract. Candidate source SHA256 is
`b0603991334b683b057c213ebc5e7bfc0f06fca70effc573f9cab354beaa4712`.
Patch SHA256 is
`eb4a3817c47491e7919825379e56b453920ed358c50c7f4aefd04b162b5235c4`.
Verdict: zero defects, zero smells and zero nitpicks in source preparation.
This bounded control final does not reset the history-source peer cap.

Both real recorder callers preserve operation assignment, additional1,
failure-before-push and successful push order. The shared helper performs
the original checked addition and actual typed Vec reservation, invalidating
both completeness mechanisms on failure. Its generic parameter has two
existing concrete callers, the proof and COW event vectors.

The ignored control uses those actual empty vectors and verifies that their
length addition succeeds before requesting usize::MAX. Nonzero event sizes
make the requested capacity impossible. This isolates the reservation branch
from arithmetic overflow without requesting an OS-sized allocation. The
control checks no event push, sticky failure after a successful ordinary1
reservation, incomplete finish and a fresh lifecycle reset. Its local guard
restores the initially inactive, clear state on normal return or unwind.
The existing module imports expose the required private items to the test.

Correctness, contract, panic handling, XML scope, tests and structure produce
no source finding. Root independently authenticated all14custody files and
47retained input pins. Inverting the exact helper, control and two caller
transformations recovers every baseline byte. The31recorded Python controls
remain source and assembly evidence only, including their lexical mutants.

No Rust compilation or execution occurred. The capacity-overflow-sized
reservation failure remains unobserved until the final-H diagnostic union
runs the ignored test. This does not prove ordinary reserve1 allocator
exhaustion, OS OOM, product refusal, observer neutrality, payload destruction,
bounded history or deciding RSS. Existing source packets and failed evidence
remain preserved. Root changed only this review record during the pass.
