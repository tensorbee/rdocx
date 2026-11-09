# S96, Root final repair review, pass30

Reviewed the frozen four-path21increment on sprint/s96 at
cec87f9501eb13d0d5d60572472b227483262620, following actual independent
Support12, SHA-256b065126e6ff4e38c86cfb121d2469afb7942969ea02b0a078ff351db6523040c.
This is Root's separate assessment of the85-line increment against Root29,
not a second independent peer or final sprint acceptance.

## Defects

None.

## Smells

None.

## Nitpicks

None.

## Source and contract assessment

Consumed investigation21, adopted F268contract, HLD03/04, the complete increment,
actual Support12, profile22and emitted-code23. Reauthenticated all4309current
source pins, with4305unchanged and exactly the plan/OXMLdocument/HLD03/HLD04changed.
Boundaryc473fadb3f069ecef31e6e4a45e3304a94cc8b7b425215099a35224871c3e5d4
and patcha293c964491858d31b7303569b4e01ceff15246f8f82c20bc01dd3b9eb38c6fa
match the independently reviewed scope.

At `crates/rdocx-oxml/src/document.rs:1551`, actual child enumeration and prior
local checks remain in place. Empty children move directly. Nonempty children
stay local while inert Enter values append above a temporary empty Finish.
The saved logical index survives reallocation, no borrow spans extension,
and the original vector is installed before any step can be popped. No model
getter, Source/Field read, callback or recoverable failure occurs in this
interval. The internal index/variant assertion follows from append-only local
ownership. Allocation unwind does not process an incomplete summary.

At `crates/rdocx-oxml/src/document.rs:1571`, completed-overbudget scheduling
remains unchanged. Active/completed summary handling, first local weighted/table
refusal, original LIFO/Finish child order, final maxima and root aggregation
remain exact. Reviewed both prospective Field-summary callers at
`crates/rdocx-oxml/src/text.rs:7005` and `:7068`, and their actual local tests.
No API, XML grammar, source/admission credit, cache, trait, generic, module,
file, unsafe code or private sibling operation is introduced.18/20remain
unchanged. No new semantic behavior is claimed whose test should fail solely
because the scheduling loop is reverted.

## Evidence assessment

Actual scoped logs prove15unique Rust passes, zero ignored, all-target
OXML/Native/binding check0, strict affected all-target/all-feature Clippy0,
fmt/prose/adapters/diff0. The initial wrong document::tests filter selected0,
receives no coverage and remains retained. Correct text::tests cases executed.
The ordinary debug build and fresh installed payload authenticate all10members,
with64core plus10picture cases passing, zero failures/errors/skips. Existing
wide513/shared-root/Field-parent/independent256/32/early-invalid-tail/pre-writer
refusal controls reject changed semantic order without inventing queue-shape tests.

The exact ten public state pairs retain original setup and prior actions,
complete input/output ZIPs/every member, values, staging and save/reopen.
All11,270,987normalized graph records, including source/Finish/pending capacity,
are identical. Source-level old lookup/new installation hooks are actual.
Main work[2] is accurately labeled queued-child cardinality, not a constructor
hook. The separate actual constructor-site observation closes that obligation:
all10positive old/new counts equal each other and the main cardinality, roots
consistently excluded. It also reproduces full input/output/parts/staging and
semantics/reopen. Its graph is intentionally disabled and does not replace
the complete graph proof. Comparison hashes are
f598006223c7bb22dd82db0e6f70dbdc7605feab0cd89748b54672091146eb65
and5a555d3aa713594dcdf21919ccabaab42cc31a42aa9a280c2dde7ea35652aa50.
Both supplements and both full diagnostics build0/restore4309/authenticate10.
The initial missing-docx observer fails before actions and remains noncoverage.

Actual optimized generated code confirms the desired compact enqueue loop,
once-after-loop checks/installation and direct empty case. It does not remove
all guard/hash/enumeration work or establish elapsed sufficiency. Whole emitted
function781to779instructions and fixed stack336to352bytes are target-specific
facts. Original vector spare capacity/lifetime and transient descriptor work
remain real costs, with no heap/peak/RSS or portable ABI claim. The current
successful public graph states are not alone proof of all refusal paths.

## Pending acceptance and handback

Latest ordinary75/79and failed secondary cold insertion remain valid failures.
No unchanged-artifact retry, timing noise waiver, threshold change or protocol
relaxation is selected.21can proceed as reviewed correctness into a new ordinary
artifact, or be composed with a separately adopted bounded increment before
that artifact. This report does not approve future24candidates. Fresh complete
compiler-free qualification still decides elapsed acceptance. Final memory,
ordinary/private calibration, Office, installed typing, hosted/archive and
full integrated-union riders remain due. Memory investigation4is proposed
contract wording only, with no omission overlay or RSS pass.

Zero findings in correctness, contract, panics, OOXML preservation, tests and
structure. No product/source/progress/plan/tracker was modified in this review.
No Cargo, runtime test, timing, Git mutation or publication occurred. This pass
writes only this report, ends here and returns control to implementation.
