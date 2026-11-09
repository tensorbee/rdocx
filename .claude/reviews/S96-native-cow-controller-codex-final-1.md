# S96 private COW controller and consumer, Root final pass 1

Root reviewed the163line consumer, finite controller adaptation and offline
controls against the accepted text controller and COW source protocol.
Verdict: one defect, zero smells and zero nitpicks. This is a new controller
review, not another history-source peer. No product or Native process ran.

## D1, serialized receipt consumption drops the execution identity joins

`consume_cow.py:154` through163checks only the protocol digest, command suffix,
matching partial identity dictionaries, inventory-file digest and seven
self-listed asset pins. It does not require the complete six-key spec or
seven-key identity, bind argv0to the binary, authenticate the actual build and
profile relationships, validate the inventory/document entry, or bind the
seven assets to the expected spec roles. The controller preflight does those
checks, but read_row does not preserve them when consuming recorded evidence.

Root called the actual read_row function with an in-memory filesystem and
otherwise valid source-shaped lane records. It returned complete=truewith no
SourceH, no binary/profile/source identities, a nonexistent argv0binary, no
build/profile assets and seven unrelated files. The identity contained only
protocol_sha256 and source_manifest_sha256. Identical empty inventory snapshots
were accepted. This is an actual offline consumer result, not a Native failure
or a claim to independently authenticate an honest producer's historical PID.

Require the full recorded spec and execution identity, concrete asset-role
links, optimized successful build/profile/source joins and argv0binary pin
when consuming the receipt. Preserve positive child PID, exact stack environment
and actual cleanup outcome. Prefer sharing the substantive pure identity
verifier with preflight in existing files, without a generic runner framework.
Add direct read_row controls for missing and mismatched identity/asset joins,
not only preflight controls. Preserve successful process completion separately
from later consumer ineligibility and all existing raw-protocol checks.

## Other inspected boundaries

The explicit artificial counts, before/after state, same-transition Vec/Arc
relations, post-call strong release, complete semantic schemas and symmetric
omission controls have no new established finding. Actual recorder operation0
is correctly retained. Normal policy assertions are not misrepresented as
emitted START/END observations. Strict JSON, exact test selection, original
spec custody, direct-child finally cleanup and operational incomplete guards
retain their intended scope. Reported86controls are offline preparation only.
No runtime, final-H, observer neutrality, payload destruction, font/text history
or RSS qualification follows. Root changed only this review record.
