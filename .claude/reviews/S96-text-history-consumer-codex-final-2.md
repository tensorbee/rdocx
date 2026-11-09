# S96 additional text-history consumer, Root final pass 2

Root read the complete finite D2 repair and added controls. Existing snapshot
links, watcher-total bounds and componentwise activation bounds are correct.
One residual D2 consistency defect remains, zero smells and zero nitpicks.
This is not a new history peer pass or a product/runtime finding.

## D2 residual, live allocations cannot increase across the pure drop edges

The new positive control at controls.py:158accepts clone_dropped and
reopen_dropped counts greater than their corresponding live counts. The
prepared source at after-document.rs:50959reads watch.alive, drops clone,
then reads the SAME unchanged watcher. Line50975does the same for reopened.
There is no intervening watcher append or producer call.

Equality is not required. Componentwise increase is nevertheless impossible
across those particular edges. Existing watched expired Arc allocations
cannot regain strong ownership through Weak, and new allocations cannot enter
the unchanged watcher vector. Likewise drop(document) at50976precedes
only_control_live at50978with no new watcher entry. These are known concrete
drop transitions, not an assumption that liveness decreases across arbitrary
publications, queries or the whole history.

Require clone_dropped<=clone_live, reopen_dropped<=reopen_live and
only_control_live<=reopen_dropped componentwise at these exact source edges.
Replace the impossible positive control with legitimate strict decreases.
Retain positives showing arbitrary earlier/later publication counts may vary
and no drop equality is guessed. This introduces no retention threshold or
payload-destruction assertion. Preserve all existing phase/custody controls.

## Other inspected boundaries

The repaired current/clone/reopen snapshot links and all per-snapshot totals
are accepted. A node may hold all three activation categories, so their sum
is correctly not capped by node count. No-Weak absence remains explicit.
No other new correctness, contract, panic, schema or structural finding was
established in this repair. The consumer/source/receipt copies remain private
preparation. Actual final-H execution, COW calibration, payload attribution
and separate font/RSS riders are unproved. Root changed no source outside
this review record and ran no product.
