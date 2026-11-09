# S96, Root final CI phase assertions, pass 12

## Scope

Reviewed the two test corrections in `crates/rdocx-py/tests/test_core.py`, six additions and one deletion. Product source is unchanged. The independent read-only investigation identified obsolete eager-inventory phase assumptions under the approved F283 initialization contract. Root inspected the actual successor diff and execution handoff `d9e2a87cbce1391469ff5b36e23240f5c7d381177f6b166d1bfb410e2bc865ee`.

Defects: 0. Smells: 0. Nitpicks: 0.

## Assessment

At `crates/rdocx-py/tests/test_core.py:12984` the deliberately mixed fixture activates checked authority before resetting work counters. The original owner-lookup assertions still require positive owner visits, zero XML bytes and zero whole-body rebuilds or reads. The mixed fixture is not widened into the finite plain-body optimization.

At `crates/rdocx-py/tests/test_core.py:15495` opening must retain zero dormant inventories for both relationship spellings. The existing alias-dependent rebuild assertion now follows the actual first header access. Later font publication, stale held owner, unchanged main part and save/reopen assertions remain intact. The correction measures the approved lifecycle rather than removing the work expectations.

A fresh isolated H10 archive, with only the test overlay, built an authenticated debug wheel with assertions and both debug counters present. The three actual cases passed with zero failures and zero skips. The actual failed CI predecessor remains two failures among 1,439 passes and two skips. This is not a complete CI rerun or qualification of the newer cold/reply product changes.

## Pending

The next reviewed checkpoint still needs its actual hosted CI run. Other packaging, performance, native and final integrated gates remain pending. No threshold, product grammar or acceptance criterion changed.
