# F-X275 archive intent support, pass 6

Verdict: zero defects, zero smells and zero nitpicks in the two affected archive
records. This is bounded archive support, not publication verification.

The actual changed sources are rdocx document and integration carriers plus
rdocx-layout paginator. Root independently recomputed both retained archive
measurements with the existing normalization helper and authenticated27and11
direct source/manifest/test members respectively. Worker authentication also
checks all75source inputs, generated lock package identities/checksums, original
manifest bytes and actual Cargo VCS identity. The first wrapper failed after
both successful packages because it assumed every row lived in root README.
The original failed wrapper and its archives are retained. A no-Cargo
postprocessor selected each actual page through MEASUREMENT_PAGES.

At scripts/readme_doctests.py:379 and382 only the two affected measured dates
advance to the actual2026-11-07UTC date. The two archive tuples at405and408
reflect actual retained measurements. Root README:44 and the crate README
crates/rdocx-layout/README.md:23 receive those exact respective rows. Other20
archive tuples/dates, speed observations and tolerances are unchanged. Both
self-containing README byte deltas are zero.

After adoption, actual Cargo packaged both affected crates again with all
local internal patches and --no-verify. Both commands exited0. Normalized
rdocx is4454838compressed/56133002member bytes/37members, exactly recorded.
Normalized rdocx-layout is458054compressed/2510055member bytes/15members,
three compressed bytes below the recorded458057and within the unchanged
64byte compression tolerance. Exact member bytes/counts and source guards
pass. No third archive pass or other20package regeneration was performed.

Evidence is under
/private/tmp/rdocx-s96-prepared-integration-readiness/word-h5-two-package-archive-recipe-1
and /private/tmp/rdocx-s96-h5-archive-postadoption-1. The complete verified
workspace publication dry run and current hosted Docs remain due. No release
tag, registry publication or universal packaging pass is inferred.
